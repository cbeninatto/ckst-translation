import io
from typing import Callable, Dict, List, Optional, Tuple

import openpyxl
from openpyxl.worksheet.cell_range import CellRange


class _Item:
    """Minimal item object compatible with translator.translate_batch(items)."""
    __slots__ = ("id", "text")

    def __init__(self, item_id: str, text: str):
        self.id = item_id
        self.text = text


def _chunk_items(
    items: List[_Item],
    max_items: int = 400,
    max_chars: int = 60000,
) -> List[List[_Item]]:
    """
    We still chunk to avoid request size limits, but keep chunks large.
    """
    out: List[List[_Item]] = []
    cur: List[_Item] = []
    cur_chars = 0

    for it in items:
        t = it.text or ""
        if cur and (len(cur) >= max_items or (cur_chars + len(t)) > max_chars):
            out.append(cur)
            cur = []
            cur_chars = 0
        cur.append(it)
        cur_chars += len(t)

    if cur:
        out.append(cur)

    return out


def _parse_print_area(print_area: str) -> List[CellRange]:
    """
    Examples:
      "'Sheet1'!$A$1:$K$41"
      "'Sheet1'!$A$1:$K$41,'Sheet1'!$M$1:$N$10"
    """
    if not print_area:
        return []

    s = str(print_area).strip()
    if not s:
        return []

    parts = [p.strip() for p in s.split(",") if p.strip()]
    ranges: List[CellRange] = []
    for p in parts:
        if "!" in p:
            p = p.split("!", 1)[1]
        p = p.replace("$", "")
        ranges.append(CellRange(p))
    return ranges


def _cell_in_ranges(r: int, c: int, ranges: List[CellRange]) -> bool:
    for cr in ranges:
        if cr.min_row <= r <= cr.max_row and cr.min_col <= c <= cr.max_col:
            return True
    return False


def translate_xlsm_bytes(
    xlsm_bytes: bytes,
    translator,
    on_progress: Optional[Callable[[str, int, int], None]] = None,
) -> bytes:
    """
    Requirements:
    - Translate ONLY inside the sheet's defined PRINT AREA.
    - Remove/clear anything OUTSIDE the print area.
    - Process all tabs (worksheets).
    - Preserve macros (.xlsm): keep_vba=True.

    NOTE: If a sheet has NO print area defined, we SKIP it (do not translate/clear).
    """

    if not hasattr(translator, "translate_batch"):
        raise RuntimeError("Translator must implement translate_batch(items) -> {id: translated_text}")

    wb = openpyxl.load_workbook(io.BytesIO(xlsm_bytes), keep_vba=True, data_only=False)

    sheets = wb.worksheets
    total_sheets = max(1, len(sheets))

    for s_idx, ws in enumerate(sheets, start=1):
        if on_progress:
            on_progress("sheet", s_idx, total_sheets)

        # openpyxl stores print area in ws.print_area (string) when set.
        print_area = getattr(ws, "print_area", None) or ""
        ranges = _parse_print_area(print_area)

        # STRICT: only use defined print area; if missing, skip sheet.
        if not ranges:
            continue

        # Collect translatable cells in print area
        items: List[_Item] = []
        coords: List[Tuple[str, int, int]] = []

        for cr in ranges:
            for row in ws.iter_rows(
                min_row=cr.min_row,
                max_row=cr.max_row,
                min_col=cr.min_col,
                max_col=cr.max_col,
            ):
                for cell in row:
                    v = cell.value

                    # Skip empty/non-text
                    if not isinstance(v, str):
                        continue
                    txt = v.strip()
                    if not txt:
                        continue

                    # Skip formulas
                    if cell.data_type == "f":
                        continue
                    if txt.startswith("="):
                        continue

                    item_id = f"{ws.title}!{cell.coordinate}"
                    items.append(_Item(item_id, v))
                    coords.append((item_id, cell.row, cell.column))

        total_cells = len(items)
        done_cells = 0
        if on_progress:
            on_progress("cells", done_cells, max(1, total_cells))

        # Translate in large chunks
        mapping: Dict[str, str] = {}
        for chunk in _chunk_items(items):
            mapping.update(translator.translate_batch(chunk))
            done_cells += len(chunk)
            if on_progress:
                on_progress("cells", min(done_cells, total_cells), max(1, total_cells))

        # Write translations back into the SAME cells
        for item_id, r, c in coords:
            new_text = mapping.get(item_id)
            if new_text is not None:
                ws.cell(row=r, column=c).value = new_text

        # Clear anything outside print area (only touching existing cells)
        for (r, c), cell in list(ws._cells.items()):
            if cell.value is None:
                continue
            if not _cell_in_ranges(r, c, ranges):
                cell.value = None

        # Keep print area as-is
        try:
            ws.print_area = ",".join([cr.coord for cr in ranges])
        except Exception:
            pass

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()
