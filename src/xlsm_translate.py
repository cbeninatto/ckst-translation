import io
from typing import Callable, Dict, List, Optional, Tuple

import openpyxl
from openpyxl.worksheet.cell_range import CellRange


# Try to reuse your existing types if they exist
try:
    from .openai_translate import TranslationItem
except Exception:
    from collections import namedtuple
    TranslationItem = namedtuple("TranslationItem", ["id", "text"])


def _chunk_items(
    items: List[TranslationItem],
    max_items: int = 200,
    max_chars: int = 20000,
) -> List[List[TranslationItem]]:
    """
    We still must chunk to avoid request size limits, but we keep chunks large.
    """
    out: List[List[TranslationItem]] = []
    cur: List[TranslationItem] = []
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


def _parse_print_areas(print_area: str) -> List[CellRange]:
    """
    openpyxl may return:
      "'Sheet'!$A$1:$K$41"  or  ""  or  None
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


def _fallback_used_range(ws) -> List[CellRange]:
    """
    If no print area exists, we fallback to the "used" range as best-effort.
    """
    dim = ws.calculate_dimension()  # e.g. "A1:K41" or "A1"
    if not dim:
        return [CellRange("A1:A1")]
    return [CellRange(dim)]


def _cell_in_any_range(r: int, c: int, ranges: List[CellRange]) -> bool:
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
    - Translates ALL sheets.
    - Only the PRINT AREA is translated/kept.
    - Anything outside the print area is CLEARED (values removed).
    - Saves as .xlsm with macros preserved (keep_vba=True).
    """
    wb = openpyxl.load_workbook(io.BytesIO(xlsm_bytes), keep_vba=True, data_only=False)

    sheets = wb.worksheets
    total_sheets = len(sheets)

    for s_idx, ws in enumerate(sheets, start=1):
        if on_progress:
            on_progress("sheet", s_idx, total_sheets)

        ranges = _parse_print_areas(getattr(ws, "print_area", None) or "")
        if not ranges:
            # If a sheet truly has no print area, we fallback to used cells.
            ranges = _fallback_used_range(ws)

        # Collect cells to translate (only inside print area)
        items: List[TranslationItem] = []
        coords: List[Tuple[str, int, int]] = []  # (item_id, row, col)

        total_cells_to_translate = 0
        for cr in ranges:
            for row in ws.iter_rows(
                min_row=cr.min_row,
                max_row=cr.max_row,
                min_col=cr.min_col,
                max_col=cr.max_col,
            ):
                for cell in row:
                    v = cell.value
                    if not isinstance(v, str):
                        continue
                    txt = v.strip()
                    if not txt:
                        continue
                    if txt.startswith("="):  # skip formulas stored as strings
                        continue

                    item_id = f"{ws.title}!{cell.coordinate}"
                    items.append(TranslationItem(item_id, v))
                    coords.append((item_id, cell.row, cell.column))

        total_cells_to_translate = len(items)
        done_cells = 0
        if on_progress:
            on_progress("cells", done_cells, max(1, total_cells_to_translate))

        # Translate in large chunks
        mapping: Dict[str, str] = {}
        for chunk in _chunk_items(items):
            mapping.update(translator.translate_batch(chunk))
            done_cells = min(total_cells_to_translate, done_cells + len(chunk))
            if on_progress:
                on_progress("cells", done_cells, max(1, total_cells_to_translate))

        # Write back translations
        for item_id, r, c in coords:
            new_text = mapping.get(item_id)
            if new_text is not None:
                ws.cell(row=r, column=c).value = new_text

        # Clear anything outside print area (values removed)
        # Use ws._cells so we only touch cells that actually exist (fast, avoids 1M scans).
        for (r, c), cell in list(ws._cells.items()):
            if cell.value is None:
                continue
            if not _cell_in_any_range(r, c, ranges):
                cell.value = None

        # Normalize print_area text (keep it)
        # If multiple ranges, set as "A1:K41,B2:C3"
        try:
            ws.print_area = ",".join([cr.coord for cr in ranges])
        except Exception:
            pass

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()
