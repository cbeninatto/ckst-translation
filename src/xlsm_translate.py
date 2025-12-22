import io
from typing import Callable, Dict, List, Optional, Tuple

import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.cell_range import CellRange

from .openai_translate import OpenAITranslator, TranslationItem, chunk_items
from .text_utils import apply_glossary_hard


def _parse_print_area(print_area: str) -> List[CellRange]:
    """
    Examples:
      "'Sheet1'!$A$1:$K$41"
      "'Sheet1'!$A$1:$K$41,'Sheet1'!$M$1:$N$10"
    Returns CellRange list WITHOUT $ and WITHOUT sheet name.
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


def _union_bounds(ranges: List[CellRange]) -> Tuple[int, int, int, int]:
    """
    Returns (min_row, max_row, min_col, max_col) union across ranges.
    """
    min_row = min(cr.min_row for cr in ranges)
    max_row = max(cr.max_row for cr in ranges)
    min_col = min(cr.min_col for cr in ranges)
    max_col = max(cr.max_col for cr in ranges)
    return min_row, max_row, min_col, max_col


def _unmerge_outside_union(ws, min_row: int, max_row: int, min_col: int, max_col: int) -> None:
    """
    If merged ranges are partially/fully outside the union, unmerge them
    to avoid broken merges after cropping.
    """
    try:
        to_unmerge = []
        for mr in list(ws.merged_cells.ranges):
            # Keep only merges fully contained in union
            if not (
                mr.min_row >= min_row and mr.max_row <= max_row and
                mr.min_col >= min_col and mr.max_col <= max_col
            ):
                to_unmerge.append(str(mr))
        for rng in to_unmerge:
            try:
                ws.unmerge_cells(rng)
            except Exception:
                pass
    except Exception:
        pass


def _crop_sheet_to_union(ws, min_row: int, max_row: int, min_col: int, max_col: int) -> None:
    """
    Deletes rows/columns outside union. After this, union becomes A1:...
    """
    # 1) Delete rows BELOW max_row (do this first so indices above don't shift)
    if ws.max_row > max_row:
        ws.delete_rows(max_row + 1, ws.max_row - max_row)

    # 2) Delete rows ABOVE min_row (do this after bottom deletion)
    if min_row > 1:
        ws.delete_rows(1, min_row - 1)

    # After deleting rows above, the kept area is now shifted up.
    # 3) Delete cols to the RIGHT of max_col (do this first)
    if ws.max_column > max_col:
        ws.delete_cols(max_col + 1, ws.max_column - max_col)

    # 4) Delete cols to the LEFT of min_col
    if min_col > 1:
        ws.delete_cols(1, min_col - 1)

    # Now the kept rectangle should start at A1
    # Reset print titles (they often reference deleted rows/cols)
    try:
        ws.print_title_rows = None
        ws.print_title_cols = None
    except Exception:
        pass

    # Reset print area to the full remaining sheet bounds
    try:
        last_col = get_column_letter(ws.max_column if ws.max_column >= 1 else 1)
        last_row = ws.max_row if ws.max_row >= 1 else 1
        ws.print_area = f"$A$1:${last_col}${last_row}"
    except Exception:
        pass


def translate_xlsm_bytes(
    xlsm_bytes: bytes,
    translator: OpenAITranslator,
    source_lang: str = "pt-BR",
    target_lang: str = "en",
    glossary: Optional[Dict[str, str]] = None,
    extra_instructions: str = "",
    on_progress: Optional[Callable[[str, int, int], None]] = None,
) -> bytes:
    """
    XLSM translator:
    - Uses each worksheet's defined PRINT AREA (strict).
    - Translates text ONLY inside PRINT AREA.
    - Clears any cells inside the UNION rectangle that are NOT in the print area ranges.
    - Deletes rows/cols outside the UNION rectangle so final sheet contains only print area region.
    - Preserves macros via keep_vba=True.
    """
    glossary = glossary or {}

    wb = openpyxl.load_workbook(io.BytesIO(xlsm_bytes), keep_vba=True, data_only=False)
    sheets = wb.worksheets
    total_sheets = max(1, len(sheets))

    if on_progress:
        on_progress("sheets", 0, total_sheets)

    for s_idx, ws in enumerate(sheets, start=1):
        if on_progress:
            on_progress("sheets", s_idx, total_sheets)

        ranges = _parse_print_area(getattr(ws, "print_area", "") or "")
        # Strict: if no print area, do not modify the sheet at all
        if not ranges:
            continue

        min_row, max_row, min_col, max_col = _union_bounds(ranges)

        # Unmerge anything that would break when cropping
        _unmerge_outside_union(ws, min_row, max_row, min_col, max_col)

        # Collect translatable cells (only within print area ranges)
        items: List[TranslationItem] = []
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
                    if not isinstance(v, str):
                        continue
                    txt = v.strip()
                    if not txt:
                        continue
                    # skip formulas
                    if cell.data_type == "f" or txt.startswith("="):
                        continue

                    item_id = f"{ws.title}!{cell.coordinate}"
                    items.append(TranslationItem(item_id, v))
                    coords.append((item_id, cell.row, cell.column))

        total_cells = len(items)
        done = 0
        if on_progress:
            on_progress("cells", 0, max(1, total_cells))

        mapping: Dict[str, str] = {}
        for ch in chunk_items(items):
            mapping.update(
                translator.translate_batch(
                    ch,
                    source_lang=source_lang,
                    target_lang=target_lang,
                    glossary=glossary,
                    extra_instructions=extra_instructions,
                )
            )
            done += len(ch)
            if on_progress:
                on_progress("cells", min(done, total_cells), max(1, total_cells))

        # Write translations back into same cells
        for item_id, r, c in coords:
            new_text = mapping.get(item_id)
            if isinstance(new_text, str):
                ws.cell(row=r, column=c).value = apply_glossary_hard(new_text, glossary)

        # Clear any cells that are inside the UNION rectangle but NOT inside the actual print area ranges
        # (So the final remaining content corresponds only to print area, even if print area had multiple blocks.)
        for r in range(min_row, max_row + 1):
            for c in range(min_col, max_col + 1):
                if not _cell_in_ranges(r, c, ranges):
                    cell = ws.cell(row=r, column=c)
                    if cell.value is not None:
                        cell.value = None

        # Crop by deleting rows/cols outside the union rectangle
        _crop_sheet_to_union(ws, min_row, max_row, min_col, max_col)

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()
