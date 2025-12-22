import io
from typing import Callable, Dict, List, Optional, Tuple

import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.cell_range import CellRange

from .openai_translate import OpenAITranslator, TranslationItem, chunk_items


def _compute_area_bounds(ws) -> Tuple[int, int, int, int, List[CellRange]]:
    """
    Returns (min_row, min_col, max_row, max_col, ranges)
    - If print_area exists, use it (can be multiple ranges).
    - Otherwise fallback to "used" bounds based on non-empty values.
    """
    pa = getattr(ws, "print_area", None)
    if pa:
        pa = pa.strip()
    if pa:
        parts = [p.strip() for p in pa.split(",") if p.strip()]
        ranges: List[CellRange] = []
        for p in parts:
            if "!" in p:
                p = p.split("!", 1)[1]
            p = p.replace("$", "")
            ranges.append(CellRange(p))
        min_row = min(r.min_row for r in ranges)
        min_col = min(r.min_col for r in ranges)
        max_row = max(r.max_row for r in ranges)
        max_col = max(r.max_col for r in ranges)
        return min_row, min_col, max_row, max_col, ranges

    min_row = min_col = None
    max_row = max_col = 0
    for row in ws.iter_rows():
        for cell in row:
            v = cell.value
            if v is None:
                continue
            if isinstance(v, str) and not v.strip():
                continue
            r = cell.row
            c = cell.column
            min_row = r if min_row is None else min(min_row, r)
            min_col = c if min_col is None else min(min_col, c)
            max_row = max(max_row, r)
            max_col = max(max_col, c)

    if min_row is None:
        min_row = min_col = 1
        max_row = max_col = 1

    return min_row, min_col, max_row, max_col, [
        CellRange(min_col=min_col, min_row=min_row, max_col=max_col, max_row=max_row)
    ]


def translate_xlsm_bytes(
    xlsm_bytes: bytes,
    translator: OpenAITranslator,
    on_progress: Optional[Callable[[str, int, int], None]] = None,
) -> bytes:
    """
    Translates all worksheets.
    Only the defined print area is kept; everything outside is removed:
    - clear cells outside print area (within bounds)
    - delete rows/cols after the print area (safe, no shifting inside)
    - keep macros (keep_vba=True)
    """
    wb = openpyxl.load_workbook(io.BytesIO(xlsm_bytes), keep_vba=True, data_only=False)

    sheets = wb.worksheets
    total_sheets = len(sheets)

    for s_idx, ws in enumerate(sheets, start=1):
        if on_progress:
            on_progress("sheet", s_idx, total_sheets)

        min_row, min_col, max_row, max_col, ranges = _compute_area_bounds(ws)

        allow = set()
        for cr in ranges:
            for r in range(cr.min_row, cr.max_row + 1):
                for c in range(cr.min_col, cr.max_col + 1):
                    allow.add((r, c))

        items: List[TranslationItem] = []
        cell_refs: List[Tuple[int, int, str]] = []
        for r in range(min_row, max_row + 1):
            for c in range(min_col, max_col + 1):
                if (r, c) not in allow:
                    continue
                cell = ws.cell(row=r, column=c)
                v = cell.value
                if not isinstance(v, str):
                    continue
                if not v.strip():
                    continue
                if v.strip().startswith("="):  # formula text
                    continue
                item_id = f"ws{s_idx}_{cell.coordinate}"
                items.append(TranslationItem(item_id, v))
                cell_refs.append((r, c, item_id))

        mapping: Dict[str, str] = {}
        for chunk in chunk_items(items, max_items=80, max_chars=11000):
            mapping.update(translator.translate_batch(chunk))

        for r, c, item_id in cell_refs:
            ws.cell(row=r, column=c).value = mapping.get(item_id, ws.cell(row=r, column=c).value)

        for r in range(min_row, max_row + 1):
            for c in range(min_col, max_col + 1):
                if (r, c) in allow:
                    continue
                ws.cell(row=r, column=c).value = None

        if ws.max_column > max_col:
            ws.delete_cols(max_col + 1, ws.max_column - max_col)
        if ws.max_row > max_row:
            ws.delete_rows(max_row + 1, ws.max_row - max_row)

        ws.print_area = f"{get_column_letter(min_col)}{min_row}:{get_column_letter(max_col)}{max_row}"

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()
