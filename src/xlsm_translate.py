import io
from typing import Callable, Dict, List, Optional, Tuple

import openpyxl
from openpyxl.worksheet.cell_range import CellRange

from .openai_translate import OpenAITranslator, TranslationItem, chunk_items
from .text_utils import apply_glossary_hard


def _parse_print_area(print_area: str) -> List[CellRange]:
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
    translator: OpenAITranslator,
    source_lang: str = "pt-BR",
    target_lang: str = "en",
    glossary: Optional[Dict[str, str]] = None,
    extra_instructions: str = "",
    on_progress: Optional[Callable[[str, int, int], None]] = None,
) -> bytes:
    """
    - Translate ONLY inside the sheet's defined print area.
    - Clear content outside print area.
    - Process all worksheets.
    - Preserve macros: keep_vba=True
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
        # strict: if no print area, skip this sheet entirely
        if not ranges:
            continue

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

        # Write back translations
        for item_id, r, c in coords:
            new_text = mapping.get(item_id)
            if isinstance(new_text, str):
                ws.cell(row=r, column=c).value = apply_glossary_hard(new_text, glossary)

        # Clear everything outside print area (existing cells only)
        for (r, c), cell in list(ws._cells.items()):
            if cell.value is None:
                continue
            if not _cell_in_ranges(r, c, ranges):
                cell.value = None

        # Keep print area unchanged
        try:
            ws.print_area = ",".join([cr.coord for cr in ranges])
        except Exception:
            pass

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()
