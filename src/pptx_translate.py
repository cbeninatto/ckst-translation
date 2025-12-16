from __future__ import annotations

from dataclasses import dataclass
from io import BytesIO
from typing import Dict, List, Tuple

from pptx import Presentation

from .openai_translate import OpenAITranslator, TranslationItem, chunk_items


@dataclass
class PptxPointer:
    # points to a paragraph-like object whose runs we can rewrite
    kind: str  # "shape_paragraph" | "cell_paragraph"
    slide_idx: int
    shape_idx: int
    para_idx: int
    row_idx: int | None = None
    col_idx: int | None = None


def extract_pptx_items(prs: Presentation) -> Tuple[List[TranslationItem], Dict[str, PptxPointer]]:
    items: List[TranslationItem] = []
    pointers: Dict[str, PptxPointer] = {}

    for s_i, slide in enumerate(prs.slides):
        for sh_i, shape in enumerate(slide.shapes):
            # Tables
            if hasattr(shape, "has_table") and shape.has_table:
                table = shape.table
                for r in range(len(table.rows)):
                    for c in range(len(table.columns)):
                        cell = table.cell(r, c)
                        if not cell.text_frame:
                            continue
                        for p_i, para in enumerate(cell.text_frame.paragraphs):
                            txt = para.text.strip()
                            if not txt:
                                continue
                            tid = f"s{s_i}_sh{sh_i}_cell{r}_{c}_p{p_i}"
                            items.append(TranslationItem(id=tid, text=txt))
                            pointers[tid] = PptxPointer(
                                kind="cell_paragraph",
                                slide_idx=s_i,
                                shape_idx=sh_i,
                                para_idx=p_i,
                                row_idx=r,
                                col_idx=c,
                            )
                continue

            # Regular text frames
            if not getattr(shape, "has_text_frame", False):
                continue
            tf = shape.text_frame
            if tf is None:
                continue
            for p_i, para in enumerate(tf.paragraphs):
                txt = para.text.strip()
                if not txt:
                    continue
                tid = f"s{s_i}_sh{sh_i}_p{p_i}"
                items.append(TranslationItem(id=tid, text=txt))
                pointers[tid] = PptxPointer(
                    kind="shape_paragraph",
                    slide_idx=s_i,
                    shape_idx=sh_i,
                    para_idx=p_i,
                )

    return items, pointers


def _rewrite_paragraph_keep_first_run(paragraph, new_text: str) -> None:
    """
    python-pptx does not support deleting runs cleanly.
    We keep formatting by writing into the first run and blanking the rest.
    If no runs exist, we add one.
    """
    runs = list(paragraph.runs)
    if not runs:
        run = paragraph.add_run()
        run.text = new_text
        return

    runs[0].text = new_text
    for r in runs[1:]:
        r.text = ""


def apply_pptx_translations(prs: Presentation, translations: Dict[str, str], pointers: Dict[str, PptxPointer]) -> None:
    for tid, translated in translations.items():
        ptr = pointers.get(tid)
        if not ptr:
            continue
        slide = prs.slides[ptr.slide_idx]
        shape = slide.shapes[ptr.shape_idx]

        if ptr.kind == "cell_paragraph":
            cell = shape.table.cell(ptr.row_idx, ptr.col_idx)  # type: ignore[arg-type]
            para = cell.text_frame.paragraphs[ptr.para_idx]
            _rewrite_paragraph_keep_first_run(para, translated)
        else:
            para = shape.text_frame.paragraphs[ptr.para_idx]
            _rewrite_paragraph_keep_first_run(para, translated)


def translate_pptx_bytes(
    pptx_bytes: bytes,
    translator: OpenAITranslator,
    source_lang: str,
    target_lang: str,
    glossary: Dict[str, str],
    extra_instructions: str,
    max_chars: int = 18000,
    max_items: int = 60,
) -> bytes:
    prs = Presentation(BytesIO(pptx_bytes))
    items, pointers = extract_pptx_items(prs)

    if not items:
        return pptx_bytes

    all_translations: Dict[str, str] = {}
    for batch in chunk_items(items, max_chars=max_chars, max_items=max_items):
        batch_map = translator.translate_batch(
            batch,
            source_lang=source_lang,
            target_lang=target_lang,
            glossary=glossary,
            extra_instructions=extra_instructions,
        )
        all_translations.update(batch_map)

    apply_pptx_translations(prs, all_translations, pointers)

    out = BytesIO()
    prs.save(out)
    return out.getvalue()
