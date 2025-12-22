from __future__ import annotations

from io import BytesIO
from typing import Callable, Dict, List, Optional, Tuple

from pptx import Presentation

from .openai_translate import OpenAITranslator, TranslationItem, chunk_items


def _rewrite_paragraph(paragraph, new_text: str) -> None:
    runs = list(paragraph.runs)
    if not runs:
        paragraph.add_run().text = new_text
        return
    runs[0].text = new_text
    for r in runs[1:]:
        r.text = ""


def _collect_slide_items(prs: Presentation, slide_idx: int) -> Tuple[List[TranslationItem], List[Tuple[str, object]]]:
    slide = prs.slides[slide_idx]
    items: List[TranslationItem] = []
    targets: List[Tuple[str, object]] = []

    for sh_i, shape in enumerate(slide.shapes):
        # Tables
        if getattr(shape, "has_table", False):
            table = shape.table
            for r in range(len(table.rows)):
                for c in range(len(table.columns)):
                    cell = table.cell(r, c)
                    tf = getattr(cell, "text_frame", None)
                    if not tf:
                        continue
                    for p_i, para in enumerate(tf.paragraphs):
                        txt = (para.text or "").strip()
                        if not txt:
                            continue
                        tid = f"s{slide_idx}_sh{sh_i}_cell{r}_{c}_p{p_i}"
                        items.append(TranslationItem(tid, txt))
                        targets.append((tid, para))
            continue

        # Text frames
        if not getattr(shape, "has_text_frame", False):
            continue
        tf = shape.text_frame
        if not tf:
            continue
        for p_i, para in enumerate(tf.paragraphs):
            txt = (para.text or "").strip()
            if not txt:
                continue
            tid = f"s{slide_idx}_sh{sh_i}_p{p_i}"
            items.append(TranslationItem(tid, txt))
            targets.append((tid, para))

    return items, targets


def translate_pptx_bytes(
    pptx_bytes: bytes,
    translator: OpenAITranslator,
    source_lang: str,
    target_lang: str,
    glossary: Dict[str, str],
    extra_instructions: str,
    on_progress: Optional[Callable[[int, int], None]] = None,  # (slide_1based, total)
) -> bytes:
    prs = Presentation(BytesIO(pptx_bytes))
    total = len(prs.slides)

    for s_i in range(total):
        if on_progress:
            on_progress(s_i + 1, total)

        items, targets = _collect_slide_items(prs, s_i)
        if not items:
            continue

        translations: Dict[str, str] = {}
        for batch in chunk_items(items):
            translations.update(
                translator.translate_batch(
                    batch,
                    source_lang=source_lang,
                    target_lang=target_lang,
                    glossary=glossary,
                    extra_instructions=extra_instructions,
                )
            )

        for tid, para in targets:
            if tid in translations:
                _rewrite_paragraph(para, translations[tid])

    out = BytesIO()
    prs.save(out)
    return out.getvalue()
