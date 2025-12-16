from __future__ import annotations

from io import BytesIO
from typing import Callable, Dict, List, Optional, Tuple

from pptx import Presentation

from .openai_translate import OpenAITranslator, TranslationItem


def _rewrite_paragraph_keep_first_run(paragraph, new_text: str) -> None:
    """
    Keeps formatting by writing into the first run and clearing the rest.
    This is the safest "in-place" replacement python-pptx can do.
    """
    runs = list(paragraph.runs)
    if not runs:
        run = paragraph.add_run()
        run.text = new_text
        return

    runs[0].text = new_text
    for r in runs[1:]:
        r.text = ""


def _collect_slide_items(
    prs: Presentation, slide_idx: int
) -> Tuple[List[TranslationItem], List[Tuple[str, object]]]:
    """
    Collect all non-empty paragraph texts on a slide, including table cells.
    Returns:
      items: TranslationItem(id, text)
      targets: list of (id, paragraph_object) for rewriting
    """
    slide = prs.slides[slide_idx]
    items: List[TranslationItem] = []
    targets: List[Tuple[str, object]] = []

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
                        txt = (para.text or "").strip()
                        if not txt:
                            continue
                        tid = f"s{slide_idx}_sh{sh_i}_cell{r}_{c}_p{p_i}"
                        items.append(TranslationItem(id=tid, text=txt))
                        targets.append((tid, para))
            continue

        # Regular text frames
        if not getattr(shape, "has_text_frame", False):
            continue
        tf = shape.text_frame
        if tf is None:
            continue

        for p_i, para in enumerate(tf.paragraphs):
            txt = (para.text or "").strip()
            if not txt:
                continue
            tid = f"s{slide_idx}_sh{sh_i}_p{p_i}"
            items.append(TranslationItem(id=tid, text=txt))
            targets.append((tid, para))

    return items, targets


def _safe_chunk(
    items: List[TranslationItem],
    max_chars: int = 18000,
    max_items: int = 60,
) -> List[List[TranslationItem]]:
    """
    Internal safety chunking (not a UI limit). Prevents hitting model context limits.
    """
    batches: List[List[TranslationItem]] = []
    cur: List[TranslationItem] = []
    cur_chars = 0

    for it in items:
        tlen = len(it.text or "")
        if cur and (len(cur) >= max_items or (cur_chars + tlen) > max_chars):
            batches.append(cur)
            cur = []
            cur_chars = 0
        cur.append(it)
        cur_chars += tlen

    if cur:
        batches.append(cur)

    return batches


def translate_pptx_bytes(
    pptx_bytes: bytes,
    translator: OpenAITranslator,
    source_lang: str,
    target_lang: str,
    glossary: Dict[str, str],
    extra_instructions: str,
    on_progress: Optional[Callable[[int, int], None]] = None,  # (slide_1based, total_slides)
) -> bytes:
    prs = Presentation(BytesIO(pptx_bytes))
    total_slides = len(prs.slides)

    for s_i in range(total_slides):
        if on_progress:
            on_progress(s_i + 1, total_slides)

        items, targets = _collect_slide_items(prs, s_i)
        if not items:
            continue

        translations: Dict[str, str] = {}

        # Translate (internally chunk only if needed)
        for batch in _safe_chunk(items):
            translations.update(
                translator.translate_batch(
                    batch,
                    source_lang=source_lang,
                    target_lang=target_lang,
                    glossary=glossary,
                    extra_instructions=extra_instructions,
                )
            )

        # Apply translations directly on the same slide
        for tid, para in targets:
            if tid in translations:
                _rewrite_paragraph_keep_first_run(para, translations[tid])

    if on_progress:
        on_progress(total_slides, total_slides)

    out = BytesIO()
    prs.save(out)
    return out.getvalue()
