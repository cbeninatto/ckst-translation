from __future__ import annotations

from io import BytesIO
from typing import Callable, Dict, List, Optional, Tuple

from pptx import Presentation

from .openai_translate import OpenAITranslator, TranslationItem


def replace_text_on_slide(slide, old_text, new_text):
    """
    Replaces old_text with new_text directly on the slide.
    """
    for shape in slide.shapes:
        if hasattr(shape, "has_text_frame") and shape.has_text_frame:
            for para in shape.text_frame.paragraphs:
                if old_text in para.text:
                    para.text = para.text.replace(old_text, new_text)


def translate_pptx_bytes(
    pptx_bytes: bytes,
    translator: OpenAITranslator,
    source_lang: str,
    target_lang: str,
    glossary: Dict[str, str],
    extra_instructions: str,
    on_progress: Optional[Callable[[int, int], None]] = None,  # (slide_idx_1based, total_slides)
) -> bytes:
    prs = Presentation(BytesIO(pptx_bytes))
    total_slides = len(prs.slides)

    for s_i in range(total_slides):
        if on_progress:
            on_progress(s_i + 1, total_slides)

        slide = prs.slides[s_i]
        items, targets = _collect_slide_items(prs, s_i)
        if not items:
            continue

        translations: Dict[str, str] = {}
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

        for tid, para in targets:
            if tid in translations:
                translated_text = translations[tid]
                replace_text_on_slide(slide, para.text, translated_text)

    if on_progress:
        on_progress(total_slides, total_slides)

    out = BytesIO()
    prs.save(out)
    return out.getvalue()
