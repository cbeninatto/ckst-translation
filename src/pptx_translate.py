import io
from typing import Callable, Optional

from pptx import Presentation

from .openai_translate import OpenAITranslator, TranslationItem, chunk_items


def translate_pptx_bytes(
    pptx_bytes: bytes,
    translator: OpenAITranslator,
    on_progress: Optional[Callable[[str, int, int], None]] = None,
) -> bytes:
    prs = Presentation(io.BytesIO(pptx_bytes))
    total = len(prs.slides)

    for s_idx, slide in enumerate(prs.slides, start=1):
        if on_progress:
            on_progress("slide", s_idx, total)

        items = []
        targets = []
        idx = 0

        for shape in slide.shapes:
            if not getattr(shape, "has_text_frame", False):
                continue
            tf = shape.text_frame
            text = tf.text
            if not text or not text.strip():
                continue

            item_id = f"s{s_idx}_t{idx}"
            items.append(TranslationItem(item_id, text))
            targets.append((shape, item_id))
            idx += 1

        if not items:
            continue

        mapping = {}
        for chunk in chunk_items(items, max_items=30, max_chars=8000):
            mapping.update(translator.translate_batch(chunk))

        for shape, item_id in targets:
            translated = mapping.get(item_id, shape.text_frame.text)
            shape.text_frame.text = translated

    out = io.BytesIO()
    prs.save(out)
    return out.getvalue()
