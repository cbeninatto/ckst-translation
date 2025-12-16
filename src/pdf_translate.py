from __future__ import annotations

from io import BytesIO
from typing import Callable, Dict, Optional

import fitz  # PyMuPDF

from .openai_translate import OpenAITranslator


def replace_text_on_page(page, old_text, new_text):
    """
    Replaces old_text with new_text directly on the same page.
    """
    text_instances = page.search_for(old_text)
    for inst in text_instances:
        page.insert_text(inst[:2], new_text, fontsize=12)  # Position the new text at the found location


def translate_pdf_bytes(
    pdf_bytes: bytes,
    translator: OpenAITranslator,
    source_lang: str,
    target_lang: str,
    glossary: Dict[str, str],
    extra_instructions: str,
    on_progress: Optional[Callable[[int, int], None]] = None,  # (page_idx, total_pages)
) -> bytes:
    src = fitz.open(stream=pdf_bytes, filetype="pdf")

    # If there's no extractable text anywhere, return original.
    any_text = False
    for i in range(src.page_count):
        if (src.load_page(i).get_text("text") or "").strip():
            any_text = True
            break
    if not any_text:
        return pdf_bytes

    out = fitz.open()

    total = src.page_count
    for i in range(total):
        if on_progress:
            on_progress(i, total)

        page = src.load_page(i)
        original_text = (page.get_text("text") or "").strip()

        # Copy original page
        out.insert_pdf(src, from_page=i, to_page=i)

        if not original_text:
            continue

        translated = translator.translate_text(
            original_text,
            source_lang=source_lang,
            target_lang=target_lang,
            glossary=glossary,
            extra_instructions=extra_instructions,
        )

        replace_text_on_page(page, original_text, translated)  # Replace the text directly

    if on_progress:
        on_progress(total, total)

    buf = BytesIO()
    out.save(buf)
    return buf.getvalue()
