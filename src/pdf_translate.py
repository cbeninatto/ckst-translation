from __future__ import annotations

from io import BytesIO
from typing import Callable, Dict, Optional

import fitz  # PyMuPDF

from .openai_translate import OpenAITranslator


def _split_text(text: str, max_chars: int = 3000):
    if len(text) <= max_chars:
        return [text]
    parts = []
    cur = []
    cur_len = 0
    for para in text.split("\n\n"):
        chunk = para.strip()
        if not chunk:
            continue
        add_len = len(chunk) + 2
        if cur and (cur_len + add_len) > max_chars:
            parts.append("\n\n".join(cur))
            cur = [chunk]
            cur_len = len(chunk)
        else:
            cur.append(chunk)
            cur_len += add_len
    if cur:
        parts.append("\n\n".join(cur))
    return parts


def _add_translation_pages(out_doc: fitz.Document, width: float, height: float, title: str, text: str) -> None:
    rect = fitz.Rect(48, 72, width - 48, height - 48)
    full = f"{title}\n\n{text}".strip()

    for fontsize in range(12, 7, -1):
        page = out_doc.new_page(width=width, height=height)
        remaining_area = page.insert_textbox(rect, full, fontsize=fontsize, fontname="helv")
        if remaining_area >= 0:
            return
        out_doc.delete_page(out_doc.page_count - 1)

    for idx, chunk in enumerate(_split_text(text, max_chars=2600), start=1):
        page = out_doc.new_page(width=width, height=height)
        chunk_title = title if idx == 1 else f"{title} (cont. {idx})"
        page.insert_textbox(rect, f"{chunk_title}\n\n{chunk}", fontsize=9, fontname="helv")


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

        # Copy original page first
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

        w = page.rect.width
        h = page.rect.height
        _add_translation_pages(
            out_doc=out,
            width=w,
            height=h,
            title=f"English translation — page {i+1}",
            text=translated,
        )

    if on_progress:
        on_progress(total, total)

    buf = BytesIO()
    out.save(buf)
    return buf.getvalue()
