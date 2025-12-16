from __future__ import annotations

from io import BytesIO
from typing import Dict, List

import fitz  # PyMuPDF

from .openai_translate import OpenAITranslator, TranslationItem, chunk_items


def extract_pdf_page_items(doc: fitz.Document) -> List[TranslationItem]:
    items: List[TranslationItem] = []
    for i in range(doc.page_count):
        page = doc.load_page(i)
        txt = (page.get_text("text") or "").strip()
        if not txt:
            continue
        items.append(TranslationItem(id=f"p{i}", text=txt))
    return items


def _split_text(text: str, max_chars: int = 3000) -> List[str]:
    if len(text) <= max_chars:
        return [text]
    parts: List[str] = []
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
    # Try to fit in one page by shrinking font a bit; otherwise split into multiple pages.
    rect = fitz.Rect(48, 72, width - 48, height - 48)
    full = f"{title}\n\n{text}".strip()

    for fontsize in range(12, 7, -1):
        page = out_doc.new_page(width=width, height=height)
        remaining_area = page.insert_textbox(rect, full, fontsize=fontsize, fontname="helv")
        if remaining_area >= 0:
            return
        # doesn't fit — remove and try smaller
        out_doc.delete_page(out_doc.page_count - 1)

    # Still doesn't fit: split
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
    max_chars: int = 18000,
    max_items: int = 25,
) -> bytes:
    src = fitz.open(stream=pdf_bytes, filetype="pdf")
    items = extract_pdf_page_items(src)

    # If there's no extractable text, it's likely scanned/image-only.
    # We still return the original; app will warn the user.
    if not items:
        return pdf_bytes

    translations: Dict[str, str] = {}
    for batch in chunk_items(items, max_chars=max_chars, max_items=max_items):
        batch_map = translator.translate_batch(
            batch,
            source_lang=source_lang,
            target_lang=target_lang,
            glossary=glossary,
            extra_instructions=extra_instructions,
        )
        translations.update(batch_map)

    out = fitz.open()
    for i in range(src.page_count):
        # Copy original page
        out.insert_pdf(src, from_page=i, to_page=i)

        # Add translation page if we have text
        tid = f"p{i}"
        if tid in translations:
            w = src.load_page(i).rect.width
            h = src.load_page(i).rect.height
            _add_translation_pages(
                out_doc=out,
                width=w,
                height=h,
                title=f"English translation — page {i+1}",
                text=translations[tid],
            )

    buf = BytesIO()
    out.save(buf)
    return buf.getvalue()
