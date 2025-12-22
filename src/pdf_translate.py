from __future__ import annotations

from io import BytesIO
from statistics import median
from typing import Callable, Dict, Optional, List, Tuple

import fitz  # PyMuPDF

from .openai_translate import OpenAITranslator


def _has_letters(s: str) -> bool:
    return any(ch.isalpha() for ch in (s or ""))


def _block_fontsize(block) -> float:
    sizes = []
    for line in block.get("lines", []):
        for span in line.get("spans", []):
            if "size" in span:
                sizes.append(float(span["size"]))
    if not sizes:
        return 10.0
    return float(median(sizes))


def _block_text(block) -> str:
    lines = []
    for line in block.get("lines", []):
        parts = [span.get("text", "") for span in line.get("spans", [])]
        line_text = "".join(parts).strip()
        if line_text:
            lines.append(line_text)
    return "\n".join(lines).strip()


def _insert_fit_text(page: fitz.Page, rect: fitz.Rect, text: str, base_size: float) -> None:
    size = max(6.0, min(14.0, base_size))
    for _ in range(12):
        remaining = page.insert_textbox(rect, text, fontsize=size, fontname="helv")
        if remaining >= 0:
            return
        size -= 0.7
        if size < 5.5:
            break
    page.insert_textbox(rect, text, fontsize=6.0, fontname="helv")


def translate_pdf_bytes(
    pdf_bytes: bytes,
    translator: OpenAITranslator,
    source_lang: str,
    target_lang: str,
    glossary: Dict[str, str],
    extra_instructions: str,
    redact_original: bool = True,
    on_progress: Optional[Callable[[int, int], None]] = None,  # (page_idx, total_pages)
) -> bytes:
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    total = doc.page_count

    # quick check for text-based PDF
    any_text = False
    for i in range(total):
        if (doc.load_page(i).get_text("text") or "").strip():
            any_text = True
            break
    if not any_text:
        return pdf_bytes

    for i in range(total):
        if on_progress:
            on_progress(i, total)

        page = doc.load_page(i)
        d = page.get_text("dict")

        blocks = [b for b in d.get("blocks", []) if b.get("type", 1) == 0]
        if not blocks:
            continue

        to_place: List[Tuple[fitz.Rect, str, float]] = []

        for b in blocks:
            txt = _block_text(b)
            if not txt or not _has_letters(txt):
                continue

            translated = translator.translate_text(
                txt,
                source_lang=source_lang,
                target_lang=target_lang,
                glossary=glossary,
                extra_instructions=extra_instructions,
            )

            bbox = b.get("bbox")
            if not bbox:
                continue
            rect = fitz.Rect(bbox)

            base_size = _block_fontsize(b)
            to_place.append((rect, translated, base_size))

            if redact_original:
                page.add_redact_annot(rect, fill=(1, 1, 1))

        if redact_original:
            page.apply_redactions()

        for rect, translated, base_size in to_place:
            _insert_fit_text(page, rect, translated, base_size)

    if on_progress:
        on_progress(total, total)

    out = BytesIO()
    doc.save(out)
    return out.getvalue()
