import io
from typing import Callable, Dict, List, Optional, Tuple

import fitz  # PyMuPDF

from .openai_translate import OpenAITranslator, TranslationItem, chunk_items
from .text_utils import apply_glossary_hard


def _iter_spans(page: fitz.Page) -> List[Tuple[fitz.Rect, str, float]]:
    """
    Returns list of (bbox, text, fontsize) for each span.
    """
    d = page.get_text("dict")
    spans = []
    for block in d.get("blocks", []):
        for line in block.get("lines", []):
            for span in line.get("spans", []):
                text = span.get("text", "")
                if not text or not text.strip():
                    continue
                # skip very short single punctuation
                if len(text.strip()) == 1 and not text.strip().isalnum():
                    continue
                bbox = fitz.Rect(span["bbox"])
                size = float(span.get("size", 10.0))
                spans.append((bbox, text, size))
    return spans


def translate_pdf_bytes(
    pdf_bytes: bytes,
    translator: OpenAITranslator,
    source_lang: str = "pt-BR",
    target_lang: str = "en",
    glossary: Optional[Dict[str, str]] = None,
    extra_instructions: str = "",
    on_progress: Optional[Callable[[str, int, int], None]] = None,
) -> bytes:
    """
    Attempts in-place PDF translation:
    - redacts original text area
    - inserts translated text into same bbox (overlay)
    This keeps the same pages (no extra pages).
    """
    glossary = glossary or {}

    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    total_pages = doc.page_count

    if on_progress:
        on_progress("pages", 0, max(1, total_pages))

    for pno in range(total_pages):
        page = doc[pno]
        if on_progress:
            on_progress("pages", pno + 1, max(1, total_pages))

        spans = _iter_spans(page)
        if not spans:
            continue

        items: List[TranslationItem] = []
        meta: List[Tuple[str, fitz.Rect, float]] = []

        for i, (bbox, text, size) in enumerate(spans):
            item_id = f"p{pno}_s{i}"
            items.append(TranslationItem(item_id, text))
            meta.append((item_id, bbox, size))

        mapping: Dict[str, str] = {}
        total_items = len(items)
        done = 0
        if on_progress:
            on_progress("spans", 0, max(1, total_items))

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
                on_progress("spans", min(done, total_items), max(1, total_items))

        # Redact and write back
        # Note: fill white; if you have colored backgrounds, we can detect/adjust later.
        for item_id, bbox, _size in meta:
            page.add_redact_annot(bbox, fill=(1, 1, 1))
        page.apply_redactions()

        for item_id, bbox, size in meta:
            new_text = mapping.get(item_id, None)
            if not new_text:
                continue
            new_text = apply_glossary_hard(new_text, glossary)

            # Insert translated text into same bounding box
            # Using a standard font; PDFs often embed fonts that can't be reused reliably.
            page.insert_textbox(
                bbox,
                new_text,
                fontsize=size,
                fontname="helv",
                color=(0, 0, 0),
                align=fitz.TEXT_ALIGN_LEFT,
            )

    out = io.BytesIO()
    doc.save(out, deflate=True, garbage=4)
    doc.close()
    return out.getvalue()
