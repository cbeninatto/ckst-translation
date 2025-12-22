import io
from typing import Callable, List, NamedTuple, Optional

import fitz  # PyMuPDF

from .openai_translate import OpenAITranslator, TranslationItem, chunk_items


class PdfLineItem(NamedTuple):
    item_id: str
    text: str
    rect: fitz.Rect
    font_size: float


def _extract_line_items(page: fitz.Page) -> List[PdfLineItem]:
    d = page.get_text("dict")
    items: List[PdfLineItem] = []
    line_idx = 0

    for block in d.get("blocks", []):
        if block.get("type") != 0:
            continue
        for line in block.get("lines", []):
            spans = line.get("spans", [])
            if not spans:
                continue

            text = "".join([s.get("text", "") for s in spans]).strip()
            if not text or len(text) <= 1:
                continue

            rect = None
            sizes = []
            for s in spans:
                bbox = s.get("bbox", None)
                if bbox:
                    r = fitz.Rect(bbox)
                    rect = r if rect is None else rect | r
                if "size" in s:
                    sizes.append(float(s["size"]))

            if rect is None:
                continue

            font_size = max(sizes) if sizes else 10.0
            item_id = f"p{page.number}_l{line_idx}"
            items.append(PdfLineItem(item_id=item_id, text=text, rect=rect, font_size=font_size))
            line_idx += 1

    return items


def _insert_translation(page: fitz.Page, rect: fitz.Rect, text: str, font_size: float) -> None:
    fs = max(4.0, float(font_size))
    for _ in range(14):
        rc = page.insert_textbox(rect, text, fontsize=fs, fontname="helv", align=0)
        if isinstance(rc, (int, float)) and rc >= 0:
            return
        fs = max(4.0, fs - 0.5)
    page.insert_textbox(rect, text, fontsize=4.0, fontname="helv", align=0)


def translate_pdf_bytes(
    pdf_bytes: bytes,
    translator: OpenAITranslator,
    on_progress: Optional[Callable[[str, int, int], None]] = None,
) -> bytes:
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    total_pages = doc.page_count

    for page_idx in range(total_pages):
        page = doc.load_page(page_idx)
        if on_progress:
            on_progress("page", page_idx + 1, total_pages)

        items = _extract_line_items(page)
        if not items:
            continue

        t_items = [TranslationItem(it.item_id, it.text) for it in items]
        mapping = {}
        for chunk in chunk_items(t_items, max_items=45, max_chars=7000):
            mapping.update(translator.translate_batch(chunk))

        for it in items:
            page.add_redact_annot(it.rect, fill=(1, 1, 1))
        page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE)

        for it in items:
            translated = mapping.get(it.item_id, it.text)
            _insert_translation(page, it.rect, translated, it.font_size)

    out = io.BytesIO()
    doc.save(out)
    doc.close()
    return out.getvalue()
