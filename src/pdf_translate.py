# src/pdf_translate.py
from __future__ import annotations

from dataclasses import dataclass
from typing import Callable, List, Optional, Tuple
import re

import fitz  # PyMuPDF
import numpy as np

from .openai_translate import OpenAITranslator


BULLET_ONLY_RE = re.compile(r"^[\s•●\u2022\-\–\—]+$")


@dataclass
class PdfLineItem:
    page_index: int
    rect: fitz.Rect
    text: str
    fontname: str
    fontsize: float
    color: Tuple[float, float, float]  # rgb floats 0..1


def _int_to_rgb_floats(color_int: int) -> Tuple[float, float, float]:
    # PyMuPDF returns 0xRRGGBB as int
    r = (color_int >> 16) & 255
    g = (color_int >> 8) & 255
    b = color_int & 255
    return (r / 255.0, g / 255.0, b / 255.0)


def _pick_base14_font(span_font_name: str) -> str:
    """
    Use Base14 fonts (work on Streamlit Cloud without font files).
    """
    fn = (span_font_name or "").lower()
    bold = "bold" in fn
    italic = ("italic" in fn) or ("oblique" in fn)

    if bold and italic:
        return "Helvetica-BoldOblique"
    if bold:
        return "Helvetica-Bold"
    if italic:
        return "Helvetica-Oblique"
    return "Helvetica"


def _safe_progress(cb: Optional[Callable], label: str, frac: Optional[float] = None) -> None:
    if not cb:
        return
    try:
        # common pattern: cb(label, frac)
        cb(label, frac)
        return
    except TypeError:
        pass
    try:
        # common pattern: cb(label)
        cb(label)
        return
    except Exception:
        return


def _render_page_rgb(page: fitz.Page, matrix: fitz.Matrix) -> np.ndarray:
    pix = page.get_pixmap(matrix=matrix, alpha=False)
    img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
    return img[:, :, :3]  # RGB


def _sample_bg_color(img_rgb: np.ndarray, rect: fitz.Rect, matrix: fitz.Matrix) -> Tuple[float, float, float]:
    """
    Estimate the background color inside rect by masking out darker pixels (likely text).
    Works well for solid-color headers like the red bar in KIT NATAL.
    """
    pr = rect * matrix
    x0, y0, x1, y1 = map(int, [pr.x0, pr.y0, pr.x1, pr.y1])
    h, w = img_rgb.shape[:2]

    x0 = max(0, min(w - 1, x0))
    x1 = max(0, min(w, x1))
    y0 = max(0, min(h - 1, y0))
    y1 = max(0, min(h, y1))

    if x1 <= x0 + 1 or y1 <= y0 + 1:
        return (1.0, 1.0, 1.0)

    crop = img_rgb[y0:y1, x0:x1]
    if crop.size == 0:
        return (1.0, 1.0, 1.0)

    brightness = crop.mean(axis=2)
    # remove very dark pixels (text strokes)
    mask = brightness > 40

    bg = crop[mask]
    if bg.size == 0:
        mean = crop.reshape(-1, 3).mean(axis=0)
    else:
        mean = bg.reshape(-1, 3).mean(axis=0)

    return (float(mean[0] / 255.0), float(mean[1] / 255.0), float(mean[2] / 255.0))


def _insert_text_fit(
    page: fitz.Page,
    rect: fitz.Rect,
    text: str,
    fontname: str,
    fontsize: float,
    color: Tuple[float, float, float],
    align: int = fitz.TEXT_ALIGN_LEFT,
) -> None:
    """
    Insert text into rect and shrink font until it fits (best effort).
    """
    fs = max(4.0, float(fontsize or 10))
    for _ in range(18):
        rc = page.insert_textbox(
            rect,
            text,
            fontname=fontname,
            fontsize=fs,
            color=color,
            align=align,
        )
        if rc >= 0:
            return
        fs -= 0.75
        if fs < 4.0:
            break

    # final attempt (may overflow slightly, but prevents blank output)
    page.insert_textbox(rect, text, fontname=fontname, fontsize=max(4.0, fs), color=color, align=align)


def _extract_pdf_line_items(doc: fitz.Document) -> List[PdfLineItem]:
    items: List[PdfLineItem] = []

    for pno in range(doc.page_count):
        page = doc.load_page(pno)
        d = page.get_text("dict")

        for block in d.get("blocks", []):
            if block.get("type") != 0:
                continue

            for line in block.get("lines", []):
                spans = line.get("spans", [])
                if not spans:
                    continue

                line_text = "".join(s.get("text", "") for s in spans).strip()
                if not line_text:
                    continue

                # Skip bullet-only items (keeps “●” markers untouched)
                if BULLET_ONLY_RE.match(line_text):
                    continue

                rect: Optional[fitz.Rect] = None
                max_size = 0.0
                fontname = "Helvetica"
                color = (0.0, 0.0, 0.0)

                for s in spans:
                    srect = fitz.Rect(s["bbox"])
                    rect = srect if rect is None else (rect | srect)
                    max_size = max(max_size, float(s.get("size", 0) or 0))
                    fontname = _pick_base14_font(s.get("font", ""))
                    color = _int_to_rgb_floats(int(s.get("color", 0)))

                if rect is None:
                    continue

                items.append(
                    PdfLineItem(
                        page_index=pno,
                        rect=rect,
                        text=line_text,
                        fontname=fontname,
                        fontsize=max_size or 10.0,
                        color=color,
                    )
                )

    return items


def _translate_texts(
    translator: OpenAITranslator,
    texts: List[str],
    source_lang: str,
    target_lang: str,
    glossary: str,
    extra_instructions: str,
) -> List[str]:
    # Try common translator APIs (keeps backward compatibility)
    if hasattr(translator, "translate_texts"):
        return translator.translate_texts(
            texts=texts,
            source_lang=source_lang,
            target_lang=target_lang,
            glossary=glossary,
            extra_instructions=extra_instructions,
        )
    if hasattr(translator, "translate_many"):
        return translator.translate_many(
            texts, source_lang=source_lang, target_lang=target_lang, glossary=glossary, extra_instructions=extra_instructions
        )

    # Fallback to per-text calls
    out: List[str] = []
    for t in texts:
        if hasattr(translator, "translate_text"):
            out.append(
                translator.translate_text(
                    text=t,
                    source_lang=source_lang,
                    target_lang=target_lang,
                    glossary=glossary,
                    extra_instructions=extra_instructions,
                )
            )
        else:
            out.append(translator.translate(t))  # last resort
    return out


def translate_pdf_bytes(
    file_bytes: bytes,
    translator: OpenAITranslator,
    source_lang: str,
    target_lang: str,
    glossary: str,
    extra_instructions: str,
    progress_callback: Optional[Callable] = None,
) -> bytes:
    """
    Translate a PDF by replacing text on the SAME page/positions.
    Fixes the common bug where apply_redactions() is called AFTER inserting translated text.
    Also preserves colored headers by sampling background color for redaction fill.
    """
    doc = fitz.open(stream=file_bytes, filetype="pdf")

    # Gather items per page (keeps layout stable, and lets us redact then write)
    all_items = _extract_pdf_line_items(doc)

    # Group items by page
    items_by_page: List[List[PdfLineItem]] = [[] for _ in range(doc.page_count)]
    for it in all_items:
        items_by_page[it.page_index].append(it)

    for pno in range(doc.page_count):
        page = doc.load_page(pno)
        page_items = items_by_page[pno]
        _safe_progress(progress_callback, f"PDF: page {pno+1}/{doc.page_count}", (pno / max(1, doc.page_count)))

        if not page_items:
            continue

        # Render once for background sampling
        mat = fitz.Matrix(2, 2)  # decent quality for color sampling
        img_rgb = _render_page_rgb(page, mat)

        texts = [it.text for it in page_items]
        translations = _translate_texts(
            translator=translator,
            texts=texts,
            source_lang=source_lang,
            target_lang=target_lang,
            glossary=glossary,
            extra_instructions=extra_instructions,
        )

        # 1) Add ALL redactions first (with sampled bg fill), 2) apply once, 3) insert translations
        for it in page_items:
            bg = _sample_bg_color(img_rgb, it.rect, mat)
            page.add_redact_annot(it.rect, fill=bg)

        # IMPORTANT: apply redactions BEFORE inserting translated text
        page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE)

        # Now place translated text
        for it, tr in zip(page_items, translations):
            _insert_text_fit(
                page=page,
                rect=it.rect,
                text=tr,
                fontname=it.fontname,
                fontsize=it.fontsize,
                color=it.color,
                align=fitz.TEXT_ALIGN_LEFT,
            )

    _safe_progress(progress_callback, "PDF: done", 1.0)
    out = doc.tobytes()
    doc.close()
    return out
