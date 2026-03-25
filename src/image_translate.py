# src/image_translate.py
from __future__ import annotations

import base64
import io
import json
import os
import re
from collections import Counter
from typing import Callable, Dict, List, Optional, Literal, Tuple

from PIL import Image, ImageDraw, ImageFont, ImageFilter
from pydantic import BaseModel, Field
from openai import OpenAI

try:
    from .text_utils import apply_glossary_hard  # type: ignore
except Exception:
    apply_glossary_hard = None  # type: ignore


class BBox(BaseModel):
    x0: float = Field(ge=0.0, le=1.0)
    y0: float = Field(ge=0.0, le=1.0)
    x1: float = Field(ge=0.0, le=1.0)
    y1: float = Field(ge=0.0, le=1.0)


class TextBlock(BaseModel):
    id: str
    text_src: str
    text_tgt: str
    bbox: BBox
    style: Literal["normal", "bold"] = "normal"
    align: Literal["left", "center", "right"] = "left"


class ImageTranslationResult(BaseModel):
    blocks: List[TextBlock]


_RX_NUMBER_UNIT = re.compile(r"(\d+(?:[.,]\d+)?\s*(?:cm|mm|m|kg|g|pcs|pc|un|usd|r\$|\$|€|£)?)", re.IGNORECASE)
_RX_ONLY_PRICEISH = re.compile(r"^[\sA-Z$€£R0-9.,:+-]+$", re.IGNORECASE)


def _to_data_url(image_bytes: bytes, mime: str) -> str:
    b64 = base64.b64encode(image_bytes).decode("utf-8")
    return f"data:{mime};base64,{b64}"


def _guess_mime(filename: str) -> str:
    f = filename.lower()
    if f.endswith(".png"):
        return "image/png"
    if f.endswith(".webp"):
        return "image/webp"
    return "image/jpeg"


def _pick_fonts() -> Tuple[Optional[str], Optional[str]]:
    candidates = [
        ("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"),
        ("/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf", "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf"),
    ]
    for reg, bold in candidates:
        if os.path.exists(reg) and os.path.exists(bold):
            return reg, bold
    return None, None


def _luminance(rgb) -> float:
    r, g, b = rgb
    return (0.2126 * r + 0.7152 * g + 0.0722 * b) / 255.0


def _quantize_rgb(rgb: Tuple[int, int, int], step: int = 16) -> Tuple[int, int, int]:
    return tuple((c // step) * step for c in rgb)  # type: ignore


def _dominant_bg_color(img_rgba: Image.Image, px_box: Tuple[int, int, int, int], pad: int = 4) -> Tuple[int, int, int, int]:
    x0, y0, x1, y1 = px_box
    w, h = img_rgba.size

    x0c = max(0, x0 - pad)
    y0c = max(0, y0 - pad)
    x1c = min(w - 1, x1 + pad)
    y1c = min(h - 1, y1 + pad)

    px = img_rgba.load()
    samples: List[Tuple[int, int, int]] = []

    for x in range(x0c, x1c + 1):
        samples.append(px[x, y0c][:3])
        samples.append(px[x, y1c][:3])
    for y in range(y0c, y1c + 1):
        samples.append(px[x0c, y][:3])
        samples.append(px[x1c, y][:3])

    if not samples:
        return (255, 255, 255, 255)

    buckets = Counter(_quantize_rgb(s, step=16) for s in samples)
    dom, _ = buckets.most_common(1)[0]
    return (int(dom[0]), int(dom[1]), int(dom[2]), 255)


def _wrap_to_width(draw: ImageDraw.ImageDraw, text: str, font: ImageFont.ImageFont, max_w: int) -> List[str]:
    if not text:
        return [""]

    try:
        if draw.textlength(text, font=font) <= max_w:
            return [text]
    except Exception:
        pass

    words = text.split()
    if not words:
        return [""]

    lines: List[str] = []
    cur = ""
    for w in words:
        test = (cur + " " + w).strip()
        try:
            length = draw.textlength(test, font=font)
        except Exception:
            length = font.getlength(test) if hasattr(font, "getlength") else len(test) * 6

        if length <= max_w:
            cur = test
        else:
            if cur:
                lines.append(cur)
            cur = w
    if cur:
        lines.append(cur)
    return lines


def _fit_text_in_box(
    draw: ImageDraw.ImageDraw,
    text: str,
    box_w: int,
    box_h: int,
    font_path: Optional[str],
    max_size: int,
    min_size: int = 10,
) -> Tuple[ImageFont.ImageFont, List[str], int]:
    max_size = max(min_size, max_size)

    for size in range(max_size, min_size - 1, -1):
        font = ImageFont.truetype(font_path, size=size) if font_path else ImageFont.load_default()
        lines = _wrap_to_width(draw, text, font, box_w)
        try:
            ascent, descent = font.getmetrics()  # type: ignore
            line_h = ascent + descent + int(size * 0.12)
        except Exception:
            line_h = int(size * 1.20)

        if line_h * len(lines) <= box_h:
            return font, lines, line_h

    font = ImageFont.truetype(font_path, size=min_size) if font_path else ImageFont.load_default()
    lines = _wrap_to_width(draw, text, font, box_w)
    try:
        ascent, descent = font.getmetrics()  # type: ignore
        line_h = ascent + descent + int(min_size * 0.12)
    except Exception:
        line_h = int(min_size * 1.20)
    return font, lines, line_h


def _apply_glossary_force(text_tgt: str, glossary: Dict[str, str]) -> str:
    if not glossary:
        return text_tgt

    if apply_glossary_hard:
        try:
            return apply_glossary_hard(text_tgt, glossary)  # type: ignore
        except Exception:
            pass

    out = text_tgt
    for src, tgt in glossary.items():
        if not src or not tgt:
            continue
        out = out.replace(src, tgt)
        out = out.replace(src.upper(), tgt)
        out = out.replace(src.lower(), tgt)
    return out


def _preserve_numbers_and_formats(text_src: str, text_tgt: str) -> str:
    src = (text_src or "").strip()
    tgt = (text_tgt or "").strip()

    if not src:
        return tgt

    if _RX_ONLY_PRICEISH.match(src) and len(re.sub(r"[^A-Za-z]", "", src)) <= 3:
        return src

    src_tokens = _RX_NUMBER_UNIT.findall(src)
    if not src_tokens:
        return tgt

    tgt_tokens = _RX_NUMBER_UNIT.findall(tgt)
    if not tgt_tokens:
        return tgt

    out = tgt
    for i, tok in enumerate(src_tokens):
        if i < len(tgt_tokens):
            out = out.replace(tgt_tokens[i], tok, 1)
    return out


def _split_multiline_blocks(blocks: List[TextBlock]) -> List[TextBlock]:
    out: List[TextBlock] = []
    for b in blocks:
        src_lines = [ln.strip() for ln in (b.text_src or "").splitlines() if ln.strip()]
        tgt_lines = [ln.strip() for ln in (b.text_tgt or "").splitlines() if ln.strip()]

        if len(src_lines) <= 1 and len(tgt_lines) <= 1:
            out.append(b)
            continue

        n = max(len(src_lines), len(tgt_lines), 1)
        y0, y1 = b.bbox.y0, b.bbox.y1
        step = (y1 - y0) / n if n else (y1 - y0)

        for i in range(n):
            line_src = src_lines[i] if i < len(src_lines) else (src_lines[-1] if src_lines else "")
            line_tgt = tgt_lines[i] if i < len(tgt_lines) else (tgt_lines[-1] if tgt_lines else "")
            out.append(
                TextBlock(
                    id=f"{b.id}_{i+1}",
                    text_src=line_src,
                    text_tgt=line_tgt,
                    bbox=BBox(
                        x0=b.bbox.x0,
                        x1=b.bbox.x1,
                        y0=max(0.0, min(1.0, y0 + step * i)),
                        y1=max(0.0, min(1.0, y0 + step * (i + 1))),
                    ),
                    style=b.style,
                    align=b.align,
                )
            )
    return out


def _normalize_bbox(b: BBox) -> BBox:
    x0 = max(0.0, min(1.0, min(b.x0, b.x1)))
    x1 = max(0.0, min(1.0, max(b.x0, b.x1)))
    y0 = max(0.0, min(1.0, min(b.y0, b.y1)))
    y1 = max(0.0, min(1.0, max(b.y0, b.y1)))
    return BBox(x0=x0, y0=y0, x1=x1, y1=y1)


def translate_image_bytes(
    image_bytes: bytes,
    filename: str,
    api_key: str,
    model: str,
    source_lang: str,
    target_lang: str,
    glossary: Dict[str, str],
    extra_instructions: str = "",
    progress_cb: Optional[Callable[[float, str], None]] = None,
) -> bytes:
    def progress(p: float, msg: str):
        if progress_cb:
            try:
                progress_cb(float(p), msg)
            except Exception:
                pass

    img = Image.open(io.BytesIO(image_bytes)).convert("RGBA")
    orig_img = img.copy()
    W, H = img.size

    progress(0.05, f"{filename}: detecting text regions…")

    client = OpenAI(api_key=api_key)

    glossary_lines = "\n".join([f"- {k} => {v}" for k, v in (glossary or {}).items() if k and v])
    if not glossary_lines.strip():
        glossary_lines = "(none)"

    extra_lang = ""
    if target_lang.lower().startswith("pt"):
        extra_lang = "Use Brazilian Portuguese (PT-BR)."

    prompt = f"""
You are translating handbag development sheets.

Translate from {source_lang} to {target_lang}. {extra_lang}

OUTPUT REQUIREMENT (VERY IMPORTANT):
- Return ONE line of text per block. Do NOT merge multiple lines into one block.
- For paragraphs/bullets, output each visible line separately as its own block with its own bbox.

For each line, output:
- text_src (exact original)
- text_tgt (translated)
- bbox (x0,y0,x1,y1) normalized 0..1, tight around THAT LINE
- style: "bold" if heading/bold
- align: left/center/right if clearly aligned

CRITICAL RULES:
- Keep all numbers, currency, and dimension formatting EXACTLY as in source (including comma decimals like 5,10).
- Keep units exactly (cm, mm, USD, etc).
- Do not invent new measurements.
- Apply glossary EXACTLY when the source term appears:
{glossary_lines}

Extra instructions:
{extra_instructions}
""".strip()

    mime = _guess_mime(filename)
    data_url = _to_data_url(image_bytes, mime)

    try:
        resp = client.responses.parse(
            model=model,
            input=[
                {
                    "role": "user",
                    "content": [
                        {"type": "input_text", "text": prompt},
                        {"type": "input_image", "image_url": data_url},
                    ],
                }
            ],
            text_format=ImageTranslationResult,
        )
        result = resp.output_parsed  # type: ignore
        if result is None:
            raise ValueError("No parsed output returned.")
    except Exception:
        resp = client.responses.create(
            model=model,
            input=[
                {
                    "role": "user",
                    "content": [
                        {"type": "input_text", "text": prompt},
                        {"type": "input_image", "image_url": data_url},
                    ],
                }
            ],
        )
        raw = getattr(resp, "output_text", None)
        if not raw:
            raw = str(resp)
        result = ImageTranslationResult(**json.loads(raw))

    blocks = []
    for b in result.blocks:
        b.bbox = _normalize_bbox(b.bbox)
        b.text_src = (b.text_src or "").replace("\r", "\n")
        b.text_tgt = (b.text_tgt or "").replace("\r", "\n")
        blocks.append(b)

    blocks = _split_multiline_blocks(blocks)

    progress(0.35, f"{filename}: erasing + overwriting text…")

    draw = ImageDraw.Draw(img)
    regular_font_path, bold_font_path = _pick_fonts()

    px_items: List[Tuple[TextBlock, Tuple[int, int, int, int], Tuple[int, int, int, int]]] = []
    for b in blocks:
        x0 = int(b.bbox.x0 * W)
        y0 = int(b.bbox.y0 * H)
        x1 = int(b.bbox.x1 * W)
        y1 = int(b.bbox.y1 * H)

        box_h = max(1, y1 - y0)
        pad = max(6, int(box_h * 0.18))
        x0p = max(0, x0 - pad)
        y0p = max(0, y0 - pad)
        x1p = min(W - 1, x1 + pad)
        y1p = min(H - 1, y1 + pad)

        if x1p <= x0p or y1p <= y0p:
            continue

        px_box = (x0p, y0p, x1p, y1p)
        bg = _dominant_bg_color(orig_img, px_box, pad=4)
        px_items.append((b, px_box, bg))

    # erase
    for _, (x0, y0, x1, y1), bg in px_items:
        crop = orig_img.crop((x0, y0, x1, y1))
        blur_r = max(6, min(18, int(min(x1 - x0, y1 - y0) * 0.12)))
        img.paste(crop.filter(ImageFilter.GaussianBlur(radius=blur_r)), (x0, y0))
        draw.rectangle((x0, y0, x1, y1), fill=bg)

    # write
    for i, (b, (x0, y0, x1, y1), bg) in enumerate(px_items, start=1):
        box_w = max(1, x1 - x0)
        box_h = max(1, y1 - y0)

        fg = (0, 0, 0, 255)
        if _luminance(bg[:3]) < 0.45:
            fg = (255, 255, 255, 255)

        text_tgt = (b.text_tgt or "").strip()
        text_tgt = _apply_glossary_force(text_tgt, glossary)
        text_tgt = _preserve_numbers_and_formats(b.text_src, text_tgt)
        if not text_tgt:
            text_tgt = (b.text_src or "").strip()

        font_path = bold_font_path if b.style == "bold" else regular_font_path
        max_size = max(10, int(box_h * 0.82))
        font, lines, line_h = _fit_text_in_box(draw, text_tgt, box_w, box_h, font_path, max_size, min_size=10)

        total_h = line_h * len(lines)
        yy = y0 + max(0, (box_h - total_h) // 2)

        for line in lines:
            try:
                line_w = int(draw.textlength(line, font=font))
            except Exception:
                line_w = len(line) * 6

            if b.align == "center":
                xx = x0 + max(0, (box_w - line_w) // 2)
            elif b.align == "right":
                xx = x0 + max(0, box_w - line_w)
            else:
                xx = x0

            draw.text((xx, yy), line, fill=fg, font=font)
            yy += line_h
            if yy > y1:
                break

        progress(0.35 + 0.60 * (i / max(1, len(px_items))), f"{filename}: blocks ({i}/{len(px_items)})")

    out = io.BytesIO()
    img.save(out, format="PNG")
    progress(1.0, f"{filename}: done")
    return out.getvalue()
