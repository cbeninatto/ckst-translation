# src/image_translate.py
from __future__ import annotations

import base64
import io
import json
import os
from typing import Callable, Dict, List, Optional, Literal, Tuple

from PIL import Image, ImageDraw, ImageFont

from pydantic import BaseModel, Field
from openai import OpenAI

# If your project already has this helper, we use it; otherwise we fallback.
try:
    from .text_utils import apply_glossary_hard  # type: ignore
except Exception:
    apply_glossary_hard = None  # type: ignore


# ----------------------------
# Structured output schema
# ----------------------------
class BBox(BaseModel):
    # normalized coords (0..1) relative to image width/height
    x0: float = Field(ge=0.0, le=1.0)
    y0: float = Field(ge=0.0, le=1.0)
    x1: float = Field(ge=0.0, le=1.0)
    y1: float = Field(ge=0.0, le=1.0)


class TextBlock(BaseModel):
    id: str
    text_pt: str
    text_en: str
    bbox: BBox
    style: Literal["normal", "bold"] = "normal"
    align: Literal["left", "center", "right"] = "left"


class ImageTranslationResult(BaseModel):
    blocks: List[TextBlock]


# ----------------------------
# Helpers
# ----------------------------
def _to_data_url(image_bytes: bytes, mime: str) -> str:
    b64 = base64.b64encode(image_bytes).decode("utf-8")
    return f"data:{mime};base64,{b64}"


def _pick_fonts() -> Tuple[Optional[str], Optional[str]]:
    """
    Choose common fonts available on Streamlit Cloud/Linux.
    Fallback to None if not found (we'll use PIL default).
    """
    candidates = [
        ("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"),
        ("/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf", "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf"),
    ]
    for reg, bold in candidates:
        if os.path.exists(reg) and os.path.exists(bold):
            return reg, bold
    return None, None


def _luminance(rgb):
    r, g, b = rgb
    return (0.2126 * r + 0.7152 * g + 0.0722 * b) / 255.0


def _sample_bg_color(img_rgba: Image.Image, px_box: Tuple[int, int, int, int], pad: int = 3) -> Tuple[int, int, int, int]:
    """
    Sample a background color around the bbox (edge pixels) and return RGBA fill.
    Uses a thin ring around the box so we don't average the text itself.
    """
    x0, y0, x1, y1 = px_box
    w, h = img_rgba.size

    x0c = max(0, x0 - pad)
    y0c = max(0, y0 - pad)
    x1c = min(w - 1, x1 + pad)
    y1c = min(h - 1, y1 + pad)

    pixels = img_rgba.load()
    samples: List[Tuple[int, int, int, int]] = []

    # top/bottom edges
    for x in range(x0c, x1c + 1):
        samples.append(pixels[x, y0c])
        samples.append(pixels[x, y1c])

    # left/right edges
    for y in range(y0c, y1c + 1):
        samples.append(pixels[x0c, y])
        samples.append(pixels[x1c, y])

    if not samples:
        return (255, 255, 255, 255)

    r = sum(p[0] for p in samples) // len(samples)
    g = sum(p[1] for p in samples) // len(samples)
    b = sum(p[2] for p in samples) // len(samples)
    return (int(r), int(g), int(b), 255)


def _wrap_to_width(draw: ImageDraw.ImageDraw, text: str, font: ImageFont.ImageFont, max_w: int) -> List[str]:
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
    """
    Find a font size that fits text in the bbox, wrapping as needed.
    Returns (font, lines, line_height).
    """
    max_size = max(min_size, max_size)

    for size in range(max_size, min_size - 1, -1):
        if font_path:
            font = ImageFont.truetype(font_path, size=size)
        else:
            font = ImageFont.load_default()

        lines = _wrap_to_width(draw, text, font, box_w)
        try:
            ascent, descent = font.getmetrics()  # type: ignore
            line_h = ascent + descent + int(size * 0.15)
        except Exception:
            line_h = int(size * 1.25)

        total_h = line_h * len(lines)
        if total_h <= box_h:
            return font, lines, line_h

    # last resort
    if font_path:
        font = ImageFont.truetype(font_path, size=min_size)
    else:
        font = ImageFont.load_default()

    lines = _wrap_to_width(draw, text, font, box_w)
    try:
        ascent, descent = font.getmetrics()  # type: ignore
        line_h = ascent + descent + int(min_size * 0.15)
    except Exception:
        line_h = int(min_size * 1.25)

    return font, lines, line_h


def _apply_glossary_force(text_en: str, glossary: Dict[str, str]) -> str:
    """
    Force glossary replacements after model translation.
    This helps enforce things like OURO BATIDO -> BRUSH GOLD even if the model misses it.
    """
    if apply_glossary_hard:
        try:
            return apply_glossary_hard(text_en, glossary)
        except Exception:
            pass

    out = text_en
    for src, tgt in (glossary or {}).items():
        if not src or not tgt:
            continue
        out = out.replace(src, tgt)
        out = out.replace(src.upper(), tgt)
        out = out.replace(src.lower(), tgt)
    return out


def _guess_mime(filename: str) -> str:
    f = filename.lower()
    if f.endswith(".png"):
        return "image/png"
    if f.endswith(".webp"):
        return "image/webp"
    return "image/jpeg"


# ----------------------------
# Core
# ----------------------------
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
    """
    Returns PNG bytes of translated image.
    Behavior: ERASE original text areas, then OVERWRITE with translated English in the SAME positions.
    """

    def progress(p: float, msg: str):
        if progress_cb:
            try:
                progress_cb(float(p), msg)
            except Exception:
                pass

    # Load image
    img = Image.open(io.BytesIO(image_bytes)).convert("RGBA")
    orig_img = img.copy()  # keep pristine for background sampling
    W, H = img.size

    progress(0.05, f"{filename}: detecting text regions…")

    client = OpenAI(api_key=api_key)

    glossary_lines = "\n".join([f"- {k} => {v}" for k, v in (glossary or {}).items() if k and v])
    if not glossary_lines.strip():
        glossary_lines = "(none)"

    prompt = f"""
You are translating Brazilian Portuguese handbag development sheets to English for Chinese factories.

TASK:
1) Identify ALL Portuguese text in the image that is relevant to the tech pack (titles, specs, bullets, measurements, notes, materials, components).
2) For each text block, output:
   - text_pt (exact original, preserve numbers/units/punctuation)
   - text_en (English translation for handbag factories; use correct handbag/material/components terminology)
   - bbox (x0,y0,x1,y1) normalized to image size (0..1), tight around the text
   - style: "bold" if heading/bold, else "normal"
   - align: left/center/right if clearly aligned

IMPORTANT RULES:
- Keep all numbers, currency, dimensions, and units exactly (cm, mm, USD, etc).
- Prefer handbag terminology: lining, edge paint, piping, binding, strap, handle, zipper, puller, magnet snap, rivet, stitching, reinforcement, webbing, canvas, synthetic, PU, PVC.
- Bounding boxes must fully cover glyphs, including accents and anti-aliased edges. Prefer slightly larger boxes over cropping letters.
- Apply glossary EXACTLY when the Portuguese source term appears:
{glossary_lines}

- If a block is already English, keep it unchanged.
- Output JSON ONLY, matching the required schema.

EXTRA INSTRUCTIONS:
{extra_instructions}
""".strip()

    mime = _guess_mime(filename)
    data_url = _to_data_url(image_bytes, mime)

    # Vision + structured output
    result: ImageTranslationResult
    try:
        response = client.responses.parse(
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
        result = response.output_parsed  # type: ignore
        if result is None:
            raise ValueError("No parsed output returned.")
    except Exception:
        # Fallback: request JSON and parse manually
        response = client.responses.create(
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

        raw = getattr(response, "output_text", None)
        if not raw:
            # Try to extract from response.output if needed
            raw = str(response)
        result = ImageTranslationResult(**json.loads(raw))

    progress(0.35, f"{filename}: erasing + overwriting text…")

    draw = ImageDraw.Draw(img)
    regular_font_path, bold_font_path = _pick_fonts()

    # --- 1) ERASE PASS (remove Portuguese everywhere first) ---
    ERASE_PAD = 6  # covers anti-aliasing / small bbox inaccuracies

    px_blocks: List[Tuple[TextBlock, Tuple[int, int, int, int], Tuple[int, int, int, int]]] = []
    # store: (block, px_box, bg_fill)

    for block in result.blocks:
        x0 = int(block.bbox.x0 * W)
        y0 = int(block.bbox.y0 * H)
        x1 = int(block.bbox.x1 * W)
        y1 = int(block.bbox.y1 * H)

        x0 = max(0, x0 - ERASE_PAD)
        y0 = max(0, y0 - ERASE_PAD)
        x1 = min(W - 1, x1 + ERASE_PAD)
        y1 = min(H - 1, y1 + ERASE_PAD)

        if x1 <= x0 or y1 <= y0:
            continue

        px_box = (x0, y0, x1, y1)
        bg = _sample_bg_color(orig_img, px_box, pad=3)  # sample from original image
        px_blocks.append((block, px_box, bg))

    # Erase all regions first
    for _, px_box, bg in px_blocks:
        draw.rectangle(px_box, fill=bg)

    # --- 2) WRITE PASS (overwrite with English in the same spot) ---
    for i, (block, px_box, bg) in enumerate(px_blocks, start=1):
        x0, y0, x1, y1 = px_box
        box_w = max(1, x1 - x0)
        box_h = max(1, y1 - y0)

        # Choose readable text color for the sampled background
        fg = (0, 0, 0, 255)
        if _luminance(bg[:3]) < 0.45:
            fg = (255, 255, 255, 255)

        text_en = (block.text_en or "").strip()
        text_en = _apply_glossary_force(text_en, glossary)

        font_path = bold_font_path if block.style == "bold" else regular_font_path
        max_size = int(box_h * 0.90)

        font, lines, line_h = _fit_text_in_box(
            draw=draw,
            text=text_en,
            box_w=box_w,
            box_h=box_h,
            font_path=font_path,
            max_size=max_size,
            min_size=10,
        )

        y = y0
        for line in lines:
            try:
                line_w = int(draw.textlength(line, font=font))
            except Exception:
                line_w = int(getattr(font, "getlength", lambda s: len(s) * 6)(line))

            if block.align == "center":
                tx = x0 + max(0, (box_w - line_w) // 2)
            elif block.align == "right":
                tx = x0 + max(0, box_w - line_w)
            else:
                tx = x0

            draw.text((tx, y), line, fill=fg, font=font)
            y += line_h
            if y > y1:
                break

        progress(0.35 + 0.60 * (i / max(1, len(px_blocks))), f"{filename}: blocks ({i}/{len(px_blocks)})")

    # Export as PNG
    out = io.BytesIO()
    img.save(out, format="PNG")
    progress(1.0, f"{filename}: done")
    return out.getvalue()
