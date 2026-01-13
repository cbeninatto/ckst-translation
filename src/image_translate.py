# src/image_translate.py
from __future__ import annotations

import base64
import io
import json
from typing import Callable, Dict, List, Optional, Literal

from PIL import Image, ImageDraw, ImageFont

from pydantic import BaseModel, Field
from openai import OpenAI

# If your project already has these utilities, we reuse them.
# Otherwise, this module works fine without them (we fall back gracefully).
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


def _pick_fonts():
    # Streamlit Cloud/Linux commonly has DejaVu installed; these paths work in most environments.
    regular = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
    bold = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"
    return regular, bold


def _luminance(rgb):
    r, g, b = rgb
    return (0.2126 * r + 0.7152 * g + 0.0722 * b) / 255.0


def _sample_bg_color(img_rgba: Image.Image, px_box: tuple[int, int, int, int], pad: int = 3):
    """Sample a background color around the bbox (edge pixels) and return an RGBA fill."""
    x0, y0, x1, y1 = px_box
    w, h = img_rgba.size
    x0c = max(0, x0 - pad)
    y0c = max(0, y0 - pad)
    x1c = min(w - 1, x1 + pad)
    y1c = min(h - 1, y1 + pad)

    # Sample a thin ring around the box
    pixels = img_rgba.load()
    samples = []

    # top/bottom edges
    for x in range(x0c, x1c):
        samples.append(pixels[x, y0c])
        samples.append(pixels[x, y1c])

    # left/right edges
    for y in range(y0c, y1c):
        samples.append(pixels[x0c, y])
        samples.append(pixels[x1c, y])

    if not samples:
        return (255, 255, 255, 255)

    # average
    r = sum(p[0] for p in samples) // len(samples)
    g = sum(p[1] for p in samples) // len(samples)
    b = sum(p[2] for p in samples) // len(samples)
    a = 255
    return (int(r), int(g), int(b), a)


def _wrap_to_width(draw: ImageDraw.ImageDraw, text: str, font: ImageFont.FreeTypeFont, max_w: int) -> List[str]:
    words = text.split()
    if not words:
        return [""]

    lines: List[str] = []
    cur = ""
    for w in words:
        test = (cur + " " + w).strip()
        if draw.textlength(test, font=font) <= max_w:
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
    font_path: str,
    max_size: int,
    min_size: int = 8,
) -> tuple[ImageFont.FreeTypeFont, List[str], int]:
    """
    Find a font size that fits text in the bbox, wrapping lines as needed.
    Returns (font, lines, line_h).
    """
    max_size = max(min_size, max_size)

    for size in range(max_size, min_size - 1, -1):
        font = ImageFont.truetype(font_path, size=size)
        lines = _wrap_to_width(draw, text, font, box_w)
        ascent, descent = font.getmetrics()
        line_h = ascent + descent + int(size * 0.15)
        total_h = line_h * len(lines)
        if total_h <= box_h:
            return font, lines, line_h

    font = ImageFont.truetype(font_path, size=min_size)
    lines = _wrap_to_width(draw, text, font, box_w)
    ascent, descent = font.getmetrics()
    line_h = ascent + descent + int(min_size * 0.15)
    return font, lines, line_h


def _apply_glossary_force(text_en: str, glossary: Dict[str, str]) -> str:
    # Force replace any glossary targets (best-effort).
    # If you have apply_glossary_hard in your project, use it.
    if apply_glossary_hard:
        try:
            return apply_glossary_hard(text_en, glossary)
        except Exception:
            pass

    # Fallback: simple replacements (case-insensitive-ish by checking variants)
    out = text_en
    for src, tgt in glossary.items():
        if not src or not tgt:
            continue
        out = out.replace(src, tgt)
        out = out.replace(src.upper(), tgt)
        out = out.replace(src.lower(), tgt)
    return out


# ----------------------------
# Core function
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
    Returns PNG bytes of the translated image (in-place overlay).
    """
    def progress(p: float, msg: str):
        if progress_cb:
            progress_cb(p, msg)

    # Load image
    img = Image.open(io.BytesIO(image_bytes)).convert("RGBA")
    W, H = img.size

    # Call OpenAI (vision + structured output)
    progress(0.05, f"{filename}: detecting text regions…")

    client = OpenAI(api_key=api_key)

    # Build glossary string for the prompt
    glossary_lines = "\n".join([f"- {k} => {v}" for k, v in glossary.items() if k and v])
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
   - style: "bold" if the original is bold/heading, else "normal"
   - align: left/center/right if clearly aligned

RULES:
- Keep all numbers, currency, dimensions, and units exactly (cm, mm, USD, etc).
- Use these handbag terms naturally (examples): lining, edge paint, piping, binding, strap, handle, zipper, puller, magnet snap, rivet, stitching, reinforcement, webbing, canvas, synthetic, PU, PVC.
- Apply glossary EXACTLY when the Portuguese source term appears:
{glossary_lines}

- If a block is already English, keep it unchanged.
- Output JSON ONLY, matching the required schema.

EXTRA INSTRUCTIONS (if any):
{extra_instructions}
""".strip()

    # Determine mime (best effort)
    mime = "image/jpeg"
    if filename.lower().endswith(".png"):
        mime = "image/png"
    elif filename.lower().endswith(".webp"):
        mime = "image/webp"

    data_url = _to_data_url(image_bytes, mime)

    # Prefer parse() if available, else fallback to create()+json.loads
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
        # Fallback: ask for json and parse manually
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
            # last resort: dig output
            raw = str(response)
        result = ImageTranslationResult(**json.loads(raw))

    progress(0.35, f"{filename}: rendering translated text…")

    # Render translations onto the same image
    draw = ImageDraw.Draw(img)
    regular_font_path, bold_font_path = _pick_fonts()

    for i, block in enumerate(result.blocks):
        # Convert normalized bbox to pixel bbox
        x0 = int(block.bbox.x0 * W)
        y0 = int(block.bbox.y0 * H)
        x1 = int(block.bbox.x1 * W)
        y1 = int(block.bbox.y1 * H)

        # Sanity clamp
        x0, y0 = max(0, x0), max(0, y0)
        x1, y1 = min(W - 1, x1), min(H - 1, y1)
        if x1 <= x0 or y1 <= y0:
            continue

        px_box = (x0, y0, x1, y1)

        # Fill background using sampled color so we "erase" Portuguese text cleanly
        bg = _sample_bg_color(img, px_box, pad=3)
        draw.rectangle(px_box, fill=bg)

        # Choose text color based on background luminance
        fg = (0, 0, 0, 255)
        if _luminance(bg[:3]) < 0.45:
            fg = (255, 255, 255, 255)

        text_en = block.text_en.strip()
        text_en = _apply_glossary_force(text_en, glossary)

        # Fit font
        box_w = max(1, x1 - x0)
        box_h = max(1, y1 - y0)
        font_path = bold_font_path if block.style == "bold" else regular_font_path
        max_size = int(box_h * 0.90)

        font, lines, line_h = _fit_text_in_box(
            draw=draw,
            text=text_en,
            box_w=box_w,
            box_h=box_h,
            font_path=font_path,
            max_size=max_size,
            min_size=8,
        )

        # Horizontal alignment
        y = y0
        for line in lines:
            line_w = int(draw.textlength(line, font=font))
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

        progress(0.35 + 0.60 * ((i + 1) / max(1, len(result.blocks))), f"{filename}: blocks ({i+1}/{len(result.blocks)})")

    # Export as PNG
    out = io.BytesIO()
    img.save(out, format="PNG")
    progress(1.0, f"{filename}: done")
    return out.getvalue()
