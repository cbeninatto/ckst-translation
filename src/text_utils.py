from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Dict, List, Tuple


# Things we almost never want translated in handbag dev files:
# - SKUs / internal codes (C40008 0003 0002 etc.)
# - dimensions (10 cm, 10cm, 10 x 20 cm, 10×20mm)
# - currencies, percentages
# - emails, URLs
PROTECT_PATTERNS: List[Tuple[str, str]] = [
    ("SKU_CODE", r"\b[A-Z]\d{4,}(?:\s?\d{2,}){0,6}\b"),            # broad: C40008 0003 0002
    ("ALNUM_CODE", r"\b[A-Z0-9]{6,}\b"),                          # long alnum tokens
    ("MEASURE", r"\b\d+(?:[.,]\d+)?\s?(?:mm|cm|m|in|inch|kg|g)\b"),
    ("DIMENSION", r"\b\d+(?:[.,]\d+)?\s?[x×]\s?\d+(?:[.,]\d+)?(?:\s?(?:mm|cm|m|in|inch))?\b"),
    ("PERCENT", r"\b\d+(?:[.,]\d+)?\s?%\b"),
    ("MONEY", r"(?:R\$|\$|€)\s?\d+(?:[.,]\d+)*(?:[.,]\d+)?"),
    ("EMAIL", r"\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b"),
    ("URL", r"\bhttps?://[^\s]+"),
]

PLACEHOLDER_FMT = "<KEEP_{n}>"


@dataclass
class ProtectedText:
    text: str
    placeholder_to_original: Dict[str, str]


def apply_glossary_hard(text: str, glossary: Dict[str, str]) -> str:
    """
    Hard-enforce a glossary with a conservative regex word boundary match.
    This is optional; many users prefer only prompt-based glossary.
    """
    if not glossary:
        return text

    out = text
    # Replace longer keys first to avoid partial overlaps
    for k in sorted(glossary.keys(), key=len, reverse=True):
        v = glossary[k]
        if not k.strip():
            continue
        # case-insensitive whole-word where possible, but allow terms with spaces
        pattern = r"\b" + re.escape(k) + r"\b" if re.match(r"^[\w\-]+$", k) else re.escape(k)
        out = re.sub(pattern, v, out, flags=re.IGNORECASE)
    return out


def protect_text(text: str) -> ProtectedText:
    """
    Replaces protected spans with placeholders <KEEP_0>, <KEEP_1>, ...
    so the model doesn't translate/alter them.
    """
    if not text:
        return ProtectedText(text="", placeholder_to_original={})

    placeholder_to_original: Dict[str, str] = {}
    matches: List[Tuple[int, int, str]] = []

    # Collect matches across patterns
    for _, pattern in PROTECT_PATTERNS:
        for m in re.finditer(pattern, text, flags=re.IGNORECASE):
            s, e = m.span()
            # Skip very short "codes" accidentally caught
            if e - s < 3:
                continue
            matches.append((s, e, text[s:e]))

    # De-overlap: keep longest spans first
    matches.sort(key=lambda x: (x[0], -(x[1] - x[0])))
    non_overlapping: List[Tuple[int, int, str]] = []
    last_end = -1
    for s, e, val in matches:
        if s >= last_end:
            non_overlapping.append((s, e, val))
            last_end = e

    if not non_overlapping:
        return ProtectedText(text=text, placeholder_to_original={})

    # Replace from end to start so indices remain valid
    out = text
    for i, (s, e, val) in enumerate(reversed(non_overlapping)):
        ph = PLACEHOLDER_FMT.format(n=i)
        placeholder_to_original[ph] = val
        out = out[:s] + ph + out[e:]

    return ProtectedText(text=out, placeholder_to_original=placeholder_to_original)


def restore_protected(text: str, placeholder_to_original: Dict[str, str]) -> str:
    if not placeholder_to_original:
        return text
    out = text
    # Replace placeholders back
    for ph, original in placeholder_to_original.items():
        out = out.replace(ph, original)
    return out


def parse_glossary_lines(raw: str) -> Dict[str, str]:
    """
    Accepts lines like:
      couro=leather
      forro=lining
    """
    glossary: Dict[str, str] = {}
    if not raw:
        return glossary

    for line in raw.splitlines():
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        if "=" not in line:
            continue
        k, v = line.split("=", 1)
        k, v = k.strip(), v.strip()
        if k and v:
            glossary[k] = v
    return glossary
