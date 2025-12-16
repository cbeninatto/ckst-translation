from __future__ import annotations

import re
from typing import Dict, NamedTuple


class ProtectedText(NamedTuple):
    text: str
    placeholder_to_original: Dict[str, str]


# Protect things that must NEVER change (SKUs, codes, sizes, % etc.)
_PROTECT_PATTERNS = [
    r"\bC\d{5}\s?\d{4}\s?\d{4}\s?[A-Z]?\b",                   # ex: C40008 0003 0002 U
    r"\b[A-Z]{2,6}\s?\d{2,6}(?:[\s\-]?\d{2,6}){1,4}\b",       # generic code blocks
    r"\b\d+(?:[.,]\d+)?\s?(?:mm|cm|m|kg|g|%|pcs|pc|un)\b",    # units
    r"\bPANTONE\s*[A-Z0-9\- ]+\b",                            # pantone
    r"\b#[0-9A-Fa-f]{6}\b",                                   # hex color
    r"\b\d+(?:[.,]\d+)?\b",                                   # numbers (last on purpose)
]


def protect_text(text: str) -> ProtectedText:
    placeholder_to_original: Dict[str, str] = {}
    out = text

    # Protect longer matches first to avoid overlaps
    matches = []
    for pat in _PROTECT_PATTERNS:
        for m in re.finditer(pat, out):
            matches.append((m.start(), m.end(), m.group(0)))
    matches.sort(key=lambda x: (x[0], -(x[1] - x[0])))

    # Apply replacements left->right without double replacing
    used = [False] * (len(out) + 1)
    kept = []
    for s, e, val in matches:
        if any(used[s:e]):
            continue
        for i in range(s, e):
            used[i] = True
        kept.append((s, e, val))

    # Replace from right to left so offsets don’t shift
    for idx, (s, e, val) in enumerate(reversed(kept)):
        ph = f"<KEEP_{idx}>"
        placeholder_to_original[ph] = val
        out = out[:s] + ph + out[e:]

    return ProtectedText(out, placeholder_to_original)


def restore_protected(text: str, placeholder_to_original: Dict[str, str]) -> str:
    out = text
    for ph, original in placeholder_to_original.items():
        out = out.replace(ph, original)
    return out


def parse_glossary_lines(raw: str) -> Dict[str, str]:
    """
    Each line: portuguese=english
    """
    out: Dict[str, str] = {}
    for line in (raw or "").splitlines():
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        if "=" not in line:
            continue
        k, v = line.split("=", 1)
        k = k.strip()
        v = v.strip()
        if k and v:
            out[k] = v
    return out
