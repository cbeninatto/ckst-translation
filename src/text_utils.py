import re
from typing import Dict, List, Tuple


def parse_glossary_lines(text: str) -> Dict[str, str]:
    """
    Parses lines like:
      pt => en
    Returns dict {pt_lower: en}
    """
    out: Dict[str, str] = {}
    if not text:
        return out

    for raw in text.splitlines():
        line = raw.strip()
        if not line or line.startswith("#"):
            continue
        if "=>" not in line:
            continue
        left, right = line.split("=>", 1)
        pt = left.strip()
        en = right.strip()
        if pt and en:
            out[pt.lower()] = en
    return out


def apply_glossary_hard(english_text: str, glossary: Dict[str, str]) -> str:
    """
    If any PT terms leaked into output, replace them hard with EN.
    Case-insensitive.
    """
    if not glossary or not english_text:
        return english_text

    out = english_text
    # Replace longer keys first to avoid partial collisions
    for pt in sorted(glossary.keys(), key=len, reverse=True):
        en = glossary[pt]
        # whole-word-ish replacement (but allow accents/spaces)
        pattern = re.compile(re.escape(pt), re.IGNORECASE)
        out = pattern.sub(en, out)

    return out


_KEEP_PATTERNS = [
    # SKUs / product codes like C40008 0003 0001, C400080003XX, etc.
    r"\b[A-Z]{1,6}\d{3,}(?:\s?\d{2,}){0,6}\b",
    r"\b[A-Z]\d{4,}[A-Z0-9]{0,}\b",
    # dimensions like 12x34x56, 12 x 34 x 56, with units
    r"\b\d+(?:[.,]\d+)?\s*(?:x|×)\s*\d+(?:[.,]\d+)?(?:\s*(?:x|×)\s*\d+(?:[.,]\d+)?)?\s*(?:mm|cm|m)\b",
    # single measurements with units
    r"\b\d+(?:[.,]\d+)?\s*(?:mm|cm|m|g|kg)\b",
    # percentages
    r"\b\d+(?:[.,]\d+)?\s*%\b",
    # currency (basic)
    r"\bR\$\s?\d+(?:[.,]\d+)?\b",
    r"\bUS\$\s?\d+(?:[.,]\d+)?\b",
    # dates
    r"\b\d{1,2}[/-]\d{1,2}[/-]\d{2,4}\b",
    # anything already in braces/brackets
    r"\{[^}]+\}",
    r"\[[^\]]+\]",
]


def protect_text(text: str) -> Tuple[str, List[str]]:
    """
    Replaces protected tokens with placeholders __KEEP0__, __KEEP1__...
    Returns (protected_text, keep_list)
    """
    if not text:
        return text, []

    keep: List[str] = []
    protected = text

    combined = re.compile("|".join(f"({p})" for p in _KEEP_PATTERNS), re.IGNORECASE)

    def _repl(m: re.Match) -> str:
        token = m.group(0)
        idx = len(keep)
        keep.append(token)
        return f"__KEEP{idx}__"

    protected = combined.sub(_repl, protected)
    return protected, keep


def restore_protected(text: str, keep_list: List[str]) -> str:
    if not keep_list or not text:
        return text

    out = text
    for i, tok in enumerate(keep_list):
        out = out.replace(f"__KEEP{i}__", tok)
    return out
