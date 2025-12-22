import re
from typing import Dict, List, NamedTuple, Tuple


class ProtectedText(NamedTuple):
    original_text: str
    protected_text: str
    placeholder_map: Dict[str, str]


def _make_placeholder(i: int) -> str:
    return f"<<KEEP_{i}>>"


_PROTECT_PATTERNS = [
    r"\bhttps?://\S+\b",
    r"\b[\w\.-]+@[\w\.-]+\.\w+\b",
    r"\b\d+([.,]\d+)?\s*%\b",
    r"\b\d+([.,]\d+)?\s*(mm|cm|m|kg|g|oz|lb|pcs|pc|un|und|u)\b",
    r"\b\d+([.,]\d+)?\s*[xX×]\s*\d+([.,]\d+)?(\s*[xX×]\s*\d+([.,]\d+)?)?\s*(mm|cm|m)?\b",
    r"\b[A-Z0-9]{6,}\b",  # long codes / SKUs / barcodes
]


def protect_text(text: str) -> ProtectedText:
    if text is None:
        text = ""
    s = str(text)

    placeholder_map: Dict[str, str] = {}
    idx = 0

    matches: List[Tuple[int, int, str]] = []
    for pat in _PROTECT_PATTERNS:
        for m in re.finditer(pat, s):
            matches.append((m.start(), m.end(), m.group(0)))

    for m in re.finditer(r"\bC\d{5}\s\d{4}\s\d{4}\b", s):
        matches.append((m.start(), m.end(), m.group(0)))

    matches.sort(key=lambda x: (x[0], -(x[1] - x[0])))

    filtered: List[Tuple[int, int, str]] = []
    last_end = -1
    for start, end, val in matches:
        if start >= last_end:
            filtered.append((start, end, val))
            last_end = end

    out = s
    for start, end, val in sorted(filtered, key=lambda x: x[0], reverse=True):
        ph = _make_placeholder(idx)
        idx += 1
        placeholder_map[ph] = val
        out = out[:start] + ph + out[end:]

    return ProtectedText(original_text=s, protected_text=out, placeholder_map=placeholder_map)


def restore_protected(text: str, prot: ProtectedText) -> str:
    out = text
    for ph, val in prot.placeholder_map.items():
        out = out.replace(ph, val)
    return out


def parse_glossary_lines(glossary_text: str) -> Dict[str, str]:
    glossary: Dict[str, str] = {}
    if not glossary_text:
        return glossary
    for raw in glossary_text.splitlines():
        line = raw.strip()
        if not line or line.startswith("#"):
            continue
        m = re.split(r"\s*(?:=>|->|=)\s*", line, maxsplit=1)
        if len(m) == 2:
            k, v = m[0].strip(), m[1].strip()
            if k and v:
                glossary[k] = v
    return glossary


def apply_glossary_hard(text: str, glossary: Dict[str, str]) -> str:
    if not glossary or not text:
        return text

    out = text
    for k in sorted(glossary.keys(), key=len, reverse=True):
        v = glossary[k]
        out = re.sub(rf"(?i)\b{re.escape(k)}\b", v, out)
    return out
