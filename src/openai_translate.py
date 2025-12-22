import json
from typing import Dict, List, Optional, Sequence

from openai import OpenAI

from .text_utils import apply_glossary_hard, protect_text, restore_protected


class TranslationItem(tuple):
    """
    Lightweight item to avoid dataclass edge-cases on some runtimes.
    item[0]=id, item[1]=text
    """
    __slots__ = ()

    def __new__(cls, id: str, text: str):
        return tuple.__new__(cls, (id, text))

    @property
    def id(self) -> str:
        return self[0]

    @property
    def text(self) -> str:
        return self[1]


def chunk_items(items: Sequence[TranslationItem], max_items: int = 60, max_chars: int = 9000) -> List[List[TranslationItem]]:
    chunks: List[List[TranslationItem]] = []
    cur: List[TranslationItem] = []
    cur_chars = 0
    for it in items:
        t = it.text or ""
        if cur and (len(cur) >= max_items or cur_chars + len(t) > max_chars):
            chunks.append(cur)
            cur = []
            cur_chars = 0
        cur.append(it)
        cur_chars += len(t)
    if cur:
        chunks.append(cur)
    return chunks


class OpenAITranslator:
    def __init__(self, api_key: str, model: str, glossary: Optional[Dict[str, str]] = None):
        self.client = OpenAI(api_key=api_key)
        self.model = model
        self.glossary = glossary or {}

    def _system_prompt(self) -> str:
        glossary_lines = "\n".join([f"- {k} -> {v}" for k, v in self.glossary.items()])
        return (
            "You are a professional technical translator for handbag and accessories tech packs.\n"
            "Translate Brazilian Portuguese to English.\n"
            "Rules:\n"
            "- Keep proper nouns, supplier names, and SKUs/codes unchanged.\n"
            "- Use handbag-specific terminology (materials, components, hardware, stitching, lining, etc.).\n"
            "- Preserve units, measurements, and numbers exactly.\n"
            "- Keep the meaning concise and production-ready for Chinese factories.\n"
            "- If a term matches the glossary, use the glossary translation.\n"
            "\n"
            "Glossary:\n"
            f"{glossary_lines if glossary_lines else '(none)'}\n"
        )

    def translate_batch(self, items: Sequence[TranslationItem]) -> Dict[str, str]:
        if not items:
            return {}

        protected = []
        for it in items:
            p = protect_text(it.text)
            protected.append((it.id, p))

        payload_items = [{"id": _id, "text": p.protected_text} for _id, p in protected]

        schema = {
            "type": "object",
            "properties": {
                "translations": {
                    "type": "array",
                    "items": {
                        "type": "object",
                        "properties": {"id": {"type": "string"}, "text": {"type": "string"}},
                        "required": ["id", "text"],
                        "additionalProperties": False,
                    },
                }
            },
            "required": ["translations"],
            "additionalProperties": False,
        }

        # IMPORTANT: no temperature here (avoids the 400 you got)
        resp = self.client.responses.create(
            model=self.model,
            input=[
                {"role": "system", "content": self._system_prompt()},
                {"role": "user", "content": "Translate the following items:\n" + json.dumps(payload_items, ensure_ascii=False)},
            ],
            text={"format": {"type": "json_schema", "name": "translation_batch", "schema": schema}},
        )

        data = json.loads(resp.output_text)
        id_to_translated: Dict[str, str] = {t["id"]: t["text"] for t in data.get("translations", [])}

        result: Dict[str, str] = {}
        for (item_id, prot) in protected:
            translated = id_to_translated.get(item_id, prot.protected_text)
            translated = restore_protected(translated, prot)
            translated = apply_glossary_hard(translated, self.glossary)
            result[item_id] = translated

        return result
