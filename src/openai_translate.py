import json
import re
from typing import Any, Dict, List, Optional

from openai import OpenAI

from .text_utils import apply_glossary_hard, protect_text, restore_protected


class TranslationItem:
    __slots__ = ("id", "text")

    def __init__(self, item_id: str, text: str):
        self.id = item_id
        self.text = text


def chunk_items(items: List[TranslationItem], max_items: int = 800, max_chars: int = 90000) -> List[List[TranslationItem]]:
    out: List[List[TranslationItem]] = []
    cur: List[TranslationItem] = []
    cur_chars = 0
    for it in items:
        t = it.text or ""
        if cur and (len(cur) >= max_items or (cur_chars + len(t)) > max_chars):
            out.append(cur)
            cur = []
            cur_chars = 0
        cur.append(it)
        cur_chars += len(t)
    if cur:
        out.append(cur)
    return out


def _extract_json(text: str) -> Optional[Dict[str, Any]]:
    if not text:
        return None
    text = text.strip()
    try:
        obj = json.loads(text)
        if isinstance(obj, dict):
            return obj
    except Exception:
        pass
    m = re.search(r"\{.*\}", text, flags=re.DOTALL)
    if m:
        try:
            obj = json.loads(m.group(0))
            if isinstance(obj, dict):
                return obj
        except Exception:
            pass
    return None


class OpenAITranslator:
    def __init__(self, api_key: str, model: str, reasoning_effort: str = "medium"):
        self.client = OpenAI(api_key=api_key)
        self.model = model
        self.reasoning_effort = reasoning_effort

    def translate_batch(
        self,
        items: List[TranslationItem],
        source_lang: str = "pt-BR",
        target_lang: str = "en",
        glossary: Optional[Dict[str, str]] = None,
        extra_instructions: str = "",
    ) -> Dict[str, str]:
        if not items:
            return {}

        glossary = glossary or {}

        protected_payload = []
        keep_maps: Dict[str, List[str]] = {}

        for it in items:
            protected_text, keep_list = protect_text(it.text)
            keep_maps[it.id] = keep_list
            protected_payload.append({"id": it.id, "text": protected_text})

        glossary_lines = "\n".join([f"- {k} => {v}" for k, v in glossary.items()]) if glossary else "(none)"

        system = (
            "You are a professional technical translator for HANDBAGS / SOFTGOODS tech packs.\n"
            f"Translate from {source_lang} to {target_lang}.\n"
            "Use industry-standard handbag terminology (materials, components, hardware, stitching, lining, reinforcement).\n"
            "Keep all __KEEP#__ placeholders EXACTLY as-is.\n"
            "Do not change numbers, measurements, SKUs, codes.\n"
            "Output ONLY valid JSON object: {\"id\": \"translated text\", ...}.\n"
            "No extra keys, no markdown, no commentary.\n"
        )
        if target_lang.lower().startswith("pt"):
            system += "Write in Brazilian Portuguese (PT-BR).\n"

        user = (
            f"Glossary (must follow exactly when relevant):\n{glossary_lines}\n\n"
            f"Extra instructions:\n{extra_instructions or '(none)'}\n\n"
            f"Translate this list of items and return JSON mapping id->translated:\n"
            f"{json.dumps(protected_payload, ensure_ascii=False)}"
        )

        text_out = self._call_model(system, user)
        obj = _extract_json(text_out) or {}

        out: Dict[str, str] = {}
        for it in items:
            raw = obj.get(it.id, None)
            if not isinstance(raw, str) or not raw.strip():
                raw = it.text
            restored = restore_protected(raw, keep_maps.get(it.id, []))
            restored = apply_glossary_hard(restored, glossary)
            out[it.id] = restored

        return out

    # Optional compatibility helper (useful for PDF modules that call translate_texts)
    def translate_texts(
        self,
        texts: List[str],
        source_lang: str = "pt-BR",
        target_lang: str = "en",
        glossary: Optional[Dict[str, str]] = None,
        extra_instructions: str = "",
    ) -> List[str]:
        items = [TranslationItem(f"t{i}", t) for i, t in enumerate(texts)]
        mapping = {}
        for ch in chunk_items(items):
            mapping.update(self.translate_batch(ch, source_lang=source_lang, target_lang=target_lang, glossary=glossary, extra_instructions=extra_instructions))
        return [mapping[f"t{i}"] for i in range(len(texts))]

    def _call_model(self, system: str, user: str) -> str:
        try:
            resp = self.client.responses.create(
                model=self.model,
                input=[
                    {"role": "system", "content": system},
                    {"role": "user", "content": user},
                ],
                reasoning={"effort": self.reasoning_effort} if self.reasoning_effort != "none" else None,
            )
            if hasattr(resp, "output_text") and resp.output_text:
                return resp.output_text
            if hasattr(resp, "output") and resp.output:
                parts = []
                for o in resp.output:
                    if hasattr(o, "content") and o.content:
                        for c in o.content:
                            if hasattr(c, "text") and c.text:
                                parts.append(c.text)
                return "\n".join(parts).strip()
        except Exception:
            pass

        cc = self.client.chat.completions.create(
            model=self.model,
            messages=[
                {"role": "system", "content": system},
                {"role": "user", "content": user},
            ],
        )
        return (cc.choices[0].message.content or "").strip()
