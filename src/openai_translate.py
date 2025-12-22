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


def chunk_items(
    items: List[TranslationItem],
    max_items: int = 800,
    max_chars: int = 90000,
) -> List[List[TranslationItem]]:
    """
    Large chunks (to satisfy "remove batching limits" as much as realistically possible),
    but still protects against API request limits.
    """
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
    """
    Try to extract a JSON object from model output.
    """
    if not text:
        return None
    text = text.strip()

    # direct json
    try:
        obj = json.loads(text)
        if isinstance(obj, dict):
            return obj
    except Exception:
        pass

    # try to find first {...} block
    m = re.search(r"\{.*\}", text, flags=re.DOTALL)
    if m:
        candidate = m.group(0)
        try:
            obj = json.loads(candidate)
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

        # Protect tokens per-item
        protected_payload = []
        keep_maps: Dict[str, List[str]] = {}

        for it in items:
            protected_text, keep_list = protect_text(it.text)
            keep_maps[it.id] = keep_list
            protected_payload.append({"id": it.id, "text": protected_text})

        # Build glossary string for prompt
        glossary_lines = ""
        if glossary:
            pairs = [f"- {k} => {v}" for k, v in glossary.items()]
            glossary_lines = "\n".join(pairs)

        system = (
            "You are a professional technical translator for HANDBAGS / SOFTGOODS tech packs.\n"
            "Translate from Brazilian Portuguese to English for Chinese factories.\n"
            "Use industry-standard handbag terminology (materials, components, hardware, stitching, lining, reinforcement).\n"
            "Keep all __KEEP#__ placeholders EXACTLY as-is.\n"
            "Do not change numbers, measurements, SKUs, codes.\n"
            "Output ONLY valid JSON object: {\"id\": \"translated text\", ...}.\n"
            "No extra keys, no markdown, no commentary."
        )

        user = (
            f"Source language: {source_lang}\n"
            f"Target language: {target_lang}\n\n"
            f"Glossary (must follow exactly when relevant):\n{glossary_lines or '(none)'}\n\n"
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
                # fallback to original
                raw = it.text
            # restore protected
            restored = restore_protected(raw, keep_maps.get(it.id, []))
            # hard glossary cleanup (if PT leaked)
            restored = apply_glossary_hard(restored, glossary)
            out[it.id] = restored

        return out

    def _call_model(self, system: str, user: str) -> str:
        """
        Prefer Responses API; fallback to Chat Completions if needed.
        Avoid parameters that some models reject (e.g., temperature).
        """
        # ---- Responses API ----
        try:
            resp = self.client.responses.create(
                model=self.model,
                input=[
                    {"role": "system", "content": system},
                    {"role": "user", "content": user},
                ],
                # reasoning_effort is ignored by some models; safe to try
                reasoning={"effort": self.reasoning_effort} if self.reasoning_effort != "none" else None,
            )
            # SDK provides output_text in many cases
            if hasattr(resp, "output_text") and resp.output_text:
                return resp.output_text
            # fallback: try to dig content
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

        # ---- Chat Completions fallback ----
        cc = self.client.chat.completions.create(
            model=self.model,
            messages=[
                {"role": "system", "content": system},
                {"role": "user", "content": user},
            ],
        )
        return (cc.choices[0].message.content or "").strip()
