from __future__ import annotations

import json
from typing import Dict, List, NamedTuple, Optional

from openai import OpenAI
from tenacity import retry, stop_after_attempt, wait_exponential, retry_if_exception_type

from .text_utils import protect_text, restore_protected


class TranslationItem(NamedTuple):
    id: str
    text: str


def chunk_items(items: List[TranslationItem], max_chars: int = 18000, max_items: int = 60) -> List[List[TranslationItem]]:
    batches: List[List[TranslationItem]] = []
    cur: List[TranslationItem] = []
    cur_chars = 0

    for it in items:
        tlen = len(it.text or "")
        if cur and (len(cur) >= max_items or (cur_chars + tlen) > max_chars):
            batches.append(cur)
            cur = []
            cur_chars = 0
        cur.append(it)
        cur_chars += tlen

    if cur:
        batches.append(cur)

    return batches


class OpenAITranslator:
    def __init__(self, api_key: Optional[str], model: str, reasoning_effort: str = "low"):
        self.client = OpenAI(api_key=api_key) if api_key else OpenAI()
        self.model = model
        self.reasoning_effort = reasoning_effort

    def _supports_reasoning(self) -> bool:
        return self.model.startswith("gpt-5") or self.model.startswith("o")

    @staticmethod
    def _schema_single():
        return {
            "type": "object",
            "properties": {"translated": {"type": "string"}},
            "required": ["translated"],
            "additionalProperties": False,
        }

    @staticmethod
    def _schema_batch():
        return {
            "type": "object",
            "properties": {
                "translations": {
                    "type": "array",
                    "items": {
                        "type": "object",
                        "properties": {"id": {"type": "string"}, "translated": {"type": "string"}},
                        "required": ["id", "translated"],
                        "additionalProperties": False,
                    },
                }
            },
            "required": ["translations"],
            "additionalProperties": False,
        }

    def _system_prompt(self, source_lang: str, target_lang: str, glossary: Dict[str, str], extra: str) -> str:
        s = (
            "You are a professional technical translator specialized in HANDBAGS / SOFTGOODS tech packs.\n"
            f"Translate from {source_lang} to {target_lang}.\n\n"
            "Rules:\n"
            "- Keep placeholders like <KEEP_0> unchanged.\n"
            "- Do NOT change numbers, units, dimensions, tolerances, weights, percentages.\n"
            "- Do NOT change SKUs, internal codes, color codes, barcodes.\n"
            "- Use handbag manufacturing terminology (lining, reinforcement, piping, hardware, strap, top handle).\n"
            "- Write factory-friendly English: short, clear, actionable; no idioms.\n"
            "- Preserve bullet structure and line breaks.\n"
        )
        if glossary:
            s += "\nPreferred glossary (must be respected):\n" + "\n".join([f"- {k} => {v}" for k, v in glossary.items()]) + "\n"
        if extra.strip():
            s += "\nExtra instructions:\n" + extra.strip() + "\n"
        return s

    def _maybe_reasoning_kwargs(self) -> Dict:
        if self._supports_reasoning() and self.reasoning_effort and self.reasoning_effort != "none":
            return {"reasoning": {"effort": self.reasoning_effort}}
        return {}

    @retry(
        reraise=True,
        stop=stop_after_attempt(4),
        wait=wait_exponential(multiplier=1, min=1, max=12),
        retry=retry_if_exception_type(Exception),
    )
    def translate_text(
        self,
        text: str,
        source_lang: str,
        target_lang: str,
        glossary: Dict[str, str],
        extra_instructions: str = "",
    ) -> str:
        pt = protect_text(text)
        sys = self._system_prompt(source_lang, target_lang, glossary, extra_instructions)

        kwargs = dict(
            model=self.model,
            input=[
                {"role": "system", "content": sys},
                {"role": "user", "content": pt.text},
            ],
            text={
                "format": {
                    "type": "json_schema",
                    "name": "translation_single",
                    "schema": self._schema_single(),
                    "strict": True,
                }
            },
            store=False,
        )
        kwargs.update(self._maybe_reasoning_kwargs())

        resp = self.client.responses.create(**kwargs)
        data = json.loads(resp.output_text)
        return restore_protected(data["translated"], pt.placeholder_to_original)

    @retry(
        reraise=True,
        stop=stop_after_attempt(4),
        wait=wait_exponential(multiplier=1, min=1, max=12),
        retry=retry_if_exception_type(Exception),
    )
    def translate_batch(
        self,
        items: List[TranslationItem],
        source_lang: str,
        target_lang: str,
        glossary: Dict[str, str],
        extra_instructions: str = "",
    ) -> Dict[str, str]:
        protected = {}
        payload = []
        for it in items:
            pt = protect_text(it.text)
            protected[it.id] = pt
            payload.append({"id": it.id, "text": pt.text})

        sys = self._system_prompt(source_lang, target_lang, glossary, extra_instructions)

        kwargs = dict(
            model=self.model,
            input=[
                {"role": "system", "content": sys},
                {"role": "user", "content": json.dumps({"items": payload}, ensure_ascii=False)},
            ],
            text={
                "format": {
                    "type": "json_schema",
                    "name": "translation_batch",
                    "schema": self._schema_batch(),
                    "strict": True,
                }
            },
            store=False,
        )
        kwargs.update(self._maybe_reasoning_kwargs())

        resp = self.client.responses.create(**kwargs)
        data = json.loads(resp.output_text)

        out: Dict[str, str] = {}
        for row in data["translations"]:
            tid = row["id"]
            out[tid] = restore_protected(row["translated"], protected[tid].placeholder_to_original)
        return out
