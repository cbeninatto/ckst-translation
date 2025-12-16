from __future__ import annotations

import json
from dataclasses import dataclass
from typing import Dict, List, Optional

from openai import OpenAI
from tenacity import retry, stop_after_attempt, wait_exponential, retry_if_exception_type

from .text_utils import ProtectedText, protect_text, restore_protected


@dataclass
class TranslationItem:
    id: str
    text: str


class OpenAITranslator:
    def __init__(self, api_key: Optional[str], model: str, reasoning_effort: str = "low"):
        self.client = OpenAI(api_key=api_key) if api_key else OpenAI()
        self.model = model
        self.reasoning_effort = reasoning_effort  # none|minimal|low|medium|high|xhigh

    @staticmethod
    def _schema_batch():
        return {
            "type": "object",
            "properties": {
                "translations": {
                    "type": "array",
                    "items": {
                        "type": "object",
                        "properties": {
                            "id": {"type": "string"},
                            "translated": {"type": "string"},
                        },
                        "required": ["id", "translated"],
                        "additionalProperties": False,
                    },
                }
            },
            "required": ["translations"],
            "additionalProperties": False,
        }

    @staticmethod
    def _schema_single():
        return {
            "type": "object",
            "properties": {"translated": {"type": "string"}},
            "required": ["translated"],
            "additionalProperties": False,
        }

    def _system_prompt(
        self,
        source_lang: str,
        target_lang: str,
        glossary: Dict[str, str],
        extra_instructions: str,
    ) -> str:
        system = (
            "You are a professional technical translator specialized in HANDBAGS / SOFTGOODS development packs.\n"
            f"Translate from {source_lang} to {target_lang}.\n\n"
            "Critical rules:\n"
            "- Keep all placeholders like <KEEP_0> exactly unchanged.\n"
            "- Do NOT change numbers, units, dimensions, tolerances, weights, percentages.\n"
            "- Do NOT change SKUs, color codes, barcodes, internal references, supplier codes.\n"
            "- Use handbag manufacturing terminology (materials, linings, padding, reinforcement, hardware, trims).\n"
            "- Prefer factory-friendly English: short, clear sentences; no idioms; actionable wording.\n"
            "- Preserve line breaks when they separate specs/components.\n"
            "- If a Portuguese term is ambiguous, choose the most common meaning in handbags.\n"
            "- If the text lists components, keep the same component order.\n"
        )

        if glossary:
            system += (
                "\nPreferred glossary (must be respected):\n"
                + "\n".join([f"- {k} => {v}" for k, v in glossary.items()])
                + "\n"
            )

        if extra_instructions.strip():
            system += "\nAdditional instructions:\n" + extra_instructions.strip() + "\n"

        return system

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
        system = self._system_prompt(source_lang, target_lang, glossary, extra_instructions)

        resp = self.client.responses.create(
            model=self.model,
            input=[
                {"role": "system", "content": system},
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
            reasoning={"effort": self.reasoning_effort},
            temperature=0,
            store=False,
        )

        data = json.loads(resp.output_text)
        translated = data["translated"]
        translated = restore_protected(translated, pt.placeholder_to_original)
        return translated

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
        protected_map: Dict[str, ProtectedText] = {}
        payload_items = []
        for it in items:
            pt = protect_text(it.text)
            protected_map[it.id] = pt
            payload_items.append({"id": it.id, "text": pt.text})

        system = self._system_prompt(source_lang, target_lang, glossary, extra_instructions)

        user_obj = {"items": payload_items}
        resp = self.client.responses.create(
            model=self.model,
            input=[
                {"role": "system", "content": system},
                {"role": "user", "content": json.dumps(user_obj, ensure_ascii=False)},
            ],
            text={
                "format": {
                    "type": "json_schema",
                    "name": "translation_batch",
                    "schema": self._schema_batch(),
                    "strict": True,
                }
            },
            reasoning={"effort": self.reasoning_effort},
            temperature=0,
            store=False,
        )

        data = json.loads(resp.output_text)
        out: Dict[str, str] = {}
        for row in data["translations"]:
            tid = row["id"]
            translated = row["translated"]
            translated = restore_protected(translated, protected_map[tid].placeholder_to_original)
            out[tid] = translated
        return out
