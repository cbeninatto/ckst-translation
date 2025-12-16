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
    def __init__(self, api_key: Optional[str], model: str):
        # If api_key is None, OpenAI() will read OPENAI_API_KEY from env automatically
        self.client = OpenAI(api_key=api_key) if api_key else OpenAI()
        self.model = model

    @staticmethod
    def _schema():
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
        """
        Returns dict id -> translated_text
        Uses Structured Outputs (JSON schema) via Responses API.
        """
        # Protect codes/measurements so the model doesn't "translate" them.
        protected_map: Dict[str, ProtectedText] = {}
        payload_items = []
        for it in items:
            pt = protect_text(it.text)
            protected_map[it.id] = pt
            payload_items.append({"id": it.id, "text": pt.text})

        system = (
            "You are a professional technical translator for handbag development files.\n"
            f"Translate from {source_lang} to {target_lang}.\n"
            "Rules:\n"
            "- Keep all placeholders like <KEEP_0> exactly unchanged.\n"
            "- Do NOT change numbers, units, dimensions, percentages, SKUs, barcodes, or codes.\n"
            "- Preserve line breaks where they help readability.\n"
            "- Keep brand names as-is.\n"
            "- Use clear, factory-friendly English.\n"
            "- If a Portuguese term is ambiguous, choose the most common manufacturing meaning.\n"
        )

        if glossary:
            system += (
                "\nGlossary (must be respected as preferred translations):\n"
                + "\n".join([f"- {k} => {v}" for k, v in glossary.items()])
                + "\n"
            )

        if extra_instructions.strip():
            system += "\nAdditional instructions:\n" + extra_instructions.strip() + "\n"

        user_obj = {
            "source_language": source_lang,
            "target_language": target_lang,
            "items": payload_items,
        }

        resp = self.client.responses.create(
            model=self.model,
            input=[
                {"role": "system", "content": system},
                {"role": "user", "content": json.dumps(user_obj, ensure_ascii=False)},
            ],
            # Structured output:
            text={
                "format": {
                    "type": "json_schema",
                    "name": "translation_batch",
                    "schema": self._schema(),
                    "strict": True,
                }
            },
            temperature=0,
            # Reduce retention (helpful for confidential docs):
            store=False,
        )

        data = json.loads(resp.output_text)
        out: Dict[str, str] = {}
        for row in data["translations"]:
            tid = row["id"]
            translated = row["translated"]
            # Restore protected tokens
            translated = restore_protected(translated, protected_map[tid].placeholder_to_original)
            out[tid] = translated
        return out


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
