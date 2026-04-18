from __future__ import annotations

import json
import os
from typing import Any

import google.generativeai as genai

from agent.prompts import SYSTEM_PROMPT, build_user_message


_DEFAULT_MODEL = "gpt-4o"
_GOOGLE_DEFAULT_MODEL = "gemini-1.5-flash-latest"

_REQUIRED_KEYS = {"Datasheet", "EBOM", "SRD", "CDD"}


class DataExtractor:
    def __init__(
        self,
        api_key: str | None = None,
        model: str | None = None,
        provider: str | None = None,
        base_url: str | None = None,
    ) -> None:

        if provider:
            self._provider = provider.lower()
        else:
            self._provider = (os.getenv("LLM_PROVIDER") or "google").lower()

        if self._provider == "google":
            self._api_key = api_key or os.getenv("GOOGLE_API_KEY")
            if not self._api_key:
                raise ValueError("GOOGLE_API_KEY not set")

            genai.configure(api_key=self._api_key)
            self._model = model or os.getenv("GOOGLE_MODEL", _GOOGLE_DEFAULT_MODEL)
            self._provider = "google_native"

        else:
            raise ValueError(f"Unsupported provider: {self._provider}")

    def extract(
        self,
        machine_name: str,
        task_type: str | None = None,
        ebom_reference: list[dict] | None = None,
    ) -> dict[str, Any]:

        user_message = build_user_message(machine_name, task_type, ebom_reference)
        raw = self._call_llm(user_message)
        return self._parse_and_validate(raw)

    def _call_llm(self, user_message: str) -> str:

        model = genai.GenerativeModel(self._model)

        prompt = f"""
{SYSTEM_PROMPT}

{user_message}

STRICT RULES:
- Output MUST be valid JSON
- Do NOT include explanations
- Do NOT include markdown
- Do NOT include ```json
- Only return raw JSON
"""

        response = model.generate_content(prompt)

        if hasattr(response, "text"):
            return response.text

        raise ValueError("No response from Gemini")

    @staticmethod
    def _parse_and_validate(raw: str) -> dict[str, Any]:

        text = raw.strip()

        if text.startswith("```"):
            lines = text.splitlines()[1:-1]
            text = "\n".join(lines)

        try:
            data = json.loads(text)
        except json.JSONDecodeError as e:
            raise ValueError(f"Invalid JSON from LLM:\n{text}") from e

        missing = _REQUIRED_KEYS - data.keys()
        if missing:
            raise ValueError(f"Missing keys: {missing}")

        return data
