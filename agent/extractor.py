from __future__ import annotations

import json
import os
import re
from typing import Any

import google.generativeai as genai

from agent.prompts import SYSTEM_PROMPT, build_user_message


_GOOGLE_DEFAULT_MODEL = "gemini-1.5-flash"
_REQUIRED_KEYS = {"Datasheet", "EBOM", "SRD", "CDD"}


class DataExtractor:
    def __init__(
        self,
        api_key: str | None = None,
        model: str | None = None,
        provider: str | None = None,
        base_url: str | None = None,
    ) -> None:

        self._provider = (provider or os.getenv("LLM_PROVIDER") or "google").lower()

        if self._provider != "google":
            raise ValueError("Only Google provider supported")

        self._api_key = api_key or os.getenv("GOOGLE_API_KEY")
        if not self._api_key:
            raise ValueError("GOOGLE_API_KEY not set")

        genai.configure(api_key=self._api_key)
        self._model = model or os.getenv("GOOGLE_MODEL", _GOOGLE_DEFAULT_MODEL)

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
- Return ONLY valid JSON
- No explanation
- No markdown
- No ```json
"""

        response = model.generate_content(prompt)

        if hasattr(response, "text") and response.text:
            print("RAW OUTPUT:", response.text)
            return response.text

        raise ValueError("Empty response from Gemini")

    @staticmethod
    def _extract_json(text: str) -> str:
        """Extract JSON even if wrapped in text"""
        match = re.search(r"\{.*\}", text, re.DOTALL)
        if match:
            return match.group(0)
        return text

    def _parse_and_validate(self, raw: str) -> dict[str, Any]:

        text = raw.strip()

        # remove markdown
        if text.startswith("```"):
            text = text.replace("```json", "").replace("```", "").strip()

        # extract JSON safely
        text = self._extract_json(text)

        try:
            data = json.loads(text)
        except Exception as e:
            print("FAILED TEXT:", text)
            raise ValueError("Invalid JSON from LLM") from e

        missing = _REQUIRED_KEYS - data.keys()
        if missing:
            raise ValueError(f"Missing keys: {missing}")

        return data
