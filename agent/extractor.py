from __future__ import annotations

import json
import os
from typing import Any

import google.generativeai as genai

try:
    from openai import OpenAI
except ImportError:
    OpenAI = None  # type: ignore

from agent.prompts import SYSTEM_PROMPT, build_user_message


# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

_DEFAULT_MODEL = "gpt-4o"

_OLLAMA_DEFAULT_BASE_URL = "http://localhost:11434/v1"
_OLLAMA_DEFAULT_MODEL = "llama3.2"

_GOOGLE_DEFAULT_MODEL = "gemini-1.5-flash"

_REQUIRED_KEYS = {"Datasheet", "EBOM", "SRD", "CDD"}

_DATASHEET_KEYS = {
    "Author", "Item Name", "HEL", "System Description",
    "Dimensional Parameters", "Other Parameters",
    "Manufacturer", "Model", "Website", "Notes", "References",
}

_PARAM_ROW_KEYS = {"Parameter", "Unit", "Value", "Reference", "Notes"}

_LIST_SHEET_COLUMNS: dict[str, set[str]] = {
    "EBOM": {
        "HEL", "Responsible person", "Task", "Machine type",
        "Specific machine", "Product website", "Product phase", "Description",
        "Height (mm)", "Length (mm)", "Width (mm)", "Mass (kg)",
        "TRL", "SRL", "MRL",
    },
    "SRD": {"HEL", "No", "Requirement", "Requirement Type"},
    "CDD": {"HEL", "No", "Statement"},
}


# ---------------------------------------------------------------------------
# DataExtractor
# ---------------------------------------------------------------------------

class DataExtractor:
    def __init__(
        self,
        api_key: str | None = None,
        model: str | None = None,
        provider: str | None = None,
        base_url: str | None = None,
    ) -> None:

        # Detect provider
        if provider:
            self._provider = provider.lower()
        else:
            self._provider = os.getenv("LLM_PROVIDER", "ollama").lower()

        # ---------------- GOOGLE (FIXED) ----------------
        if self._provider == "google":
            self._api_key = api_key or os.getenv("GOOGLE_API_KEY")
            if not self._api_key:
                raise ValueError("GOOGLE_API_KEY not set")

            genai.configure(api_key=self._api_key)

            self._model = model or os.getenv("GOOGLE_MODEL", _GOOGLE_DEFAULT_MODEL)
            self._provider = "google_native"

        # ---------------- OLLAMA ----------------
        elif self._provider == "ollama":
            self._base_url = base_url or _OLLAMA_DEFAULT_BASE_URL
            self._api_key = "ollama"
            self._model = model or _OLLAMA_DEFAULT_MODEL

        # ---------------- OPENAI ----------------
        else:
            self._provider = "openai"
            self._api_key = api_key or os.getenv("OPENAI_API_KEY")
            self._model = model or _DEFAULT_MODEL

    # ------------------------------------------------------------------

    def extract(
        self,
        machine_name: str,
        task_type: str | None = None,
        ebom_reference: list[dict] | None = None,
    ) -> dict[str, Any]:

        user_message = build_user_message(machine_name, task_type, ebom_reference)

        raw = self._call_llm(user_message)

        return self._parse_and_validate(raw)

    # ------------------------------------------------------------------

    def _call_llm(self, user_message: str) -> str:

        # ✅ GOOGLE (FIXED)
        if self._provider == "google_native":
            try:
                model = genai.GenerativeModel(self._model)

                response = model.generate_content(
                    f"{SYSTEM_PROMPT}\n\n{user_message}"
                )

                if hasattr(response, "text"):
                    return response.text

                return "No response from Gemini"

            except Exception as e:
                print("GOOGLE ERROR:", str(e))
                raise

        # ---------------- OLLAMA / OPENAI ----------------
        if OpenAI is None:
            raise RuntimeError("Install openai package")

        client = OpenAI(api_key=self._api_key)

        response = client.chat.completions.create(
            model=self._model,
            messages=[
                {"role": "system", "content": SYSTEM_PROMPT},
                {"role": "user", "content": user_message},
            ],
            temperature=0.2,
        )

        return response.choices[0].message.content or ""

    # ------------------------------------------------------------------

    @staticmethod
    def _parse_and_validate(raw: str) -> dict[str, Any]:

        text = raw.strip()

        if text.startswith("```"):
            lines = text.splitlines()[1:-1]
            text = "\n".join(lines)

        try:
            data = json.loads(text)
        except json.JSONDecodeError as exc:
            raise ValueError(f"Invalid JSON: {exc}\n\nRaw:\n{raw}")

        missing = _REQUIRED_KEYS - data.keys()
        if missing:
            raise ValueError(f"Missing keys: {missing}")

        return data
