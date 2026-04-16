"""
DataExtractor — calls an LLM to extract and map machine data.

Supports five free/paid backends selected via the ``LLM_PROVIDER`` env var:

  ollama  (default when no key is set)
      Runs a local model via Ollama — completely free, no API key needed.
      Install Ollama from https://ollama.com, then pull a model:
          ollama pull llama3.2
      Set env vars (or leave at defaults):
          LLM_PROVIDER=ollama
          OLLAMA_BASE_URL=http://localhost:11434/v1   # default
          OLLAMA_MODEL=llama3.2                        # default

  google
      Uses Google Gemini via Google AI Studio (free tier, generous limits).
      Get a free key at https://aistudio.google.com/apikey
          LLM_PROVIDER=google
          GOOGLE_API_KEY=AIza...
          GOOGLE_MODEL=gemini-2.0-flash                # default

  deepseek
      Uses DeepSeek's API (very cheap; free credits on sign-up).
      Get a key at https://platform.deepseek.com
          LLM_PROVIDER=deepseek
          DEEPSEEK_API_KEY=sk-...
          DEEPSEEK_MODEL=deepseek-chat                 # default

  groq
      Uses Groq's free cloud API (generous free tier, needs a free key).
      Sign up at https://console.groq.com to get a free API key.
          LLM_PROVIDER=groq
          GROQ_API_KEY=gsk_...
          GROQ_MODEL=llama-3.1-70b-versatile           # default

  openai
      Uses OpenAI's paid API (existing behaviour).
          LLM_PROVIDER=openai
          OPENAI_API_KEY=sk-...
          OPENAI_MODEL=gpt-4o                          # default

Usage
-----
    from agent.extractor import DataExtractor

    extractor = DataExtractor()
    data = extractor.extract(
        machine_name="Siemens 1LE1 15 kW Induction Motor",
        task_type="maintenance",
    )
    # data is a dict with keys: Datasheet (dict), EBOM, SRD, CDD (lists)
"""

from __future__ import annotations

import json
import os
from typing import Any

try:
    from openai import OpenAI
except ImportError:  # pragma: no cover — openai not installed
    OpenAI = None  # type: ignore[assignment,misc]

from agent.prompts import SYSTEM_PROMPT, build_user_message

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

_DEFAULT_MODEL = "gpt-4o"

# Ollama defaults (local, completely free)
_OLLAMA_DEFAULT_BASE_URL = "http://localhost:11434/v1"
_OLLAMA_DEFAULT_MODEL = "llama3.2"

# Google Gemini defaults (free tier — OpenAI-compatible endpoint)
_GOOGLE_BASE_URL = "https://generativelanguage.googleapis.com/v1beta/openai/"
_GOOGLE_DEFAULT_MODEL = "gemini-2.0-flash"

# DeepSeek defaults (very cheap / free credits — OpenAI-compatible endpoint)
_DEEPSEEK_BASE_URL = "https://api.deepseek.com/v1"
_DEEPSEEK_DEFAULT_MODEL = "deepseek-chat"

# Groq defaults (free-tier cloud)
_GROQ_BASE_URL = "https://api.groq.com/openai/v1"
_GROQ_DEFAULT_MODEL = "llama-3.1-70b-versatile"

# Top-level keys required in every valid agent response.
_REQUIRED_KEYS = {"Datasheet", "EBOM", "SRD", "CDD"}

# Required keys for the Datasheet dict.
_DATASHEET_KEYS = {
    "Author", "Item Name", "HEL", "System Description",
    "Dimensional Parameters", "Other Parameters",
    "Manufacturer", "Model", "Website", "Notes", "References",
}

# Required columns for each parameter-row list inside Datasheet.
_PARAM_ROW_KEYS = {"Parameter", "Unit", "Value", "Reference", "Notes"}

# Required columns for tabular sheets (lists of dicts).
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

# Expose for tests / other modules
_REQUIRED_COLUMNS = _LIST_SHEET_COLUMNS  # backward-compat alias


# ---------------------------------------------------------------------------
# DataExtractor
# ---------------------------------------------------------------------------

class DataExtractor:
    """Extract and map engineering data for a machine via an LLM.

    The backend is selected by the ``LLM_PROVIDER`` environment variable
    (or the *provider* constructor argument).  Supported values:

    * ``"ollama"``    — local Ollama server (free, no API key needed)
    * ``"google"``    — Google Gemini API (free tier, needs ``GOOGLE_API_KEY``)
    * ``"deepseek"``  — DeepSeek API (very cheap / free credits, needs ``DEEPSEEK_API_KEY``)
    * ``"groq"``      — Groq cloud API (free tier, needs ``GROQ_API_KEY``)
    * ``"openai"``    — OpenAI API (paid, needs ``OPENAI_API_KEY``)

    When neither ``LLM_PROVIDER`` nor any API key is set the default is
    ``"ollama"``.

    Parameters
    ----------
    api_key:
        API key for the selected provider.  Falls back to the appropriate
        environment variable (``GOOGLE_API_KEY``, ``DEEPSEEK_API_KEY``,
        ``GROQ_API_KEY``, or ``OPENAI_API_KEY``).  Ignored for Ollama.
    model:
        Model name to use.  Falls back to the provider-specific env var,
        then to the provider default.
    provider:
        LLM backend: ``"ollama"``, ``"google"``, ``"deepseek"``, ``"groq"``,
        or ``"openai"``.  Defaults to the ``LLM_PROVIDER`` env var; if that
        is also unset, ``"ollama"`` is chosen when no key is present.
    base_url:
        Override the API base URL (useful for custom Ollama hosts or other
        OpenAI-compatible endpoints).
    """

    def __init__(
        self,
        api_key: str | None = None,
        model: str | None = None,
        provider: str | None = None,
        base_url: str | None = None,
    ) -> None:
        # Determine provider ---------------------------------------------------
        # Explicit arg > env var > auto-detect (ollama when no key is set)
        if provider:
            self._provider = provider.lower()
        else:
            env_provider = os.getenv("LLM_PROVIDER", "").strip().lower()
            if env_provider:
                self._provider = env_provider
            elif os.getenv("OPENAI_API_KEY"):
                self._provider = "openai"
            elif os.getenv("GOOGLE_API_KEY"):
                self._provider = "google"
            elif os.getenv("DEEPSEEK_API_KEY"):
                self._provider = "deepseek"
            elif os.getenv("GROQ_API_KEY"):
                self._provider = "groq"
            else:
                self._provider = "ollama"

        # Configure per provider ----------------------------------------------
        if self._provider == "ollama":
            self._base_url = base_url or os.getenv("OLLAMA_BASE_URL", _OLLAMA_DEFAULT_BASE_URL)
            self._api_key = api_key or "ollama"  # openai client needs a non-empty key
            self._model = model or os.getenv("OLLAMA_MODEL", _OLLAMA_DEFAULT_MODEL)

        elif self._provider == "google":
            self._base_url = base_url or _GOOGLE_BASE_URL
            self._api_key = api_key or os.getenv("GOOGLE_API_KEY")
            self._model = model or os.getenv("GOOGLE_MODEL", _GOOGLE_DEFAULT_MODEL)

        elif self._provider == "deepseek":
            self._base_url = base_url or _DEEPSEEK_BASE_URL
            self._api_key = api_key or os.getenv("DEEPSEEK_API_KEY")
            self._model = model or os.getenv("DEEPSEEK_MODEL", _DEEPSEEK_DEFAULT_MODEL)

        elif self._provider == "groq":
            self._base_url = base_url or _GROQ_BASE_URL
            self._api_key = api_key or os.getenv("GROQ_API_KEY")
            self._model = model or os.getenv("GROQ_MODEL", _GROQ_DEFAULT_MODEL)

        else:  # "openai" (default paid path)
            self._provider = "openai"
            self._base_url = base_url or os.getenv("OPENAI_BASE_URL")
            self._api_key = api_key or os.getenv("OPENAI_API_KEY")
            self._model = model or os.getenv("OPENAI_MODEL", _DEFAULT_MODEL)

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    def extract(
        self,
        machine_name: str,
        task_type: str | None = None,
        ebom_reference: list[dict] | None = None,
    ) -> dict[str, Any]:
        """Return filled template data for *machine_name*.

        Parameters
        ----------
        machine_name:
            Descriptive name / model of the machine.
        task_type:
            Optional task context (e.g. "installation", "maintenance").
        ebom_reference:
            Optional EBOM reference rows from an existing XLSX file.

        Returns
        -------
        dict
            Keys:
              - ``Datasheet`` — dict with form fields and parameter lists
              - ``EBOM`` — list of row dicts
              - ``SRD``  — list of row dicts
              - ``CDD``  — list of row dicts

        Raises
        ------
        ValueError
            If the LLM response is not valid JSON or is missing required keys.
        RuntimeError
            If the OpenAI API call fails.
        """
        user_message = build_user_message(machine_name, task_type, ebom_reference)
        raw = self._call_llm(user_message)
        return self._parse_and_validate(raw)

    # ------------------------------------------------------------------
    # Private helpers
    # ------------------------------------------------------------------

    def _call_llm(self, user_message: str) -> str:
        """Call the configured LLM backend and return raw text."""
        if OpenAI is None:
            raise RuntimeError(
                "The 'openai' package is required. "
                "Install it with: pip install openai"
            )

        if self._provider != "ollama" and not self._api_key:
            _key_map = {
                "google": "GOOGLE_API_KEY",
                "deepseek": "DEEPSEEK_API_KEY",
                "groq": "GROQ_API_KEY",
                "openai": "OPENAI_API_KEY",
            }
            _key_var = _key_map.get(self._provider, "API_KEY")
            raise ValueError(
                f"No API key provided for provider '{self._provider}'. "
                f"Set {_key_var} in your environment or pass api_key= to DataExtractor."
            )

        client_kwargs: dict[str, Any] = {"api_key": self._api_key}
        if self._base_url:
            client_kwargs["base_url"] = self._base_url

        client = OpenAI(**client_kwargs)
        response = client.chat.completions.create(
            model=self._model,
            messages=[
                {"role": "system", "content": SYSTEM_PROMPT},
                {"role": "user", "content": user_message},
            ],
            temperature=0.2,
        )
        return response.choices[0].message.content or ""

    @staticmethod
    def _parse_and_validate(raw: str) -> dict[str, Any]:
        """Parse *raw* JSON and validate the required structure."""
        # Strip markdown code fences if the model wraps output in them.
        text = raw.strip()
        if text.startswith("```"):
            lines = text.splitlines()
            lines = lines[1:] if lines[0].startswith("```") else lines
            if lines and lines[-1].startswith("```"):
                lines = lines[:-1]
            text = "\n".join(lines)

        try:
            data = json.loads(text)
        except json.JSONDecodeError as exc:
            raise ValueError(
                f"LLM response is not valid JSON: {exc}\n\nRaw response:\n{raw}"
            ) from exc

        if not isinstance(data, dict):
            raise ValueError(f"Expected a JSON object; got {type(data).__name__}.")

        missing = _REQUIRED_KEYS - data.keys()
        if missing:
            raise ValueError(
                f"LLM response is missing required top-level keys: {missing}"
            )

        # -- Validate Datasheet (dict) ------------------------------------
        ds = data["Datasheet"]
        if not isinstance(ds, dict):
            raise ValueError(f"'Datasheet' must be a JSON object; got {type(ds).__name__}.")
        missing_ds = _DATASHEET_KEYS - ds.keys()
        if missing_ds:
            raise ValueError(f"'Datasheet' is missing required keys: {missing_ds}")

        for list_field in ("Dimensional Parameters", "Other Parameters"):
            rows = ds.get(list_field, [])
            if not isinstance(rows, list):
                raise ValueError(
                    f"'Datasheet.{list_field}' must be a JSON array; "
                    f"got {type(rows).__name__}."
                )
            for i, row in enumerate(rows):
                if not isinstance(row, dict):
                    raise ValueError(
                        f"Row {i} in 'Datasheet.{list_field}' is not a JSON object."
                    )
                missing_cols = _PARAM_ROW_KEYS - row.keys()
                if missing_cols:
                    raise ValueError(
                        f"Row {i} in 'Datasheet.{list_field}' is missing keys: {missing_cols}"
                    )

        # -- Validate tabular sheets (lists of dicts) ---------------------
        for sheet, required_cols in _LIST_SHEET_COLUMNS.items():
            rows = data.get(sheet, [])
            if not isinstance(rows, list):
                raise ValueError(f"'{sheet}' must be a JSON array; got {type(rows).__name__}.")
            for i, row in enumerate(rows):
                if not isinstance(row, dict):
                    raise ValueError(f"Row {i} in '{sheet}' is not a JSON object.")
                missing_cols = required_cols - row.keys()
                if missing_cols:
                    raise ValueError(
                        f"Row {i} in '{sheet}' is missing columns: {missing_cols}"
                    )

        return data
