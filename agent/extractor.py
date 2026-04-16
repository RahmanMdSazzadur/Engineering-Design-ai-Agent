"""
DataExtractor — calls the OpenAI chat API to extract and map machine data.

Usage
-----
    from agent.extractor import DataExtractor

    extractor = DataExtractor()          # reads OPENAI_API_KEY from env
    data = extractor.extract(
        machine_name="Siemens 1LE1 15 kW Induction Motor",
        task_type="maintenance",
    )
    # data is a dict with keys: Datasheet (dict), EBOM, SRD, CDD (lists)

Environment variables
---------------------
    OPENAI_API_KEY   — required (unless api_key is passed directly)
    OPENAI_MODEL     — optional, default "gpt-4o"
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

    Parameters
    ----------
    api_key:
        OpenAI API key.  Falls back to the ``OPENAI_API_KEY`` environment
        variable when not provided.
    model:
        OpenAI model to use.  Falls back to the ``OPENAI_MODEL`` environment
        variable, then to ``gpt-4o``.
    """

    def __init__(
        self,
        api_key: str | None = None,
        model: str | None = None,
    ) -> None:
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
        """Call the OpenAI chat-completions endpoint and return raw text."""
        if OpenAI is None:
            raise RuntimeError(
                "The 'openai' package is required. "
                "Install it with: pip install openai"
            )

        if not self._api_key:
            raise ValueError(
                "No OpenAI API key provided. "
                "Set OPENAI_API_KEY in your environment or pass api_key= to DataExtractor."
            )

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
