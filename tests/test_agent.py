"""
Tests for the AI datasheet-filling agent.

Run with:
    python -m pytest tests/ -v
"""

from __future__ import annotations

import json
import os
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import MagicMock, patch

# Make the repo root importable regardless of how pytest is invoked.
sys.path.insert(0, str(Path(__file__).parent.parent))

from agent.extractor import DataExtractor, _REQUIRED_COLUMNS, _REQUIRED_KEYS
from agent.prompts import SYSTEM_PROMPT, build_user_message
from utils.excel_handler import (
    SHEET_COLUMNS,
    create_template,
    fill_template,
    read_ebom_reference,
    read_template_structure,
)


# ---------------------------------------------------------------------------
# Shared fixture data
# ---------------------------------------------------------------------------

_SAMPLE_DATA: dict = {
    "DATASHEET": [
        {
            "Parameter": "Power Rating",
            "Value": "15",
            "Unit": "kW",
            "Description": "Rated mechanical output power",
            "Source": "Manufacturer Datasheet",
        },
        {
            "Parameter": "Voltage",
            "Value": "400",
            "Unit": "V",
            "Description": "Nominal supply voltage",
            "Source": "Manufacturer Datasheet",
        },
    ],
    "EBOM": [
        {
            "Component Name": "Stator Winding",
            "Quantity": "1",
            "Specification": "Copper, Class F insulation",
            "Material": "Copper",
            "Supplier": "Motor OEM",
            "Notes": "Rewound at site",
        }
    ],
    "SRD": [
        {
            "Req ID": "REQ-01",
            "Requirement Description": "Motor shall deliver 15 kW continuous at rated voltage.",
            "Type": "Functional",
            "Priority": "High",
            "Source": "Client Spec",
            "Validation Method": "Load Test",
        }
    ],
    "CDD": [
        {
            "Section": "1",
            "Title": "System Overview",
            "Description": "15 kW three-phase induction motor for pump drive.",
        }
    ],
}


# ---------------------------------------------------------------------------
# Tests for agent.prompts
# ---------------------------------------------------------------------------


class TestPrompts(unittest.TestCase):
    def test_system_prompt_is_nonempty(self):
        self.assertGreater(len(SYSTEM_PROMPT), 100)

    def test_system_prompt_contains_json_key_names(self):
        for key in ["DATASHEET", "EBOM", "SRD", "CDD"]:
            self.assertIn(key, SYSTEM_PROMPT)

    def test_build_user_message_contains_machine_name(self):
        msg = build_user_message("Test Motor 5 kW")
        self.assertIn("Test Motor 5 kW", msg)

    def test_build_user_message_with_task_type(self):
        msg = build_user_message("Test Motor", task_type="maintenance")
        self.assertIn("maintenance", msg)

    def test_build_user_message_with_ebom_reference(self):
        ref = [{"Component Name": "Bearing", "Quantity": "2"}]
        msg = build_user_message("Test Motor", ebom_reference=ref)
        self.assertIn("Bearing", msg)

    def test_build_user_message_no_optional_args(self):
        msg = build_user_message("Motor X")
        self.assertNotIn("Task Type", msg)
        self.assertNotIn("EBOM Reference", msg)


# ---------------------------------------------------------------------------
# Tests for agent.extractor (DataExtractor)
# ---------------------------------------------------------------------------


class TestDataExtractorValidation(unittest.TestCase):
    """Test _parse_and_validate without calling the real LLM."""

    def _valid_json(self):
        return json.dumps(_SAMPLE_DATA)

    def test_valid_response_parsed_correctly(self):
        result = DataExtractor._parse_and_validate(self._valid_json())
        self.assertEqual(set(result.keys()), _REQUIRED_KEYS)
        self.assertEqual(result["DATASHEET"][0]["Parameter"], "Power Rating")

    def test_strips_markdown_fences(self):
        fenced = "```json\n" + self._valid_json() + "\n```"
        result = DataExtractor._parse_and_validate(fenced)
        self.assertIn("DATASHEET", result)

    def test_strips_plain_backtick_fences(self):
        fenced = "```\n" + self._valid_json() + "\n```"
        result = DataExtractor._parse_and_validate(fenced)
        self.assertIn("EBOM", result)

    def test_raises_on_invalid_json(self):
        with self.assertRaises(ValueError) as ctx:
            DataExtractor._parse_and_validate("not json at all")
        self.assertIn("not valid JSON", str(ctx.exception))

    def test_raises_on_missing_top_level_key(self):
        data = dict(_SAMPLE_DATA)
        del data["CDD"]
        with self.assertRaises(ValueError) as ctx:
            DataExtractor._parse_and_validate(json.dumps(data))
        self.assertIn("CDD", str(ctx.exception))

    def test_raises_on_non_dict_response(self):
        with self.assertRaises(ValueError):
            DataExtractor._parse_and_validate("[1, 2, 3]")

    def test_raises_on_missing_column_in_row(self):
        data = dict(_SAMPLE_DATA)
        data["DATASHEET"] = [{"Parameter": "X"}]  # missing Value, Unit, Description, Source
        with self.assertRaises(ValueError) as ctx:
            DataExtractor._parse_and_validate(json.dumps(data))
        self.assertIn("DATASHEET", str(ctx.exception))

    def test_raises_when_no_api_key(self):
        extractor = DataExtractor(api_key=None)
        env = {k: v for k, v in os.environ.items() if k != "OPENAI_API_KEY"}
        with patch.dict(os.environ, env, clear=True):
            with self.assertRaises((ValueError, RuntimeError)):
                extractor.extract("Test Motor")

    def test_extract_calls_openai_and_returns_data(self):
        """Mock the OpenAI client so we don't make real API calls."""
        mock_response = MagicMock()
        mock_response.choices[0].message.content = json.dumps(_SAMPLE_DATA)

        mock_client = MagicMock()
        mock_client.chat.completions.create.return_value = mock_response

        with patch("agent.extractor.OpenAI", return_value=mock_client):
            extractor = DataExtractor(api_key="test-key-123")
            result = extractor.extract("Siemens Motor 15 kW", task_type="maintenance")

        self.assertIn("DATASHEET", result)
        self.assertIn("EBOM", result)
        self.assertEqual(result["DATASHEET"][0]["Value"], "15")

    def test_extractor_uses_env_model(self):
        mock_response = MagicMock()
        mock_response.choices[0].message.content = json.dumps(_SAMPLE_DATA)
        mock_client = MagicMock()
        mock_client.chat.completions.create.return_value = mock_response

        with patch.dict(os.environ, {"OPENAI_MODEL": "gpt-4-turbo", "OPENAI_API_KEY": "k"}):
            with patch("agent.extractor.OpenAI", return_value=mock_client):
                extractor = DataExtractor()
                extractor.extract("Motor")
                call_kwargs = mock_client.chat.completions.create.call_args[1]
                self.assertEqual(call_kwargs["model"], "gpt-4-turbo")


# ---------------------------------------------------------------------------
# Tests for utils.excel_handler
# ---------------------------------------------------------------------------


class TestExcelHandler(unittest.TestCase):
    def setUp(self):
        self._tmpdir = tempfile.mkdtemp()
        self._template_path = Path(self._tmpdir) / "template.xlsx"
        self._output_path = Path(self._tmpdir) / "filled.xlsx"

    def test_create_template_creates_file(self):
        create_template(self._template_path)
        self.assertTrue(self._template_path.exists())

    def test_create_template_has_four_sheets(self):
        import openpyxl
        create_template(self._template_path)
        wb = openpyxl.load_workbook(self._template_path)
        self.assertEqual(set(wb.sheetnames), {"DATASHEET", "EBOM", "SRD", "CDD"})

    def test_create_template_has_correct_headers(self):
        create_template(self._template_path)
        structure = read_template_structure(self._template_path)
        for sheet, expected_cols in SHEET_COLUMNS.items():
            self.assertEqual(structure[sheet], expected_cols,
                             f"Header mismatch in sheet '{sheet}'")

    def test_read_template_structure(self):
        create_template(self._template_path)
        structure = read_template_structure(self._template_path)
        self.assertIn("DATASHEET", structure)
        self.assertIn("Parameter", structure["DATASHEET"])

    def test_fill_template_writes_data(self):
        import openpyxl
        create_template(self._template_path)
        fill_template(_SAMPLE_DATA, self._template_path, self._output_path)
        self.assertTrue(self._output_path.exists())

        wb = openpyxl.load_workbook(self._output_path)
        ws = wb["DATASHEET"]
        # Row 2 should have the first data row
        self.assertEqual(ws.cell(row=2, column=1).value, "Power Rating")
        self.assertEqual(ws.cell(row=2, column=2).value, "15")

    def test_fill_template_srd_rows(self):
        import openpyxl
        create_template(self._template_path)
        fill_template(_SAMPLE_DATA, self._template_path, self._output_path)
        wb = openpyxl.load_workbook(self._output_path)
        ws = wb["SRD"]
        self.assertEqual(ws.cell(row=2, column=1).value, "REQ-01")

    def test_fill_template_cdd_rows(self):
        import openpyxl
        create_template(self._template_path)
        fill_template(_SAMPLE_DATA, self._template_path, self._output_path)
        wb = openpyxl.load_workbook(self._output_path)
        ws = wb["CDD"]
        self.assertEqual(ws.cell(row=2, column=2).value, "System Overview")

    def test_read_ebom_reference_from_filled_template(self):
        create_template(self._template_path)
        fill_template(_SAMPLE_DATA, self._template_path, self._output_path)
        rows = read_ebom_reference(self._output_path, sheet_name="EBOM")
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["Component Name"], "Stator Winding")

    def test_read_ebom_reference_missing_sheet(self):
        create_template(self._template_path)
        # EBOM sheet exists but let's test a non-existent sheet name
        rows = read_ebom_reference(self._template_path, sheet_name="NONEXISTENT")
        self.assertEqual(rows, [])

    def test_fill_template_creates_output_dir(self):
        create_template(self._template_path)
        deep_output = Path(self._tmpdir) / "nested" / "deep" / "filled.xlsx"
        fill_template(_SAMPLE_DATA, self._template_path, deep_output)
        self.assertTrue(deep_output.exists())

    def test_create_template_creates_parent_dirs(self):
        deep_template = Path(self._tmpdir) / "a" / "b" / "template.xlsx"
        create_template(deep_template)
        self.assertTrue(deep_template.exists())


# ---------------------------------------------------------------------------
# Tests for utils.pdf_converter
# ---------------------------------------------------------------------------


class TestPdfConverter(unittest.TestCase):
    def setUp(self):
        self._tmpdir = tempfile.mkdtemp()
        self._pdf_path = Path(self._tmpdir) / "output.pdf"

    def test_generate_pdf_creates_file(self):
        from utils.pdf_converter import generate_pdf
        generate_pdf(_SAMPLE_DATA, self._pdf_path, machine_name="Test Motor")
        self.assertTrue(self._pdf_path.exists())
        self.assertGreater(self._pdf_path.stat().st_size, 1000)

    def test_generate_pdf_creates_parent_dirs(self):
        from utils.pdf_converter import generate_pdf
        deep_pdf = Path(self._tmpdir) / "a" / "b" / "report.pdf"
        generate_pdf(_SAMPLE_DATA, deep_pdf)
        self.assertTrue(deep_pdf.exists())

    def test_generate_pdf_with_empty_sections(self):
        from utils.pdf_converter import generate_pdf
        sparse_data = {
            "DATASHEET": _SAMPLE_DATA["DATASHEET"],
            "EBOM": [],
            "SRD": [],
            "CDD": [],
        }
        generate_pdf(sparse_data, self._pdf_path)
        self.assertTrue(self._pdf_path.exists())

    def test_generate_pdf_returns_path(self):
        from utils.pdf_converter import generate_pdf
        result = generate_pdf(_SAMPLE_DATA, self._pdf_path)
        self.assertEqual(result, self._pdf_path)


# ---------------------------------------------------------------------------
# Integration-style test for main.run (mocked LLM)
# ---------------------------------------------------------------------------


class TestMainRun(unittest.TestCase):
    def setUp(self):
        self._tmpdir = tempfile.mkdtemp()

    def test_run_full_pipeline(self):
        """Integration test: mock LLM, run full pipeline, verify outputs."""
        from main import run

        mock_response = MagicMock()
        mock_response.choices[0].message.content = json.dumps(_SAMPLE_DATA)
        mock_client = MagicMock()
        mock_client.chat.completions.create.return_value = mock_response

        output_dir = Path(self._tmpdir) / "output"
        template_path = Path(self._tmpdir) / "templates" / "template.xlsx"

        with patch("agent.extractor.OpenAI", return_value=mock_client):
            with patch.dict(os.environ, {"OPENAI_API_KEY": "test-key"}):
                result = run([
                    "--machine", "Test Motor 5 kW",
                    "--task", "maintenance",
                    "--template", str(template_path),
                    "--output", str(output_dir),
                ])

        self.assertIn("DATASHEET", result)
        # Check XLSX was created
        xlsx_files = list(output_dir.glob("*.xlsx"))
        self.assertEqual(len(xlsx_files), 1)
        # Check PDF was created
        pdf_files = list(output_dir.glob("*.pdf"))
        self.assertEqual(len(pdf_files), 1)

    def test_run_no_pdf_flag(self):
        """--no-pdf should skip PDF generation."""
        from main import run

        mock_response = MagicMock()
        mock_response.choices[0].message.content = json.dumps(_SAMPLE_DATA)
        mock_client = MagicMock()
        mock_client.chat.completions.create.return_value = mock_response

        output_dir = Path(self._tmpdir) / "output_nopdf"
        template_path = Path(self._tmpdir) / "templates" / "template.xlsx"

        with patch("agent.extractor.OpenAI", return_value=mock_client):
            with patch.dict(os.environ, {"OPENAI_API_KEY": "test-key"}):
                run([
                    "--machine", "Test Motor",
                    "--template", str(template_path),
                    "--output", str(output_dir),
                    "--no-pdf",
                ])

        pdf_files = list(output_dir.glob("*.pdf"))
        self.assertEqual(len(pdf_files), 0)

    def test_run_saves_json_when_requested(self):
        from main import run

        mock_response = MagicMock()
        mock_response.choices[0].message.content = json.dumps(_SAMPLE_DATA)
        mock_client = MagicMock()
        mock_client.chat.completions.create.return_value = mock_response

        output_dir = Path(self._tmpdir) / "output_json"
        template_path = Path(self._tmpdir) / "templates2" / "template.xlsx"
        json_out = Path(self._tmpdir) / "out.json"

        with patch("agent.extractor.OpenAI", return_value=mock_client):
            with patch.dict(os.environ, {"OPENAI_API_KEY": "test-key"}):
                run([
                    "--machine", "Test Motor",
                    "--template", str(template_path),
                    "--output", str(output_dir),
                    "--no-pdf",
                    "--json-out", str(json_out),
                ])

        self.assertTrue(json_out.exists())
        loaded = json.loads(json_out.read_text())
        self.assertIn("DATASHEET", loaded)


if __name__ == "__main__":
    unittest.main()
