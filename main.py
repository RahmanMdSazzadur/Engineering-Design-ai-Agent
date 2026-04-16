#!/usr/bin/env python3
"""
main.py — CLI entry point for the AI datasheet-filling agent.

Usage
-----
    python main.py --machine "Siemens SIMOTICS 1LE1 15 kW Induction Motor" \\
                   --task maintenance \\
                   --output output/

Optional flags
--------------
    --template PATH     Path to template XLSX (default: templates/template.xlsx)
    --ebom-ref  PATH    Path to EBOM reference XLSX to guide component selection
    --no-pdf            Skip PDF generation
    --json-out  PATH    Also save the raw JSON to this file

Environment
-----------
    OPENAI_API_KEY   — required for LLM extraction
    OPENAI_MODEL     — optional (default: gpt-4o)
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

# Load .env if present
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

from agent.extractor import DataExtractor
from utils.excel_handler import (
    create_template,
    fill_template,
    read_ebom_reference,
)
from utils.pdf_converter import generate_pdf


_DEFAULT_TEMPLATE = Path(__file__).parent / "templates" / "template.xlsx"
_DEFAULT_OUTPUT_DIR = Path("output")


def _parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="AI agent that extracts machine data and fills engineering templates.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    parser.add_argument(
        "--machine", "-m",
        required=True,
        help="Machine name / model (e.g. 'Siemens 1LE1 15 kW Induction Motor')",
    )
    parser.add_argument(
        "--task", "-t",
        default=None,
        help="Optional task type context (e.g. maintenance, installation, commissioning)",
    )
    parser.add_argument(
        "--template",
        default=str(_DEFAULT_TEMPLATE),
        help=f"Path to template XLSX (default: {_DEFAULT_TEMPLATE})",
    )
    parser.add_argument(
        "--ebom-ref",
        default=None,
        help="Path to EBOM reference XLSX file",
    )
    parser.add_argument(
        "--output", "-o",
        default=str(_DEFAULT_OUTPUT_DIR),
        help=f"Output directory (default: {_DEFAULT_OUTPUT_DIR})",
    )
    parser.add_argument(
        "--no-pdf",
        action="store_true",
        help="Skip PDF generation",
    )
    parser.add_argument(
        "--json-out",
        default=None,
        help="Optional path to also save the raw JSON output",
    )
    return parser.parse_args(argv)


def run(argv: list[str] | None = None) -> dict:
    """Execute the full agent pipeline.

    Parameters
    ----------
    argv:
        Command-line argument list (uses sys.argv when None).

    Returns
    -------
    dict
        The extracted and mapped data dict (DATASHEET, EBOM, SRD, CDD).
    """
    args = _parse_args(argv)

    # ------------------------------------------------------------------ #
    # 1. Ensure the template XLSX exists
    # ------------------------------------------------------------------ #
    template_path = Path(args.template)
    if not template_path.exists():
        print(f"[info] Template not found at '{template_path}' — creating a blank one.")
        create_template(template_path)
        print(f"[info] Blank template created: {template_path}")

    # ------------------------------------------------------------------ #
    # 2. Optionally read EBOM reference data
    # ------------------------------------------------------------------ #
    ebom_reference: list[dict] | None = None
    if args.ebom_ref:
        ebom_ref_path = Path(args.ebom_ref)
        if not ebom_ref_path.exists():
            print(f"[warn] EBOM reference file not found: {ebom_ref_path}", file=sys.stderr)
        else:
            ebom_reference = read_ebom_reference(ebom_ref_path)
            print(f"[info] Loaded {len(ebom_reference)} EBOM reference rows from '{ebom_ref_path}'")

    # ------------------------------------------------------------------ #
    # 3. Extract machine data via LLM
    # ------------------------------------------------------------------ #
    print(f"[info] Extracting data for: {args.machine!r}")
    if args.task:
        print(f"[info] Task type: {args.task!r}")

    extractor = DataExtractor()
    data = extractor.extract(
        machine_name=args.machine,
        task_type=args.task,
        ebom_reference=ebom_reference,
    )
    print("[info] Data extraction complete.")

    # ------------------------------------------------------------------ #
    # 4. Save JSON if requested
    # ------------------------------------------------------------------ #
    if args.json_out:
        json_path = Path(args.json_out)
        json_path.parent.mkdir(parents=True, exist_ok=True)
        json_path.write_text(json.dumps(data, indent=2, ensure_ascii=False))
        print(f"[info] JSON saved: {json_path}")

    # ------------------------------------------------------------------ #
    # 5. Fill XLSX template
    # ------------------------------------------------------------------ #
    output_dir = Path(args.output)
    machine_slug = args.machine.replace(" ", "_").replace("/", "-")[:50]
    xlsx_path = output_dir / f"{machine_slug}_filled.xlsx"

    filled_xlsx = fill_template(data, template_path, xlsx_path)
    print(f"[info] Filled XLSX saved: {filled_xlsx}")

    # ------------------------------------------------------------------ #
    # 6. Generate PDF
    # ------------------------------------------------------------------ #
    if not args.no_pdf:
        pdf_path = output_dir / f"{machine_slug}_report.pdf"
        generate_pdf(data, pdf_path, machine_name=args.machine)
        print(f"[info] PDF report saved: {pdf_path}")

    return data


if __name__ == "__main__":
    run()
