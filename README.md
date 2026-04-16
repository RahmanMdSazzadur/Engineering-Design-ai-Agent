# AI Agent for Filling Engineering Datasheets

An AI-powered engineering agent that extracts or infers technical data for any machine and automatically fills four standardised template documents — **DATASHEET**, **EBOM**, **SRD**, and **CDD** — then exports them as a filled XLSX file and a combined PDF report.

---

## Features

| Capability | Detail |
|---|---|
| **Machine data extraction** | Uses OpenAI GPT to extract or infer electrical/mechanical properties |
| **Template filling** | Fills four sheets: DATASHEET · EBOM · SRD · CDD |
| **XLSX output** | Styled, ready-to-share Excel file |
| **PDF output** | Single combined PDF document with all four sections |
| **EBOM reference** | Optionally seeds component list from an existing EBOM XLSX |
| **Strict JSON schema** | LLM output is validated against required column definitions |

---

## Project Structure

```
.
├── main.py                  # CLI entry point
├── requirements.txt
├── .env.example
├── agent/
│   ├── extractor.py         # DataExtractor — calls OpenAI and validates response
│   └── prompts.py           # System prompt and user-message builder
├── utils/
│   ├── excel_handler.py     # Create / fill / read XLSX templates
│   └── pdf_converter.py     # Generate combined PDF from JSON data
├── templates/
│   └── template.xlsx        # Blank four-sheet template (auto-created if missing)
└── tests/
    └── test_agent.py        # 34 unit + integration tests
```

---

## Quick Start

### 1. Install dependencies

```bash
pip install -r requirements.txt
```

### 2. Configure your OpenAI API key

```bash
cp .env.example .env
# edit .env and set OPENAI_API_KEY=sk-...
```

### 3. Run the agent

```bash
python main.py --machine "Siemens SIMOTICS 1LE1 15 kW Induction Motor" \
               --task maintenance \
               --output output/
```

This creates:
- `output/Siemens_SIMOTICS_1LE1_15_kW_Induction_Motor_filled.xlsx`
- `output/Siemens_SIMOTICS_1LE1_15_kW_Induction_Motor_report.pdf`

---

## CLI Reference

```
usage: main.py [-h] --machine MACHINE [--task TASK] [--template TEMPLATE]
               [--ebom-ref EBOM_REF] [--output OUTPUT] [--no-pdf]
               [--json-out JSON_OUT]

Arguments:
  --machine, -m     Machine name / model (required)
  --task, -t        Optional task type: maintenance | installation | commissioning
  --template        Path to template XLSX (default: templates/template.xlsx)
  --ebom-ref        Path to EBOM reference XLSX for seeding component list
  --output, -o      Output directory (default: output/)
  --no-pdf          Skip PDF generation
  --json-out        Also save the raw JSON to this path
```

---

## Template Sheets

### DATASHEET
Electrical and mechanical properties of the machine.

| Parameter | Value | Unit | Description | Source |
|---|---|---|---|---|
| Power Rating | 15 | kW | Rated mechanical output | Manufacturer Datasheet |
| Voltage | 400 | V | Nominal supply voltage | Manufacturer Datasheet |
| … | | | | |

### EBOM (Engineering Bill of Materials)
Major physical components.

| Component Name | Quantity | Specification | Material | Supplier | Notes |
|---|---|---|---|---|---|
| Stator Winding | 1 | Class F insulation | Copper | Motor OEM | |

### SRD (System Requirements Document)
Functional and non-functional requirements.

| Req ID | Requirement Description | Type | Priority | Source | Validation Method |
|---|---|---|---|---|---|
| REQ-01 | Motor shall deliver 15 kW continuous | Functional | High | Client Spec | Load Test |

### CDD (Concept Design Document)
Structured design narrative.

| Section | Title | Description |
|---|---|---|
| 1 | System Overview | 15 kW three-phase induction motor for pump drive |
| 2 | Architecture | … |

---

## Programmatic Usage

```python
from agent.extractor import DataExtractor
from utils.excel_handler import create_template, fill_template
from utils.pdf_converter import generate_pdf

extractor = DataExtractor(api_key="sk-...")
data = extractor.extract(
    machine_name="ABB ACS880 Variable Speed Drive 22 kW",
    task_type="commissioning",
)

create_template("templates/template.xlsx")
fill_template(data, "templates/template.xlsx", "output/filled.xlsx")
generate_pdf(data, "output/report.pdf", machine_name="ABB ACS880 22 kW VSD")
```

---

## Environment Variables

| Variable | Required | Default | Description |
|---|---|---|---|
| `OPENAI_API_KEY` | Yes | — | OpenAI API key |
| `OPENAI_MODEL` | No | `gpt-4o` | Model to use for extraction |

---

## Running Tests

```bash
pip install pytest
python -m pytest tests/ -v
```

All 34 tests run without an API key (the LLM is mocked).

---

## Output Format (JSON Schema)

The agent always produces and validates this structure before writing files:

```json
{
  "DATASHEET": [
    {"Parameter": "", "Value": "", "Unit": "", "Description": "", "Source": ""}
  ],
  "EBOM": [
    {"Component Name": "", "Quantity": "", "Specification": "",
     "Material": "", "Supplier": "", "Notes": ""}
  ],
  "SRD": [
    {"Req ID": "", "Requirement Description": "", "Type": "",
     "Priority": "", "Source": "", "Validation Method": ""}
  ],
  "CDD": [
    {"Section": "", "Title": "", "Description": ""}
  ]
}
```
