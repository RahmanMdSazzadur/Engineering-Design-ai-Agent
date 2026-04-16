# AI Agent for Filling Engineering Datasheets

An AI-powered engineering agent that extracts or infers technical data for any machine and automatically fills four standardised documents — **Datasheet**, **EBOM**, **SRD**, and **CDD** — using the provided **Form.xlsx** template, then exports a filled XLSX and a combined PDF report.

---

## Features

| Capability | Detail |
|---|---|
| **Machine data extraction** | Uses OpenAI GPT to extract or infer electrical/mechanical properties |
| **Template filling** | Fills Form.xlsx exactly: Datasheet · EBOM · SRD · CDD |
| **XLSX output** | Form.xlsx copy filled with machine-specific data |
| **PDF output** | Single combined PDF document with all four sections |
| **EBOM reference** | Optionally seeds component list from an existing EBOM XLSX |
| **Strict JSON schema** | LLM output is validated against Form.xlsx column definitions |

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
│   ├── excel_handler.py     # Fill Form.xlsx cells by address; read EBOM references
│   └── pdf_converter.py     # Generate combined PDF from structured data
├── templates/
│   └── Form.xlsx            # Primary template (four sheets: Datasheet, EBOM, SRD, CDD)
└── tests/
    └── test_agent.py        # 43 unit + integration tests
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
  --template        Path to template XLSX (default: templates/Form.xlsx)
  --ebom-ref        Path to EBOM reference XLSX for seeding component list
  --output, -o      Output directory (default: output/)
  --no-pdf          Skip PDF generation
  --json-out        Also save the raw JSON to this path
```

---

## Form.xlsx Template Sheets

### Datasheet
Structured form — filled by cell address:

| Field | Cell | Description |
|---|---|---|
| Author | B3 | `Author: <name>` |
| Item Name | E5 | Machine name |
| HEL | E6 | Human Engineering Label |
| System Description | B9 | Free-text description |
| Dimensional Parameters | rows 17-20 | Height, Width, Length, Mass (B=name, D=unit, F=value, H=ref, J=notes) |
| Other Parameters | rows 22-30 | Temperature, pressure, power, etc. |
| Manufacturer | E33 | Manufacturer name |
| Model | E34 | Product model |
| Website | E35 | URL |
| Notes | B38 | Numbered notes |
| References | B41 | Citations |

### EBOM (Engineering Bill of Materials)
Tabular — columns: HEL, Responsible person, Task, Machine type, Specific machine, Product website, Product phase, Description, Height(mm), Length(mm), Width(mm), Mass(kg), TRL, SRL, MRL

### SRD (System Requirements Document)
Tabular — columns: HEL, No, Requirement, Requirement Type

### CDD (Context Definition Document)
Tabular — columns: HEL, No, Statement

---

## Programmatic Usage

```python
from agent.extractor import DataExtractor
from utils.excel_handler import fill_template
from utils.pdf_converter import generate_pdf

extractor = DataExtractor(api_key="sk-...")
data = extractor.extract(
    machine_name="ABB ACS880 Variable Speed Drive 22 kW",
    task_type="commissioning",
)

fill_template(data, "templates/Form.xlsx", "output/filled.xlsx")
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

All 43 tests run without an API key (the LLM is mocked).

---

## Output Format (JSON Schema)

The agent validates this structure before writing any files:

```json
{
  "Datasheet": {
    "Author": "",
    "Item Name": "",
    "HEL": "",
    "System Description": "",
    "Dimensional Parameters": [
      {"Parameter": "", "Unit": "", "Value": "", "Reference": "", "Notes": ""}
    ],
    "Other Parameters": [
      {"Parameter": "", "Unit": "", "Value": "", "Reference": "", "Notes": ""}
    ],
    "Manufacturer": "",
    "Model": "",
    "Website": "",
    "Notes": "",
    "References": ""
  },
  "EBOM": [
    {
      "HEL": "", "Responsible person": "", "Task": "", "Machine type": "",
      "Specific machine": "", "Product website": "", "Product phase": "",
      "Description": "", "Height (mm)": "", "Length (mm)": "", "Width (mm)": "",
      "Mass (kg)": "", "TRL": "", "SRL": "", "MRL": ""
    }
  ],
  "SRD": [
    {"HEL": "", "No": "", "Requirement": "", "Requirement Type": ""}
  ],
  "CDD": [
    {"HEL": "", "No": "", "Statement": ""}
  ]
}
```
