# AI Agent for Filling Engineering Datasheets

An AI-powered engineering agent that extracts or infers technical data for any machine and automatically fills four standardised documents — **Datasheet**, **EBOM**, **SRD**, and **CDD** — using the provided **Form.xlsx** template, then exports a filled XLSX and a combined PDF report.

---

## Features

| Capability | Detail |
|---|---|
| **Machine data extraction** | Uses a local or cloud LLM to extract or infer electrical/mechanical properties |
| **Fully free option** | Run with Ollama (local) — no API key, no cost |
| **Free cloud option** | Run with Groq's free tier — just sign up for a free key |
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
│   ├── extractor.py         # DataExtractor — calls LLM and validates response
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

## Quick Start — Fully Free (Ollama)

### 1. Install Ollama

Download and install from **https://ollama.com** (Windows, macOS, Linux).

### 2. Pull a free model

```bash
ollama pull llama3.2
```

### 3. Install Python dependencies

```bash
pip install -r requirements.txt
```

### 4. Configure the provider

```bash
cp .env.example .env
# The default .env.example already sets LLM_PROVIDER=ollama — no changes needed.
```

### 5. Run the agent

```bash
python main.py --machine "Siemens SIMOTICS 1LE1 15 kW Induction Motor" \
               --task maintenance \
               --output output/
```

This creates:
- `output/Siemens_SIMOTICS_1LE1_15_kW_Induction_Motor_filled.xlsx`
- `output/Siemens_SIMOTICS_1LE1_15_kW_Induction_Motor_report.pdf`

---

## Quick Start — Free Cloud (Groq)

Groq offers a **free tier** with no credit card required and very generous rate limits.

### 1. Get a free Groq API key

Sign up at **https://console.groq.com** → API Keys → Create key.

### 2. Configure

```bash
cp .env.example .env
# Edit .env:
LLM_PROVIDER=groq
GROQ_API_KEY=gsk_your_key_here
```

### 3. Run

```bash
python main.py --machine "ABB ACS880 Variable Speed Drive 22 kW" --task commissioning
```

---

## Quick Start — OpenAI (Paid)

```bash
cp .env.example .env
# Edit .env:
LLM_PROVIDER=openai
OPENAI_API_KEY=sk-your_key_here
OPENAI_MODEL=gpt-4o        # optional, gpt-4o-mini is cheaper
```

```bash
python main.py --machine "Siemens SIMOTICS 1LE1 15 kW Induction Motor"
```

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

## Environment Variables

### Provider selection

| Variable | Default | Description |
|---|---|---|
| `LLM_PROVIDER` | `ollama` (auto-detected) | LLM backend: `ollama`, `groq`, or `openai` |

### Ollama (free, local)

| Variable | Default | Description |
|---|---|---|
| `OLLAMA_BASE_URL` | `http://localhost:11434/v1` | Ollama server URL |
| `OLLAMA_MODEL` | `llama3.2` | Model name (must be pulled first) |

### Groq (free tier)

| Variable | Default | Description |
|---|---|---|
| `GROQ_API_KEY` | — | Free key from console.groq.com |
| `GROQ_MODEL` | `llama-3.1-70b-versatile` | Groq model to use |

### OpenAI (paid)

| Variable | Default | Description |
|---|---|---|
| `OPENAI_API_KEY` | — | OpenAI API key |
| `OPENAI_MODEL` | `gpt-4o` | Model to use |

> **Auto-detection:** If `LLM_PROVIDER` is not set, the agent picks `openai` when
> `OPENAI_API_KEY` is present, `groq` when `GROQ_API_KEY` is present, and defaults
> to `ollama` otherwise.

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

# Free local (Ollama)
extractor = DataExtractor(provider="ollama")

# Free cloud (Groq)
# extractor = DataExtractor(provider="groq", api_key="gsk_...")

# Paid (OpenAI)
# extractor = DataExtractor(provider="openai", api_key="sk-...")

data = extractor.extract(
    machine_name="ABB ACS880 Variable Speed Drive 22 kW",
    task_type="commissioning",
)

fill_template(data, "templates/Form.xlsx", "output/filled.xlsx")
generate_pdf(data, "output/report.pdf", machine_name="ABB ACS880 22 kW VSD")
```

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

