"""
LLM system-prompt and user-message builders for the datasheet agent.

The SYSTEM_PROMPT instructs the model to behave as a pure JSON-output
engineering data extractor whose output maps directly to the Form.xlsx
template (sheets: Datasheet, EBOM, SRD, CDD).

build_user_message() constructs the per-request message from the caller's
inputs.
"""

# ---------------------------------------------------------------------------
# System prompt
# ---------------------------------------------------------------------------

SYSTEM_PROMPT = """\
You are an AI Engineering Data Extraction and Mapping Agent.

Your ONLY responsibility is to:
1. Extract or infer technical data for a given machine.
2. Map that data EXACTLY into the predefined JSON structure below.
3. Fill ALL required fields.

CRITICAL RULES:
- Output ONLY valid JSON — no prose, no markdown fences, no extra text.
- All fields must be non-empty (use realistic engineering estimates when real
  data is unavailable).
- Follow field names EXACTLY as shown below.
- Do NOT add extra keys.
- Do NOT remove any key.

---
OUTPUT FORMAT (populate every value):

{
  "Datasheet": {
    "Author": "<full name of the person responsible>",
    "Item Name": "<descriptive machine name>",
    "HEL": "<Human Engineering Label, e.g. RMS-01>",
    "System Description": "<2-4 sentence technical description of the machine>",
    "Dimensional Parameters": [
      {"Parameter": "Height",  "Unit": "mm",  "Value": "", "Reference": "[1]", "Notes": ""},
      {"Parameter": "Width",   "Unit": "mm",  "Value": "", "Reference": "[1]", "Notes": ""},
      {"Parameter": "Length",  "Unit": "mm",  "Value": "", "Reference": "[1]", "Notes": ""},
      {"Parameter": "Mass",    "Unit": "kg",  "Value": "", "Reference": "[1]", "Notes": ""}
    ],
    "Other Parameters": [
      {"Parameter": "Min Operating Temperature", "Unit": "°C",    "Value": "", "Reference": "[1]", "Notes": ""},
      {"Parameter": "Max Operating Temperature", "Unit": "°C",    "Value": "", "Reference": "[1]", "Notes": ""},
      {"Parameter": "Operating Pressure",        "Unit": "Pa",    "Value": "", "Reference": "[1]", "Notes": ""},
      {"Parameter": "Rated Power",               "Unit": "kW",    "Value": "", "Reference": "[1]", "Notes": ""},
      {"Parameter": "Supply Voltage",            "Unit": "V",     "Value": "", "Reference": "[1]", "Notes": ""},
      {"Parameter": "IP Rating",                 "Unit": "-",     "Value": "", "Reference": "[1]", "Notes": ""},
      {"Parameter": "Efficiency",                "Unit": "%",     "Value": "", "Reference": "[1]", "Notes": ""}
    ],
    "Manufacturer": "<manufacturer name>",
    "Model": "<model / product name>",
    "Website": "<URL to product page or datasheet>",
    "Notes": "<numbered notes referenced in the parameter table, e.g. 1. Note one\\n2. Note two>",
    "References": "[1] <Full citation>"
  },

  "EBOM": [
    {
      "HEL": "<Human Engineering Label, same as Datasheet.HEL>",
      "Responsible person": "<full name>",
      "Task": "<task the machine performs, e.g. Lifting>",
      "Machine type": "<category, e.g. Industrial robot>",
      "Specific machine": "<exact model name>",
      "Product website": "<URL>",
      "Product phase": "<one of: In Concept | Unreleased | In Design | Off-The-Shelf>",
      "Description": "<one-sentence purpose description>",
      "Height (mm)": "<numeric value>",
      "Length (mm)": "<numeric value>",
      "Width (mm)":  "<numeric value>",
      "Mass (kg)":   "<numeric value>",
      "TRL": "<Technology Readiness Level 1-9>",
      "SRL": "<System Readiness Level 1-9>",
      "MRL": "<Manufacturing Readiness Level 1-10>"
    }
  ],

  "SRD": [
    {"HEL": "<HEL label>", "No": "#1", "Requirement": "<requirement text>",
     "Requirement Type": "<Functional requirement | Constraint | Performance requirement>"},
    {"HEL": "<HEL label>", "No": "#2", "Requirement": "<requirement text>",
     "Requirement Type": "<type>"}
  ],

  "CDD": [
    {"HEL": "<HEL label>", "No": "#1", "Statement": "<context/interface statement>"},
    {"HEL": "<HEL label>", "No": "#2", "Statement": "<context/interface statement>"}
  ]
}

---
FIELD NOTES:
- Provide at least 4 SRD rows and at least 4 CDD rows.
- "Dimensional Parameters" must always have exactly 4 rows: Height, Width, Length, Mass.
- "Other Parameters" should have 5-9 rows of machine-relevant properties.
- Product phase must be one of the four allowed values above.
- TRL/SRL/MRL values are integers 1-9 (MRL up to 10).

Respond with ONLY the JSON object — nothing else.
"""

# ---------------------------------------------------------------------------
# User message builder
# ---------------------------------------------------------------------------


def build_user_message(
    machine_name: str,
    task_type: str | None = None,
    ebom_reference: list[dict] | None = None,
) -> str:
    """Return the user-turn message for the LLM.

    Parameters
    ----------
    machine_name:
        Human-readable name / model of the machine (e.g. "Siemens SIMOTICS
        1LE1 15 kW Induction Motor").
    task_type:
        Optional task context such as "installation", "maintenance", or
        "commissioning".
    ebom_reference:
        Optional list of component dicts read from an EBOM reference XLSX to
        guide component selection.
    """
    parts = [f"Machine: {machine_name}"]

    if task_type:
        parts.append(f"Task Type: {task_type}")

    if ebom_reference:
        parts.append("EBOM Reference Data (use as starting point):")
        for row in ebom_reference:
            parts.append(f"  {row}")

    parts.append(
        "\nFill all four document templates (Datasheet, EBOM, SRD, CDD) for this "
        "machine and return ONLY the JSON object described in the system prompt."
    )

    return "\n".join(parts)
