"""
LLM system-prompt and user-message builders for the datasheet agent.

The SYSTEM_PROMPT instructs the model to behave as a pure JSON-output
engineering data extractor.  build_user_message() constructs the
per-request message from the caller's inputs.
"""

SYSTEM_PROMPT = """\
You are an AI Engineering Data Extraction and Mapping Agent.

Your ONLY responsibility is to:
1. Extract or infer technical data for a given machine.
2. Map that data EXACTLY into a predefined JSON structure.
3. Fill all required fields.

CRITICAL RULES:
- Output ONLY valid JSON — no prose, no markdown fences, no extra text.
- All fields must be filled (use realistic engineering estimates when real data
  is unavailable; mark Source as "Estimated" in those cases).
- Follow the column names EXACTLY as given below.
- Do NOT add new fields.
- Do NOT remove fields.

OUTPUT FORMAT (return this structure verbatim, populated with real values):
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

DATASHEET — capture every key electrical/mechanical property of the machine.
Typical rows: power_rating, voltage, frequency, efficiency,
temperature_range, weight, dimensions, ip_rating, insulation_class,
speed_rating, torque, noise_level, cooling_method, duty_cycle.

EBOM — list major physical components needed to assemble / maintain the machine.
Use Req ID format REQ-01, REQ-02 … for SRD rows.

CDD sections should cover:
  1. System Overview
  2. Architecture
  3. Components
  4. Working Principle
  5. Constraints
  6. Assumptions
  7. Future Improvements

Respond with ONLY the JSON object — nothing else.
"""


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
        "\nFill all four templates (DATASHEET, EBOM, SRD, CDD) for this machine "
        "and return ONLY the JSON object."
    )

    return "\n".join(parts)
