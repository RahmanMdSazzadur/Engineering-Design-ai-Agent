"""
Excel handler utilities — create templates, fill them with data, and read
reference EBOM files.

Functions
---------
create_template(path)
    Write a blank four-sheet XLSX template to *path*.

fill_template(data, template_path, output_path)
    Copy the template and fill it with the provided JSON data dict.

read_template_structure(path)
    Return the column headers for every sheet in the template.

read_ebom_reference(path, sheet_name)
    Return EBOM rows from an existing XLSX file as a list of dicts.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# Template column definitions (single source of truth)
# ---------------------------------------------------------------------------

SHEET_COLUMNS: dict[str, list[str]] = {
    "DATASHEET": [
        "Parameter",
        "Value",
        "Unit",
        "Description",
        "Source",
    ],
    "EBOM": [
        "Component Name",
        "Quantity",
        "Specification",
        "Material",
        "Supplier",
        "Notes",
    ],
    "SRD": [
        "Req ID",
        "Requirement Description",
        "Type",
        "Priority",
        "Source",
        "Validation Method",
    ],
    "CDD": [
        "Section",
        "Title",
        "Description",
    ],
}

# Header row styling
_HEADER_FILL = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
_HEADER_FONT = Font(color="FFFFFF", bold=True, name="Calibri", size=11)
_CELL_FONT = Font(name="Calibri", size=10)
_ALT_FILL = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")

# Approximate column widths (characters) per sheet
_COL_WIDTHS: dict[str, list[int]] = {
    "DATASHEET": [25, 20, 12, 45, 18],
    "EBOM": [30, 10, 35, 20, 25, 30],
    "SRD": [12, 60, 15, 12, 20, 25],
    "CDD": [20, 30, 80],
}


# ---------------------------------------------------------------------------
# Public helpers
# ---------------------------------------------------------------------------


def create_template(path: str | Path) -> Path:
    """Create a blank four-sheet XLSX template at *path*.

    If the file already exists it is overwritten.

    Parameters
    ----------
    path:
        Destination file path.

    Returns
    -------
    Path
        Resolved path of the created file.
    """
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)

    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # remove default "Sheet"

    for sheet_name, columns in SHEET_COLUMNS.items():
        ws = wb.create_sheet(title=sheet_name)
        _write_header_row(ws, columns, _COL_WIDTHS.get(sheet_name, []))

    wb.save(path)
    return path


def fill_template(
    data: dict[str, list[dict[str, Any]]],
    template_path: str | Path,
    output_path: str | Path,
) -> Path:
    """Fill *template_path* with *data* and save to *output_path*.

    Parameters
    ----------
    data:
        Dict with keys DATASHEET, EBOM, SRD, CDD — each a list of row dicts.
    template_path:
        Path to the blank template XLSX.
    output_path:
        Destination path for the filled XLSX.

    Returns
    -------
    Path
        Resolved path of the written file.
    """
    template_path = Path(template_path)
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    wb = openpyxl.load_workbook(template_path)

    for sheet_name, columns in SHEET_COLUMNS.items():
        rows = data.get(sheet_name, [])
        if sheet_name not in wb.sheetnames:
            ws = wb.create_sheet(title=sheet_name)
            _write_header_row(ws, columns, _COL_WIDTHS.get(sheet_name, []))
        else:
            ws = wb[sheet_name]

        for row_idx, row_data in enumerate(rows, start=2):
            is_even_row = row_idx % 2 == 0
            for col_idx, col_name in enumerate(columns, start=1):
                cell = ws.cell(row=row_idx, column=col_idx)
                cell.value = str(row_data.get(col_name, ""))
                cell.font = _CELL_FONT
                cell.alignment = Alignment(wrap_text=True, vertical="top")
                if is_even_row:
                    cell.fill = _ALT_FILL

    wb.save(output_path)
    return output_path


def read_template_structure(path: str | Path) -> dict[str, list[str]]:
    """Return the header columns for every sheet in the template XLSX.

    Parameters
    ----------
    path:
        Path to the template XLSX.

    Returns
    -------
    dict
        Mapping of sheet name → list of column header strings.
    """
    path = Path(path)
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    structure: dict[str, list[str]] = {}
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        headers: list[str] = []
        for row in ws.iter_rows(min_row=1, max_row=1, values_only=True):
            headers = [str(cell) for cell in row if cell is not None]
        structure[sheet_name] = headers
    wb.close()
    return structure


def read_ebom_reference(
    path: str | Path,
    sheet_name: str = "EBOM",
) -> list[dict[str, Any]]:
    """Read EBOM reference data from an XLSX file.

    Parameters
    ----------
    path:
        Path to the XLSX file containing EBOM reference data.
    sheet_name:
        Name of the sheet to read (default: "EBOM").

    Returns
    -------
    list of dict
        One dict per data row, keyed by column header.
    """
    path = Path(path)
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)

    if sheet_name not in wb.sheetnames:
        wb.close()
        return []

    ws = wb[sheet_name]
    rows_iter = ws.iter_rows(values_only=True)

    # First row = headers
    try:
        headers = [str(h) if h is not None else "" for h in next(rows_iter)]
    except StopIteration:
        wb.close()
        return []

    records: list[dict[str, Any]] = []
    for row in rows_iter:
        record = {headers[i]: (row[i] if i < len(row) else None)
                  for i in range(len(headers))}
        records.append(record)

    wb.close()
    return records


# ---------------------------------------------------------------------------
# Private helpers
# ---------------------------------------------------------------------------


def _write_header_row(
    ws: openpyxl.worksheet.worksheet.Worksheet,
    columns: list[str],
    col_widths: list[int],
) -> None:
    """Write a styled header row to *ws*."""
    for col_idx, col_name in enumerate(columns, start=1):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        cell.fill = _HEADER_FILL
        cell.font = _HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")

        # Set column width
        if col_idx <= len(col_widths):
            ws.column_dimensions[get_column_letter(col_idx)].width = col_widths[col_idx - 1]
        else:
            ws.column_dimensions[get_column_letter(col_idx)].width = 20

    ws.row_dimensions[1].height = 20
