"""
PDF generator — converts structured JSON data (Datasheet dict + EBOM/SRD/CDD
lists) into a single combined PDF using ReportLab.

The output mirrors the Form.xlsx layout:
  1. Datasheet — rendered as a structured form with a parameter table
  2. EBOM      — tabular
  3. SRD       — tabular
  4. CDD       — tabular

Usage
-----
    from utils.pdf_converter import generate_pdf

    generate_pdf(data, "output/report.pdf", machine_name="Siemens Motor")
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

__all__ = ["generate_pdf"]

# ---------------------------------------------------------------------------
# Color palette  (R, G, B — 0..1)
# ---------------------------------------------------------------------------
_DARK_BLUE  = (0.122, 0.286, 0.490)   # #1F497D
_LIGHT_BLUE = (0.863, 0.902, 0.949)   # #DCE6F1
_WHITE      = (1.0, 1.0, 1.0)
_BLACK      = (0.0, 0.0, 0.0)
_LIGHT_GREY = (0.95, 0.95, 0.95)

# Column-width fractions for tabular sheets (must match EBOM column order)
_COL_FRACS: dict[str, list[float]] = {
    "EBOM": [0.07, 0.10, 0.08, 0.10, 0.12, 0.08, 0.09, 0.15,
             0.05, 0.05, 0.05, 0.05, 0.05, 0.05, 0.06],
    "SRD":  [0.08, 0.05, 0.60, 0.27],
    "CDD":  [0.08, 0.05, 0.87],
}

_DATASHEET_PARAM_FRACS = [0.32, 0.10, 0.18, 0.22, 0.18]


def generate_pdf(
    data: dict[str, Any],
    output_path: str | Path,
    machine_name: str = "",
) -> Path:
    """Generate a single combined PDF from *data* and save to *output_path*.

    Parameters
    ----------
    data:
        Dict with keys ``Datasheet`` (dict), ``EBOM``, ``SRD``, ``CDD``
        (lists of dicts).
    output_path:
        Destination file path for the PDF.
    machine_name:
        Optional machine name shown in the document header.

    Returns
    -------
    Path
        Resolved path of the created PDF.
    """
    try:
        from reportlab.lib import colors
        from reportlab.lib.pagesizes import A4, landscape
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        from reportlab.lib.units import mm
        from reportlab.platypus import (
            SimpleDocTemplate,
            Table,
            TableStyle,
            Paragraph,
            Spacer,
            PageBreak,
            KeepTogether,
        )
        from reportlab.lib.enums import TA_CENTER, TA_LEFT
    except ImportError as exc:
        raise RuntimeError(
            "The 'reportlab' package is required for PDF generation. "
            "Install it with: pip install reportlab"
        ) from exc

    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    page_w, _page_h = landscape(A4)
    usable_w = page_w - 30 * mm

    doc = SimpleDocTemplate(
        str(output_path),
        pagesize=landscape(A4),
        leftMargin=15 * mm, rightMargin=15 * mm,
        topMargin=15 * mm,  bottomMargin=15 * mm,
    )

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "DocTitle", parent=styles["Title"], fontSize=16,
        textColor=colors.Color(*_DARK_BLUE), spaceAfter=4, alignment=TA_CENTER,
    )
    section_style = ParagraphStyle(
        "Section", parent=styles["Heading1"], fontSize=12,
        textColor=colors.Color(*_DARK_BLUE), spaceBefore=10, spaceAfter=4,
        alignment=TA_LEFT,
    )
    sub_style = ParagraphStyle(
        "Sub", parent=styles["Heading2"], fontSize=10,
        textColor=colors.Color(*_DARK_BLUE), spaceBefore=6, spaceAfter=2,
        alignment=TA_LEFT,
    )
    cell_style = ParagraphStyle(
        "Cell", parent=styles["Normal"], fontSize=8, leading=10, wordWrap="CJK",
    )
    label_style = ParagraphStyle(
        "Label", parent=styles["Normal"], fontSize=8, leading=10,
        textColor=colors.Color(0.3, 0.3, 0.3),
    )
    body_style = ParagraphStyle(
        "Body", parent=styles["Normal"], fontSize=9, leading=12,
    )

    dark_blue_c  = colors.Color(*_DARK_BLUE)
    light_blue_c = colors.Color(*_LIGHT_BLUE)
    white_c      = colors.Color(*_WHITE)
    black_c      = colors.Color(*_BLACK)
    light_grey_c = colors.Color(*_LIGHT_GREY)

    story: list = []

    # ------------------------------------------------------------------ #
    # Document title
    # ------------------------------------------------------------------ #
    header = "Engineering Documentation Report"
    if machine_name:
        header += f"<br/><font size='11' color='grey'>{machine_name}</font>"
    story.append(Paragraph(header, title_style))
    story.append(Spacer(1, 6 * mm))

    # ------------------------------------------------------------------ #
    # 1. DATASHEET section
    # ------------------------------------------------------------------ #
    ds = data.get("Datasheet", {})
    if ds:
        story.append(Paragraph("Technical Datasheet", section_style))

        def kv_row(label: str, value: str) -> list:
            return [
                Paragraph(label, label_style),
                Paragraph(str(value), body_style),
            ]

        # Catalogue info table
        cat_data = [
            kv_row("Author",    ds.get("Author", "")),
            kv_row("Item Name", ds.get("Item Name", "")),
            kv_row("HEL",       ds.get("HEL", "")),
        ]
        cat_table = Table(cat_data, colWidths=[usable_w * 0.20, usable_w * 0.80])
        cat_table.setStyle(TableStyle([
            ("VALIGN",       (0, 0), (-1, -1), "TOP"),
            ("TEXTCOLOR",    (0, 0), (0, -1),  dark_blue_c),
            ("FONTNAME",     (0, 0), (0, -1),  "Helvetica-Bold"),
            ("FONTSIZE",     (0, 0), (-1, -1), 8),
            ("TOPPADDING",   (0, 0), (-1, -1), 3),
            ("BOTTOMPADDING",(0, 0), (-1, -1), 3),
            ("ROWBACKGROUND",(0, 0), (-1, -1), [white_c, light_grey_c]),
        ]))
        story.append(cat_table)
        story.append(Spacer(1, 3 * mm))

        # System description
        story.append(Paragraph("System Description", sub_style))
        story.append(Paragraph(ds.get("System Description", ""), body_style))
        story.append(Spacer(1, 3 * mm))

        # Parameter table header
        story.append(Paragraph("Key Design Information", sub_style))
        param_header = [
            Paragraph("<b>Parameter</b>", cell_style),
            Paragraph("<b>Unit</b>",      cell_style),
            Paragraph("<b>Value</b>",     cell_style),
            Paragraph("<b>Reference</b>", cell_style),
            Paragraph("<b>Notes</b>",     cell_style),
        ]
        col_widths_p = [usable_w * f for f in _DATASHEET_PARAM_FRACS]
        param_rows = [param_header]

        dim_params   = ds.get("Dimensional Parameters", [])
        other_params = ds.get("Other Parameters", [])
        all_params   = dim_params + other_params

        for p in all_params:
            param_rows.append([
                Paragraph(str(p.get("Parameter", "")), cell_style),
                Paragraph(str(p.get("Unit", "")),      cell_style),
                Paragraph(str(p.get("Value", "")),     cell_style),
                Paragraph(str(p.get("Reference", "")), cell_style),
                Paragraph(str(p.get("Notes", "")),     cell_style),
            ])

        param_table = Table(param_rows, colWidths=col_widths_p, repeatRows=1)
        param_table.setStyle(TableStyle([
            ("BACKGROUND",   (0, 0), (-1, 0), dark_blue_c),
            ("TEXTCOLOR",    (0, 0), (-1, 0), white_c),
            ("FONTNAME",     (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE",     (0, 0), (-1, 0), 8),
            ("ALIGN",        (0, 0), (-1, 0), "CENTER"),
            ("VALIGN",       (0, 0), (-1, -1), "TOP"),
            ("GRID",         (0, 0), (-1, -1), 0.4, black_c),
            ("ROWBACKGROUND",(0, 1), (-1, -1), [white_c, light_blue_c]),
            ("TOPPADDING",   (0, 0), (-1, -1), 3),
            ("BOTTOMPADDING",(0, 0), (-1, -1), 3),
            ("LEFTPADDING",  (0, 0), (-1, -1), 3),
            ("RIGHTPADDING", (0, 0), (-1, -1), 3),
        ]))
        story.append(param_table)
        story.append(Spacer(1, 3 * mm))

        # Manufacturer info
        story.append(Paragraph("Manufacturer Information", sub_style))
        mfr_data = [
            kv_row("Manufacturer", ds.get("Manufacturer", "")),
            kv_row("Model",        ds.get("Model", "")),
            kv_row("Website",      ds.get("Website", "")),
        ]
        mfr_table = Table(mfr_data, colWidths=[usable_w * 0.20, usable_w * 0.80])
        mfr_table.setStyle(TableStyle([
            ("VALIGN",       (0, 0), (-1, -1), "TOP"),
            ("TEXTCOLOR",    (0, 0), (0, -1),  dark_blue_c),
            ("FONTNAME",     (0, 0), (0, -1),  "Helvetica-Bold"),
            ("FONTSIZE",     (0, 0), (-1, -1), 8),
            ("TOPPADDING",   (0, 0), (-1, -1), 3),
            ("BOTTOMPADDING",(0, 0), (-1, -1), 3),
            ("ROWBACKGROUND",(0, 0), (-1, -1), [white_c, light_grey_c]),
        ]))
        story.append(mfr_table)
        story.append(Spacer(1, 3 * mm))

        if ds.get("Notes"):
            story.append(Paragraph("Notes", sub_style))
            story.append(Paragraph(ds.get("Notes", ""), body_style))
            story.append(Spacer(1, 2 * mm))

        if ds.get("References"):
            story.append(Paragraph("References", sub_style))
            story.append(Paragraph(ds.get("References", ""), body_style))

        story.append(PageBreak())

    # ------------------------------------------------------------------ #
    # 2-4. Tabular sheets: EBOM, SRD, CDD
    # ------------------------------------------------------------------ #
    tabular_sections = [
        ("EBOM", "Engineering Bill of Materials (EBOM)"),
        ("SRD",  "System Requirements Document (SRD)"),
        ("CDD",  "Context Definition Document (CDD)"),
    ]

    for sheet_key, sheet_title in tabular_sections:
        rows = data.get(sheet_key, [])
        if not rows:
            continue

        story.append(Paragraph(sheet_title, section_style))

        columns = list(rows[0].keys())
        fracs = _COL_FRACS.get(sheet_key, [1.0 / len(columns)] * len(columns))
        total = sum(fracs[: len(columns)])
        col_widths_t = [usable_w * (f / total) for f in fracs[: len(columns)]]

        header_row = [Paragraph(f"<b>{c}</b>", cell_style) for c in columns]
        table_data  = [header_row]
        for row_dict in rows:
            table_data.append(
                [Paragraph(str(row_dict.get(c, "")), cell_style) for c in columns]
            )

        tbl = Table(table_data, colWidths=col_widths_t, repeatRows=1)
        tbl.setStyle(TableStyle([
            ("BACKGROUND",   (0, 0), (-1, 0), dark_blue_c),
            ("TEXTCOLOR",    (0, 0), (-1, 0), white_c),
            ("FONTNAME",     (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE",     (0, 0), (-1, 0), 8),
            ("ALIGN",        (0, 0), (-1, 0), "CENTER"),
            ("VALIGN",       (0, 0), (-1, -1), "TOP"),
            ("GRID",         (0, 0), (-1, -1), 0.4, black_c),
            ("ROWBACKGROUND",(0, 1), (-1, -1), [white_c, light_blue_c]),
            ("TOPPADDING",   (0, 0), (-1, -1), 3),
            ("BOTTOMPADDING",(0, 0), (-1, -1), 3),
            ("LEFTPADDING",  (0, 0), (-1, -1), 3),
            ("RIGHTPADDING", (0, 0), (-1, -1), 3),
        ]))

        story.append(KeepTogether([tbl]))
        story.append(PageBreak())

    # Remove the trailing PageBreak if present.
    if story and isinstance(story[-1], PageBreak):
        story.pop()

    doc.build(story)
    return output_path
