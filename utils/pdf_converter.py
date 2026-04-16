"""
PDF generator — converts structured JSON data (DATASHEET, EBOM, SRD, CDD)
into a single combined PDF using ReportLab.

Usage
-----
    from utils.pdf_converter import generate_pdf

    generate_pdf(data, "output/report.pdf", machine_name="Siemens Motor")
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

# ReportLab imports — kept lazy so the module can be imported even when
# reportlab is not installed (tests can still exercise non-PDF code paths).

__all__ = ["generate_pdf"]

# Colour palette
_DARK_BLUE = (0.122, 0.286, 0.490)   # #1F497D  — header background
_LIGHT_BLUE = (0.863, 0.902, 0.949)  # #DCE6F1  — alternating row
_WHITE = (1.0, 1.0, 1.0)
_BLACK = (0.0, 0.0, 0.0)
_TITLE_GREY = (0.3, 0.3, 0.3)

# Column layout: proportion of total table width per sheet
_COL_FRACS: dict[str, list[float]] = {
    "DATASHEET": [0.18, 0.14, 0.09, 0.35, 0.14],
    "EBOM":      [0.22, 0.07, 0.27, 0.15, 0.18, 0.11],
    "SRD":       [0.08, 0.42, 0.11, 0.09, 0.15, 0.15],
    "CDD":       [0.15, 0.22, 0.63],
}

_SHEET_ORDER = ["DATASHEET", "EBOM", "SRD", "CDD"]


def generate_pdf(
    data: dict[str, list[dict[str, Any]]],
    output_path: str | Path,
    machine_name: str = "",
) -> Path:
    """Generate a single combined PDF from *data* and save to *output_path*.

    Parameters
    ----------
    data:
        Dict with keys DATASHEET, EBOM, SRD, CDD — each a list of row dicts.
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

    page_width, page_height = landscape(A4)
    usable_width = page_width - 30 * mm  # 15 mm margin on each side

    doc = SimpleDocTemplate(
        str(output_path),
        pagesize=landscape(A4),
        leftMargin=15 * mm,
        rightMargin=15 * mm,
        topMargin=15 * mm,
        bottomMargin=15 * mm,
    )

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "DocTitle",
        parent=styles["Title"],
        fontSize=16,
        textColor=colors.Color(*_DARK_BLUE),
        spaceAfter=6,
        alignment=TA_CENTER,
    )
    sheet_title_style = ParagraphStyle(
        "SheetTitle",
        parent=styles["Heading1"],
        fontSize=13,
        textColor=colors.Color(*_DARK_BLUE),
        spaceBefore=10,
        spaceAfter=4,
        alignment=TA_LEFT,
    )
    cell_style = ParagraphStyle(
        "CellText",
        parent=styles["Normal"],
        fontSize=8,
        leading=10,
        wordWrap="CJK",
    )

    story: list = []

    # ---- Document title ------------------------------------------------
    title_text = f"Engineering Documentation Report"
    if machine_name:
        title_text += f"<br/><font size='11' color='grey'>{machine_name}</font>"
    story.append(Paragraph(title_text, title_style))
    story.append(Spacer(1, 6 * mm))

    # ---- One section per sheet ----------------------------------------
    for sheet_name in _SHEET_ORDER:
        rows = data.get(sheet_name, [])
        if not rows:
            continue

        story.append(Paragraph(sheet_name, sheet_title_style))

        columns = list(rows[0].keys())
        fracs = _COL_FRACS.get(sheet_name, [1.0 / len(columns)] * len(columns))

        # Normalize fracs so they sum to 1.0
        total = sum(fracs[: len(columns)])
        col_widths = [usable_width * (f / total) for f in fracs[: len(columns)]]

        # Build table data
        header_row = [
            Paragraph(f"<b>{col}</b>", cell_style) for col in columns
        ]
        table_data = [header_row]

        for row_dict in rows:
            table_data.append(
                [
                    Paragraph(str(row_dict.get(col, "")), cell_style)
                    for col in columns
                ]
            )

        # Build table style commands
        header_bg = colors.Color(*_DARK_BLUE)
        alt_bg = colors.Color(*_LIGHT_BLUE)
        white = colors.Color(*_WHITE)
        black = colors.Color(*_BLACK)

        style_commands = [
            ("BACKGROUND", (0, 0), (-1, 0), header_bg),
            ("TEXTCOLOR", (0, 0), (-1, 0), white),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, 0), 9),
            ("ALIGN", (0, 0), (-1, 0), "CENTER"),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("GRID", (0, 0), (-1, -1), 0.5, black),
            ("ROWBACKGROUND", (0, 1), (-1, -1), [white, alt_bg]),
            ("TOPPADDING", (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ("LEFTPADDING", (0, 0), (-1, -1), 4),
            ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ]

        table = Table(table_data, colWidths=col_widths, repeatRows=1)
        table.setStyle(TableStyle(style_commands))

        story.append(KeepTogether([table]))
        story.append(PageBreak())

    # Remove trailing PageBreak
    if story and isinstance(story[-1], PageBreak):
        story.pop()

    doc.build(story)
    return output_path
