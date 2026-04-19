from __future__ import annotations

import logging
import os
import re
import uuid
from pathlib import Path

from flask import Flask, abort, jsonify, render_template, request, send_file

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

from agent.extractor import DataExtractor
from utils.excel_handler import fill_template
from utils.image_fetcher import fetch_machine_image
from utils.pdf_converter import generate_pdf


logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = Flask(__name__, template_folder="web_templates")


_TEMPLATE_PATH = Path(__file__).parent / "templates" / "Form.xlsx"
_OUTPUT_DIR = Path(__file__).parent / "output"

_download_registry: dict[str, tuple[Path, str]] = {}
_TOKEN_RE = re.compile(r"^[0-9a-f]{32}$")


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/generate", methods=["POST"])
def generate():
    machine_name = (request.form.get("machine_name") or "").strip()
    task_type = (request.form.get("task_type") or "").strip() or None
    forms_raw = (request.form.get("forms") or "all").strip()

    _ALL_FORMS = ["Datasheet", "EBOM", "SRD", "CDD"]
    if forms_raw == "all" or forms_raw not in _ALL_FORMS:
        forms = _ALL_FORMS
    else:
        forms = [forms_raw]

    if not machine_name:
        return jsonify({"error": "Machine name is required."}), 400

    if not _TEMPLATE_PATH.exists():
        return jsonify({
            "error": "Template file templates/Form.xlsx not found on the server.",
        }), 500

    try:
        extractor = DataExtractor()
        data = extractor.extract(machine_name=machine_name, task_type=task_type)
    except Exception as e:
        logger.exception("LLM extraction failed")
        return jsonify({"error": str(e)}), 500

    image_path = fetch_machine_image(machine_name)

    job_id = uuid.uuid4().hex[:8]
    machine_slug = re.sub(r"[^A-Za-z0-9_-]", "_", machine_name)[:40]
    base_name = f"{machine_slug}_{job_id}"

    _OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    xlsx_path = _OUTPUT_DIR / f"{base_name}_filled.xlsx"
    pdf_path = _OUTPUT_DIR / f"{base_name}_report.pdf"

    try:
        fill_template(data, _TEMPLATE_PATH, xlsx_path, forms=forms, image_path=image_path)
    except Exception as e:
        logger.exception("Excel generation failed")
        return jsonify({"error": str(e)}), 500

    try:
        generate_pdf(data, pdf_path, machine_name=machine_name, forms=forms, image_path=image_path)
    except Exception as e:
        logger.exception("PDF generation failed")
        return jsonify({"error": str(e)}), 500

    xlsx_token = uuid.uuid4().hex
    pdf_token = uuid.uuid4().hex
    _download_registry[xlsx_token] = (xlsx_path, f"{base_name}_filled.xlsx")
    _download_registry[pdf_token] = (pdf_path, f"{base_name}_report.pdf")

    return jsonify({
        "success": True,
        "xlsx_token": xlsx_token,
        "pdf_token": pdf_token,
        "machine_name": machine_name,
    })


@app.route("/download/<token>")
def download(token: str):
    if not _TOKEN_RE.match(token):
        abort(400)

    entry = _download_registry.get(token)
    if entry is None:
        abort(404)

    file_path, download_name = entry
    if not file_path.is_file():
        abort(404)

    return send_file(file_path, as_attachment=True, download_name=download_name)


if __name__ == "__main__":
    port = int(os.getenv("PORT", "5000"))
    app.run(host="0.0.0.0", port=port, debug=False)
