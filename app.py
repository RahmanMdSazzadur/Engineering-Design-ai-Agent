"""
app.py — Flask web interface for the AI Datasheet-Filling Agent.

Run locally
-----------
    python app.py

Deploy online (Heroku / Railway / Render)
-----------------------------------------
    Set the environment variables from .env.example on your platform, then
    the included Procfile starts the app automatically via gunicorn.

Environment variables
---------------------
    PORT            HTTP port (default: 5000)
    LLM_PROVIDER    One of: ollama, google, deepseek, groq, openai
    <PROVIDER>_API_KEY  The key for the selected provider (e.g. GROQ_API_KEY)
"""

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
from utils.pdf_converter import generate_pdf

# ---------------------------------------------------------------------------
# App setup
# ---------------------------------------------------------------------------

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = Flask(__name__, template_folder="web_templates")

_TEMPLATE_PATH = Path(__file__).parent / "templates" / "Form.xlsx"
_OUTPUT_DIR = Path(__file__).parent / "output"

# Only permit filenames consisting of alphanumerics, underscores, hyphens and
# dots — no path separators or other special characters.
_SAFE_FILENAME_RE = re.compile(r"^[A-Za-z0-9_.%-]+$")


# ---------------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------------

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/generate", methods=["POST"])
def generate():
    """Accept form data, run the agent, return download URLs as JSON."""
    machine_name = (request.form.get("machine_name") or "").strip()
    task_type = (request.form.get("task_type") or "").strip() or None

    if not machine_name:
        return jsonify({"error": "Machine name is required."}), 400

    if not _TEMPLATE_PATH.exists():
        return jsonify({
            "error": "Template file templates/Form.xlsx not found on the server.",
        }), 500

    try:
        extractor = DataExtractor()
        data = extractor.extract(machine_name=machine_name, task_type=task_type)
    except Exception:
        logger.exception("LLM extraction failed for machine %r", machine_name)
        return jsonify({"error": "LLM extraction failed. Check server logs for details."}), 500

    # Unique job suffix to avoid filename collisions between concurrent users.
    job_id = uuid.uuid4().hex[:8]
    machine_slug = re.sub(r"[^A-Za-z0-9_-]", "_", machine_name)[:40]
    base_name = f"{machine_slug}_{job_id}"

    _OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    xlsx_path = _OUTPUT_DIR / f"{base_name}_filled.xlsx"
    pdf_path = _OUTPUT_DIR / f"{base_name}_report.pdf"

    try:
        fill_template(data, _TEMPLATE_PATH, xlsx_path)
    except Exception:
        logger.exception("Excel generation failed")
        return jsonify({"error": "Excel generation failed. Check server logs for details."}), 500

    try:
        generate_pdf(data, pdf_path, machine_name=machine_name)
    except Exception:
        logger.exception("PDF generation failed")
        return jsonify({"error": "PDF generation failed. Check server logs for details."}), 500

    return jsonify({
        "success": True,
        "xlsx_file": xlsx_path.name,
        "pdf_file": pdf_path.name,
        "machine_name": machine_name,
    })


@app.route("/download/<path:filename>")
def download(filename: str):
    """Serve a generated file from the output directory."""
    # Take only the basename — strip any directory components supplied by the
    # caller before doing anything else.
    safe_name = Path(filename).name

    # Reject names that contain characters outside our explicit allowlist.
    if not _SAFE_FILENAME_RE.match(safe_name):
        abort(400)

    requested = (_OUTPUT_DIR / safe_name).resolve()

    # Verify the resolved path is strictly inside _OUTPUT_DIR.
    try:
        requested.relative_to(_OUTPUT_DIR.resolve())
    except ValueError:
        abort(400)

    if not requested.is_file():
        abort(404)

    return send_file(requested, as_attachment=True)


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    port = int(os.getenv("PORT", "5000"))
    app.run(host="0.0.0.0", port=port, debug=False)
