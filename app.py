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

import os
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

app = Flask(__name__, template_folder="web_templates")

_TEMPLATE_PATH = Path(__file__).parent / "templates" / "Form.xlsx"
_OUTPUT_DIR = Path(__file__).parent / "output"


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
    except Exception as exc:
        return jsonify({"error": f"LLM extraction failed: {exc}"}), 500

    # Unique job suffix to avoid filename collisions between concurrent users.
    job_id = uuid.uuid4().hex[:8]
    machine_slug = machine_name.replace(" ", "_").replace("/", "-")[:40]
    base_name = f"{machine_slug}_{job_id}"

    _OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    xlsx_path = _OUTPUT_DIR / f"{base_name}_filled.xlsx"
    pdf_path = _OUTPUT_DIR / f"{base_name}_report.pdf"

    try:
        fill_template(data, _TEMPLATE_PATH, xlsx_path)
    except Exception as exc:
        return jsonify({"error": f"Excel generation failed: {exc}"}), 500

    try:
        generate_pdf(data, pdf_path, machine_name=machine_name)
    except Exception as exc:
        return jsonify({"error": f"PDF generation failed: {exc}"}), 500

    return jsonify({
        "success": True,
        "xlsx_file": xlsx_path.name,
        "pdf_file": pdf_path.name,
        "machine_name": machine_name,
    })


@app.route("/download/<path:filename>")
def download(filename: str):
    """Serve a generated file from the output directory."""
    # Resolve and verify the path is strictly inside _OUTPUT_DIR to prevent
    # directory traversal attacks.
    requested = (_OUTPUT_DIR / Path(filename).name).resolve()
    if not str(requested).startswith(str(_OUTPUT_DIR.resolve())):
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
