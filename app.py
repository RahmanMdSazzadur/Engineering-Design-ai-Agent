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
        print("REAL ERROR:", str(e))   # 👈 IMPORTANT
        return jsonify({"error": str(e)}), 500   # 👈 SHOW REAL ERROR

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
        print("EXCEL ERROR:", str(e))
        return jsonify({"error": str(e)}), 500

    try:
        generate_pdf(data, pdf_path, machine_name=machine_name, forms=forms, image_path=image_path)
    except Exception as e:
        print("PDF ERROR:", str(e))
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
