from flask import Flask, render_template, request, send_file
import os
from extractor import (
    extract_page_text,
    process_pages,
    export_to_excel,
    EOB_TYPES
)

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "output"
OUTPUT_FILENAME = "EOB_Extracted.xlsx"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)


@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        eob_type = request.form.get("eob_type")
        file = request.files.get("pdf")

        if not file or file.filename == "":
            return render_template(
                "index.html",
                eob_types=EOB_TYPES,
                error="Please upload a PDF"
            )

        pdf_path = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(pdf_path)

        pages = extract_page_text(pdf_path)
        records = process_pages(pages, eob_type)

        output_file = os.path.join(OUTPUT_FOLDER, OUTPUT_FILENAME)
        export_to_excel(records, output_file)

        return render_template(
            "index.html",
            eob_types=EOB_TYPES,
            success=True,
            download_ready=True
        )

    # GET request
    return render_template(
        "index.html",
        eob_types=EOB_TYPES
    )


@app.route("/download")
def download():
    output_file = os.path.join(OUTPUT_FOLDER, OUTPUT_FILENAME)

    if not os.path.exists(output_file):
        return "No file available for download. Please process a PDF first.", 404

    return send_file(
        output_file,
        as_attachment=True,
        download_name=OUTPUT_FILENAME
    )


if __name__ == "__main__":
    app.run(debug=True)
