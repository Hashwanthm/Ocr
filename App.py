from pathlib import Path
import re
import uuid

import pdfplumber
from flask import Flask, render_template, request, jsonify, send_from_directory
from werkzeug.utils import secure_filename

from openpyxl import Workbook
from openpyxl.styles import Border, Side, Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


app = Flask(__name__)

BASE_DIR = Path(__file__).resolve().parent
UPLOAD_DIR = BASE_DIR / "uploads"
OUTPUT_DIR = BASE_DIR / "outputs"

UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)


# These lines will be skipped so logo/title text is not added to Excel
SKIP_TEXTS = [
    "EMPLOYEES' PROVIDENT FUND ORGANISATION",
    "कर्मचारी भविष्य निधि संगठन",
    "Ministry of Labour",
    "EPFO",
]


def clean_text(value):
    if value is None:
        return ""

    value = str(value)
    value = value.replace("\n", " ")
    value = value.replace("\xa0", " ")
    value = re.sub(r"\s+", " ", value)
    return value.strip()


def should_skip_line(text):
    text_lower = text.lower()

    for skip in SKIP_TEXTS:
        if skip.lower() in text_lower:
            return True

    return False


def extract_tables(page):
    """
    First tries to extract tables using PDF lines.
    If that fails, tries text-based extraction.
    """

    line_settings = {
        "vertical_strategy": "lines",
        "horizontal_strategy": "lines",
        "snap_tolerance": 3,
        "join_tolerance": 3,
        "intersection_tolerance": 5,
        "text_x_tolerance": 2,
        "text_y_tolerance": 3,
    }

    text_settings = {
        "vertical_strategy": "text",
        "horizontal_strategy": "text",
        "snap_tolerance": 3,
        "join_tolerance": 3,
        "intersection_tolerance": 5,
        "text_x_tolerance": 2,
        "text_y_tolerance": 3,
    }

    tables = page.extract_tables(table_settings=line_settings)

    if not tables:
        tables = page.extract_tables(table_settings=text_settings)

    final_tables = []

    for table in tables:
        cleaned_table = []

        for row in table:
            cleaned_row = [clean_text(cell) for cell in row]

            if any(cleaned_row):
                cleaned_table.append(cleaned_row)

        if cleaned_table:
            final_tables.append(cleaned_table)

    return final_tables


def extract_pdf_to_rows(pdf_path):
    """
    Converts PDF text/tables into rows.
    These rows are used for Excel and right-side preview.
    """

    final_rows = []

    with pdfplumber.open(pdf_path) as pdf:
        for page_number, page in enumerate(pdf.pages, start=1):
            tables = extract_tables(page)

            if tables:
                for table_number, table in enumerate(tables, start=1):
                    final_rows.append([f"Page {page_number} - Table {table_number}"])

                    for row in table:
                        cleaned_row = []

                        for cell in row:
                            text = clean_text(cell)

                            if should_skip_line(text):
                                text = ""

                            cleaned_row.append(text)

                        if any(cleaned_row):
                            final_rows.append(cleaned_row)

                    final_rows.append([""])

            else:
                text = page.extract_text(x_tolerance=2, y_tolerance=3) or ""

                final_rows.append([f"Page {page_number} - Text"])

                for line in text.splitlines():
                    line = clean_text(line)

                    if not line:
                        continue

                    if should_skip_line(line):
                        continue

                    final_rows.append([line])

                final_rows.append([""])

    return final_rows


def create_excel(rows, excel_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "PDF Extract"

    thin = Side(style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    header_fill = PatternFill("solid", fgColor="D9EAF7")
    section_fill = PatternFill("solid", fgColor="B7DEE8")

    if not rows:
        rows = [["No data extracted"]]

    max_cols = max(len(row) for row in rows)

    for row_index, row_data in enumerate(rows, start=1):
        for col_index in range(1, max_cols + 1):
            value = ""

            if col_index <= len(row_data):
                value = row_data[col_index - 1]

            cell = ws.cell(row=row_index, column=col_index, value=value)
            cell.border = border
            cell.alignment = Alignment(
                horizontal="left",
                vertical="top",
                wrap_text=True,
            )

            if len(row_data) == 1 and str(value).startswith("Page "):
                cell.font = Font(bold=True)
                cell.fill = section_fill

    for col_index in range(1, max_cols + 1):
        col_letter = get_column_letter(col_index)
        max_length = 12

        for row_index in range(1, ws.max_row + 1):
            value = ws.cell(row=row_index, column=col_index).value

            if value:
                max_length = max(max_length, len(str(value)))

        ws.column_dimensions[col_letter].width = min(max_length + 3, 45)

    ws.freeze_panes = "A2"
    wb.save(excel_path)


@app.route("/")
def home():
    return render_template("index.html")


@app.route("/convert", methods=["POST"])
def convert_pdf():
    if "pdf" not in request.files:
        return jsonify({"error": "No PDF file received"}), 400

    file = request.files["pdf"]

    if file.filename == "":
        return jsonify({"error": "Please select a PDF file"}), 400

    if not file.filename.lower().endswith(".pdf"):
        return jsonify({"error": "Only PDF files are allowed"}), 400

    safe_name = secure_filename(file.filename)
    unique_id = uuid.uuid4().hex

    uploaded_pdf_name = f"{unique_id}_{safe_name}"
    uploaded_pdf_path = UPLOAD_DIR / uploaded_pdf_name

    file.save(uploaded_pdf_path)

    try:
        rows = extract_pdf_to_rows(uploaded_pdf_path)

        if not rows:
            return jsonify({
                "error": "No text found. This PDF may be scanned/image-based and may need OCR."
            }), 400

        excel_name = f"{Path(safe_name).stem}_converted_{unique_id}.xlsx"
        excel_path = OUTPUT_DIR / excel_name

        create_excel(rows, excel_path)

        return jsonify({
            "message": "PDF converted successfully",
            "rows": rows,
            "download_url": f"/download/{excel_name}"
        })

    except Exception as error:
        return jsonify({"error": str(error)}), 500


@app.route("/download/<filename>")
def download_excel(filename):
    return send_from_directory(
        OUTPUT_DIR,
        filename,
        as_attachment=True
    )


if __name__ == "__main__":
    app.run(debug=True)
