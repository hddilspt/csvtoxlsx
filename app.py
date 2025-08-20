import os
import io
import tempfile
import subprocess
import shutil

from flask import Flask, request, send_file
from werkzeug.utils import secure_filename
import pandas as pd

app = Flask(__name__)

def force_landscape_xlsx(xlsx_path: str) -> None:
    """
    Sets all sheets to landscape and fits to 1 page wide (unlimited height).
    Modifies the .xlsx file in place.
    """
    from openpyxl import load_workbook

    wb = load_workbook(xlsx_path)
    for ws in wb.worksheets:
        ws.page_setup.orientation = "landscape"
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0  # don't force height
        ws.print_options.horizontalCentered = True
    wb.save(xlsx_path)


@app.route('/', methods=['GET'])
def home():
    return "CSV→XLSX and XLSX→PDF API is running!"


@app.route('/convert', methods=['POST'])
def convert():
    try:
        # Prefer standard multipart 'file'; otherwise accept raw body (Power Automate HTTP).
        uploaded = request.files.get('file')
        if uploaded and uploaded.filename:
            filename = secure_filename(uploaded.filename)
            data = uploaded.read()
            mode = 'multipart'
        else:
            data = request.get_data() or b""
            if not data:
                return {"error": "No file part in the request"}, 400
            # filename can be provided via query string or header; both are optional
            filename = request.args.get('filename') or request.headers.get('X-File-Name') or 'upload'
            filename = secure_filename(filename)
            mode = 'raw'

        # Basic type detection
        ext = os.path.splitext(filename)[1].lower()
        head = data[:8]
        is_xlsx = head.startswith(b'PK')  # XLSX is a ZIP (starts with PK)

        # --- CSV -> XLSX ---
        if (ext == '.csv') and not is_xlsx:
            df = pd.read_csv(io.BytesIO(data))
            out = io.BytesIO()
            with pd.ExcelWriter(out, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name="Sheet1")
            out.seek(0)
            resp = send_file(
                out,
                mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                as_attachment=True,
                download_name="converted.xlsx",
            )
            resp.headers['X-Debug-Path'] = f'{mode}:csv->xlsx'
            return resp

        # --- XLSX -> PDF (ignore .xls by request) ---
        if is_xlsx or ext == '.xlsx':
            if shutil.which("soffice") is None:
                return {"error": "LibreOffice (soffice) not found in PATH."}, 500

            with tempfile.TemporaryDirectory() as tmpdir:
                in_path = os.path.join(tmpdir, 'upload.xlsx')
                with open(in_path, 'wb') as fh:
                    fh.write(data)

                # Force landscape before exporting
                try:
                    force_landscape_xlsx(in_path)
                except Exception:
                    # If this fails for any reason, still proceed with conversion.
                    pass

                # Convert using LibreOffice headless
                cmd = [
                    "soffice",
                    "--headless",
                    "--convert-to", "pdf",
                    "--outdir", tmpdir,
                    in_path
                ]
                proc = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                if proc.returncode != 0:
                    return {
                        "error": "PDF conversion failed.",
                        "stdout": proc.stdout.decode('utf-8', 'ignore')[:800],
                        "stderr": proc.stderr.decode('utf-8', 'ignore')[:800],
                    }, 500

                pdf_path = os.path.splitext(in_path)[0] + ".pdf"
                if not os.path.exists(pdf_path):
                    candidates = [f for f in os.listdir(tmpdir) if f.lower().endswith(".pdf")]
                    if not candidates:
                        return {"error": "Converted PDF not found."}, 500
                    pdf_path = os.path.join(tmpdir, candidates[0])

                with open(pdf_path, "rb") as f:
                    buf = io.BytesIO(f.read())
                buf.seek(0)

                out_name = os.path.splitext(os.path.basename(filename))[0] + ".pdf"
                resp = send_file(
                    buf,
                    mimetype="application/pdf",
                    as_attachment=True,
                    download_name=out_name
                )
                resp.headers['X-Debug-Path'] = f'{mode}:xlsx->pdf'
                return resp

        return {"error": "Unsupported file type. Upload .csv or .xlsx."}, 400

    except Exception as e:
        return {"error": str(e)}, 500


if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)
