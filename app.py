import os
import io
import tempfile
import subprocess
import shutil

from flask import Flask, request, send_file
from werkzeug.utils import secure_filename
import pandas as pd

app = Flask(__name__)

@app.route('/', methods=['GET'])
def home():
    return "CSV→XLSX and XLSX→PDF API is running!"

@app.route('/convert', methods=['POST'])
def convert():
    try:
        if 'file' not in request.files:
            return {"error": "No file part in the request"}, 400

        f = request.files['file']
        if not f or f.filename == '':
            return {"error": "No selected file"}, 400

        ext = os.path.splitext(f.filename)[1].lower()

        # CSV -> XLSX (uses your existing pandas/xlsxwriter)
        if ext == '.csv':
            df = pd.read_csv(f)
            out = io.BytesIO()
            with pd.ExcelWriter(out, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name="Sheet1")
            out.seek(0)
            return send_file(
                out,
                mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                as_attachment=True,
                download_name="converted.xlsx"
            )

        # XLSX/XLS -> PDF (via LibreOffice headless)
        elif ext in ('.xlsx', '.xls'):
            if shutil.which("soffice") is None:
                return {"error": "LibreOffice (soffice) not found in PATH."}, 500

            with tempfile.TemporaryDirectory() as tmp:
                in_path = os.path.join(tmp, secure_filename(f.filename))
                f.save(in_path)

                cmd = ["soffice", "--headless", "--convert-to", "pdf", "--outdir", tmp, in_path]
                proc = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                if proc.returncode != 0:
                    return {
                        "error": "PDF conversion failed.",
                        "stdout": proc.stdout.decode('utf-8', 'ignore')[:400],
                        "stderr": proc.stderr.decode('utf-8', 'ignore')[:400]
                    }, 500

                pdf_path = os.path.splitext(in_path)[0] + ".pdf"
                if not os.path.exists(pdf_path):
                    # Fallback to first generated PDF
                    candidates = [p for p in os.listdir(tmp) if p.lower().endswith(".pdf")]
                    if not candidates:
                        return {"error": "Converted PDF not found."}, 500
                    pdf_path = os.path.join(tmp, candidates[0])

                with open(pdf_path, "rb") as fh:
                    buf = io.BytesIO(fh.read())
                buf.seek(0)
                return send_file(buf, mimetype="application/pdf",
                                 as_attachment=True, download_name=os.path.basename(pdf_path))

        else:
            return {"error": "Unsupported file type. Upload .csv, .xlsx or .xls."}, 400

    except Exception as e:
        return {"error": str(e)}, 500

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)
