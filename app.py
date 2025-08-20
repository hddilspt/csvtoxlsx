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
        # 1) Prefer multipart 'file' (Postman/browsers)
        f = request.files.get('file')
        if f and f.filename:
            filename = secure_filename(f.filename)
            data = f.read()
            mode = 'multipart'
        else:
            # 2) Raw body fallback (Power Automate HTTP)
            data = request.get_data() or b''
            if not data:
                return {"error": "No file part in the request"}, 400
            filename = request.args.get('filename') or 'upload'
            mode = 'raw'

        # detect by content/extension
        head = data[:8]
        is_xlsx = head.startswith(b'PK')
        is_xls  = head.startswith(b'\xD0\xCF\x11\xE0')
        ext = os.path.splitext(filename)[1].lower()

        # CSV -> XLSX
        if (ext == '.csv') and not (is_xlsx or is_xls):
            df = pd.read_csv(io.BytesIO(data))
            out = io.BytesIO()
            with pd.ExcelWriter(out, engine='xlsxwriter') as w:
                df.to_excel(w, index=False, sheet_name="Sheet1")
            out.seek(0)
            resp = send_file(out,
                mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                as_attachment=True, download_name="converted.xlsx")
            resp.headers['X-Debug-Path'] = f'{mode}:csv->xlsx'
            return resp

        # XLS/XLSX -> PDF
        if is_xlsx or is_xls or ext in ('.xlsx', '.xls'):
            if shutil.which("soffice") is None:
                return {"error": "LibreOffice (soffice) not found in PATH."}, 500
            with tempfile.TemporaryDirectory() as tmp:
                tmp_ext = '.xlsx' if (is_xlsx or ext == '.xlsx') else '.xls'
                in_path = os.path.join(tmp, 'upload' + tmp_ext)
                with open(in_path, 'wb') as fh: fh.write(data)
                cmd = ["soffice","--headless","--convert-to","pdf","--outdir",tmp,in_path]
                p = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                if p.returncode != 0:
                    return {"error": "PDF conversion failed.",
                            "stdout": p.stdout.decode('utf-8','ignore')[:400],
                            "stderr": p.stderr.decode('utf-8','ignore')[:400]}, 500
                pdf_path = os.path.splitext(in_path)[0] + ".pdf"
                if not os.path.exists(pdf_path):
                    cands = [x for x in os.listdir(tmp) if x.lower().endswith('.pdf')]
                    if not cands: return {"error": "Converted PDF not found."}, 500
                    pdf_path = os.path.join(tmp, cands[0])
                with open(pdf_path, 'rb') as fh:
                    buf = io.BytesIO(fh.read())
                buf.seek(0)
                resp = send_file(buf, mimetype="application/pdf",
                                 as_attachment=True, download_name=os.path.basename(pdf_path))
                resp.headers['X-Debug-Path'] = f'{mode}:xlsx->pdf'
                return resp

        return {"error": "Unsupported file type. Upload .csv, .xlsx or .xls."}, 400
    except Exception as e:
        return {"error": str(e)}, 500

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)
