"""
API endpoint: POST /api/pdf_to_word
Converts PDF to DOCX using pdf2docx.
Accepts multipart/form-data with field 'file'.
Returns the DOCX file as a download.
"""
import os
import sys
import tempfile
import traceback
from http.server import BaseHTTPRequestHandler
import cgi
import io

# ── Try importing pdf2docx ────────────────────────────────────────────────────
try:
    from pdf2docx import Converter
    PDF2DOCX_OK = True
except ImportError:
    PDF2DOCX_OK = False


def _parse_multipart(handler):
    """Parse multipart/form-data from the request and return file bytes + filename."""
    content_type = handler.headers.get("Content-Type", "")
    if "multipart/form-data" not in content_type:
        return None, None, "Content-Type must be multipart/form-data"

    # Use cgi.FieldStorage to parse
    env = {
        "REQUEST_METHOD": "POST",
        "CONTENT_TYPE": content_type,
        "CONTENT_LENGTH": handler.headers.get("Content-Length", "0"),
    }
    body = handler.rfile.read(int(handler.headers.get("Content-Length", 0)))
    fs = cgi.FieldStorage(
        fp=io.BytesIO(body),
        environ=env,
        keep_blank_values=True
    )

    if "file" not in fs:
        return None, None, "No 'file' field in request"

    field = fs["file"]
    return field.file.read(), field.filename, None


class handler(BaseHTTPRequestHandler):

    def do_OPTIONS(self):
        self._cors()
        self.end_headers()

    def do_POST(self):
        self._cors()

        if not PDF2DOCX_OK:
            self._error(500, "pdf2docx not installed")
            return

        try:
            file_bytes, filename, err = _parse_multipart(self)
            if err:
                self._error(400, err)
                return

            if not filename or not filename.lower().endswith(".pdf"):
                self._error(400, "Please upload a PDF file")
                return

            # Write PDF to temp file, convert, read result
            with tempfile.TemporaryDirectory() as tmpdir:
                pdf_path  = os.path.join(tmpdir, "input.pdf")
                docx_path = os.path.join(tmpdir, "output.docx")

                with open(pdf_path, "wb") as f:
                    f.write(file_bytes)

                cv = Converter(pdf_path)
                cv.convert(docx_path, start=0, end=None)
                cv.close()

                with open(docx_path, "rb") as f:
                    docx_bytes = f.read()

            out_name = filename.rsplit(".", 1)[0] + ".docx"
            self.send_response(200)
            self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            self.send_header("Content-Disposition", f'attachment; filename="{out_name}"')
            self.send_header("Content-Length", str(len(docx_bytes)))
            self._cors_headers()
            self.end_headers()
            self.wfile.write(docx_bytes)

        except Exception as e:
            self._error(500, f"Conversion failed: {str(e)}\n{traceback.format_exc()}")

    def _cors(self):
        pass

    def _cors_headers(self):
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")

    def _error(self, code, msg):
        body = f'{{"error": "{msg}"}}'.encode()
        self.send_response(code)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(body)))
        self._cors_headers()
        self.end_headers()
        self.wfile.write(body)
