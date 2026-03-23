"""
API endpoint: POST /api/pdf_to_excel
Extracts tables from PDF to XLSX using pdfplumber + openpyxl.
"""
import os
import tempfile
import traceback
import io
import cgi
from http.server import BaseHTTPRequestHandler

try:
    import pdfplumber
    import openpyxl
    DEPS_OK = True
except ImportError:
    DEPS_OK = False


def _parse_multipart(handler):
    content_type = handler.headers.get("Content-Type", "")
    if "multipart/form-data" not in content_type:
        return None, None, "Content-Type must be multipart/form-data"
    env = {
        "REQUEST_METHOD": "POST",
        "CONTENT_TYPE": content_type,
        "CONTENT_LENGTH": handler.headers.get("Content-Length", "0"),
    }
    body = handler.rfile.read(int(handler.headers.get("Content-Length", 0)))
    fs = cgi.FieldStorage(fp=io.BytesIO(body), environ=env, keep_blank_values=True)
    if "file" not in fs:
        return None, None, "No 'file' field in request"
    field = fs["file"]
    return field.file.read(), field.filename, None


class handler(BaseHTTPRequestHandler):

    def do_OPTIONS(self):
        self._cors_headers()
        self.end_headers()

    def do_POST(self):
        if not DEPS_OK:
            self._error(500, "pdfplumber or openpyxl not installed")
            return
        try:
            file_bytes, filename, err = _parse_multipart(self)
            if err:
                self._error(400, err)
                return
            if not filename or not filename.lower().endswith(".pdf"):
                self._error(400, "Please upload a PDF file")
                return

            wb = openpyxl.Workbook()
            wb.remove(wb.active)  # remove default sheet

            with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
                total_tables = 0
                for page_num, page in enumerate(pdf.pages, 1):
                    tables = page.extract_tables()
                    if not tables:
                        # No tables — extract text as single-column fallback
                        text = page.extract_text()
                        if text:
                            ws = wb.create_sheet(title=f"Pag{page_num}_texto")
                            for line in text.split("\n"):
                                ws.append([line])
                        continue

                    for t_idx, table in enumerate(tables):
                        total_tables += 1
                        sheet_name = f"Pag{page_num}"
                        if len(tables) > 1:
                            sheet_name += f"_T{t_idx+1}"
                        # Sheet names max 31 chars
                        ws = wb.create_sheet(title=sheet_name[:31])

                        for row in table:
                            # Replace None with empty string
                            ws.append([cell if cell is not None else "" for cell in row])

                        # Auto-width columns
                        for col in ws.columns:
                            max_len = 0
                            col_letter = col[0].column_letter
                            for cell in col:
                                try:
                                    if cell.value:
                                        max_len = max(max_len, len(str(cell.value)))
                                except Exception:
                                    pass
                            ws.column_dimensions[col_letter].width = min(max_len + 2, 50)

            if not wb.sheetnames:
                wb.create_sheet("Sin_tablas")
                wb.active.append(["No se encontraron tablas en este PDF"])

            out_buf = io.BytesIO()
            wb.save(out_buf)
            xlsx_bytes = out_buf.getvalue()

            out_name = filename.rsplit(".", 1)[0] + ".xlsx"
            self.send_response(200)
            self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            self.send_header("Content-Disposition", f'attachment; filename="{out_name}"')
            self.send_header("Content-Length", str(len(xlsx_bytes)))
            self._cors_headers()
            self.end_headers()
            self.wfile.write(xlsx_bytes)

        except Exception as e:
            self._error(500, f"Conversion failed: {str(e)}\n{traceback.format_exc()}")

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
