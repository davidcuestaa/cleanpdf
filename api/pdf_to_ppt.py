"""
API endpoint: POST /api/pdf_to_ppt
Converts each PDF page to a PowerPoint slide using pdfplumber (text) + python-pptx.
"""
import os
import io
import cgi
import traceback
from http.server import BaseHTTPRequestHandler

try:
    import pdfplumber
    from pptx import Presentation
    from pptx.util import Inches, Pt, Emu
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN
    DEPS_OK = True
except ImportError as e:
    DEPS_OK = False
    IMPORT_ERR = str(e)


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


def pdf_to_pptx(file_bytes):
    prs = Presentation()
    prs.slide_width  = Inches(13.33)
    prs.slide_height = Inches(7.5)

    blank_layout = prs.slide_layouts[6]  # completely blank

    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for page_num, page in enumerate(pdf.pages):
            slide = prs.slides.add_slide(blank_layout)

            # Background rectangle (white)
            bg = slide.shapes.add_shape(
                1,  # MSO_SHAPE_TYPE.RECTANGLE
                0, 0,
                prs.slide_width, prs.slide_height
            )
            bg.fill.solid()
            bg.fill.fore_color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            bg.line.color.rgb = RGBColor(0xE8, 0xEA, 0xF0)

            # Page number badge
            badge = slide.shapes.add_shape(
                1,
                Inches(12.5), Inches(0.15),
                Inches(0.7), Inches(0.3)
            )
            badge.fill.solid()
            badge.fill.fore_color.rgb = RGBColor(0x43, 0x61, 0xEE)
            badge.line.color.rgb = RGBColor(0x43, 0x61, 0xEE)
            tf = badge.text_frame
            tf.text = str(page_num + 1)
            tf.paragraphs[0].alignment = PP_ALIGN.CENTER
            run = tf.paragraphs[0].runs[0]
            run.font.size = Pt(9)
            run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            run.font.bold = True

            # Extract text
            text = page.extract_text() or ""
            lines = [l for l in text.split("\n") if l.strip()]

            if not lines:
                lines = ["(página sin texto extraíble)"]

            # Title = first line
            title_text = lines[0][:120]
            body_lines = lines[1:40]  # up to 39 more lines

            # Title text box
            title_box = slide.shapes.add_textbox(
                Inches(0.5), Inches(0.2),
                Inches(11.8), Inches(1.0)
            )
            tf = title_box.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.text = title_text
            p.alignment = PP_ALIGN.LEFT
            run = p.runs[0]
            run.font.size = Pt(20)
            run.font.bold = True
            run.font.color.rgb = RGBColor(0x14, 0x16, 0x1F)

            # Body text box
            if body_lines:
                body_box = slide.shapes.add_textbox(
                    Inches(0.5), Inches(1.4),
                    Inches(12.3), Inches(5.8)
                )
                tf = body_box.text_frame
                tf.word_wrap = True

                for i, line in enumerate(body_lines):
                    p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
                    p.text = line[:200]
                    p.alignment = PP_ALIGN.LEFT
                    if p.runs:
                        run = p.runs[0]
                        run.font.size = Pt(12)
                        run.font.color.rgb = RGBColor(0x14, 0x16, 0x1F)

    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


class handler(BaseHTTPRequestHandler):

    def do_OPTIONS(self):
        self._cors_headers()
        self.end_headers()

    def do_POST(self):
        if not DEPS_OK:
            self._error(500, f"Missing dependency: {IMPORT_ERR}")
            return
        try:
            file_bytes, filename, err = _parse_multipart(self)
            if err:
                self._error(400, err)
                return
            if not filename or not filename.lower().endswith(".pdf"):
                self._error(400, "Please upload a PDF file")
                return

            pptx_bytes = pdf_to_pptx(file_bytes)
            out_name = filename.rsplit(".", 1)[0] + ".pptx"

            self.send_response(200)
            self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.presentationml.presentation")
            self.send_header("Content-Disposition", f'attachment; filename="{out_name}"')
            self.send_header("Content-Length", str(len(pptx_bytes)))
            self._cors_headers()
            self.end_headers()
            self.wfile.write(pptx_bytes)

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
