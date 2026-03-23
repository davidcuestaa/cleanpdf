"""
API endpoint: POST /api/office_to_pdf
Converts Word (.docx), Excel (.xlsx), PowerPoint (.pptx) to PDF.
Uses python-docx + reportlab for DOCX, openpyxl + reportlab for XLSX,
python-pptx + reportlab for PPTX.
"""
import os
import io
import cgi
import traceback
from http.server import BaseHTTPRequestHandler

try:
    from docx import Document
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import cm
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_LEFT, TA_CENTER
    import openpyxl
    from pptx import Presentation
    from pptx.util import Pt
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


# ── DOCX → PDF ────────────────────────────────────────────────────────────────
def docx_to_pdf(file_bytes):
    doc = Document(io.BytesIO(file_bytes))
    buf = io.BytesIO()
    pdf = SimpleDocTemplate(buf, pagesize=A4,
                            leftMargin=2*cm, rightMargin=2*cm,
                            topMargin=2*cm, bottomMargin=2*cm)
    styles = getSampleStyleSheet()
    story = []

    heading_sizes = {1: 18, 2: 15, 3: 13, 4: 12, 5: 11, 6: 10}

    for para in doc.paragraphs:
        text = para.text
        if not text.strip():
            story.append(Spacer(1, 6))
            continue

        style_name = para.style.name if para.style else "Normal"

        # Detect headings
        if style_name.startswith("Heading"):
            level = 1
            try:
                level = int(style_name.split()[-1])
            except Exception:
                pass
            size = heading_sizes.get(level, 12)
            hs = ParagraphStyle(
                f"H{level}", parent=styles["Normal"],
                fontSize=size, fontName="Helvetica-Bold",
                spaceAfter=6, spaceBefore=10,
            )
            story.append(Paragraph(_esc(text), hs))
        else:
            # Bold / italic detection from runs
            is_bold   = any(r.bold   for r in para.runs if r.bold   is not None)
            is_italic = any(r.italic for r in para.runs if r.italic is not None)
            font_name = "Helvetica"
            if is_bold and is_italic:
                font_name = "Helvetica-BoldOblique"
            elif is_bold:
                font_name = "Helvetica-Bold"
            elif is_italic:
                font_name = "Helvetica-Oblique"

            ps = ParagraphStyle(
                "body", parent=styles["Normal"],
                fontName=font_name, fontSize=10,
                leading=14, spaceAfter=4,
            )
            story.append(Paragraph(_esc(text), ps))

    # Tables
    for table in doc.tables:
        data = []
        for row in table.rows:
            data.append([_esc(cell.text) for cell in row.cells])
        if data:
            col_count = max(len(r) for r in data)
            col_width = (A4[0] - 4*cm) / col_count
            ts = TableStyle([
                ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#4361ee")),
                ("TEXTCOLOR",  (0,0), (-1,0), colors.white),
                ("FONTNAME",   (0,0), (-1,0), "Helvetica-Bold"),
                ("FONTSIZE",   (0,0), (-1,-1), 9),
                ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.white, colors.HexColor("#f7f8fc")]),
                ("GRID", (0,0), (-1,-1), 0.5, colors.HexColor("#e8eaf0")),
                ("LEFTPADDING",  (0,0), (-1,-1), 6),
                ("RIGHTPADDING", (0,0), (-1,-1), 6),
                ("TOPPADDING",   (0,0), (-1,-1), 4),
                ("BOTTOMPADDING",(0,0), (-1,-1), 4),
            ])
            t = Table(data, colWidths=[col_width]*col_count, style=ts)
            story.append(Spacer(1, 8))
            story.append(t)
            story.append(Spacer(1, 8))

    if not story:
        story.append(Paragraph("Documento vacío", styles["Normal"]))

    pdf.build(story)
    return buf.getvalue()


# ── XLSX → PDF ────────────────────────────────────────────────────────────────
def xlsx_to_pdf(file_bytes):
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    buf = io.BytesIO()
    pdf = SimpleDocTemplate(buf, pagesize=A4,
                            leftMargin=1.5*cm, rightMargin=1.5*cm,
                            topMargin=2*cm, bottomMargin=2*cm)
    styles = getSampleStyleSheet()
    story = []

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]

        # Sheet title
        story.append(Paragraph(f"<b>{_esc(sheet_name)}</b>", styles["Heading2"]))
        story.append(Spacer(1, 6))

        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            story.append(Paragraph("(hoja vacía)", styles["Normal"]))
            story.append(PageBreak())
            continue

        # Build table data — limit columns to keep readable
        max_cols = min(len(rows[0]) if rows[0] else 1, 12)
        data = []
        for row in rows:
            data.append([str(cell) if cell is not None else "" for cell in row[:max_cols]])

        if data:
            col_w = (A4[0] - 3*cm) / max_cols
            ts = TableStyle([
                ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#4361ee")),
                ("TEXTCOLOR",  (0,0), (-1,0), colors.white),
                ("FONTNAME",   (0,0), (-1,0), "Helvetica-Bold"),
                ("FONTSIZE",   (0,0), (-1,-1), 7),
                ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.white, colors.HexColor("#f7f8fc")]),
                ("GRID", (0,0), (-1,-1), 0.4, colors.HexColor("#e8eaf0")),
                ("LEFTPADDING",  (0,0), (-1,-1), 4),
                ("RIGHTPADDING", (0,0), (-1,-1), 4),
                ("TOPPADDING",   (0,0), (-1,-1), 3),
                ("BOTTOMPADDING",(0,0), (-1,-1), 3),
            ])
            t = Table(data, colWidths=[col_w]*max_cols, style=ts, repeatRows=1)
            story.append(t)

        if sheet_name != wb.sheetnames[-1]:
            story.append(PageBreak())

    if not story:
        story.append(Paragraph("Documento vacío", styles["Normal"]))

    pdf.build(story)
    return buf.getvalue()


# ── PPTX → PDF ────────────────────────────────────────────────────────────────
def pptx_to_pdf(file_bytes):
    prs = Presentation(io.BytesIO(file_bytes))
    buf = io.BytesIO()
    # Use landscape A4 for slides
    from reportlab.lib.pagesizes import landscape
    page_w, page_h = landscape(A4)
    pdf = SimpleDocTemplate(buf, pagesize=(page_w, page_h),
                            leftMargin=1.5*cm, rightMargin=1.5*cm,
                            topMargin=1.5*cm, bottomMargin=1.5*cm)
    styles = getSampleStyleSheet()
    story = []

    slide_style = ParagraphStyle(
        "slide_num", parent=styles["Normal"],
        fontSize=8, textColor=colors.HexColor("#7c8099"),
        spaceAfter=4,
    )
    title_style = ParagraphStyle(
        "slide_title", parent=styles["Normal"],
        fontSize=18, fontName="Helvetica-Bold",
        textColor=colors.HexColor("#14161f"),
        spaceBefore=4, spaceAfter=8, leading=22,
    )
    body_style = ParagraphStyle(
        "slide_body", parent=styles["Normal"],
        fontSize=11, leading=16, spaceAfter=4,
        textColor=colors.HexColor("#14161f"),
    )

    for i, slide in enumerate(prs.slides):
        story.append(Paragraph(f"Diapositiva {i+1}", slide_style))

        # Extract text from shapes
        title_text = ""
        body_texts = []

        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            for j, para in enumerate(shape.text_frame.paragraphs):
                text = para.text.strip()
                if not text:
                    continue
                # Heuristic: first non-empty paragraph with large font = title
                try:
                    run_size = para.runs[0].font.size
                    is_big = run_size and run_size >= Pt(20)
                except Exception:
                    is_big = False

                if j == 0 and (is_big or not title_text):
                    if not title_text:
                        title_text = text
                    else:
                        body_texts.append(text)
                else:
                    body_texts.append(text)

        if title_text:
            story.append(Paragraph(_esc(title_text), title_style))
        for bt in body_texts:
            story.append(Paragraph(f"• {_esc(bt)}", body_style))

        if i < len(prs.slides) - 1:
            story.append(PageBreak())

    if not story:
        story.append(Paragraph("Presentación vacía", styles["Normal"]))

    pdf.build(story)
    return buf.getvalue()


def _esc(text):
    """Escape special XML chars for ReportLab Paragraph."""
    return (str(text)
            .replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
            .replace('"', "&quot;"))


CONVERTERS = {
    ".docx": ("application/pdf", docx_to_pdf),
    ".xlsx": ("application/pdf", xlsx_to_pdf),
    ".pptx": ("application/pdf", pptx_to_pdf),
}


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

            ext = os.path.splitext(filename or "")[1].lower()
            if ext not in CONVERTERS:
                self._error(400, f"Unsupported format: {ext}. Use .docx, .xlsx, or .pptx")
                return

            mime, convert_fn = CONVERTERS[ext]
            pdf_bytes = convert_fn(file_bytes)

            out_name = filename.rsplit(".", 1)[0] + ".pdf"
            self.send_response(200)
            self.send_header("Content-Type", "application/pdf")
            self.send_header("Content-Disposition", f'attachment; filename="{out_name}"')
            self.send_header("Content-Length", str(len(pdf_bytes)))
            self._cors_headers()
            self.end_headers()
            self.wfile.write(pdf_bytes)

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
