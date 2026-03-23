"""
API endpoint: POST /api/pdf_to_ppt
Convierte páginas de PDF a diapositivas PowerPoint editables.

Entorno: Vercel serverless.
Librerías: pdfplumber + python-pptx (100% Python, sin binarios del sistema)
           pdf2image NO funciona en Vercel (necesita poppler)

Estrategia:
- Extrae texto con coordenadas reales (pdfplumber)
- Clasifica bloques por tamaño de fuente → título / subtítulo / cuerpo
- Posiciona cajas de texto en la diapositiva respetando el layout original
- Detecta y convierte tablas del PDF a tablas PowerPoint nativas
- Diseño profesional con cabecera, número de diapositiva y separadores
"""
import io
import os
import cgi
import json
import traceback
from http.server import BaseHTTPRequestHandler

try:
    import pdfplumber
    PDFPLUMBER_OK = True
except ImportError:
    PDFPLUMBER_OK = False

try:
    from pdfminer.high_level import extract_pages
    from pdfminer.layout import LTTextContainer, LTChar
    PDFMINER_OK = True
except ImportError:
    PDFMINER_OK = False

try:
    from pptx import Presentation
    from pptx.util import Inches, Pt, Emu
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN
    from pptx.enum.shapes import PP_PLACEHOLDER
    PPTX_OK = True
except ImportError:
    PPTX_OK = False


def _parse_multipart(handler):
    ct = handler.headers.get("Content-Type", "")
    if "multipart/form-data" not in ct:
        return None, None, "Content-Type must be multipart/form-data"
    env = {"REQUEST_METHOD": "POST", "CONTENT_TYPE": ct,
           "CONTENT_LENGTH": handler.headers.get("Content-Length", "0")}
    body = handler.rfile.read(int(handler.headers.get("Content-Length", 0)))
    fs = cgi.FieldStorage(fp=io.BytesIO(body), environ=env, keep_blank_values=True)
    if "file" not in fs:
        return None, None, "No 'file' field in request"
    f = fs["file"]
    return f.file.read(), f.filename, None


# ── Colores y constantes ──────────────────────────────────────────────────────

C_ACCENT = RGBColor(0x43, 0x61, 0xEE)
C_DARK   = RGBColor(0x14, 0x16, 0x1F)
C_MUTED  = RGBColor(0x70, 0x79, 0x99)
C_WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
C_BG     = RGBColor(0xF7, 0xF8, 0xFC)
C_BODY   = RGBColor(0x2A, 0x2D, 0x3E)
C_SUB    = RGBColor(0x43, 0x61, 0xEE)
C_BORDER = RGBColor(0xE0, 0xE3, 0xF0)

SLIDE_W = Inches(13.33)  # 16:9 widescreen
SLIDE_H = Inches(7.5)


# ── Helpers de forma ─────────────────────────────────────────────────────────

def _rect(slide, x, y, w, h, fill=None, line=None, line_w=0.5):
    shape = slide.shapes.add_shape(1, x, y, w, h)
    if fill:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill
    else:
        shape.fill.background()
    if line:
        shape.line.color.rgb = line
        shape.line.width = Pt(line_w)
    else:
        shape.line.fill.background()
    return shape


def _textbox(slide, x, y, w, h, text, size, bold=False,
             color=None, align=PP_ALIGN.LEFT, wrap=True):
    color = color or C_DARK
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = wrap
    p  = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size  = Pt(size)
    run.font.bold  = bold
    run.font.color.rgb = color
    return tb


def _set_bg(slide, rgb):
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = rgb


# ── Extracción de contenido ───────────────────────────────────────────────────

def _extract_page_content(page):
    """
    Extrae bloques de texto con posición, tamaño y negrita.
    Coordenadas normalizadas 0-1 (0,0 = esquina superior izquierda).
    """
    blocks = []
    pw = float(page.width)
    ph = float(page.height)

    # Extraer palabras con atributos de fuente
    try:
        words = page.extract_words(
            x_tolerance=3, y_tolerance=3,
            keep_blank_chars=False,
            use_text_flow=True,
            extra_attrs=["size", "fontname"],
        )
    except Exception:
        words = page.extract_words(x_tolerance=3, y_tolerance=3)

    if not words:
        return blocks

    # Agrupar palabras en líneas por proximidad vertical
    lines = []
    curr  = []
    last_y = None

    for word in sorted(words, key=lambda w: (-w["top"], w["x0"])):
        if last_y is None or abs(word["top"] - last_y) < 6:
            curr.append(word)
            last_y = word["top"]
        else:
            if curr:
                lines.append(curr)
            curr   = [word]
            last_y = word["top"]
    if curr:
        lines.append(curr)

    for line_words in lines:
        if not line_words:
            continue
        text = " ".join(w["text"] for w in line_words)
        x0   = min(w["x0"]     for w in line_words) / pw
        y0   = min(w["top"]    for w in line_words) / ph
        x1   = max(w["x1"]     for w in line_words) / pw
        y1   = max(w["bottom"] for w in line_words) / ph
        size = max(w.get("size", 11) for w in line_words)
        bold = any("Bold" in str(w.get("fontname", "")) for w in line_words)

        blocks.append({
            "text": text, "x0": x0, "y0": y0,
            "x1": x1, "y1": y1, "size": size, "bold": bold,
        })

    return blocks


def _classify(blocks):
    """Clasifica bloques en título / subtítulo / cuerpo."""
    if not blocks:
        return "", "", []

    sorted_b = sorted(blocks, key=lambda b: b["y0"])
    max_size = max(b["size"] for b in sorted_b)

    title = subtitle = ""
    body  = []

    for b in sorted_b:
        text = b["text"].strip()
        if not text:
            continue
        size = b["size"]
        bold = b["bold"]
        top  = b["y0"]

        if not title and (size >= max_size * 0.85 or
                          (bold and size >= max_size * 0.65)):
            title = text
        elif not subtitle and top < 0.4 and size >= max_size * 0.55:
            subtitle = text
        else:
            body.append(b)

    return title, subtitle, body


def _extract_tables(page):
    """Extrae tablas de la página con múltiples estrategias."""
    for settings in [
        {"vertical_strategy": "lines_strict",
         "horizontal_strategy": "lines_strict", "snap_tolerance": 4},
        {"vertical_strategy": "lines",
         "horizontal_strategy": "lines", "snap_tolerance": 5},
    ]:
        try:
            tables = page.extract_tables(settings)
            if tables:
                valid = [
                    t for t in tables
                    if any(any(c for c in row if c) for row in t)
                ]
                if valid:
                    return valid
        except Exception:
            pass
    return []


# ── Construcción de diapositivas ──────────────────────────────────────────────

def _add_pptx_table(slide, table_data, x, y, w, h):
    """Añade una tabla nativa de PowerPoint."""
    if not table_data:
        return

    rows = len(table_data)
    cols = max(len(r) for r in table_data) if table_data else 1

    shape = slide.shapes.add_table(rows, cols, x, y, w, h)
    tbl   = shape.table

    # Ancho de columnas uniforme
    for c in range(cols):
        tbl.columns[c].width = w // cols

    for r_idx, row in enumerate(table_data):
        for c_idx in range(cols):
            val  = row[c_idx] if c_idx < len(row) else ""
            cell = tbl.cell(r_idx, c_idx)
            cell.text = str(val) if val else ""
            # Estilos de celda
            tf = cell.text_frame
            tf.word_wrap = True
            for para in tf.paragraphs:
                for run in para.runs:
                    run.font.size = Pt(9)
                    run.font.bold = (r_idx == 0)
                    run.font.color.rgb = C_WHITE if r_idx == 0 else C_BODY
            # Fondo de cabecera
            if r_idx == 0:
                cell.fill.solid()
                cell.fill.fore_color.rgb = C_ACCENT
            elif r_idx % 2 == 0:
                cell.fill.solid()
                cell.fill.fore_color.rgb = RGBColor(0xF4, 0xF5, 0xFB)
            else:
                cell.fill.solid()
                cell.fill.fore_color.rgb = C_WHITE


def build_slide(prs, slide_num, n_slides, title, subtitle, body_blocks,
                page_tables=None):
    """Construye una diapositiva completa."""
    blank  = prs.slide_layouts[6]
    slide  = prs.slides.add_slide(blank)

    _set_bg(slide, RGBColor(0xFF, 0xFF, 0xFF))

    # Barra de acento superior
    _rect(slide, 0, 0, SLIDE_W, Inches(0.07), fill=C_ACCENT)

    # Zona de cabecera
    _rect(slide, 0, Inches(0.07), SLIDE_W, Inches(1.25),
          fill=RGBColor(0xF7, 0xF8, 0xFC))

    # Número de diapositiva (badge)
    badge_w = Inches(0.50)
    badge_h = Inches(0.28)
    badge   = _rect(slide,
                    SLIDE_W - badge_w - Inches(0.28),
                    Inches(0.16),
                    badge_w, badge_h, fill=C_ACCENT)
    tf = badge.text_frame
    tf.text = f"{slide_num}/{n_slides}"
    tf.paragraphs[0].alignment = PP_ALIGN.CENTER
    for run in tf.paragraphs[0].runs:
        run.font.size  = Pt(8)
        run.font.bold  = True
        run.font.color.rgb = C_WHITE

    # Título
    if title:
        _textbox(slide,
                 Inches(0.55), Inches(0.16),
                 SLIDE_W - Inches(1.5), Inches(0.8),
                 title, 22, bold=True, color=C_DARK, align=PP_ALIGN.LEFT)

    # Subtítulo
    if subtitle:
        _textbox(slide,
                 Inches(0.55), Inches(0.88),
                 SLIDE_W - Inches(1.2), Inches(0.4),
                 subtitle, 13, bold=False, color=C_SUB, align=PP_ALIGN.LEFT)

    # Línea divisoria
    _rect(slide, Inches(0.55), Inches(1.33),
          SLIDE_W - Inches(1.1), Inches(0.015), fill=C_BORDER)

    # Área de cuerpo
    body_y = Inches(1.42)
    body_h = SLIDE_H - body_y - Inches(0.35)
    margin = Inches(0.55)
    body_w = SLIDE_W - margin * 2

    # Tablas en la diapositiva
    if page_tables:
        table_area_h = body_h * 0.55 if body_blocks else body_h
        table_area_h = max(table_area_h, Inches(1.5))
        table_y      = body_y

        for t_idx, tbl_data in enumerate(page_tables[:2]):  # máx 2 tablas
            t_h = table_area_h // len(page_tables)
            _add_pptx_table(slide, tbl_data,
                            margin, table_y + t_idx * t_h,
                            body_w, t_h - Inches(0.1))

        body_y = body_y + table_area_h + Inches(0.1)
        body_h = SLIDE_H - body_y - Inches(0.35)

    # Texto de cuerpo
    if body_blocks and body_h > Inches(0.3):
        tb = slide.shapes.add_textbox(margin, body_y, body_w, body_h)
        tf = tb.text_frame
        tf.word_wrap = True

        first = True
        for b in body_blocks:
            text = b["text"].strip()
            if not text:
                continue
            size = b["size"]
            bold = b["bold"]

            p = tf.paragraphs[0] if first else tf.add_paragraph()
            first = False
            p.alignment = PP_ALIGN.LEFT

            is_heading = bold and size >= 13
            prefix = "" if is_heading else "• "

            run = p.add_run()
            run.text = prefix + text
            run.font.size  = Pt(min(max(float(size), 9), 20))
            run.font.bold  = is_heading
            run.font.color.rgb = C_DARK if is_heading else C_BODY

            p.space_after = Pt(4 if is_heading else 2)

    # Pie de página
    footer_y = SLIDE_H - Inches(0.3)
    _rect(slide, 0, footer_y, SLIDE_W, Inches(0.015), fill=C_BORDER)

    return slide


# ── Conversión principal ──────────────────────────────────────────────────────

def pdf_to_pptx(file_bytes):
    if not PPTX_OK:
        raise RuntimeError("python-pptx no está instalado")

    prs = Presentation()
    prs.slide_width  = SLIDE_W
    prs.slide_height = SLIDE_H

    if PDFPLUMBER_OK:
        pages_content = []
        with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
            for page in pdf.pages:
                blocks = _extract_page_content(page)
                tables = _extract_tables(page)
                pages_content.append((blocks, tables))
    elif PDFMINER_OK:
        pages_content = []
        for page_layout in extract_pages(io.BytesIO(file_bytes)):
            pw = float(page_layout.width)
            ph = float(page_layout.height)
            blocks = []
            for element in page_layout:
                if isinstance(element, LTTextContainer):
                    text = element.get_text().strip()
                    if not text:
                        continue
                    size, bold = 11.0, False
                    for line in element:
                        for char in line:
                            if isinstance(char, LTChar):
                                size = char.size
                                bold = "Bold" in (char.fontname or "")
                                break
                        break
                    blocks.append({
                        "text": text,
                        "x0": element.x0 / pw,
                        "y0": 1 - element.y1 / ph,
                        "x1": element.x1 / pw,
                        "y1": 1 - element.y0 / ph,
                        "size": size, "bold": bold,
                    })
            pages_content.append((blocks, []))
    else:
        raise RuntimeError(
            "Instala pdfplumber: pip install pdfplumber python-pptx"
        )

    n_slides = len(pages_content)

    for idx, (blocks, tables) in enumerate(pages_content):
        title, subtitle, body = _classify(blocks)
        build_slide(prs, idx + 1, n_slides,
                    title, subtitle, body, tables)

    if not pages_content:
        # PDF vacío — slide de aviso
        blank = prs.slide_layouts[6]
        slide = prs.slides.add_slide(blank)
        _set_bg(slide, RGBColor(0xFF, 0xFF, 0xFF))
        _textbox(slide, SLIDE_W//4, SLIDE_H//3,
                 SLIDE_W//2, Inches(1),
                 "PDF sin contenido extraíble", 18,
                 bold=True, color=C_MUTED, align=PP_ALIGN.CENTER)

    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


# ── HTTP handler ──────────────────────────────────────────────────────────────

class handler(BaseHTTPRequestHandler):

    def do_OPTIONS(self):
        self._cors()
        self.end_headers()

    def do_POST(self):
        try:
            file_bytes, filename, err = _parse_multipart(self)
            if err:
                self._error(400, err)
                return
            if not filename or not filename.lower().endswith(".pdf"):
                self._error(400, "Por favor sube un archivo PDF (.pdf)")
                return
            if not file_bytes:
                self._error(400, "El archivo está vacío")
                return

            pptx_bytes = pdf_to_pptx(file_bytes)
            out_name   = filename.rsplit(".", 1)[0] + ".pptx"

            self.send_response(200)
            self.send_header(
                "Content-Type",
                "application/vnd.openxmlformats-officedocument"
                ".presentationml.presentation",
            )
            self.send_header("Content-Disposition",
                             f'attachment; filename="{out_name}"')
            self.send_header("Content-Length", str(len(pptx_bytes)))
            self._cors()
            self.end_headers()
            self.wfile.write(pptx_bytes)

        except Exception as e:
            self._error(500, f"Error al convertir: {str(e).splitlines()[0]}")

    def _cors(self):
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")

    def _error(self, code, msg):
        body = json.dumps({"error": msg}).encode()
        self.send_response(code)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(body)))
        self._cors()
        self.end_headers()
        self.wfile.write(body)

    def log_message(self, fmt, *args):
        pass
