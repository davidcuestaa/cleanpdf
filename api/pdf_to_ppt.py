"""
API endpoint: POST /api/pdf_to_ppt
Convierte páginas de PDF a diapositivas PowerPoint.

Entorno: Vercel serverless (100% Python, sin binarios del sistema).

Estrategia (por orden de calidad):
  1. pymupdf (fitz)  → renderiza cada página como imagen PNG de alta resolución
                       y la incrusta como diapositiva. Fidelidad 100%.
                       Funciona en Vercel: es una wheel Python pura.
                       pip install pymupdf

  2. pdfplumber + python-pptx → fallback si pymupdf no está disponible.
                       Extrae texto con coordenadas reales, bullets nativos,
                       detección de líneas horizontales.

Para PDFs que son presentaciones exportadas, la opción 1 es la correcta:
preserva fuentes, imágenes, gráficos, colores y layout exactos.
El resultado es editable en cuanto a mover/redimensionar la imagen,
pero el contenido está como imagen (igual que Google Slides al exportar a PDF).
"""
import io
import os
import cgi
import json
import re
from http.server import BaseHTTPRequestHandler

# ── Método 1: pymupdf (mejor calidad, recomendado) ────────────────────────────
try:
    import fitz  # pymupdf
    PYMUPDF_OK = True
except ImportError:
    PYMUPDF_OK = False

# ── Método 2: pdfplumber (fallback texto) ─────────────────────────────────────
try:
    import pdfplumber
    PDFPLUMBER_OK = True
except ImportError:
    PDFPLUMBER_OK = False

try:
    from pdfminer.high_level import extract_pages
    from pdfminer.layout import LTTextContainer, LTChar, LTLine, LTRect
    PDFMINER_OK = True
except ImportError:
    PDFMINER_OK = False

try:
    from pptx import Presentation
    from pptx.util import Inches, Pt, Emu
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN
    from pptx.oxml.ns import qn
    from lxml import etree
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


# ─────────────────────────────────────────────────────────────────────────────
# MÉTODO 1: pymupdf → imagen por página (fidelidad perfecta)
# ─────────────────────────────────────────────────────────────────────────────

def pdf_to_pptx_images(file_bytes):
    """
    Renderiza cada página del PDF como imagen PNG y la incrusta
    como diapositiva a pantalla completa.

    Ventajas: fidelidad 100%, funciona con cualquier PDF.
    Resultado: diapositivas con imagen (como Google Slides exportado a PDF).
    """
    doc = fitz.open(stream=file_bytes, filetype="pdf")

    prs = Presentation()

    # Detectar orientación de la primera página para ajustar tamaño de slide
    if len(doc) > 0:
        first = doc[0].rect
        aspect = first.width / first.height if first.height > 0 else 16/9
    else:
        aspect = 16 / 9

    # Fijar slide en 33.87 cm × alto proporcional (máx estándar PowerPoint)
    slide_w_emu = int(12192000)  # 33.87 cm en EMU — máximo de PowerPoint
    slide_h_emu = int(slide_w_emu / aspect)
    prs.slide_width  = slide_w_emu
    prs.slide_height = slide_h_emu

    blank_layout = prs.slide_layouts[6]  # completamente en blanco

    for page_num in range(len(doc)):
        page = doc[page_num]

        # Renderizar a 150 DPI — buen balance calidad/tamaño de archivo
        # (96 DPI = rápido/pequeño, 200 DPI = alta calidad/grande)
        mat  = fitz.Matrix(150 / 72, 150 / 72)
        pix  = page.get_pixmap(matrix=mat, alpha=False)
        img_bytes = pix.tobytes("png")

        slide = prs.slides.add_slide(blank_layout)

        # Insertar imagen a pantalla completa
        img_stream = io.BytesIO(img_bytes)
        slide.shapes.add_picture(
            img_stream,
            left=0, top=0,
            width=slide_w_emu,
            height=slide_h_emu,
        )

    doc.close()

    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


# ─────────────────────────────────────────────────────────────────────────────
# MÉTODO 2: pdfplumber → texto con layout (fallback, contenido editable)
# ─────────────────────────────────────────────────────────────────────────────

C_ACCENT = RGBColor(0x43, 0x61, 0xEE)
C_DARK   = RGBColor(0x14, 0x16, 0x1F)
C_MUTED  = RGBColor(0x70, 0x79, 0x99)
C_WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
C_BODY   = RGBColor(0x2A, 0x2D, 0x3E)
C_SUB    = RGBColor(0x43, 0x61, 0xEE)
C_BORDER = RGBColor(0xD8, 0xDB, 0xEC)

SLIDE_W = Inches(13.33)
SLIDE_H = Inches(7.5)

RE_BULLET = re.compile(
    r"^[\u2022\u2023\u2043\u25cf\u25e6\u25aa\u25ab\u25fb\u25fc"
    r"\u2043\u204c\u204d\u2219\u00b7\u25b8\u25b9\u25ba\u25bb"
    r"\u2192\u2794\u27a2\u2714\u2713\*\-\+]\s+"
)


def _rect(slide, x, y, w, h, fill=None):
    shape = slide.shapes.add_shape(1, x, y, w, h)
    if fill:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill
    else:
        shape.fill.background()
    shape.line.fill.background()
    return shape


def _textbox(slide, x, y, w, h, text, size, bold=False,
             color=None, align=PP_ALIGN.LEFT):
    color = color or C_DARK
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
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


def _add_native_bullet(para, color_hex="4361EE"):
    pPr = para._p.get_or_add_pPr()
    for child in list(pPr):
        if child.tag in (qn("a:buNone"), qn("a:buChar")):
            pPr.remove(child)
    buChar = etree.SubElement(pPr, qn("a:buChar"))
    buChar.set("char", "\u2022")
    pPr.set("marL", "342900")
    pPr.set("indent", "-342900")
    buClr = etree.SubElement(pPr, qn("a:buClr"))
    srgb  = etree.SubElement(buClr, qn("a:srgbClr"))
    srgb.set("val", color_hex)


def _extract_plumber(page):
    pw = float(page.width)
    ph = float(page.height)
    blocks = []

    # Líneas horizontales
    for col in (getattr(page, "lines", []), getattr(page, "rects", [])):
        for obj in col:
            try:
                w = abs(obj["x1"] - obj["x0"])
                h = abs(obj["bottom"] - obj["top"])
                if w > pw * 0.25 and h < 4:
                    blocks.append({
                        "text": "", "x0": 0, "y0": round(obj["top"] / ph, 3),
                        "x1": 1, "y1": round(obj["top"] / ph, 3),
                        "size": 0, "bold": False,
                        "is_bullet": False, "is_hr": True,
                    })
            except Exception:
                pass

    try:
        words = page.extract_words(
            x_tolerance=3, y_tolerance=3,
            keep_blank_chars=False, use_text_flow=True,
            extra_attrs=["size", "fontname"],
        )
    except Exception:
        words = page.extract_words(x_tolerance=3, y_tolerance=3)

    if words:
        lines, curr, last_y = [], [], None
        for word in sorted(words, key=lambda w: (w["top"], w["x0"])):
            if last_y is None or abs(word["top"] - last_y) < 6:
                curr.append(word)
                last_y = word["top"]
            else:
                lines.append(curr)
                curr   = [word]
                last_y = word["top"]
        if curr:
            lines.append(curr)

        for lw in lines:
            if not lw:
                continue
            text = " ".join(w["text"] for w in lw)
            x0   = min(w["x0"]     for w in lw) / pw
            y0   = min(w["top"]    for w in lw) / ph
            x1   = max(w["x1"]     for w in lw) / pw
            y1   = max(w["bottom"] for w in lw) / ph
            size = max(w.get("size", 11) for w in lw)
            bold = any("Bold" in str(w.get("fontname", "")) for w in lw)

            is_bullet = bool(RE_BULLET.match(text))
            if is_bullet:
                text = RE_BULLET.sub("", text).strip()
                if not text:
                    continue
            if not is_bullet and x0 > 0.08:
                is_bullet = True

            blocks.append({
                "text": text, "x0": x0, "y0": y0, "x1": x1, "y1": y1,
                "size": size, "bold": bold,
                "is_bullet": is_bullet, "is_hr": False,
            })

    blocks.sort(key=lambda b: b["y0"])
    return blocks


def _classify(blocks):
    text_b   = [b for b in blocks if not b["is_hr"] and not b["is_bullet"] and b["text"].strip()]
    max_size = max((b["size"] for b in text_b), default=11)

    title = subtitle = ""
    body  = []

    for b in sorted(blocks, key=lambda b: b["y0"]):
        if b["is_hr"] or b["is_bullet"]:
            body.append(b)
            continue
        text = b["text"].strip()
        if not text:
            continue
        if not title and (b["size"] >= max_size * 0.85 or
                          (b["bold"] and b["size"] >= max_size * 0.65)):
            title = text
        elif not subtitle and b["y0"] < 0.4 and b["size"] >= max_size * 0.55:
            subtitle = text
        else:
            body.append(b)

    return title, subtitle, body


def _build_slide(prs, slide_num, n_slides, title, subtitle, body_blocks):
    blank = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank)
    _set_bg(slide, RGBColor(0xFF, 0xFF, 0xFF))

    _rect(slide, 0, 0, SLIDE_W, Inches(0.07), fill=C_ACCENT)
    _rect(slide, 0, Inches(0.07), SLIDE_W, Inches(1.25),
          fill=RGBColor(0xF7, 0xF8, 0xFC))

    bw, bh = Inches(0.55), Inches(0.28)
    badge  = _rect(slide, SLIDE_W - bw - Inches(0.28), Inches(0.16), bw, bh,
                   fill=C_ACCENT)
    badge.text_frame.text = f"{slide_num}/{n_slides}"
    badge.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    for run in badge.text_frame.paragraphs[0].runs:
        run.font.size = Pt(8); run.font.bold = True
        run.font.color.rgb = C_WHITE

    if title:
        _textbox(slide, Inches(0.55), Inches(0.16),
                 SLIDE_W - Inches(1.5), Inches(0.8),
                 title, 22, bold=True, color=C_DARK)
    if subtitle:
        _textbox(slide, Inches(0.55), Inches(0.88),
                 SLIDE_W - Inches(1.2), Inches(0.4),
                 subtitle, 13, color=C_SUB)

    _rect(slide, Inches(0.55), Inches(1.33),
          SLIDE_W - Inches(1.1), Inches(0.015), fill=C_BORDER)

    margin = Inches(0.55)
    body_w = SLIDE_W - margin * 2
    body_y = Inches(1.45)
    body_h = SLIDE_H - body_y - Inches(0.35)

    if body_blocks and body_h > Inches(0.25):
        groups, current = [], []
        for b in body_blocks:
            if b["is_hr"]:
                if current: groups.append(("text", current)); current = []
                groups.append(("hr", None))
            else:
                current.append(b)
        if current:
            groups.append(("text", current))

        n_text = sum(1 for k, _ in groups if k == "text")
        n_hr   = sum(1 for k, _ in groups if k == "hr")
        hr_h_v = Inches(0.03)
        hr_gap = Inches(0.1)
        text_h = (body_h - n_hr * (hr_h_v + hr_gap * 2)) / max(n_text, 1)
        cur_y  = body_y

        for kind, data in groups:
            if kind == "hr":
                _rect(slide, margin + Inches(0.15), cur_y + hr_gap,
                      body_w - Inches(0.3), hr_h_v, fill=C_BORDER)
                cur_y += hr_h_v + hr_gap * 2
            else:
                if not data: continue
                tb = slide.shapes.add_textbox(margin, cur_y, body_w, text_h)
                tf = tb.text_frame; tf.word_wrap = True
                first = True
                for b in data:
                    text = b["text"].strip()
                    if not text: continue
                    p = tf.paragraphs[0] if first else tf.add_paragraph()
                    first = False
                    p.alignment = PP_ALIGN.LEFT
                    if b["is_bullet"]:
                        _add_native_bullet(p)
                    run = p.add_run()
                    run.text = text
                    is_h = b["bold"] and b["size"] >= 13 and not b["is_bullet"]
                    run.font.size = Pt(min(max(float(b["size"]), 9), 20))
                    run.font.bold = is_h
                    run.font.color.rgb = C_DARK if is_h else C_BODY
                    p.space_after = Pt(4 if is_h else 2)
                cur_y += text_h

    _rect(slide, 0, SLIDE_H - Inches(0.28), SLIDE_W, Inches(0.015),
          fill=C_BORDER)
    return slide


def pdf_to_pptx_text(file_bytes):
    prs = Presentation()
    prs.slide_width  = SLIDE_W
    prs.slide_height = SLIDE_H

    if PDFPLUMBER_OK:
        pages = []
        with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
            for page in pdf.pages:
                pages.append(_extract_plumber(page))
    elif PDFMINER_OK:
        pages = []
        for pl in extract_pages(io.BytesIO(file_bytes)):
            pw, ph = float(pl.width), float(pl.height)
            blocks = []
            for el in pl:
                if isinstance(el, LTTextContainer):
                    text = el.get_text().strip()
                    if not text: continue
                    size, bold = 11.0, False
                    for line in el:
                        for ch in line:
                            if isinstance(ch, LTChar):
                                size = ch.size
                                bold = "Bold" in (ch.fontname or "")
                                break
                        break
                    x0, y0 = el.x0/pw, 1-el.y1/ph
                    is_b = bool(RE_BULLET.match(text))
                    if is_b: text = RE_BULLET.sub("", text).strip()
                    if not text: continue
                    if not is_b and x0 > 0.08: is_b = True
                    blocks.append({"text": text, "x0": x0, "y0": y0,
                                   "x1": el.x1/pw, "y1": 1-el.y0/ph,
                                   "size": size, "bold": bold,
                                   "is_bullet": is_b, "is_hr": False})
            blocks.sort(key=lambda b: b["y0"])
            pages.append(blocks)
    else:
        raise RuntimeError("Instala pymupdf: pip install pymupdf")

    n = len(pages)
    for idx, blocks in enumerate(pages):
        title, subtitle, body = _classify(blocks)
        _build_slide(prs, idx+1, n, title, subtitle, body)

    if not pages:
        blank = prs.slide_layouts[6]
        slide = prs.slides.add_slide(blank)
        _set_bg(slide, RGBColor(0xFF, 0xFF, 0xFF))
        _textbox(slide, SLIDE_W//4, SLIDE_H//3, SLIDE_W//2, Inches(1),
                 "PDF sin contenido extraíble", 18, bold=True, color=C_MUTED,
                 align=PP_ALIGN.CENTER)

    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


# ─────────────────────────────────────────────────────────────────────────────
# Dispatcher principal
# ─────────────────────────────────────────────────────────────────────────────

def pdf_to_pptx(file_bytes):
    if not PPTX_OK:
        raise RuntimeError("python-pptx no está instalado")

    # Preferir imagen (fidelidad perfecta) si pymupdf está disponible
    if PYMUPDF_OK:
        return pdf_to_pptx_images(file_bytes)

    # Fallback: extracción de texto con layout
    return pdf_to_pptx_text(file_bytes)


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
