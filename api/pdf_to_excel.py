"""
API endpoint: POST /api/pdf_to_excel
Extrae tablas de PDF a XLSX.

Entorno: Vercel serverless.
Librería: pdfplumber (100% Python, funciona en Vercel)
          camelot-py NO funciona en Vercel (necesita OpenCV + Ghostscript)

Estrategia de extracción en 2 pasadas:
  1. lines_strict  — tablas con líneas explícitas (más preciso)
  2. text-based    — tablas separadas por espacios (facturas, listados)

Styling XLSX: cabecera azul, filas alternas, columnas auto-ajustadas,
primera fila congelada, bordes.
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
    from pdfminer.high_level import extract_text as pm_extract_text
    PDFMINER_OK = True
except ImportError:
    PDFMINER_OK = False

try:
    import openpyxl
    from openpyxl.styles import (Font, PatternFill, Alignment,
                                  Border, Side, GradientFill)
    from openpyxl.utils import get_column_letter
    OPENPYXL_OK = True
except ImportError:
    OPENPYXL_OK = False


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


# ── Estilos XLSX ──────────────────────────────────────────────────────────────

HEADER_FILL = PatternFill("solid", fgColor="4361EE")
ALT_FILL    = PatternFill("solid", fgColor="F4F5FB")
HEADER_FONT = Font(bold=True, color="FFFFFF", name="Calibri", size=10)
BODY_FONT   = Font(name="Calibri", size=10)
CENTER_AL   = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT_AL     = Alignment(horizontal="left",   vertical="center", wrap_text=True)
THIN        = Side(style="thin", color="D0D4E8")
BORDER      = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def _style_ws(ws, data):
    """Aplica estilo profesional a una hoja."""
    if not data:
        return
    for r_idx, row in enumerate(data, 1):
        is_hdr = r_idx == 1
        is_alt = r_idx % 2 == 0
        for c_idx, val in enumerate(row, 1):
            cell = ws.cell(row=r_idx, column=c_idx, value=val)
            cell.border    = BORDER
            cell.alignment = CENTER_AL if is_hdr else LEFT_AL
            if is_hdr:
                cell.fill = HEADER_FILL
                cell.font = HEADER_FONT
            elif is_alt:
                cell.fill = ALT_FILL
                cell.font = BODY_FONT
            else:
                cell.font = BODY_FONT

    # Auto-ancho de columnas
    for col in ws.columns:
        letter  = get_column_letter(col[0].column)
        max_len = max(
            (len(str(c.value).split("\n")[0]) for c in col if c.value),
            default=0
        )
        ws.column_dimensions[letter].width = min(max(max_len + 2, 8), 60)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


def _safe_name(name, used):
    name = str(name)[:28]
    for ch in r'\/*?[]':
        name = name.replace(ch, "_")
    base = name
    i = 1
    while name in used:
        name = f"{base}_{i}"
        i += 1
    used.add(name)
    return name


# ── Extracción con pdfplumber ─────────────────────────────────────────────────

# Configuraciones de extracción (se prueban en orden)
TABLE_SETTINGS = [
    # 1. Líneas estrictas — mejor para tablas con bordes
    {
        "vertical_strategy":   "lines_strict",
        "horizontal_strategy": "lines_strict",
        "snap_tolerance": 4,
        "join_tolerance": 3,
        "edge_min_length": 3,
    },
    # 2. Líneas normales — tablas con bordes parciales
    {
        "vertical_strategy":   "lines",
        "horizontal_strategy": "lines",
        "snap_tolerance": 5,
        "join_tolerance": 4,
    },
    # 3. Texto — tablas separadas por espacios (facturas, listados, etc.)
    {
        "vertical_strategy":   "text",
        "horizontal_strategy": "text",
        "snap_tolerance": 3,
        "min_words_vertical": 2,
        "min_words_horizontal": 1,
    },
]


def extract_with_pdfplumber(file_bytes):
    """
    Extrae tablas con 2 pasadas:
    - Pasada 1: configuración estricta (líneas)
    - Pasada 2: si no hay tablas en la página, prueba modo texto
    Devuelve lista de (sheet_name, [[row_data]])
    """
    results   = []
    used_names = set()

    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for page_num, page in enumerate(pdf.pages, 1):
            page_tables = []

            # Prueba cada configuración
            for settings in TABLE_SETTINGS:
                tables = page.extract_tables(settings)
                if tables:
                    # Filtrar tablas vacías
                    valid = [
                        t for t in tables
                        if any(any(c for c in row if c) for row in t)
                    ]
                    if valid:
                        page_tables = valid
                        break

            if page_tables:
                for t_idx, table in enumerate(page_tables):
                    # Limpiar None y normalizar
                    data = [
                        [str(c).strip() if c is not None else "" for c in row]
                        for row in table
                        if any(c for c in row if c)
                    ]
                    if not data:
                        continue

                    suffix = f"_T{t_idx+1}" if len(page_tables) > 1 else ""
                    name = _safe_name(f"Pag{page_num}{suffix}", used_names)
                    results.append((name, data))
            else:
                # Sin tablas: extraer texto como columna única
                text = page.extract_text() or ""
                lines = [l.strip() for l in text.split("\n") if l.strip()]
                if lines:
                    # Intentar detectar si el texto parece una tabla (columnas por espacios)
                    data = _text_to_rows(lines)
                    name = _safe_name(f"Pag{page_num}_texto", used_names)
                    results.append((name, data))

    return results


def _text_to_rows(lines):
    """
    Intenta estructurar líneas de texto en filas/columnas.
    Si las líneas tienen múltiples palabras separadas por ≥3 espacios,
    las divide en columnas. Si no, las deja como una sola columna.
    """
    import re
    # Detectar si hay separadores de columna (≥2 espacios consecutivos)
    has_multi_col = any(re.search(r'  +', line) for line in lines[:10])

    if has_multi_col:
        rows = []
        for line in lines:
            cols = re.split(r'  +', line)
            rows.append([c.strip() for c in cols if c.strip()])
        # Normalizar número de columnas
        max_c = max(len(r) for r in rows) if rows else 1
        rows = [r + [""] * (max_c - len(r)) for r in rows]
        return rows
    else:
        return [[line] for line in lines]


def extract_fallback_pdfminer(file_bytes):
    """Fallback de texto puro con pdfminer."""
    text = pm_extract_text(io.BytesIO(file_bytes)) or ""
    lines = [l.strip() for l in text.split("\n") if l.strip()]
    if not lines:
        return []
    return [("Texto", [[l] for l in lines])]


# ── Conversión principal ──────────────────────────────────────────────────────

def pdf_to_xlsx(file_bytes):
    if not OPENPYXL_OK:
        raise RuntimeError("openpyxl no está instalado")

    if PDFPLUMBER_OK:
        sheets = extract_with_pdfplumber(file_bytes)
    elif PDFMINER_OK:
        sheets = extract_fallback_pdfminer(file_bytes)
    else:
        raise RuntimeError(
            "Instala pdfplumber: pip install pdfplumber openpyxl"
        )

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    if not sheets:
        ws = wb.create_sheet("Sin_tablas")
        ws["A1"] = "No se encontraron tablas en este PDF."
        ws["A1"].font = Font(italic=True, color="888888", name="Calibri")
        ws.column_dimensions["A"].width = 45
    else:
        for sheet_name, data in sheets:
            ws = wb.create_sheet(title=sheet_name)
            _style_ws(ws, data)

    buf = io.BytesIO()
    wb.save(buf)
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

            xlsx_bytes = pdf_to_xlsx(file_bytes)
            out_name   = filename.rsplit(".", 1)[0] + ".xlsx"

            self.send_response(200)
            self.send_header(
                "Content-Type",
                "application/vnd.openxmlformats-officedocument"
                ".spreadsheetml.sheet",
            )
            self.send_header("Content-Disposition",
                             f'attachment; filename="{out_name}"')
            self.send_header("Content-Length", str(len(xlsx_bytes)))
            self._cors()
            self.end_headers()
            self.wfile.write(xlsx_bytes)

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
