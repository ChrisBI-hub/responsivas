"""
Generador de Etiquetas QR para CPU
==================================
- Conecta a SQL Server y obtiene solo equipos CPU
- Genera un código QR por CPU
- Agrupa las etiquetas en hojas tamaño carta para imprimir y recortar
"""

import qrcode
import pyodbc
import json
import io
import os
from PIL import Image
from reportlab.lib.pagesizes import LETTER
from reportlab.lib.units import cm, mm
from reportlab.pdfgen import canvas
from reportlab.lib import colors
from reportlab.lib.utils import ImageReader

# ─────────────────────────────────────────────
#  CONFIGURACIÓN
# ─────────────────────────────────────────────
SQL_CONFIG = {
    "server": "150.1.1.152",
    "database": "BI",
    "username": "ConsultaBD",
    "password": "5D$bc#kM&5W2T8J40?s%",
    "driver": "{ODBC Driver 17 for SQL Server}",
}

TABLAS_SQL = {
    "CPU": {
        "tabla_principal":   "Inventario.CPU",
        "tabla_software":    "Inventario.CPU_Software",
        "tabla_perifericos": "Inventario.CPU_Perifericos",
        "tabla_relaciones":  "Inventario.CPU_Relaciones",
    },
}

OUTPUT_DIR = "qr_etiquetas_output"
PDF_SALIDA = os.path.join(OUTPUT_DIR, "etiquetas_inventario.pdf")

# ── Medidas CPU (etiqueta con QR) ─────────────────────
QR_SIZE_CM    = 4.0   # QR cuadrado
CPU_LBL_W_CM  = 4.2   # Ancho total etiqueta CPU
CPU_LBL_H_CM  = 4.2   # Alto total etiqueta CPU (QR + texto)

# ── Layout de página ──────────────────────────────────
MARGIN_CM     = 1.0   # Márgenes de página
GAP_CM        = 0.5  # Espacio entre etiquetas


# ─────────────────────────────────────────────
#  CONEXIÓN A BASE DE DATOS
# ─────────────────────────────────────────────
def get_connection():
    conn_str = (
        f"DRIVER={SQL_CONFIG['driver']};"
        f"SERVER={SQL_CONFIG['server']};"
        f"DATABASE={SQL_CONFIG['database']};"
        f"UID={SQL_CONFIG['username']};"
        f"PWD={SQL_CONFIG['password']};"
        "TrustServerCertificate=yes;"
    )
    return pyodbc.connect(conn_str, timeout=10)


def query_to_dicts(cursor, sql):
    cursor.execute(sql)
    cols = [d[0] for d in cursor.description]
    return [dict(zip(cols, row)) for row in cursor.fetchall()]


# ─────────────────────────────────────────────
#  OBTENER DATOS DE EQUIPOS
# ─────────────────────────────────────────────
def obtener_equipos_cpu(conn):
    cur = conn.cursor()
    equipos = []

    registros = query_to_dicts(cur, f"SELECT * FROM {TABLAS_SQL['CPU']['tabla_principal']}")
    if not registros:
        return equipos

    primer = registros[0]
    id_field = next(
        (k for k in primer if "id" in k.lower() or "serial" in k.lower()),
        list(primer.keys())[0]
    )

    for reg in registros:
        equipo_id = reg.get(id_field)
        datos = {"TIPO": "CPU"}
        datos.update({k: v for k, v in reg.items() if v is not None and str(v).strip() != ""})

        # Software
        try:
            sw = query_to_dicts(
                cur,
                f"SELECT * FROM {TABLAS_SQL['CPU']['tabla_software']} WHERE {id_field} = '{equipo_id}'"
            )
            if sw:
                datos["Software"] = [list(s.values()) for s in sw]
        except Exception:
            pass

        # Periféricos
        try:
            per = query_to_dicts(
                cur,
                f"SELECT * FROM {TABLAS_SQL['CPU']['tabla_perifericos']} WHERE {id_field} = '{equipo_id}'"
            )
            if per:
                datos["Perifericos"] = [list(p.values()) for p in per]
        except Exception:
            pass

        # Relaciones
        try:
            rel = query_to_dicts(
                cur,
                f"SELECT * FROM {TABLAS_SQL['CPU']['tabla_relaciones']} WHERE {id_field} = '{equipo_id}'"
            )
            if rel:
                datos["Relaciones"] = [list(r.values()) for r in rel]
        except Exception:
            pass

        equipos.append(datos)

    return equipos


def dict_to_qr_text(datos: dict) -> str:
    """Convierte el dict del equipo a texto compacto para el QR."""
    lines = []
    for k, v in datos.items():
        if isinstance(v, list):
            lines.append(f"{k}: {json.dumps(v, ensure_ascii=False, default=str)}")
        else:
            lines.append(f"{k}: {v}")
    return "\n".join(lines)


def generar_qr_imagen(texto: str, size_px: int = 350) -> Image.Image:
    qr = qrcode.QRCode(
        version=None,
        error_correction=qrcode.constants.ERROR_CORRECT_M,
        box_size=10,
        border=1,
    )
    qr.add_data(texto)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white").convert("RGB")
    img = img.resize((size_px, size_px), Image.LANCZOS)
    return img


def extraer_nombre_equipo(datos: dict) -> str:
    """Intenta sacar un nombre/ID corto para mostrar en la etiqueta."""
    for clave in ["Nombre", "Hostname", "NombreEquipo", "nombre", "hostname",
                  "NOMBRE", "HOSTNAME", "Serial", "serial", "ID", "id", "Equipo"]:
        if clave in datos and datos[clave]:
            return str(datos[clave])[:30]
    for v in datos.values():
        if v and isinstance(v, (str, int)):
            return str(v)[:30]
    return "SIN NOMBRE"


def dibujar_guias_corte(c, x, y, lbl_w, lbl_h):
    """Dibuja el borde gris y las líneas de corte en las 4 esquinas."""
    cut = 3 * mm

    # Borde de etiqueta
    c.setStrokeColor(colors.HexColor("#BBBBBB"))
    c.setLineWidth(0.4)
    c.rect(x, y, lbl_w, lbl_h, fill=0, stroke=1)

    # Marcas de esquina
    c.setStrokeColor(colors.HexColor("#999999"))
    c.setLineWidth(0.25)
    for sx, sy in [(x, y + lbl_h), (x + lbl_w, y + lbl_h),
                   (x, y),         (x + lbl_w, y)]:
        dx = cut if sx == x else -cut
        c.line(sx + dx, sy, sx, sy)
        dy = -cut if sy == y + lbl_h else cut
        c.line(sx, sy + dy, sx, sy)


# ─────────────────────────────────────────────
#  GENERAR PDF EN HOJAS CARTA
# ─────────────────────────────────────────────
def generar_pdf(equipos_cpu: list):
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    page_w, page_h = LETTER  # puntos
    margin = MARGIN_CM * cm
    gap    = GAP_CM    * cm

    # ── Métricas CPU ──────────────────────────────────────
    cpu_lbl_w = CPU_LBL_W_CM * cm
    cpu_lbl_h = CPU_LBL_H_CM * cm
    qr_size   = QR_SIZE_CM   * cm

    cpu_cols     = int((page_w - 2 * margin + gap) / (cpu_lbl_w + gap))
    cpu_rows     = int((page_h - 2 * margin + gap) / (cpu_lbl_h + gap))
    cpu_per_page = cpu_cols * cpu_rows

    c = canvas.Canvas(PDF_SALIDA, pagesize=LETTER)

    total_cpu  = len(equipos_cpu)

    print(f"\n  ── Equipos CPU  ──────────────────────────")
    print(f"  Total:            {total_cpu}")
    print(f"  Etiquetas/hoja:   {cpu_per_page}  ({cpu_cols} col × {cpu_rows} fil)")
    print(f"  Páginas:          {-(-total_cpu // cpu_per_page) if total_cpu else 0}")
    print()

    # ══════════════════════════════════════════════════════
    #  SECCIÓN 1: CPU  →  Código QR
    # ══════════════════════════════════════════════════════
    primera_pagina = True
    for idx, datos in enumerate(equipos_cpu):
        page_idx = idx % cpu_per_page

        if page_idx == 0:
            if not primera_pagina:
                c.showPage()
            else:
                primera_pagina = False

        col = page_idx % cpu_cols
        row = page_idx // cpu_cols

        x = margin + col * (cpu_lbl_w + gap)
        y = page_h - margin - (row + 1) * cpu_lbl_h - row * gap

        dibujar_guias_corte(c, x, y, cpu_lbl_w, cpu_lbl_h)

        # ── QR ──────────────────────────────────────────
        texto_qr = dict_to_qr_text(datos)
        img_qr   = generar_qr_imagen(texto_qr, size_px=350)

        qr_x = x + (cpu_lbl_w - qr_size) / 2
        qr_y = y + cpu_lbl_h - qr_size - (1 * mm)

        buf = io.BytesIO()
        img_qr.save(buf, format="PNG")
        buf.seek(0)
        c.drawImage(ImageReader(buf), qr_x, qr_y, width=qr_size, height=qr_size)

        # ── Texto inferior ───────────────────────────────
        nombre = extraer_nombre_equipo(datos)

        c.setStrokeColor(colors.HexColor("#DDDDDD"))
        c.setLineWidth(0.3)
        c.line(x + 2*mm, qr_y - 1*mm, x + cpu_lbl_w - 2*mm, qr_y - 1*mm)

        c.setFont("Helvetica", 5)
        c.setFillColor(colors.HexColor("#888888"))
        c.drawCentredString(x + cpu_lbl_w / 2, qr_y - 5*mm, "CPU")

        c.setFont("Helvetica-Bold", 6.5)
        c.setFillColor(colors.black)
        c.drawCentredString(x + cpu_lbl_w / 2, qr_y - 9.5*mm, nombre)

        if (idx + 1) % 10 == 0 or idx + 1 == total_cpu:
            print(f"  [CPU] ✓ {idx + 1}/{total_cpu} procesados...", end="\r")

    if total_cpu:
        print()

    c.save()
    print(f"\n  ✅ PDF generado: {PDF_SALIDA}")


# ─────────────────────────────────────────────
#  MAIN
# ─────────────────────────────────────────────
def main():
    print("=" * 55)
    print("  GENERADOR DE ETIQUETAS QR - CPU")
    print("  CPU → QR")
    print("=" * 55)

    print("\n[1/3] Conectando a SQL Server...")
    try:
        conn = get_connection()
        print("  ✅ Conexión exitosa")
    except Exception as e:
        print(f"  ❌ Error de conexión: {e}")
        return

    print("\n[2/3] Obteniendo equipos...")
    equipos_cpu = []

    try:
        equipos_cpu = obtener_equipos_cpu(conn)
        print(f"  • CPU:           {len(equipos_cpu)} equipos")
    except Exception as e:
        print(f"  ⚠ Error leyendo CPU: {e}")

    conn.close()

    if not equipos_cpu:
        print("\n  ❌ No se encontraron equipos CPU. Verifica la tabla.")
        return

    print("\n[3/3] Generando PDF con etiquetas...")
    generar_pdf(equipos_cpu)

    print("\n" + "=" * 55)
    print(f"  Archivo listo: {PDF_SALIDA}")
    print("  Imprime en tamaño carta, sin escalar (100%)")
    print("  CPU:   etiqueta QR         2.5 × 2.5 cm")
    print("=" * 55)


if __name__ == "__main__":
    main()
