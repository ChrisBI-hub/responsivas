import argparse
from pathlib import Path
import re

import pyodbc
from jinja2 import Environment, FileSystemLoader
from weasyprint import HTML


BASE_DIR = Path(__file__).resolve().parent
TEMPLATE_NAME = "index.html"
OUTPUT_DIR = BASE_DIR / "responsivas_pdf"

SQL_CONFIG = {
    "server": "150.1.1.152",
    "database": "BI",
    "username": "ConsultaBD",
    "password": "5D$bc#kM&5W2T8J40?s%",
    "driver": "{ODBC Driver 17 for SQL Server}",
}


def build_test_document():
    return {
        "empresa": "ABC",
        "Empresa": "ABC LOGISTICA",
        "Fecha_Creacion": "2026-06-08",
        "Fecha_Modificacion": "2026-06-08",
        "Colaborador": "USUARIO DE PRUEBA",
        "puesto": "ANALISTA DE OPERACIONES",
        "Sub_área": "INFRAESTRUCTURA Y CIBERSEGURIDAD",
        "usuario_red": "uprueba",
        "Host": "CPU-PRUEBA",
        "No_Serie": "SERIE123",
        "Estado": "BUENO",
        "Marca": "DELL",
        "Modelo": "OPTIPLEX 7090",
        "Procesador": "INTEL CORE I5",
        "RAM": "16 GB",
        "Capacidad_Disco": "512 GB SSD",
        "proyecto": "ABC",
        "routing": "RT-001",
        "Observaciones": "Equipo de prueba para validar el formato PDF.",
        "SO": "Windows 11 Pro",
        "Office": "Microsoft 365",
        "Lector_PDF": "Adobe Acrobat Reader",
        "Antivirus": "Microsoft Defender",
        "ERP": "N/A",
        "Otro_1": "VPN corporativa",
        "Otro_2": "Navegador corporativo",
        "Otro_3": "Herramientas internas",
        "hardware": [
            {
                "Tipo": "Monitor",
                "No_Serie": "MON123456",
                "Marca": "Dell",
                "Modelo": "P2422H",
                "Estado": "Bueno",
                "Observaciones": "Sin observaciones",
            },
            {
                "Tipo": "Teclado",
                "No_Serie": "TEC123456",
                "Marca": "Logitech",
                "Modelo": "K120",
                "Estado": "Bueno",
                "Observaciones": "Sin observaciones",
            },
            {
                "Tipo": "Mouse",
                "No_Serie": "MOU123456",
                "Marca": "Logitech",
                "Modelo": "M90",
                "Estado": "Bueno",
                "Observaciones": "Sin observaciones",
            },
        ],
    }


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


def load_sql(filename):
    sql = (BASE_DIR / filename).read_text(encoding="utf-8").strip()
    if not sql:
        raise ValueError(f"El archivo {filename} esta vacio.")
    return sql.rstrip().rstrip(";")


def query_to_dicts(cursor, sql, params=()):
    cursor.execute(sql, params)
    columns = [column[0] for column in cursor.description]
    return [dict(zip(columns, row)) for row in cursor.fetchall()]


def normalize_key(name):
    return re.sub(r"[^a-z0-9]", "", str(name).lower())


def build_lookup(data):
    return {normalize_key(key): value for key, value in data.items()}


def first_value(lookup, *candidates, default=""):
    for candidate in candidates:
        key = normalize_key(candidate)
        if key in lookup and lookup[key] not in (None, ""):
            return lookup[key]
    return default


def merge_aliases(cpu_row, software_row):
    context = {}
    context.update(cpu_row)
    context.update(software_row)

    lookup = build_lookup(context)
    aliases = {
        "empresa": ("Empresa", "empresa"),
        "Empresa": ("Empresa", "empresa"),
        "Fecha_Creacion": ("Fecha_Creacion", "fecha_creacion"),
        "Fecha_Modificacion": ("Fecha_Modificacion", "fecha_modificacion", "Timestamp"),
        "Colaborador": ("Colaborador", "colaborador", "Usuario", "Nombre_Usuario"),
        "puesto": ("puesto", "Puesto"),
        "Sub_área": ("Sub_área", "Sub_area", "Departamento", "departamento"),
        "usuario_red": ("usuario_red", "Usuario_Red", "usuario"),
        "Host": ("Host", "Host_CPU", "host"),
        "No_Serie": ("No_Serie", "serie", "Serial", "Numero_Serie"),
        "Estado": ("Estado", "estado"),
        "Marca": ("Marca", "marca"),
        "Modelo": ("Modelo", "modelo"),
        "Procesador": ("Procesador", "procesador", "CPU"),
        "RAM": ("RAM", "ram", "Memoria_RAM"),
        "Capacidad_Disco": ("Capacidad_Disco", "disco", "Disco", "Disco_Duro"),
        "proyecto": ("proyecto", "Proyecto"),
        "routing": ("routing", "Routing", "No_Routing"),
        "Observaciones": ("Observaciones", "obs_cpu", "Observacion"),
        "SO": ("SO", "Sistema_Operativo", "so"),
        "Office": ("Office", "Ofimatica", "ofimatica"),
        "Lector_PDF": ("Lector_PDF", "PDF", "pdf"),
        "Antivirus": ("Antivirus", "antivirus"),
        "ERP": ("ERP", "erp"),
        "Otro_1": ("Otro_1", "otros"),
        "Otro_2": ("Otro_2",),
        "Otro_3": ("Otro_3",),
    }

    for target, candidates in aliases.items():
        if target not in context or context[target] in (None, ""):
            context[target] = first_value(lookup, *candidates)

    return context


def build_filtered_query(sql, filter_column):
    return f"SELECT * FROM ({sql}) AS src WHERE [{filter_column}] = ?"


def fetch_data():
    cpu_sql = load_sql("cpu.sql")
    software_sql = load_sql("software.sql")
    periferico_sql = load_sql("periferico.sql")

    conn = get_connection()
    try:
        cursor = conn.cursor()
        cpu_rows = query_to_dicts(cursor, cpu_sql)
        documentos = []

        for cpu_row in cpu_rows:
            host = cpu_row.get("Host") or cpu_row.get("Host_CPU")
            if not host:
                continue

            software_rows = query_to_dicts(
                cursor,
                build_filtered_query(software_sql, "Host_CPU"),
                (host,),
            )
            perifericos = query_to_dicts(
                cursor,
                build_filtered_query(periferico_sql, "Host_CPU"),
                (host,),
            )

            software_row = software_rows[0] if software_rows else {}
            context = merge_aliases(cpu_row, software_row)
            context["hardware"] = perifericos
            documentos.append(context)

        return documentos
    finally:
        conn.close()


def safe_filename(value):
    text = str(value or "sin_nombre").strip()
    text = re.sub(r"[^\w.-]+", "_", text, flags=re.UNICODE)
    return text[:120] or "sin_nombre"


def render_pdfs(documentos):
    env = Environment(loader=FileSystemLoader(str(BASE_DIR)))
    template = env.get_template(TEMPLATE_NAME)
    OUTPUT_DIR.mkdir(exist_ok=True)

    generated_files = []
    for context in documentos:
        host = context.get("Host") or context.get("Host_CPU")
        serie = context.get("No_Serie")
        filename = safe_filename(f"{host}_{serie}.pdf")
        output_path = OUTPUT_DIR / filename

        html = template.render(context)
        HTML(string=html, base_url=str(BASE_DIR)).write_pdf(output_path)
        generated_files.append(output_path)

    return generated_files


def main():
    parser = argparse.ArgumentParser(description="Genera responsivas PDF.")
    parser.add_argument(
        "--prueba",
        action="store_true",
        help="Genera un PDF de prueba sin conectarse a SQL Server.",
    )
    args = parser.parse_args()

    if args.prueba:
        print("Generando PDF de prueba...")
        generated_files = render_pdfs([build_test_document()])
        print(f"PDF de prueba generado: {generated_files[0]}")
        return

    print("Generando responsivas PDF...")
    try:
        documentos = fetch_data()
    except pyodbc.Error as exc:
        print(f"Error de conexion o consulta SQL: {exc}")
        return
    except Exception as exc:
        print(f"Error generando responsivas: {exc}")
        return

    if not documentos:
        print("No se encontraron CPUs para generar responsivas.")
        return

    generated_files = render_pdfs(documentos)
    print(f"PDF generados: {len(generated_files)}")
    print(f"Directorio de salida: {OUTPUT_DIR}")


if __name__ == "__main__":
    main()
