"""Consulta masiva de medidas correctivas - Policía Nacional (RNMC).

El sitio no tiene CAPTCHA: la aplicación llena tipo de documento, cédula y
fecha de expedición, consulta, guarda el certificado en PDF y continúa.
"""

from __future__ import annotations

# Certificados CA en macOS para undetected-chromedriver.
import os as _os

try:
    import certifi as _certifi

    _os.environ.setdefault("SSL_CERT_FILE", _certifi.where())
    _os.environ.setdefault("REQUESTS_CA_BUNDLE", _certifi.where())
except Exception:
    pass

import argparse
import base64
import csv
import json
import logging
import platform
import time
from datetime import datetime
from pathlib import Path

from selenium.common.exceptions import TimeoutException, WebDriverException
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select, WebDriverWait

from config import Settings, load_settings
from sources import load_users
from utils import configure_console, safe_filename

configure_console()

logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

MEDIDAS_URL = "https://srvcnpc.policia.gov.co/PSC/frm_cnp_consulta.aspx"

ID_TIPO_DOC = "ctl00_ContentPlaceHolder3_ddlTipoDoc"
ID_CEDULA = "ctl00_ContentPlaceHolder3_txtExpediente"
ID_FECHA_EXP = "txtFechaexp"
ID_BTN_CONSULTAR = "ctl00_ContentPlaceHolder3_btnConsultar2"
ID_BTN_NUEVA = "ctl00_ContentPlaceHolder3_btnNuevo2"
ID_BTN_IMPRIMIR = "ctl00_ContentPlaceHolder3_btnImprimir2"
ID_LOADER = "ctl00_ContentPlaceHolder3_UpdateProgress1"

CHROME_UA = (
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/154.0.0.0 Safari/537.36"
)
if platform.system() == "Windows":
    CHROME_UA = (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/154.0.0.0 Safari/537.36"
    )

PROJECT_DIR = Path(__file__).resolve().parent

PROFILE_FOLDER_MAP = {
    "AUXILIAR": "AUXILIAR",
    "ENFERMERO": "ENFERMERO",
    "MEDICO": "MEDICO",
}

SETTINGS: Settings = load_settings()

EXCEL_PATH = Path(SETTINGS.excel_path)
BASE_OUTPUT_DIR = Path(SETTINGS.output_dir)
STATE_FILE = Path(SETTINGS.state_file)
TEMP_DIR = Path(SETTINGS.temp_dir)
ERROR_DIR = Path(SETTINGS.error_dir)

LAST_STATS: dict = {}
_STOP_CHECKER = lambda: False


def set_stop_checker(checker=None) -> None:
    global _STOP_CHECKER
    _STOP_CHECKER = checker or (lambda: False)


def stop_requested() -> bool:
    try:
        return bool(_STOP_CHECKER())
    except Exception:
        return False


def apply_settings(settings: Settings) -> Settings:
    global SETTINGS, EXCEL_PATH, BASE_OUTPUT_DIR, STATE_FILE, TEMP_DIR, ERROR_DIR

    SETTINGS = settings
    EXCEL_PATH = Path(settings.excel_path)
    BASE_OUTPUT_DIR = Path(settings.output_dir)
    STATE_FILE = Path(settings.state_file)
    TEMP_DIR = Path(settings.temp_dir)
    ERROR_DIR = Path(settings.error_dir)
    return SETTINGS


# ---------------------------------------------------------------------------
# Estado
# ---------------------------------------------------------------------------

def load_state():
    if not STATE_FILE.exists():
        return {"procesados": [], "fallidos": [], "ultima_actualizacion": None}
    try:
        with open(STATE_FILE, "r", encoding="utf-8") as file:
            data = json.load(file)
            data.setdefault("procesados", [])
            data.setdefault("fallidos", [])
            return data
    except Exception:
        return {"procesados": [], "fallidos": [], "ultima_actualizacion": None}


def save_state(state):
    state["ultima_actualizacion"] = datetime.now().isoformat(timespec="seconds")
    STATE_FILE.parent.mkdir(parents=True, exist_ok=True)
    with open(STATE_FILE, "w", encoding="utf-8") as file:
        json.dump(state, file, ensure_ascii=False, indent=2)


# ---------------------------------------------------------------------------
# Carpetas y navegador
# ---------------------------------------------------------------------------

def configure_folders():
    TEMP_DIR.mkdir(parents=True, exist_ok=True)
    ERROR_DIR.mkdir(parents=True, exist_ok=True)
    BASE_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    for folder in PROFILE_FOLDER_MAP.values():
        (BASE_OUTPUT_DIR / folder).mkdir(parents=True, exist_ok=True)


def create_driver(headless: bool = True, backend: str | None = None):
    backend = (backend or SETTINGS.driver_backend or "auto").strip().lower()
    args = [
        "--no-sandbox",
        "--disable-dev-shm-usage",
        "--disable-gpu",
        "--window-size=1400,1000",
        "--lang=es-CO",
    ]
    if headless:
        args.append("--headless=new")

    if backend in ("auto", "uc"):
        try:
            import undetected_chromedriver as uc

            options = uc.ChromeOptions()
            for arg in args:
                options.add_argument(arg)
            options.add_argument(f"--user-agent={CHROME_UA}")
            driver = uc.Chrome(options=options)
            logging.info("Driver: undetected-chromedriver")
            return driver
        except Exception as exc:
            if backend == "uc":
                raise
            logging.warning("undetected-chromedriver no disponible (%s). Usando Selenium...", str(exc)[:160])

    from selenium import webdriver

    options = webdriver.ChromeOptions()
    for arg in args:
        options.add_argument(arg)
    options.add_argument(f"--user-agent={CHROME_UA}")
    driver = webdriver.Chrome(options=options)
    logging.info("Driver: Selenium estándar")
    return driver


# ---------------------------------------------------------------------------
# Formulario
# ---------------------------------------------------------------------------

def ensure_form(driver, wait):
    if driver.find_elements(By.ID, ID_TIPO_DOC):
        return True
    driver.get(MEDIDAS_URL)
    wait.until(EC.presence_of_element_located((By.ID, ID_TIPO_DOC)))
    return True


def select_document_type(driver, wait, tipo: str):
    select = Select(wait.until(EC.element_to_be_clickable((By.ID, ID_TIPO_DOC))))
    select.select_by_value(tipo or "55")
    wait.until(EC.visibility_of_element_located((By.ID, ID_FECHA_EXP)))


def fill_form(driver, wait, cedula: str, fecha: str):
    cedula_input = wait.until(EC.visibility_of_element_located((By.ID, ID_CEDULA)))
    driver.execute_script("arguments[0].value=''; arguments[0].focus();", cedula_input)
    cedula_input.send_keys(cedula)
    if cedula_input.get_attribute("value") != cedula:
        driver.execute_script(
            "arguments[0].value=arguments[1];"
            "arguments[0].dispatchEvent(new Event('input', {bubbles:true}));"
            "arguments[0].dispatchEvent(new Event('change', {bubbles:true}));",
            cedula_input, cedula,
        )

    fecha_input = wait.until(EC.visibility_of_element_located((By.ID, ID_FECHA_EXP)))
    driver.execute_script("arguments[0].value=''; arguments[0].focus();", fecha_input)
    fecha_input.send_keys(fecha)
    if fecha_input.get_attribute("value") != fecha:
        driver.execute_script(
            "arguments[0].value=arguments[1];"
            "arguments[0].dispatchEvent(new Event('input', {bubbles:true}));"
            "arguments[0].dispatchEvent(new Event('change', {bubbles:true}));",
            fecha_input, fecha,
        )


def click_consultar(driver, wait):
    button = wait.until(EC.element_to_be_clickable((By.ID, ID_BTN_CONSULTAR)))
    button.click()


def _clean_site_error(body: str) -> str:
    """Extrae solo el mensaje del modal de error, sin toda la página."""
    text = " ".join(str(body or "").split())
    if "Error" in text and "Aceptar" in text:
        start = text.find("Error") + len("Error")
        end = text.find("Aceptar", start)
        if end > start:
            return text[start:end].strip(" :.-")
    return text[:300]


def wait_result(driver, wait, documento: str, timeout: int = 35):
    """Espera el certificado y devuelve (estado, detalle)."""
    deadline = time.time() + timeout
    while time.time() < deadline:
        if stop_requested():
            return "detenido", "El usuario detuvo el proceso."

        try:
            body = driver.find_element(By.TAG_NAME, "body").text
        except Exception:
            body = ""

        # Errores del sitio (modal de ASP.NET). Van antes del resultado porque
        # la página de error conserva los botones Nueva Búsqueda / Imprimir.
        if "Debe registrar la fecha" in body or "no es correcta" in body.lower():
            return "fecha_invalida", _clean_site_error(body)
        if "no corresponde" in body.lower() or "no coincide" in body.lower():
            return "datos_no_coinciden", _clean_site_error(body)
        if "Error" in body and "Aceptar" in body:
            return "error_sitio", _clean_site_error(body)

        tiene_resultado = bool(
            driver.find_elements(By.ID, ID_BTN_NUEVA)
            or driver.find_elements(By.ID, ID_BTN_IMPRIMIR)
        )
        if tiene_resultado and documento in body:
            return "ok", body[:300]

        if "no se encontr" in body.lower() and documento in body:
            return "sin_resultado", body[:300]

        time.sleep(0.8)

    return "sin_resultado", "No apareció el certificado en el tiempo esperado."


# ---------------------------------------------------------------------------
# Salida
# ---------------------------------------------------------------------------

def build_filename(user, pattern: str) -> str:
    doc = user["documento"]
    nombre = user["nombre"]
    apellido = user["apellido"]

    if pattern == "documento_primero":
        stem = f"{doc} {nombre} {apellido}"
    else:
        stem = f"{apellido} {nombre} - {doc}"

    return f"{safe_filename(stem, max_stem=120)}.pdf"


def build_output_path(user, pattern: str) -> Path:
    folder = BASE_OUTPUT_DIR / PROFILE_FOLDER_MAP[user["perfil_grupo"]]
    folder.mkdir(parents=True, exist_ok=True)
    target = folder / build_filename(user, pattern)

    if not target.exists():
        return target

    stem, suffix = target.stem, target.suffix
    i = 2
    while True:
        candidate = folder / f"{stem} ({i}){suffix}"
        if not candidate.exists():
            return candidate
        i += 1


def save_pdf(driver, user, pattern: str) -> Path:
    target = build_output_path(user, pattern)

    driver.execute_script("""
        var style = document.getElementById('print-fix');
        if (!style) {
            style = document.createElement('style');
            style.id = 'print-fix';
            document.head.appendChild(style);
        }
        style.innerHTML = `
            @media print {
                body { background: #fff !important; }
                .panel-heading.panel-titulo {
                    background: #fff !important;
                    color: #142749 !important;
                }
                #ctl00_ContentPlaceHolder3_btnNuevo2,
                #ctl00_ContentPlaceHolder3_btnImprimir2,
                .btn { display: none !important; }
            }
        `;
    """)

    result = driver.execute_cdp_cmd(
        "Page.printToPDF",
        {
            "printBackground": True,
            "landscape": False,
            "displayHeaderFooter": False,
        },
    )
    pdf_data = base64.b64decode(result["data"])
    if len(pdf_data) < 5000:
        raise RuntimeError("El PDF salió demasiado pequeño; probablemente la página no cargó.")
    target.write_bytes(pdf_data)
    return target


def save_error_screenshot(driver, user, label="error"):
    try:
        name = f"{label}_{user.get('documento', 'sin_doc')}_{int(time.time())}.png"
        path = ERROR_DIR / name
        driver.save_screenshot(str(path))
        return str(path)
    except Exception:
        return ""


# ---------------------------------------------------------------------------
# Flujo por persona
# ---------------------------------------------------------------------------

def process_user(driver, wait, user, settings: Settings, on_status=None):
    attempts = []
    last_reason, last_detail = "sin_resultado", ""

    for attempt in range(1, max(1, settings.max_attempts) + 1):
        if stop_requested():
            return False, attempts, "detenido", "El usuario detuvo el proceso.", ""

        try:
            driver.get(MEDIDAS_URL)
            ensure_form(driver, wait)

            select_document_type(driver, wait, user.get("tipo_documento", "55"))
            fill_form(driver, wait, user["documento"], user["fecha_expedicion"])

            click_consultar(driver, wait)
            status, detail = wait_result(driver, wait, user["documento"])
            attempts.append({"intento": attempt, "resultado": status, "detalle": detail})

            if status == "ok":
                path = save_pdf(driver, user, settings.filename_pattern)
                logging.info(
                    "[OK] Guardado: %s | Perfil: %s | Doc: %s (intento %s)",
                    path.name, user["perfil_grupo"], user["documento"], attempt,
                )
                return True, attempts, "", "", str(path)

            last_reason, last_detail = status, detail
            logging.warning(
                "[WARN] Intento %s/%s falló para %s: %s",
                attempt, settings.max_attempts, user["documento"], detail,
            )

            if status in ("datos_no_coinciden", "fecha_invalida"):
                break

        except (TimeoutException, WebDriverException) as exc:
            last_reason = "error_navegador"
            last_detail = f"{type(exc).__name__}: {str(exc)[:200]}"
            attempts.append({"intento": attempt, "resultado": last_reason, "detalle": last_detail})
            logging.error("Error de navegador con %s: %s", user["documento"], last_detail)
        except Exception as exc:
            last_reason = "error_inesperado"
            last_detail = f"{type(exc).__name__}: {str(exc)[:200]}"
            attempts.append({"intento": attempt, "resultado": last_reason, "detalle": last_detail})
            logging.error("Error inesperado con %s: %s", user["documento"], last_detail)

    screenshot = save_error_screenshot(driver, user, label=last_reason)
    return False, attempts, last_reason, last_detail, screenshot


# ---------------------------------------------------------------------------
# Informe
# ---------------------------------------------------------------------------

MOTIVOS_REPORTE = {
    "dato_incompleto": (
        "Faltan datos en el Excel.",
        "Completa la fila y vuelve a intentar.",
    ),
    "fecha_invalida": (
        "La fecha de expedición no es correcta.",
        "Corrige la fecha en la columna Expedicion (DD/MM/AAAA) y vuelve a intentar.",
    ),
    "datos_no_coinciden": (
        "La fecha o el documento no coinciden con el registro.",
        "Revisa la cédula y la fecha de expedición en el Excel.",
    ),
    "sin_resultado": (
        "El sitio no mostró el certificado.",
        "Reintenta; si persiste, revisa la captura.",
    ),
    "error_sitio": (
        "El sitio mostró un error.",
        "Revisa el mensaje del sitio y la captura.",
    ),
    "error_navegador": (
        "Falló el navegador.",
        "Reintenta; puede ser un timeout o un bloqueo temporal.",
    ),
    "error_inesperado": (
        "Error inesperado.",
        "Revisa la captura y el detalle.",
    ),
    "detenido": (
        "Proceso detenido.",
        "Vuelve a iniciar para continuar.",
    ),
}


def _motivo_info(motivo: str):
    return MOTIVOS_REPORTE.get(motivo, ("Error no clasificado.", "Revisa el detalle y la captura."))


def write_failure_report(fallos, incomplete_rows, timestamp):
    if not fallos and not incomplete_rows:
        return None, None, None

    rows = []
    for fila in incomplete_rows:
        rows.append({
            "fila_excel": fila.get("fila", ""),
            "documento": fila.get("documento", ""),
            "nombre": fila.get("nombre", ""),
            "perfil": "",
            "fecha_expedicion": fila.get("fecha_expedicion", ""),
            "motivo": fila.get("motivo", "dato_incompleto"),
            "detalle": " ".join(str(fila.get("detalle", "")).split()),
            "intentos": 0,
            "captura": "",
        })

    for fallo in fallos:
        rows.append({
            "fila_excel": fallo.get("fila_excel", ""),
            "documento": fallo.get("documento", ""),
            "nombre": fallo.get("nombre", ""),
            "perfil": fallo.get("perfil", ""),
            "fecha_expedicion": fallo.get("fecha_expedicion", ""),
            "motivo": fallo.get("motivo", "sin_resultado"),
            "detalle": _clean_site_error(fallo.get("detalle", "")),
            "intentos": len(fallo.get("intentos", [])),
            "captura": fallo.get("captura", ""),
        })

    rows.sort(key=lambda r: (int(r["fila_excel"]) if str(r["fila_excel"]).isdigit() else 10**9))

    csv_path = BASE_OUTPUT_DIR / f"REPORTE_FALLOS_{timestamp}.csv"
    txt_path = BASE_OUTPUT_DIR / f"REPORTE_FALLOS_{timestamp}.txt"
    xlsx_path = BASE_OUTPUT_DIR / f"REPORTE_FALLOS_{timestamp}.xlsx"

    headers = [
        "Fila Excel", "Documento", "Nombre", "Perfil", "Fecha expedición",
        "Motivo", "Qué pasó", "Qué hacer", "Mensaje del sitio", "Intentos", "Captura",
    ]

    # CSV limpio: una fila por persona, sin saltos de línea dentro de las celdas.
    with open(csv_path, "w", newline="", encoding="utf-8-sig") as handle:
        writer = csv.writer(handle)
        writer.writerow(headers)
        for row in rows:
            explicacion, que_hacer = _motivo_info(row["motivo"])
            writer.writerow([
                row["fila_excel"], row["documento"], row["nombre"], row["perfil"],
                row["fecha_expedicion"], row["motivo"], explicacion, que_hacer,
                row["detalle"], row["intentos"], row["captura"],
            ])

    # TXT resumido, un bloque corto por persona.
    fecha_txt = timestamp
    try:
        fecha_txt = datetime.strptime(timestamp, "%Y%m%d_%H%M%S").strftime("%d/%m/%Y %H:%M")
    except Exception:
        pass

    from collections import Counter

    contador = Counter(row["motivo"] for row in rows)

    lineas = [
        "REPORTE DE FALLOS - MEDIDAS CORRECTIVAS (RNMC)",
        f"Generado: {fecha_txt}",
        f"Excel: {EXCEL_PATH}",
        f"Carpeta de salida: {BASE_OUTPUT_DIR}",
        "",
        f"Total de fallos: {len(rows)}",
        "",
        "RESUMEN",
    ]
    for motivo, cantidad in contador.most_common():
        explicacion, _ = _motivo_info(motivo)
        lineas.append(f"  {motivo}: {cantidad}  ({explicacion})")

    lineas += ["", "DETALLE (en el orden del Excel)", ""]
    for index, row in enumerate(rows, start=1):
        explicacion, que_hacer = _motivo_info(row["motivo"])
        lineas.append(
            f"{index}. Fila {row['fila_excel']} | {row['documento']} | "
            f"{row['nombre']} | {row['perfil'] or '(sin perfil)'}"
        )
        lineas.append(f"   Fecha expedición: {row['fecha_expedicion'] or '(vacía)'}")
        lineas.append(f"   Motivo: {row['motivo']}")
        lineas.append(f"   Qué pasó: {explicacion}")
        lineas.append(f"   Qué hacer: {que_hacer}")
        if row["detalle"]:
            lineas.append(f"   Mensaje del sitio: {row['detalle']}")
        lineas.append(f"   Intentos: {row['intentos']}")
        if row["captura"]:
            lineas.append(f"   Captura: {row['captura']}")
        lineas.append("")

    txt_path.write_text("\n".join(lineas) + "\n", encoding="utf-8")

    # Excel con tabla formateada y hoja de resumen.
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
        from openpyxl.utils import get_column_letter

        wb = Workbook()
        ws = wb.active
        ws.title = "Fallos"

        header_fill = PatternFill("solid", fgColor="142749")
        header_font = Font(color="FFFFFF", bold=True)
        thin = Side(style="thin", color="D9D9D9")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)

        ws.append(headers)
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = border

        for row in rows:
            explicacion, que_hacer = _motivo_info(row["motivo"])
            captura_nombre = Path(row["captura"]).name if row["captura"] else ""
            ws.append([
                row["fila_excel"], row["documento"], row["nombre"], row["perfil"],
                row["fecha_expedicion"], row["motivo"], explicacion, que_hacer,
                row["detalle"], row["intentos"], captura_nombre,
            ])
            row_index = ws.max_row
            for cell in ws[row_index]:
                cell.alignment = Alignment(vertical="top", wrap_text=True)
                cell.border = border
            if row["captura"]:
                cell = ws.cell(row=row_index, column=11)
                cell.hyperlink = "file://" + str(Path(row["captura"]).resolve())
                cell.style = "Hyperlink"

        for index, width in enumerate([9, 14, 34, 12, 15, 18, 45, 48, 60, 9, 45], start=1):
            ws.column_dimensions[get_column_letter(index)].width = width
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

        resumen = wb.create_sheet("Resumen")
        resumen.append(["Motivo", "Cantidad", "Qué significa", "Qué hacer"])
        for cell in resumen[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = border
        for motivo, cantidad in contador.most_common():
            explicacion, que_hacer = _motivo_info(motivo)
            resumen.append([motivo, cantidad, explicacion, que_hacer])
        for row in resumen.iter_rows(min_row=2):
            for cell in row:
                cell.alignment = Alignment(vertical="top", wrap_text=True)
                cell.border = border
        for index, width in enumerate([22, 10, 55, 60], start=1):
            resumen.column_dimensions[get_column_letter(index)].width = width
        resumen.freeze_panes = "A2"

        wb.save(xlsx_path)
    except Exception as exc:
        logging.warning("No se pudo generar el reporte Excel: %s", exc)
        xlsx_path = None

    return csv_path, txt_path, xlsx_path


# ---------------------------------------------------------------------------
# Lectura del Excel
# ---------------------------------------------------------------------------

def load_users_from_excel(excel_path, selected_profiles=None, sheet_name=None,
                          header_row=None, column_map=None):
    sheet = SETTINGS.sheet_name if sheet_name is None else sheet_name
    header = SETTINGS.header_row if header_row is None else header_row
    mapping = SETTINGS.column_map if column_map is None else column_map

    users, incomplete, stats = load_users(
        excel_path,
        sheet_name=sheet,
        header_row=header,
        column_map=mapping,
        selected_profiles=selected_profiles,
    )
    globals()["LAST_STATS"] = stats
    logging.info(
        "Columnas detectadas -> Documento: %r | Fecha: %r | Nombres: %r | Apellidos: %r | Perfil: %r",
        stats["columnas"].get("documento"), stats["columnas"].get("fecha"),
        stats["columnas"].get("nombre"), stats["columnas"].get("apellido"),
        stats["columnas"].get("perfil"),
    )
    return users, incomplete


# ---------------------------------------------------------------------------
# CLI / ejecución
# ---------------------------------------------------------------------------

def choose_profiles():
    print("\nPerfiles para procesar:")
    print("  1. AUXILIAR")
    print("  2. ENFERMERO")
    print("  3. MEDICO")
    print("  4. TODOS")
    option = input("Selecciona opción (1-4) [4]: ").strip() or "4"
    return {"1": ["AUXILIAR"], "2": ["ENFERMERO"], "3": ["MEDICO"], "4": None}.get(option, None)


def choose_batch_size(total_pending):
    value = input(f"Cantidad por tanda (0 = todos, máximo {total_pending}) [0]: ").strip() or "0"
    try:
        size = int(value)
    except ValueError:
        size = 0
    return total_pending if size <= 0 else min(size, total_pending)


def build_arg_parser():
    parser = argparse.ArgumentParser(
        description="Consulta masiva de medidas correctivas - Policía Nacional (RNMC)."
    )
    parser.add_argument("--config", help="Archivo config.json alternativo")
    parser.add_argument("--excel", help="Ruta del Excel de caracterización")
    parser.add_argument("--salida", help="Carpeta base de salida")
    parser.add_argument("--hoja", help="Nombre o índice de la hoja")
    parser.add_argument("--fila-encabezado", type=int, help="Fila (1-based) de encabezados")
    parser.add_argument("--perfil", choices=["AUXILIAR", "ENFERMERO", "MEDICO", "TODOS"])
    parser.add_argument("--tanda", type=int, help="Cantidad de registros en esta corrida")
    parser.add_argument("--si", dest="no_prompts", action="store_true", help="No preguntar nada")
    parser.add_argument("--reiniciar-estado", action="store_true", help="Ignorar el estado previo")
    parser.add_argument("--dry-run", action="store_true", help="Solo validar el Excel")
    parser.add_argument("--visible", action="store_true", help="Abrir Chrome visible")
    return parser


def _future_dates(users):
    hoy = datetime.now().date()
    futuras = []
    for user in users:
        try:
            fecha = datetime.strptime(user["fecha_expedicion"], "%d/%m/%Y").date()
        except Exception:
            continue
        if fecha > hoy:
            futuras.append(user)
    return futuras


def dry_run(settings: Settings):
    apply_settings(settings)
    print("\n" + "=" * 70)
    print("  VALIDACIÓN DE DATOS (sin navegador)")
    print("=" * 70)
    print(f"Excel:       {EXCEL_PATH}")
    print(f"Hoja:        {settings.sheet_name or '(primera)'}")
    print(f"Encabezados: fila {settings.header_row}")
    print("=" * 70)

    try:
        users, incomplete, stats = load_users(
            EXCEL_PATH,
            sheet_name=settings.sheet_name,
            header_row=settings.header_row,
            column_map=settings.column_map,
            selected_profiles=settings.profiles,
        )
    except Exception as exc:
        print(f"\n[ERROR] {exc}")
        return 2

    print(f"Hoja leída:      {stats['hoja']}")
    print(f"Filas con datos: {stats['total_filas']}")
    print(f"Válidos:         {stats['validos']}")
    print(f"Incompletos:     {stats['incompletos']}")
    print(f"Duplicados:      {stats['duplicados']}")
    print(f"Filtrados:       {stats['filtrados_por_perfil']}")
    print("\nColumnas usadas:")
    for clave, columna in stats["columnas"].items():
        print(f"  - {clave:14s} -> {columna}")

    if users:
        print(f"\nPrimero: {users[0]['apellido']} {users[0]['nombre']} ({users[0]['documento']}) "
              f"exp: {users[0]['fecha_expedicion']}")
        print(f"Último:  {users[-1]['apellido']} {users[-1]['nombre']} ({users[-1]['documento']}) "
              f"exp: {users[-1]['fecha_expedicion']}")

    futuras = _future_dates(users)
    if futuras:
        print(f"\n[WARN] {len(futuras)} fechas de expedición son futuras (posible error del Excel):")
        for user in futuras[:10]:
            print(f"  fila {user['fila_excel']}: {user['documento']} -> {user['fecha_expedicion']}")

    if incomplete:
        print("\nPrimeros incompletos:")
        for fila in incomplete[:10]:
            print(f"  fila {fila['fila']}: {fila['detalle']}")

    print("\n[OK] Validación terminada.")
    return 0


def run(settings: Settings, on_progress=None, on_status=None, should_stop=None):
    apply_settings(settings)
    set_stop_checker(should_stop)

    interactive = bool(settings.interactive)
    detenido = False

    print("\n" + "=" * 70)
    print("  MEDIDAS CORRECTIVAS - POLICÍA NACIONAL (RNMC)")
    print("=" * 70)
    print(f"Sistema:      {platform.system()} {platform.machine()}")
    print(f"Excel fuente: {EXCEL_PATH}")
    print(f"Salida base:  {BASE_OUTPUT_DIR}")
    print("=" * 70)

    configure_folders()

    if not Path(EXCEL_PATH).exists():
        print(f"\n[ERROR] No existe el Excel: {EXCEL_PATH}")
        return {"codigo": 2, "success": 0, "failed": 0, "total": 0, "detenido": False}

    if settings.profiles is not None:
        selected_profiles = None if not settings.profiles else list(settings.profiles)
    elif not interactive:
        selected_profiles = None
    else:
        selected_profiles = choose_profiles()

    try:
        users, incomplete_rows = load_users_from_excel(
            EXCEL_PATH,
            selected_profiles=selected_profiles,
            sheet_name=settings.sheet_name,
            header_row=settings.header_row,
            column_map=settings.column_map,
        )
    except Exception as exc:
        print(f"\n[ERROR] Error leyendo el Excel: {exc}")
        return {"codigo": 2, "success": 0, "failed": 0, "total": 0, "detenido": False}

    if not users:
        print("\n[ERROR] No se encontraron usuarios válidos para ese filtro.")
        return {"codigo": 2, "success": 0, "failed": 0, "total": 0, "detenido": False}

    state = load_state()
    processed_set = set(state.get("procesados", []))

    print(f"\nUsuarios válidos: {len(users)}")
    print(f"Filas con datos incompletos: {len(incomplete_rows)}")
    print(f"Procesados previos en estado: {len(processed_set)}")

    futuras = _future_dates(users)
    if futuras:
        print(f"[WARN] {len(futuras)} fechas de expedición son futuras; posible error del Excel.")

    if not settings.reset_state:
        if interactive:
            if settings.resume:
                answer = input("¿Reanudar omitiendo procesados previos? [S/n]: ").strip().lower()
                resume = answer != "n"
            else:
                answer = input("¿Reanudar omitiendo procesados previos? [s/N]: ").strip().lower()
                resume = answer == "s"
        else:
            resume = bool(settings.resume)

        if resume:
            users = [u for u in users if u["documento"] not in processed_set]

    if not users:
        print("\n[OK] No hay pendientes para los perfiles seleccionados.")
        return {"codigo": 0, "success": 0, "failed": 0, "total": 0, "detenido": False}

    if settings.batch_size is not None:
        batch_size = settings.batch_size if settings.batch_size > 0 else len(users)
    elif not interactive:
        batch_size = len(users)
    else:
        batch_size = choose_batch_size(len(users))
    users = users[:batch_size]

    total = len(users)
    headless = bool(settings.headless)
    print(f"\n[INFO] Esta tanda procesará {total} usuarios.")
    print(f"[INFO] Chrome: {'headless' if headless else 'visible'}")

    driver = None
    success = 0
    failed = 0
    fallos = []
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    if on_progress:
        try:
            on_progress(0, total, 0, 0)
        except Exception:
            pass

    try:
        driver = create_driver(headless=headless)
        wait = WebDriverWait(driver, 30)

        for idx, user in enumerate(users, start=1):
            if stop_requested():
                detenido = True
                print("\n[INFO] Detenido por el usuario.")
                break

            print(
                f"\n[{idx}/{total}] {user['apellido']} {user['nombre']} "
                f"({user['documento']}) - {user['perfil_grupo']} - exp: {user['fecha_expedicion']}"
            )
            if on_status:
                try:
                    on_status(f"[{idx}/{total}] {user['apellido']} {user['nombre']}")
                except Exception:
                    pass

            ok, intentos, motivo, detalle, captura = process_user(
                driver, wait, user, settings, on_status=on_status
            )

            if motivo == "detenido":
                detenido = True
                print("\n[INFO] Detenido por el usuario.")
                break

            if ok:
                success += 1
                processed_set.add(user["documento"])
                state.setdefault("fallidos", [])
                if user["documento"] in state["fallidos"]:
                    state["fallidos"].remove(user["documento"])
            else:
                failed += 1
                fallos.append({
                    "documento": user["documento"],
                    "nombre": f"{user['nombre']} {user['apellido']}",
                    "perfil": user["perfil_grupo"],
                    "fila_excel": user.get("fila_excel", ""),
                    "fecha_expedicion": user.get("fecha_expedicion", ""),
                    "motivo": motivo or "sin_resultado",
                    "detalle": detalle,
                    "intentos": intentos,
                    "captura": captura,
                })
                state.setdefault("fallidos", [])
                if user["documento"] not in state["fallidos"]:
                    state["fallidos"].append(user["documento"])

            state["procesados"] = sorted(processed_set)
            save_state(state)

            if on_progress:
                try:
                    on_progress(idx, total, success, failed)
                except Exception:
                    pass

            time.sleep(0.5)

    except KeyboardInterrupt:
        detenido = True
        print("\n\n[INFO] Proceso interrumpido por el usuario.")
    except Exception as exc:
        print(f"\n[ERROR] Error general: {exc}")
    finally:
        if driver:
            try:
                driver.quit()
            except Exception:
                pass

    csv_path, txt_path, xlsx_path = write_failure_report(fallos, incomplete_rows, timestamp)

    print("\n" + "=" * 70)
    print("RESUMEN")
    print(f"[OK] Exitosos: {success}")
    print(f"[ERROR] Fallidos: {failed}")
    print(f"Carpeta base: {BASE_OUTPUT_DIR}")
    print(f"Estado: {STATE_FILE}")
    if txt_path:
        print(f"Informe de fallos: {txt_path}")
        print(f"Informe (CSV):     {csv_path}")
        if xlsx_path:
            print(f"Informe (Excel):   {xlsx_path}")
    else:
        print("[OK] No hubo fallos que reportar.")
    print("=" * 70)

    return {
        "codigo": 0,
        "success": success,
        "failed": failed,
        "total": total,
        "detenido": detenido,
        "csv": str(csv_path) if csv_path else "",
        "txt": str(txt_path) if txt_path else "",
        "xlsx": str(xlsx_path) if xlsx_path else "",
        "carpeta": str(BASE_OUTPUT_DIR),
    }


def main(argv=None):
    args = build_arg_parser().parse_args(argv)
    settings = load_settings(Path(args.config) if args.config else None)

    if args.excel:
        settings.excel_path = Path(args.excel)
    if args.salida:
        settings.output_dir = Path(args.salida)
    if args.hoja is not None:
        settings.sheet_name = args.hoja
    if args.fila_encabezado is not None:
        settings.header_row = max(1, args.fila_encabezado)
    if args.perfil:
        settings.profiles = None if args.perfil == "TODOS" else [args.perfil]
    if args.tanda is not None:
        settings.batch_size = args.tanda
    if args.reiniciar_estado:
        settings.reset_state = True
    if args.no_prompts:
        settings.interactive = False
    if args.visible:
        settings.headless = False

    if args.dry_run:
        return dry_run(settings)

    return run(settings)["codigo"]


if __name__ == "__main__":
    raise SystemExit(main())
