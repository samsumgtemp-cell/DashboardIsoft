"""
Exportador de casos de Jira -> Excel "Detalle de Casos" (TODOS los campos)
============================================================================

Trae TODOS los campos disponibles de cada ticket del proyecto MDI
(estándar + personalizados), filtrado solo a casos CREADOS EN 2026.

MODO INCREMENTAL (para uso diario / atado al dashboard)
---------------------------------------------------------
La primera corrida trae todos los casos de 2026 (puede tardar, ya que
consulta el changelog de cada ticket uno por uno). Las siguientes
corridas solo traen los tickets actualizados desde la última vez
(guardada en 'ultima_actualizacion.json'), y actualizan solo esas filas
en el Excel existente.

CONFIGURACIÓN
--------------
Define estas variables de entorno antes de correr:
    JIRA_DOMAIN    -> isoft-ste.atlassian.net
    JIRA_EMAIL     -> tu correo de Atlassian
    JIRA_API_TOKEN -> el token de API generado (scope read:jira-work)

Correr normal:
    python exportar_casos_jira.py
Forzar recarga completa (ignora la última fecha guardada):
    python exportar_casos_jira.py --full
"""

import os
import sys
import json
import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
from pathlib import Path
from datetime import datetime, timezone, timedelta
import openpyxl
from openpyxl.styles import Font, Alignment
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE


def limpiar_para_excel(valor):
    """Quita caracteres de control que Excel/XML no acepta en celdas."""
    if isinstance(valor, str):
        return ILLEGAL_CHARACTERS_RE.sub("", valor)
    return valor


# ------------------------------------------------------------------
# 0. SESIÓN HTTP CON REINTENTOS AUTOMÁTICOS
# ------------------------------------------------------------------

def crear_sesion():
    sesion = requests.Session()
    reintentos = Retry(
        total=6,
        backoff_factor=2,
        status_forcelist=[429, 500, 502, 503, 504],
        allowed_methods=["GET"],
    )
    adaptador = HTTPAdapter(max_retries=reintentos)
    sesion.mount("https://", adaptador)
    sesion.mount("http://", adaptador)
    return sesion


SESION = crear_sesion()

# ------------------------------------------------------------------
# 1. CONFIGURACIÓN Y CREDENCIALES
# ------------------------------------------------------------------

JIRA_DOMAIN = os.environ.get("JIRA_DOMAIN", "isoft-ste.atlassian.net")
JIRA_EMAIL = os.environ.get("JIRA_EMAIL")
JIRA_API_TOKEN = os.environ.get("JIRA_API_TOKEN")
PROJECT_KEY = "MDI"

# Filtro de año. Cambia esto si en el futuro quieres otro año, o ponlo en
# None para traer todos los años (sin filtro de fecha).
ANIO_FILTRO = 2026

CARPETA_SCRIPT = Path(__file__).resolve().parent
RUTA_EXCEL = CARPETA_SCRIPT / f"Detalle_Casos_{PROJECT_KEY}_{ANIO_FILTRO or 'TODOS'}.xlsx"
RUTA_ESTADO = CARPETA_SCRIPT / "ultima_actualizacion.json"

FORZAR_COMPLETO = "--full" in sys.argv

if not JIRA_EMAIL or not JIRA_API_TOKEN:
    print("ERROR: Debes definir las variables de entorno JIRA_EMAIL y JIRA_API_TOKEN")
    sys.exit(1)


def obtener_cloud_id(dominio):
    resp = requests.get(f"https://{dominio}/_edge/tenant_info")
    resp.raise_for_status()
    return resp.json()["cloudId"]


print("Obteniendo Cloud ID del sitio...")
CLOUD_ID = obtener_cloud_id(JIRA_DOMAIN)
print(f"  Cloud ID: {CLOUD_ID}")

BASE_URL = f"https://api.atlassian.com/ex/jira/{CLOUD_ID}/rest/api/3"
AUTH = (JIRA_EMAIL, JIRA_API_TOKEN)
HEADERS = {"Accept": "application/json"}


# ------------------------------------------------------------------
# 2. ESTADO DE ÚLTIMA CORRIDA (modo incremental)
# ------------------------------------------------------------------

def leer_ultima_actualizacion():
    """
    Lee la última sincronización y retrocede 10 minutos como margen de seguridad.
    Esto evita perder cambios de Jira que ocurran cerca del corte entre corridas.
    Como las filas se consolidan por Ticket, volver a recibir un ticket no lo duplica.
    """
    if FORZAR_COMPLETO or not RUTA_ESTADO.exists():
        return None
    try:
        with open(RUTA_ESTADO, "r", encoding="utf-8") as f:
            valor = json.load(f).get("ultima_actualizacion")

        if not valor:
            return None

        fecha = datetime.strptime(valor, "%Y-%m-%d %H:%M").replace(tzinfo=timezone.utc)
        fecha_segura = fecha - timedelta(minutes=10)
        return fecha_segura.strftime("%Y-%m-%d %H:%M")
    except Exception as e:
        print(f"ADVERTENCIA: no se pudo leer la última actualización ({e}).")
        return None


def guardar_ultima_actualizacion(momento_iso):
    with open(RUTA_ESTADO, "w", encoding="utf-8") as f:
        json.dump({"ultima_actualizacion": momento_iso}, f)


# ------------------------------------------------------------------
# 3. LISTA COMPLETA DE CAMPOS (id -> nombre) DE LA INSTANCIA DE JIRA
# ------------------------------------------------------------------

def obtener_mapa_de_campos():
    """
    Trae TODOS los campos existentes en la instancia (estándar + personalizados)
    y devuelve un dict {id_campo: nombre_campo}, ej. {'customfield_10045':
    'Distribuidor', 'summary': 'Summary', ...}
    """
    resp = SESION.get(f"{BASE_URL}/field", auth=AUTH, headers=HEADERS, timeout=30)
    resp.raise_for_status()
    campos = resp.json()
    mapa = {c["id"]: c["name"] for c in campos}
    print(f"  Total de campos detectados en la instancia: {len(mapa)}")
    return mapa


# ------------------------------------------------------------------
# 4. BÚSQUEDA DE ISSUES (paginada, con soporte incremental y filtro de año)
# ------------------------------------------------------------------

def buscar_issues(anio=None, actualizado_desde=None):
    jql = f'project = {PROJECT_KEY}'
    if anio:
        jql += f' AND created >= "{anio}-01-01" AND created <= "{anio}-12-31"'
    if actualizado_desde:
        jql += f' AND updated >= "{actualizado_desde}"'
    jql += " ORDER BY created ASC"

    print(f"\nJQL usado: {jql}\n")

    todos_issues = []
    next_page_token = None
    max_results = 100

    while True:
        params = {
            "jql": jql,
            "maxResults": max_results,
            "fields": "*all",  # traer TODOS los campos de cada issue
        }
        if next_page_token:
            params["nextPageToken"] = next_page_token

        resp = SESION.get(f"{BASE_URL}/search/jql", auth=AUTH, headers=HEADERS,
                           params=params, timeout=30)
        resp.raise_for_status()
        data = resp.json()

        todos_issues.extend(data["issues"])
        print(f"  Descargados {len(todos_issues)} issues hasta ahora...")

        next_page_token = data.get("nextPageToken")
        if not next_page_token:
            break

    return todos_issues


# ------------------------------------------------------------------
# 5. CHANGELOG: horas por estado
# ------------------------------------------------------------------

def obtener_changelog(issue_key):
    cambios = []
    start_at = 0
    max_results = 100

    while True:
        resp = SESION.get(
            f"{BASE_URL}/issue/{issue_key}/changelog",
            auth=AUTH, headers=HEADERS,
            params={"startAt": start_at, "maxResults": max_results},
            timeout=30,
        )
        resp.raise_for_status()
        data = resp.json()

        for historial in data["values"]:
            fecha = historial["created"]
            for item in historial["items"]:
                if item["field"] == "status":
                    cambios.append((fecha, item.get("fromString"), item.get("toString")))

        start_at += max_results
        if start_at >= data["total"]:
            break

    return cambios


def calcular_horas_por_estado(cambios, fecha_creacion):
    horas_por_estado = {}
    if not cambios:
        return horas_por_estado

    puntos = [(fecha_creacion, None, cambios[0][1] or "Abierto")]
    for fecha, desde, hacia in cambios:
        puntos.append((fecha, desde, hacia))

    for i in range(len(puntos) - 1):
        fecha_ini = datetime.strptime(puntos[i][0][:19], "%Y-%m-%dT%H:%M:%S")
        fecha_fin = datetime.strptime(puntos[i + 1][0][:19], "%Y-%m-%dT%H:%M:%S")
        estado = puntos[i][2]
        horas = (fecha_fin - fecha_ini).total_seconds() / 3600
        horas_por_estado[estado] = horas_por_estado.get(estado, 0) + horas

    return horas_por_estado


# ------------------------------------------------------------------
# 6. EXTRAER VALOR DE CUALQUIER TIPO DE CAMPO
# ------------------------------------------------------------------

def extraer_valor_campo(valor):
    """
    Normaliza cualquier tipo de valor de campo de Jira a un string simple:
    - None -> ""
    - texto/número -> tal cual
    - descripción en formato ADF (documento con 'content') -> texto plano
    - objetos de selección {value:...} / {name:...} -> ese texto
    - usuarios {displayName:...} -> el nombre
    - listas -> unidas con "; "
    """
    if valor is None or valor == "":
        return ""
    if isinstance(valor, (str, int, float, bool)):
        return str(valor)
    if isinstance(valor, list):
        return "; ".join(extraer_valor_campo(v) for v in valor if v not in (None, ""))
    if isinstance(valor, dict):
        # Documento ADF (descripciones, comentarios largos, etc.)
        if valor.get("type") == "doc" and "content" in valor:
            try:
                textos = []
                for bloque in valor.get("content", []):
                    for nodo in bloque.get("content", []) or []:
                        if "text" in nodo:
                            textos.append(nodo["text"])
                return " ".join(textos)
            except Exception:
                return ""
        for clave in ("displayName", "value", "name", "key"):
            if valor.get(clave):
                return str(valor[clave])
        return ""
    return str(valor)


# ------------------------------------------------------------------
# 7. PROCESAR UN ISSUE -> DICT {nombre_campo: valor}
# ------------------------------------------------------------------

def procesar_issue(issue, mapa_campos):
    key = issue["key"]
    fields = issue["fields"]

    fila = {"Ticket": key}
    for campo_id, valor in fields.items():
        nombre_campo = mapa_campos.get(campo_id, campo_id)
        fila[nombre_campo] = extraer_valor_campo(valor)

    fecha_creacion = fields.get("created")
    cambios = obtener_changelog(key)
    horas_por_estado = calcular_horas_por_estado(cambios, fecha_creacion)
    fila["Horas por Estado (detalle)"] = "; ".join(
        f"{k}: {v:.1f}h" for k, v in horas_por_estado.items()
    )

    return fila


# ------------------------------------------------------------------
# 8. LEER / ESCRIBIR EXCEL (columnas dinámicas, modo incremental)
# ------------------------------------------------------------------

COLUMNAS_PRIORITARIAS = [
    "Ticket", "Issue Type", "Status", "Distribuidor", "Assignee", "Reporter",
    "Created", "Resolved", "Description",
]


def calcular_orden_columnas(todas_las_columnas):
    """Pone primero las columnas más relevantes, luego el resto alfabético,
    y 'Horas por Estado (detalle)' siempre al final."""
    resto = sorted(c for c in todas_las_columnas
                    if c not in COLUMNAS_PRIORITARIAS and c != "Horas por Estado (detalle)")
    orden = [c for c in COLUMNAS_PRIORITARIAS if c in todas_las_columnas] + resto
    if "Horas por Estado (detalle)" in todas_las_columnas:
        orden.append("Horas por Estado (detalle)")
    return orden


def cargar_filas_existentes():
    """Devuelve (columnas, {ticket: {col: valor}}) a partir del Excel existente."""
    if not RUTA_EXCEL.exists():
        return [], {}

    wb = openpyxl.load_workbook(RUTA_EXCEL)
    ws = wb["Detalle de Casos"]
    encabezados = [c.value for c in ws[1]]
    filas = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0]:
            filas[row[0]] = dict(zip(encabezados, row))
    return encabezados, filas


def guardar_excel(filas_dict):
    """filas_dict: {ticket: {columna: valor}}. Reescribe el Excel completo."""
    todas_columnas = set()
    for fila in filas_dict.values():
        todas_columnas.update(fila.keys())
    columnas = calcular_orden_columnas(todas_columnas)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Detalle de Casos"

    ws.append(columnas)
    for celda in ws[1]:
        celda.font = Font(bold=True)
        celda.alignment = Alignment(horizontal="center")

    for ticket in sorted(filas_dict.keys()):
        fila = filas_dict[ticket]
        ws.append([limpiar_para_excel(fila.get(col, "")) for col in columnas])

    for columna in ws.columns:
        max_len = max((len(str(c.value)) for c in columna if c.value), default=10)
        ws.column_dimensions[columna[0].column_letter].width = min(max_len + 2, 60)

    wb.save(RUTA_EXCEL)
    print(f"\nArchivo guardado en: {RUTA_EXCEL}")


# ------------------------------------------------------------------
# 9. FLUJO PRINCIPAL
# ------------------------------------------------------------------

def main():
    momento_inicio = datetime.now(timezone.utc)
    ultima_actualizacion = leer_ultima_actualizacion()

    if ultima_actualizacion:
        print(f"Modo INCREMENTAL: trayendo solo casos actualizados desde {ultima_actualizacion}")
    else:
        print(f"Modo COMPLETO: trayendo TODOS los casos de {ANIO_FILTRO or 'todos los años'}")
        print("(primera corrida o --full). Puede tardar, ya que consulta el changelog")
        print("de cada ticket uno por uno.")

    print("\nObteniendo mapa de campos (estándar + personalizados)...")
    mapa_campos = obtener_mapa_de_campos()

    print(f"\nBuscando issues del proyecto {PROJECT_KEY} (año={ANIO_FILTRO})...")
    issues = buscar_issues(anio=ANIO_FILTRO, actualizado_desde=ultima_actualizacion)
    print(f"\nTotal de issues a procesar: {len(issues)}")

    _, filas_existentes = cargar_filas_existentes()
    print(f"Filas ya existentes en el Excel: {len(filas_existentes)}")

    for i, issue in enumerate(issues, 1):
        key = issue["key"]
        print(f"  [{i}/{len(issues)}] Procesando {key} - trayendo changelog...")
        try:
            fila = procesar_issue(issue, mapa_campos)
            filas_existentes[key] = fila
        except Exception as e:
            print(f"    ADVERTENCIA: no se pudo procesar {key} ({e}). Se omite por ahora.")
            continue

        if i % 100 == 0:
            guardar_excel(filas_existentes)
            print(f"    (progreso guardado: {i}/{len(issues)})")

    guardar_excel(filas_existentes)
    guardar_ultima_actualizacion(momento_inicio.strftime("%Y-%m-%d %H:%M"))
    print(f"\nListo. Próxima corrida solo traerá cambios desde {momento_inicio.strftime('%Y-%m-%d %H:%M')} UTC.")


if __name__ == "__main__":
    main()
