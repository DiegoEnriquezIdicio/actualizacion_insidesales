"""
Actualiza el Sales Rep (Inside Sales) del Sales Team en transacciones de NetSuite.

Soporta tres tipos de transaccion: salesorder, quote y opportunity.
Lee un Excel de la carpeta data/ con las columnas id_transaccion e id_insidesales,
hace un PATCH por cada fila al REST Record API de NetSuite y genera un reporte de
auditoria en Excel dentro de la carpeta logs/.

Ejecucion manual:
    python actualizar_inside_sales.py
    python actualizar_inside_sales.py --tipo quote --archivo data/quotes.xlsx
    python actualizar_inside_sales.py --tipo salesorder --dry-run
"""

import argparse
import json
import os
import sys
import time
import traceback
from datetime import datetime

import pandas as pd
import requests
from dotenv import load_dotenv
from requests_oauthlib import OAuth1

# ---------------------------------------------------------------------------
# CONFIGURACION - esto es lo unico que se cambia entre corridas
# ---------------------------------------------------------------------------

# Tipo de transaccion a procesar: "salesorder" | "quote" | "opportunity"
TIPO_TRANSACCION = "salesorder"

# Excel de entrada. None = toma el archivo mas reciente de la carpeta data/
ARCHIVO_EXCEL = None

# Archivo de variables de entorno con las credenciales de NetSuite
ARCHIVO_ENV = ".env"

# True = no envia nada a NetSuite, solo valida el Excel y arma los bodies
DRY_RUN = False

# Pausa en segundos entre peticiones (para no saturar la cuota de NetSuite)
PAUSA_SEGUNDOS = 0.3

# Reintentos ante errores temporales (429 / 5xx / timeout)
REINTENTOS = 2

# Timeout de cada peticion HTTP en segundos
TIMEOUT = 60

# ---------------------------------------------------------------------------
# Mapeo de tipos de transaccion -> recurso REST de NetSuite
# ---------------------------------------------------------------------------
# El recurso es el segmento que va en la URL:
#   {ENDPOINT_NETSUITE}record/v1/{recurso}/{id_transaccion}?replace=salesTeam
# En el REST Record API de NetSuite el Quote se llama "estimate".
#
# El parametro ?replace=salesTeam le indica a NetSuite que REEMPLACE el sublist
# completo del Sales Team en lugar de hacer merge con los miembros existentes.
# Sin el, el nuevo Inside Sales se agregaria a los que ya estaban.
REPLACE_SUBLIST = "salesTeam"
TRANSACCIONES = {
    "salesorder": {"recurso": "salesOrder", "contribution": False},
    "quote": {"recurso": "estimate", "contribution": False},
    "opportunity": {"recurso": "opportunity", "contribution": True},
}

# Nombres aceptados para las columnas del Excel (se normalizan a minusculas sin espacios)
COLUMNA_TRANSACCION = ["id_transaccion", "idtransaccion", "id_transaction", "internalid", "id"]
COLUMNA_INSIDESALES = ["id_insidesales", "id_inside_sales", "idinsidesales", "inside_sales", "employee_id"]

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
LOGS_DIR = os.path.join(BASE_DIR, "logs")


# ---------------------------------------------------------------------------
# Autenticacion NetSuite (OAuth 1.0 - Token Based Authentication)
# ---------------------------------------------------------------------------
def cargar_entorno(ruta_env):
    """Carga el .env con las credenciales de NetSuite y valida que esten completas."""
    # Se buscan varias ubicaciones para que el script funcione tanto dentro del
    # proyecto como copiado a otra maquina con su propio .env al lado.
    candidatos = []
    if ruta_env:
        candidatos.append(ruta_env if os.path.isabs(ruta_env) else os.path.join(BASE_DIR, ruta_env))
    candidatos.append(os.path.join(BASE_DIR, ".env"))
    candidatos.append(os.path.join(BASE_DIR, ".env_evonet_prod"))
    # Ultimo recurso: entorno del proyecto principal, si el script vive dentro de el
    candidatos.append(os.path.join(BASE_DIR, "..", "REP_005-BTCT_005", ".env_evonet_prod"))

    dotenv_path = None
    for candidato in candidatos:
        candidato = os.path.normpath(candidato)
        if os.path.exists(candidato):
            dotenv_path = candidato
            break

    if dotenv_path is None:
        rutas = "\n  - ".join(os.path.normpath(c) for c in candidatos)
        raise FileNotFoundError(
            "No se encontro el archivo de credenciales. Se busco en:\n  - " + rutas +
            "\n\nCopia .env.example como .env en esta carpeta y llena los valores, "
            "o indica la ruta con --env."
        )

    load_dotenv(dotenv_path)
    print(f"Entorno cargado desde: {dotenv_path}")

    requeridas = [
        "ENDPOINT_NETSUITE",
        "CONSUMER_KEY_NETSUITE_API_EVONET",
        "CONSUMER_SECRET_NETSUITE_API_EVONET",
        "TOKEN_KEY_NETSUITE_API_EVONET",
        "TOKEN_SECRET_NETSUITE_API_EVONET",
        "REALM_NETSUITE_API_EVONET",
    ]
    faltantes = [v for v in requeridas if not os.environ.get(v)]
    if faltantes:
        raise RuntimeError(f"Faltan variables de entorno: {', '.join(faltantes)}")


def crear_auth():
    """Crea el objeto OAuth1 (HMAC-SHA256 + realm) que usa NetSuite REST."""
    return OAuth1(
        os.environ.get("CONSUMER_KEY_NETSUITE_API_EVONET"),
        client_secret=os.environ.get("CONSUMER_SECRET_NETSUITE_API_EVONET"),
        resource_owner_key=os.environ.get("TOKEN_KEY_NETSUITE_API_EVONET"),
        resource_owner_secret=os.environ.get("TOKEN_SECRET_NETSUITE_API_EVONET"),
        signature_method="HMAC-SHA256",
        realm=os.environ.get("REALM_NETSUITE_API_EVONET"),
    )


# ---------------------------------------------------------------------------
# Lectura y validacion del Excel de entrada
# ---------------------------------------------------------------------------
def resolver_archivo(archivo):
    """Devuelve la ruta del Excel a procesar; si no se indica, el mas reciente de data/."""
    if archivo:
        ruta = archivo if os.path.isabs(archivo) else os.path.join(BASE_DIR, archivo)
        ruta = os.path.normpath(ruta)
        if not os.path.exists(ruta):
            raise FileNotFoundError(f"No se encontro el Excel: {ruta}")
        return ruta

    candidatos = [
        os.path.join(DATA_DIR, f)
        for f in os.listdir(DATA_DIR)
        if f.lower().endswith((".xlsx", ".xls")) and not f.startswith("~$")
    ]
    if not candidatos:
        raise FileNotFoundError(f"No hay archivos Excel en {DATA_DIR}")

    ruta = max(candidatos, key=os.path.getmtime)
    print(f"Archivo autodetectado (mas reciente): {os.path.basename(ruta)}")
    return ruta


def normalizar(nombre):
    return str(nombre).strip().lower().replace(" ", "_").replace("-", "_")


def buscar_columna(columnas_normalizadas, alias):
    for a in alias:
        if a in columnas_normalizadas:
            return columnas_normalizadas[a]
    return None


def limpiar_id(valor):
    """Convierte el valor de la celda a un id entero en texto ('448551.0' -> '448551')."""
    if valor is None:
        return ""
    texto = str(valor).strip()
    if texto.lower() in ("", "nan", "none", "nat"):
        return ""
    if texto.endswith(".0") and texto[:-2].isdigit():
        texto = texto[:-2]
    return texto


def leer_excel(ruta):
    """Lee el Excel y devuelve una lista de dicts con fila, id_transaccion e id_insidesales."""
    df = pd.read_excel(ruta, dtype=str)
    columnas_normalizadas = {normalizar(c): c for c in df.columns}

    col_tx = buscar_columna(columnas_normalizadas, COLUMNA_TRANSACCION)
    col_is = buscar_columna(columnas_normalizadas, COLUMNA_INSIDESALES)

    if not col_tx or not col_is:
        raise RuntimeError(
            "El Excel debe tener las columnas 'id_transaccion' e 'id_insidesales'. "
            f"Columnas encontradas: {list(df.columns)}"
        )

    filas = []
    for indice, registro in df.iterrows():
        filas.append(
            {
                # +2 = encabezado + indice base 0, para que coincida con la fila real del Excel
                "fila_excel": indice + 2,
                "id_transaccion": limpiar_id(registro[col_tx]),
                "id_insidesales": limpiar_id(registro[col_is]),
            }
        )
    return filas


# ---------------------------------------------------------------------------
# Construccion del body y llamada a NetSuite
# ---------------------------------------------------------------------------
def construir_body(id_insidesales, incluir_contribution):
    """Arma el body del PATCH con el nuevo Sales Rep como primario del Sales Team."""
    item = {
        "employee": {"id": int(id_insidesales)},
        "isPrimary": True,
        "salesrole": {"id": "-2"},
    }
    if incluir_contribution:
        item["contribution"] = 100.0
    return {"salesTeam": {"items": [item]}}


def extraer_error(respuesta):
    """Obtiene un mensaje de error legible de la respuesta de NetSuite."""
    try:
        data = respuesta.json()
    except ValueError:
        return (respuesta.text or "").strip()[:800]

    detalles = data.get("o:errorDetails") or []
    if detalles:
        partes = []
        for d in detalles:
            codigo = d.get("o:errorCode", "")
            detalle = d.get("detail", "")
            partes.append(f"[{codigo}] {detalle}" if codigo else detalle)
        return " | ".join(partes)[:800]

    return (data.get("title") or json.dumps(data))[:800]


def actualizar_transaccion(session, auth, url, body):
    """Ejecuta el PATCH con reintentos. Devuelve (codigo_http, mensaje_error)."""
    intento = 0
    while True:
        try:
            respuesta = session.patch(
                url,
                headers={"Content-Type": "application/json"},
                json=body,
                auth=auth,
                timeout=TIMEOUT,
            )
        except requests.RequestException as e:
            if intento < REINTENTOS:
                intento += 1
                time.sleep(2 * intento)
                continue
            return None, f"Excepcion de red: {type(e).__name__}: {e}"

        if respuesta.status_code == 204:
            return 204, ""

        # 429 (limite de concurrencia) y 5xx son temporales: reintentar
        if respuesta.status_code in (429, 500, 502, 503, 504) and intento < REINTENTOS:
            intento += 1
            time.sleep(2 * intento)
            continue

        return respuesta.status_code, extraer_error(respuesta)


# ---------------------------------------------------------------------------
# Reporte de auditoria
# ---------------------------------------------------------------------------
def escribir_reporte(resultados, tipo, archivo_entrada):
    os.makedirs(LOGS_DIR, exist_ok=True)
    marca = datetime.now().strftime("%Y%m%d_%H%M%S")
    ruta = os.path.join(LOGS_DIR, f"reporte_{tipo}_{marca}.xlsx")

    df = pd.DataFrame(
        resultados,
        columns=[
            "fila_excel",
            "id_transaccion",
            "id_insidesales",
            "estatus",
            "codigo_http",
            "error",
            "url",
            "fecha_hora",
        ],
    )

    exitos = int((df["estatus"] == "EXITO").sum())
    errores = int((df["estatus"] == "ERROR").sum())
    omitidos = int((df["estatus"] == "OMITIDO").sum())

    resumen = pd.DataFrame(
        [
            {"concepto": "tipo_transaccion", "valor": tipo},
            {"concepto": "archivo_entrada", "valor": os.path.basename(archivo_entrada)},
            {"concepto": "fecha_ejecucion", "valor": datetime.now().strftime("%Y-%m-%d %H:%M:%S")},
            {"concepto": "modo", "valor": "DRY_RUN" if DRY_RUN else "REAL"},
            {"concepto": "total_filas", "valor": len(df)},
            {"concepto": "exitos", "valor": exitos},
            {"concepto": "errores", "valor": errores},
            {"concepto": "omitidos", "valor": omitidos},
        ]
    )

    with pd.ExcelWriter(ruta, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="detalle", index=False)
        resumen.to_excel(writer, sheet_name="resumen", index=False)

    return ruta, exitos, errores, omitidos


# ---------------------------------------------------------------------------
# Proceso principal
# ---------------------------------------------------------------------------
def main():
    global TIPO_TRANSACCION, ARCHIVO_EXCEL, DRY_RUN

    parser = argparse.ArgumentParser(description="Actualiza el Inside Sales del Sales Team en NetSuite.")
    parser.add_argument("--tipo", choices=sorted(TRANSACCIONES.keys()), help="Tipo de transaccion a procesar")
    parser.add_argument("--archivo", help="Ruta del Excel de entrada (por defecto, el mas reciente de data/)")
    parser.add_argument("--env", help="Ruta del archivo .env con las credenciales")
    parser.add_argument("--dry-run", action="store_true", help="No envia nada a NetSuite, solo valida")
    args = parser.parse_args()

    tipo = (args.tipo or TIPO_TRANSACCION).lower()
    if tipo not in TRANSACCIONES:
        print(f"Tipo de transaccion invalido: {tipo}. Validos: {', '.join(sorted(TRANSACCIONES))}")
        return 1

    if args.dry_run:
        DRY_RUN = True

    configuracion = TRANSACCIONES[tipo]
    ruta_env = args.env or ARCHIVO_ENV

    # Se crean solas si no existen: no hay que prepararlas a mano
    os.makedirs(DATA_DIR, exist_ok=True)
    os.makedirs(LOGS_DIR, exist_ok=True)

    try:
        cargar_entorno(ruta_env)
        ruta_excel = resolver_archivo(args.archivo or ARCHIVO_EXCEL)
        filas = leer_excel(ruta_excel)
    except (FileNotFoundError, RuntimeError) as e:
        # Errores esperados (falta el .env, falta el Excel, faltan columnas):
        # el mensaje ya es claro, no hace falta el traceback.
        print(f"\nERROR de configuracion: {e}\n")
        return 1
    except Exception as e:
        print(f"\nERROR inesperado: {e}\n")
        traceback.print_exc()
        return 1

    endpoint = os.environ.get("ENDPOINT_NETSUITE")
    if not endpoint.endswith("/"):
        endpoint += "/"
    base_url = f"{endpoint}record/v1/{configuracion['recurso']}/"
    sufijo_url = f"?replace={REPLACE_SUBLIST}" if REPLACE_SUBLIST else ""

    print("-" * 70)
    print(f"Tipo de transaccion : {tipo} (recurso REST: {configuracion['recurso']})")
    print(f"Archivo de entrada  : {os.path.basename(ruta_excel)}")
    print(f"Filas a procesar    : {len(filas)}")
    print(f"URL base            : {base_url}{{id_transaccion}}{sufijo_url}")
    print(f"Modo                : {'DRY RUN (no se envia nada)' if DRY_RUN else 'REAL'}")
    print("-" * 70)

    auth = crear_auth()
    session = requests.Session()
    resultados = []

    for fila in filas:
        id_tx = fila["id_transaccion"]
        id_is = fila["id_insidesales"]
        ahora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        # Validacion previa: ambos ids deben existir y ser numericos
        if not id_tx or not id_is or not id_tx.isdigit() or not id_is.isdigit():
            motivo = "id_transaccion e id_insidesales deben ser numericos y no estar vacios"
            print(f"[fila {fila['fila_excel']}] OMITIDO - {motivo}")
            resultados.append(
                {
                    "fila_excel": fila["fila_excel"],
                    "id_transaccion": id_tx,
                    "id_insidesales": id_is,
                    "estatus": "OMITIDO",
                    "codigo_http": "",
                    "error": motivo,
                    "url": "",
                    "fecha_hora": ahora,
                }
            )
            continue

        url = base_url + id_tx + sufijo_url
        body = construir_body(id_is, configuracion["contribution"])

        if DRY_RUN:
            print(f"[fila {fila['fila_excel']}] DRY RUN PATCH {url} -> {json.dumps(body)}")
            resultados.append(
                {
                    "fila_excel": fila["fila_excel"],
                    "id_transaccion": id_tx,
                    "id_insidesales": id_is,
                    "estatus": "OMITIDO",
                    "codigo_http": "",
                    "error": "DRY RUN: no se envio la peticion",
                    "url": url,
                    "fecha_hora": ahora,
                }
            )
            continue

        codigo, error = actualizar_transaccion(session, auth, url, body)
        exito = codigo == 204

        if exito:
            print(f"[fila {fila['fila_excel']}] EXITO  {tipo} {id_tx} -> inside sales {id_is}")
        else:
            print(f"[fila {fila['fila_excel']}] ERROR  {tipo} {id_tx} ({codigo}): {error}")

        resultados.append(
            {
                "fila_excel": fila["fila_excel"],
                "id_transaccion": id_tx,
                "id_insidesales": id_is,
                "estatus": "EXITO" if exito else "ERROR",
                "codigo_http": codigo if codigo is not None else "",
                "error": error,
                "url": url,
                "fecha_hora": ahora,
            }
        )

        if PAUSA_SEGUNDOS:
            time.sleep(PAUSA_SEGUNDOS)

    ruta_reporte, exitos, errores, omitidos = escribir_reporte(resultados, tipo, ruta_excel)

    print("-" * 70)
    print(f"Total procesado : {len(resultados)}")
    print(f"Exitos          : {exitos}")
    print(f"Errores         : {errores}")
    print(f"Omitidos        : {omitidos}")
    print(f"Reporte         : {ruta_reporte}")
    print("-" * 70)

    return 0 if errores == 0 else 2


if __name__ == "__main__":
    sys.exit(main())
