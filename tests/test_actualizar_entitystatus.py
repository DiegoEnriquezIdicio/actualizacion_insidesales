"""
Pruebas de actualizar_entitystatus.py. No se conectan a NetSuite: las peticiones HTTP
las responde un adaptador falso y el script trabaja sobre una carpeta temporal.

Ejecutar desde la raíz del proyecto:
    python -m unittest discover -s tests -v
"""

import contextlib
import io
import json
import os
import sys
import tempfile
import time
import unittest
from unittest import mock

import openpyxl
import pandas as pd
import requests
from requests.adapters import BaseAdapter, HTTPAdapter

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import actualizar_entitystatus as script  # noqa: E402

URL_BASE = "https://cuenta-prueba.suitetalk.api.netsuite.com/services/rest/record/v1/"

ENV_FALSO = """\
ENDPOINT_NETSUITE=https://cuenta-prueba.suitetalk.api.netsuite.com/services/rest/
REALM_NETSUITE_API_EVONET=CUENTA_PRUEBA
CONSUMER_KEY_NETSUITE_API_EVONET=consumer-key
CONSUMER_SECRET_NETSUITE_API_EVONET=consumer-secret
TOKEN_KEY_NETSUITE_API_EVONET=token-key
TOKEN_SECRET_NETSUITE_API_EVONET=token-secret
"""

SesionReal = requests.Session

# Cuerpos de error con la forma que devuelve el REST Record API de NetSuite
ERROR_STATUS_INVALIDO = {
    "type": "https://www.rfc-editor.org/rfc/rfc9110.html#section-15.5.1",
    "title": "Bad Request",
    "status": 400,
    "o:errorDetails": [
        {
            "detail": "Invalid Field Value 99 for the following field: entitystatus",
            "o:errorPath": "entitystatus",
            "o:errorCode": "INVALID_KEY_OR_REF",
        }
    ],
}
ERROR_CONCURRENCIA = {
    "type": "https://www.rfc-editor.org/rfc/rfc6585.html#section-4",
    "title": "Too Many Requests",
    "status": 429,
    "o:errorDetails": [
        {
            "detail": "Concurrent request limit exceeded. Request blocked.",
            "o:errorCode": "CONCURRENCY_LIMIT_EXCEEDED",
        }
    ],
}
ERROR_SERVICIO = {
    "type": "https://www.rfc-editor.org/rfc/rfc9110.html#section-15.6.4",
    "title": "Service Unavailable",
    "status": 503,
    "o:errorDetails": [
        {
            "detail": "The service is temporarily unavailable.",
            "o:errorCode": "SERVICE_UNAVAILABLE",
        }
    ],
}


class NetSuiteFalso(BaseAdapter):
    """Responde en lugar de NetSuite y guarda cada petición que recibe.

    Cada respuesta es (codigo_http, cuerpo_json) o una excepción de red para lanzar.
    """

    def __init__(self):
        super().__init__()
        self.respuestas = []
        self.peticiones = []

    def send(self, request, **kwargs):
        self.peticiones.append(request)
        if not self.respuestas:
            raise AssertionError(f"Petición no esperada: {request.method} {request.url}")
        siguiente = self.respuestas.pop(0)
        if isinstance(siguiente, Exception):
            raise siguiente
        codigo, cuerpo = siguiente
        respuesta = requests.Response()
        respuesta.status_code = codigo
        respuesta.request = request
        respuesta.url = request.url
        respuesta._content = b""
        if cuerpo is not None:
            respuesta.headers["Content-Type"] = "application/vnd.oracle.resource+json; type=error"
            respuesta._content = json.dumps(cuerpo).encode()
        return respuesta

    def close(self):
        pass


def bloquear_red(*args, **kwargs):
    raise AssertionError("Las pruebas no deben salir a la red")


def header(peticion, nombre):
    """Valor de un header como texto: al firmar, requests_oauthlib los deja en bytes."""
    valor = peticion.headers[nombre]
    return valor.decode() if isinstance(valor, bytes) else valor


class ExcelTest(unittest.TestCase):
    def setUp(self):
        temporal = tempfile.TemporaryDirectory()
        self.addCleanup(temporal.cleanup)
        self.dir = temporal.name

    def crear_excel(self, filas, nombre="estatus.xlsx"):
        """Crea un Excel cuya primera fila son los encabezados."""
        libro = openpyxl.Workbook()
        for fila in filas:
            libro.active.append(fila)
        ruta = os.path.join(self.dir, nombre)
        libro.save(ruta)
        return ruta


class LeerExcelTest(ExcelTest):
    def test_limpia_los_ids_y_numera_las_filas_como_en_excel(self):
        ruta = self.crear_excel(
            [
                ["id_transaccion", "id_entitystatus"],
                [1234567, "13.0"],
                [" 1234568 ", None],
            ]
        )
        self.assertEqual(
            script.leer_excel(ruta),
            [
                {"fila_excel": 2, "id_transaccion": "1234567", "id_entitystatus": "13"},
                {"fila_excel": 3, "id_transaccion": "1234568", "id_entitystatus": ""},
            ],
        )

    def test_reconoce_encabezados_escritos_a_mano(self):
        ruta = self.crear_excel([["Internal ID", "Entity Status"], [1001, 13]])
        self.assertEqual(
            script.leer_excel(ruta),
            [{"fila_excel": 2, "id_transaccion": "1001", "id_entitystatus": "13"}],
        )

    def test_rechaza_un_excel_de_inside_sales(self):
        # Leyendo por posición, los ids de empleado se enviarían como status
        ruta = self.crear_excel([["id_transaccion", "id_insidesales"], [1001, 361]])
        with self.assertRaises(RuntimeError) as contexto:
            script.leer_excel(ruta)
        self.assertIn("id_entitystatus", str(contexto.exception))


class EjecucionTest(ExcelTest):
    def setUp(self):
        super().setUp()
        # El script trabaja en la carpeta temporal con un .env falso: nunca ve las
        # credenciales reales ni escribe en el logs/ del proyecto.
        with open(os.path.join(self.dir, ".env"), "w") as archivo:
            archivo.write(ENV_FALSO)
        self.parchar(script, "BASE_DIR", self.dir)
        self.parchar(script, "DATA_DIR", os.path.join(self.dir, "data"))
        self.parchar(script, "LOGS_DIR", os.path.join(self.dir, "logs"))
        # Las pausas entre peticiones y entre reintentos solo alargarían las pruebas
        self.parchar(time, "sleep", lambda segundos: None)

        # load_dotenv no pisa variables ya definidas: se quitan las de NetSuite que
        # hubiera en el entorno, y patch.dict lo restaura todo al terminar.
        entorno = mock.patch.dict(os.environ)
        entorno.start()
        self.addCleanup(entorno.stop)
        for linea in ENV_FALSO.splitlines():
            os.environ.pop(linea.split("=")[0], None)

        self.netsuite = NetSuiteFalso()

        def crear_sesion():
            sesion = SesionReal()
            sesion.mount("https://", self.netsuite)
            sesion.mount("http://", self.netsuite)
            return sesion

        self.parchar(requests, "Session", crear_sesion)
        self.parchar(HTTPAdapter, "send", bloquear_red)

    def parchar(self, objeto, nombre, valor):
        parche = mock.patch.object(objeto, nombre, valor)
        parche.start()
        self.addCleanup(parche.stop)

    def ejecutar(self, *argumentos, respuestas=()):
        self.netsuite.respuestas = list(respuestas)
        with contextlib.redirect_stdout(io.StringIO()):
            return script.main(list(argumentos))

    def reportes(self):
        carpeta = os.path.join(self.dir, "logs")
        return sorted(os.listdir(carpeta)) if os.path.isdir(carpeta) else []

    def leer_reporte(self):
        reportes = self.reportes()
        self.assertEqual(len(reportes), 1, reportes)
        hojas = pd.read_excel(
            os.path.join(self.dir, "logs", reportes[0]), sheet_name=None, dtype=str, keep_default_na=False
        )
        resumen = dict(zip(hojas["resumen"]["concepto"], hojas["resumen"]["valor"]))
        return hojas["detalle"], resumen

    def test_sin_tipo_no_procesa_nada(self):
        ruta = self.crear_excel([["id_transaccion", "id_entitystatus"], [1001, 13]])
        self.parchar(script, "TIPO_TRANSACCION", None)
        self.assertEqual(self.ejecutar("--archivo", ruta), 1)
        self.assertEqual(self.netsuite.peticiones, [])
        self.assertEqual(self.reportes(), [])

    def test_dry_run_no_envia_nada_y_deja_las_urls_en_el_reporte(self):
        ruta = self.crear_excel([["id_transaccion", "id_entitystatus"], [1001, 13], [1002, 14]])
        self.assertEqual(self.ejecutar("--tipo", "quote", "--archivo", ruta, "--dry-run"), 0)
        self.assertEqual(self.netsuite.peticiones, [])
        detalle, resumen = self.leer_reporte()
        self.assertEqual(list(detalle["estatus"]), ["OMITIDO", "OMITIDO"])
        self.assertEqual(list(detalle["url"]), [URL_BASE + "estimate/1001", URL_BASE + "estimate/1002"])
        self.assertEqual(resumen["modo"], "DRY_RUN")

    def test_excel_sin_filas_termina_sin_reporte(self):
        ruta = self.crear_excel([["id_transaccion", "id_entitystatus"]])
        self.assertEqual(self.ejecutar("--tipo", "opportunity", "--archivo", ruta), 1)
        self.assertEqual(self.reportes(), [])

    def correr_opportunities(self):
        """Fila 2 sale bien, fila 3 la rechaza NetSuite, fila 4 trae un número de
        documento en vez del internal id y fila 5 no tiene status: esas dos no se envían."""
        ruta = self.crear_excel(
            [
                ["id_transaccion", "id_entitystatus"],
                [1001, 13],
                [1002, 99],
                ["OP-1003", 13],
                [1004, None],
            ]
        )
        return self.ejecutar(
            "--tipo", "opportunity", "--archivo", ruta,
            respuestas=[(204, None), (400, ERROR_STATUS_INVALIDO)],
        )

    def test_envia_un_patch_firmado_por_cada_fila_valida(self):
        self.correr_opportunities()
        enviadas = [(p.method, p.url, json.loads(p.body)) for p in self.netsuite.peticiones]
        self.assertEqual(
            enviadas,
            [
                ("PATCH", URL_BASE + "opportunity/1001", {"entitystatus": {"id": "13"}}),
                ("PATCH", URL_BASE + "opportunity/1002", {"entitystatus": {"id": "99"}}),
            ],
        )
        for peticion in self.netsuite.peticiones:
            # Sin este header NetSuite ignora un campo mal escrito y responde 204 igual
            self.assertEqual(header(peticion, "X-NetSuite-PropertyNameValidation"), "error")
            self.assertTrue(header(peticion, "Authorization").startswith('OAuth realm="CUENTA_PRUEBA"'))

    def test_reporta_cada_fila_y_sale_con_codigo_2_si_hubo_errores(self):
        self.assertEqual(self.correr_opportunities(), 2)
        detalle, resumen = self.leer_reporte()
        self.assertEqual(list(detalle["estatus"]), ["EXITO", "ERROR", "OMITIDO", "OMITIDO"])
        self.assertEqual(list(detalle["codigo_http"]), ["204", "400", "", ""])
        self.assertEqual(
            detalle["error"][1],
            "[INVALID_KEY_OR_REF] Invalid Field Value 99 for the following field: entitystatus",
        )
        self.assertEqual(
            (resumen["modo"], resumen["exitos"], resumen["errores"], resumen["omitidos"]),
            ("REAL", "1", "1", "2"),
        )

    def correr_una_quote(self, respuestas):
        ruta = self.crear_excel([["id_transaccion", "id_entitystatus"], [1001, 13]])
        return self.ejecutar("--tipo", "quote", "--archivo", ruta, respuestas=respuestas)

    def test_reintenta_si_netsuite_limita_la_concurrencia(self):
        self.assertEqual(self.correr_una_quote([(429, ERROR_CONCURRENCIA), (204, None)]), 0)
        self.assertEqual([p.url for p in self.netsuite.peticiones], [URL_BASE + "estimate/1001"] * 2)
        detalle, _ = self.leer_reporte()
        self.assertEqual(list(detalle["estatus"]), ["EXITO"])

    def test_reintenta_si_se_cae_la_conexion(self):
        caida = requests.ConnectionError("Connection reset by peer")
        self.assertEqual(self.correr_una_quote([caida, (204, None)]), 0)
        self.assertEqual(len(self.netsuite.peticiones), 2)
        detalle, _ = self.leer_reporte()
        self.assertEqual(list(detalle["estatus"]), ["EXITO"])

    def test_se_rinde_tras_dos_reintentos(self):
        self.assertEqual(self.correr_una_quote([(503, ERROR_SERVICIO)] * 3), 2)
        self.assertEqual(len(self.netsuite.peticiones), 3)
        detalle, _ = self.leer_reporte()
        self.assertEqual(list(detalle["estatus"]), ["ERROR"])
        self.assertEqual(detalle["error"][0], "[SERVICE_UNAVAILABLE] The service is temporarily unavailable.")


if __name__ == "__main__":
    unittest.main()
