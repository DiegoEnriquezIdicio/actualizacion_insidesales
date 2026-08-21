# Cambios de Inside Sales en NetSuite

Actualiza masivamente el **Sales Rep (Inside Sales)** del Sales Team en transacciones de
NetSuite a partir de un Excel, y genera un reporte de auditoria en Excel.

Funciona para tres tipos de transaccion: **Sales Orders, Quotes y Opportunities**.

---

## Que se necesita

| | |
|---|---|
| **Python** | 3.8 o superior |
| **Librerias** | pandas, requests, requests-oauthlib, python-dotenv, openpyxl (se instalan con `requirements.txt`) |
| **Credenciales** | 6 variables de NetSuite en un archivo `.env` (ver paso 3) |
| **Excel de entrada** | Con las columnas `id_transaccion` e `id_insidesales` (ver paso 4) |

Las carpetas `data/` y `logs/` se crean solas en la primera ejecucion.

---

## Paso a paso

### 1. Clonar el repositorio

```bash
git clone <url-del-repo>
cd cambios_inside_sales
```

### 2. Instalar las dependencias

```bash
python -m venv .venv
```

Activar el entorno virtual:

```bash
source .venv/bin/activate        # Linux / macOS
.venv\Scripts\activate           # Windows
```

Instalar:

```bash
pip install -r requirements.txt
```

### 3. Configurar las credenciales

Copiar la plantilla:

```bash
cp .env.example .env             # Windows: copy .env.example .env
```

Abrir el `.env` y llenar las 6 variables. **No estan en el repositorio**, hay que pedirlas
al responsable de la integracion:

```
ENDPOINT_NETSUITE=https://<cuenta>.suitetalk.api.netsuite.com/services/rest/
REALM_NETSUITE_API_EVONET=<id de la cuenta>
CONSUMER_KEY_NETSUITE_API_EVONET=
CONSUMER_SECRET_NETSUITE_API_EVONET=
TOKEN_KEY_NETSUITE_API_EVONET=
TOKEN_SECRET_NETSUITE_API_EVONET=
```

El `.env` esta en `.gitignore`: no se sube al repositorio.

### 4. Preparar el Excel de entrada

Guardar el archivo dentro de la carpeta `data/`. En la **primera hoja**, con los
encabezados en la **fila 1** y estas dos columnas:

| id_transaccion | id_insidesales |
|----------------|----------------|
| 1234567        | 448551         |
| 1234568        | 343            |

- `id_transaccion`: **internal id** de la transaccion en NetSuite (numerico). No es el
  numero de documento tipo `SO12345`.
- `id_insidesales`: **internal id del empleado** que sera el nuevo Inside Sales.
- Un archivo por tipo de transaccion, sin mezclar Sales Orders con Quotes.
- Las filas vacias o con valores no numericos no se envian: quedan marcadas como `OMITIDO`
  en el reporte.
- **Cerrar el archivo en Excel** antes de ejecutar el script.

En `data/plantilla_inside_sales.xlsx` hay un ejemplo listo para copiar.

### 5. Probar sin enviar nada (recomendado)

```bash
python actualizar_inside_sales.py --tipo salesorder --dry-run
```

Con `--dry-run` el script lee el Excel, valida las filas y muestra en pantalla exactamente
lo que enviaria, **sin tocar NetSuite**. Sirve para confirmar que el archivo esta bien
armado y que los ids son los correctos.

Cambiar `--tipo` segun la transaccion a actualizar:

```bash
python actualizar_inside_sales.py --tipo salesorder --dry-run
python actualizar_inside_sales.py --tipo quote --dry-run
python actualizar_inside_sales.py --tipo opportunity --dry-run
```

### 6. Ejecutar de verdad

La misma linea, **sin** `--dry-run`:

```bash
python actualizar_inside_sales.py --tipo salesorder
```

El script recorre el Excel fila por fila y hace una peticion a NetSuite por cada renglon.
El avance se ve en pantalla:

```
[fila 2] EXITO  salesorder 1234567 -> inside sales 448551
[fila 3] ERROR  salesorder 1234568 (400): [INVALID_KEY_OR_REF] Invalid employee reference.
```

Si una fila falla, las siguientes continuan.

> Antes de una carga masiva, correr primero un Excel de 1 o 2 filas y verificar el
> resultado en la interfaz de NetSuite.

### 7. Revisar el reporte

Al terminar imprime un resumen y la ruta del reporte:

```
Total procesado : 150
Exitos          : 148
Errores         : 2
Omitidos        : 0
Reporte         : .../logs/reporte_salesorder_20260821_143012.xlsx
```

El reporte queda en `logs/` (ver seccion siguiente).

---

## Comandos disponibles

```bash
# Tipo de transaccion (obligatorio salvo que se edite TIPO_TRANSACCION en el script)
python actualizar_inside_sales.py --tipo salesorder
python actualizar_inside_sales.py --tipo quote
python actualizar_inside_sales.py --tipo opportunity

# Simulacion, no envia nada a NetSuite
python actualizar_inside_sales.py --tipo quote --dry-run

# Elegir un Excel especifico (por defecto toma el mas reciente de data/)
python actualizar_inside_sales.py --tipo quote --archivo data/quotes_agosto.xlsx

# Usar otro archivo de credenciales (por ejemplo un ambiente de pruebas)
python actualizar_inside_sales.py --tipo quote --env .env_test

# Ver todas las opciones
python actualizar_inside_sales.py --help
```

---

## Reporte de auditoria

Se genera un archivo por corrida en `logs/reporte_{tipo}_{YYYYMMDD_HHMMSS}.xlsx`.

Hoja **detalle**, una fila por registro procesado:

| Columna | Descripcion |
|---------|-------------|
| `fila_excel` | Numero de fila en el Excel de entrada |
| `id_transaccion` | Internal id de la transaccion |
| `id_insidesales` | Internal id del nuevo Inside Sales |
| `estatus` | `EXITO` (HTTP 204), `ERROR` u `OMITIDO` |
| `codigo_http` | Codigo devuelto por NetSuite |
| `error` | Codigo y detalle del error, ej. `[NONEXISTENT_ID] The record instance does not exist.` |
| `url` | URL exacta a la que se hizo la peticion |
| `fecha_hora` | Momento de la peticion |

Hoja **resumen**: tipo de transaccion, archivo de entrada, fecha, modo y totales.

Para reprocesar solo lo fallido: filtrar la hoja `detalle` por `estatus = ERROR`, armar un
Excel nuevo con esas dos columnas y volver a ejecutar.

---

## Que hace por cada fila

`PATCH {ENDPOINT_NETSUITE}record/v1/{recurso}/{id_transaccion}?replace=salesTeam`

| `--tipo` | Path | Body |
|----------|------|------|
| `salesorder`   | `/salesOrder/{id}?replace=salesTeam`  | employee + isPrimary + salesrole `-2` |
| `quote`        | `/estimate/{id}?replace=salesTeam`    | employee + isPrimary + salesrole `-2` |
| `opportunity`  | `/opportunity/{id}?replace=salesTeam` | employee + isPrimary + salesrole `-2` + `contribution: 100.0` |

Ejemplo de body enviado para `opportunity`:

```json
{"salesTeam": {"items": [{"employee": {"id": 448551}, "isPrimary": true, "salesrole": {"id": "-2"}, "contribution": 100.0}]}}
```

`?replace=salesTeam` reemplaza el sublist completo del Sales Team. Sin ese parametro el
nuevo Inside Sales se agregaria a los miembros existentes en lugar de sustituirlos.

La autenticacion es OAuth 1.0 TBA (HMAC-SHA256 + realm). El rol dueno de los tokens debe
tener permiso de `REST Web Services` y de edicion sobre Sales Orders, Quotes y
Opportunities. Los cambios quedan registrados en NetSuite a nombre de ese usuario de
integracion.

---

## Notas

- El PATCH deja un unico miembro en el Sales Team: el nuevo Inside Sales como primario.
  Los demas miembros que tuviera la transaccion se pierden.
- Si una fila falla, las siguientes continuan. No hay rollback.
- Reintentos automaticos ante errores 429 y 5xx, con pausa de 0.3 s entre peticiones.
- Se puede volver a ejecutar con el mismo Excel sin efectos secundarios.
- Codigo de salida: `0` si no hubo errores, `2` si hubo al menos uno.
