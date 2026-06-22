"""
SmartControl ERP — Microservicio de procesamiento PDF/Excel bancario
Desplegado en Render.com (free tier)

Endpoint principal:
  POST /procesar
    - archivo: PDF o Excel del banco
    - ruc: RUC de la empresa (contraseña PDF BCP)
    - banco: código del banco (opcional, auto-detecta)

Respuesta JSON:
  {
    "ok": true,
    "banco": "bcp",
    "cabecera": { "numero_cuenta": "...", "periodo": "...", "saldo_inicial": 0.0, "saldo_final": 0.0, "moneda": "PEN" },
    "movimientos": [
      { "fecha_operacion": "2026-04-01", "referencia": "12345", "descripcion": "...",
        "tipo": "cargo|abono", "importe": 150.00, "saldo_banco": 5000.00, "moneda": "PEN" }
    ],
    "total": 45,
    "errores": []
  }
"""

import os
import tempfile
import traceback
from pathlib import Path

from fastapi import FastAPI, File, Form, UploadFile, HTTPException, Security
from fastapi.security.api_key import APIKeyHeader
from fastapi.responses import JSONResponse
from dotenv import load_dotenv

from parsers.factory import ParserFactory

load_dotenv()

app = FastAPI(
    title="SmartControl PDF Bancario",
    description="Microservicio de extracción de movimientos bancarios",
    version="1.0.0",
    docs_url=None,   # deshabilitar docs en producción
    redoc_url=None,
)

# === Seguridad por API Key ===
API_KEY        = os.getenv("API_KEY", "smartcontrol_2026_secret")
api_key_header = APIKeyHeader(name="X-API-Key", auto_error=False)

def verificar_api_key(key: str = Security(api_key_header)):
    if key != API_KEY:
        raise HTTPException(status_code=403, detail="API Key inválida o ausente")
    return key


@app.get("/")
def health():
    """Health check para Render.com"""
    return {"status": "ok", "servicio": "SmartControl PDF Bancario v1.0"}


@app.post("/procesar")
async def procesar(
    archivo: UploadFile = File(...),
    ruc: str            = Form(...),
    banco: str          = Form(""),
    _key: str           = Security(verificar_api_key),
):
    """
    Procesa un estado de cuenta bancario (PDF o Excel) y retorna
    los movimientos en formato JSON normalizado.
    """
    nombre    = archivo.filename or "archivo"
    extension = Path(nombre).suffix.lower().lstrip(".")

    if extension not in ("pdf", "xlsx", "xls", "csv"):
        raise HTTPException(status_code=400, detail=f"Extensión no permitida: {extension}")

    # Guardar archivo temporalmente
    contenido = await archivo.read()
    if len(contenido) == 0:
        raise HTTPException(status_code=400, detail="El archivo está vacío")
    if len(contenido) > 50 * 1024 * 1024:   # 50 MB máx
        raise HTTPException(status_code=400, detail="Archivo demasiado grande (máx 50MB)")

    with tempfile.NamedTemporaryFile(
        suffix=f".{extension}", delete=False, prefix="bco_"
    ) as tmp:
        tmp.write(contenido)
        ruta_tmp = tmp.name

    try:
        factory   = ParserFactory()
        resultado = factory.procesar(ruta_tmp, extension, ruc.strip(), banco.strip())
        return JSONResponse(content=resultado)

    except ValueError as e:
        raise HTTPException(status_code=422, detail=str(e))
    except Exception as e:
        tb = traceback.format_exc()
        raise HTTPException(status_code=500, detail=f"{str(e)}\n{tb}")
    finally:
        try:
            os.unlink(ruta_tmp)
        except Exception:
            pass
