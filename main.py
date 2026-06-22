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


from fastapi import Body
from pydantic import BaseModel

class ProcesarBase64Request(BaseModel):
    archivo_base64: str
    extension: str
    ruc: str
    banco: str = ""

@app.post("/procesar-base64")
async def procesar_base64(
    payload: ProcesarBase64Request,
    _key: str = Security(verificar_api_key),
):
    """
    Procesa un archivo enviado en base64 (evita corrupción en multipart).
    """
    import base64 as b64
    
    extension = payload.extension.lower().lstrip(".")
    if extension not in ("pdf", "xlsx", "xls", "csv"):
        raise HTTPException(status_code=400, detail=f"Extensión no permitida: {extension}")

    try:
        contenido = b64.b64decode(payload.archivo_base64)
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Base64 inválido: {str(e)}")

    if len(contenido) == 0:
        raise HTTPException(status_code=400, detail="El archivo está vacío")

    with tempfile.NamedTemporaryFile(
        suffix=f".{extension}", delete=False, prefix="bco_"
    ) as tmp:
        tmp.write(contenido)
        ruta_tmp = tmp.name

    try:
        factory   = ParserFactory()
        resultado = factory.procesar(ruta_tmp, extension, payload.ruc.strip(), payload.banco.strip())
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


@app.post("/diagnostico")
async def diagnostico(
    payload: ProcesarBase64Request,
    _key: str = Security(verificar_api_key),
):
    """Diagnóstico del PDF recibido."""
    import base64 as b64
    import hashlib
    
    try:
        contenido = b64.b64decode(payload.archivo_base64)
    except Exception as e:
        return {"error": f"Base64 inválido: {str(e)}"}

    info = {
        "tamano_bytes": len(contenido),
        "md5": hashlib.md5(contenido).hexdigest(),
        "primeros_bytes": contenido[:20].hex(),
        "es_pdf": contenido[:4] == b'%PDF',
        "version_pdf": contenido[:8].decode('latin-1', errors='replace'),
        "ruc_recibido": payload.ruc,
        "extension": payload.extension,
    }
    
    # Intentar abrir con pdfplumber
    with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
        tmp.write(contenido)
        ruta = tmp.name
    
    try:
        # Sin password
        with pdfplumber.open(ruta) as pdf:
            info["pdfplumber_sin_password"] = f"OK - {len(pdf.pages)} páginas"
    except Exception as e:
        info["pdfplumber_sin_password"] = f"ERROR: {str(e)}"
    
    try:
        # Con password = RUC
        with pdfplumber.open(ruta, password=payload.ruc) as pdf:
            info["pdfplumber_con_ruc"] = f"OK - {len(pdf.pages)} páginas"
            texto = pdf.pages[0].extract_text() or ""
            info["texto_pagina1"] = texto[:200]
    except Exception as e:
        info["pdfplumber_con_ruc"] = f"ERROR: {str(e)}"

    # Intentar con pypdf
    try:
        from pypdf import PdfReader
        reader = PdfReader(ruta)
        info["pypdf_encriptado"] = reader.is_encrypted
        if reader.is_encrypted:
            result = reader.decrypt(payload.ruc)
            info["pypdf_decrypt_result"] = str(result)
    except Exception as e:
        info["pypdf_error"] = str(e)
    
    os.unlink(ruta)
    return info
