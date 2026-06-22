# SmartControl — Microservicio PDF Bancario

Procesa estados de cuenta PDF/Excel de bancos peruanos y retorna JSON.

## Bancos soportados
- BCP (PDF encriptado con RUC + Excel)
- BBVA (PDF + Excel)
- Scotiabank (PDF + Excel)
- Interbank (PDF + Excel)
- Banco de la Nación (PDF + Excel)
- Genérico (cualquier banco, auto-detecta columnas)

---

## Despliegue en Render.com (GRATIS)

### 1. Subir a GitHub
```bash
git init
git add .
git commit -m "SmartControl PDF Bancario v1.0"
git branch -M main
git remote add origin https://github.com/TU_USUARIO/smartcontrol-pdf-bancario.git
git push -u origin main
```

### 2. Crear servicio en Render.com
1. Ir a https://render.com → New → Web Service
2. Conectar el repositorio de GitHub
3. Configurar:
   - **Name:** smartcontrol-pdf-bancario
   - **Environment:** Python 3
   - **Build Command:** `pip install -r requirements.txt`
   - **Start Command:** `uvicorn main:app --host 0.0.0.0 --port $PORT`
   - **Plan:** Free
4. Agregar variable de entorno:
   - `API_KEY` = (una clave secreta larga, ej: `sc_pdf_2026_xK9mQ2nR5p`)
5. Click en **Create Web Service**

Render.com dará una URL como:
`https://smartcontrol-pdf-bancario.onrender.com`

---

## Configurar en el ERP (Ferozo)

En `config/config.php` del ERP agregar:
```php
define('PDF_SERVICE_URL', 'https://smartcontrol-pdf-bancario.onrender.com');
define('PDF_SERVICE_KEY', 'sc_pdf_2026_xK9mQ2nR5p');  // misma que en Render
```

---

## Uso del endpoint

```
POST /procesar
Headers:
  X-API-Key: tu_api_key
Body (multipart/form-data):
  archivo: [archivo PDF o Excel]
  ruc: 20615570673          (contraseña del PDF BCP)
  banco: bcp                (opcional, auto-detecta)
```

### Respuesta exitosa:
```json
{
  "ok": true,
  "banco": "bcp",
  "cabecera": {
    "numero_cuenta": "1930123456789",
    "periodo": "202604",
    "saldo_inicial": 15000.00,
    "saldo_final": 12345.67,
    "moneda": "PEN"
  },
  "movimientos": [
    {
      "fecha_operacion": "2026-04-01",
      "fecha_valor": "2026-04-01",
      "referencia": "1234567890",
      "descripcion": "PAGO SUNAT RUC 20615570673",
      "tipo": "cargo",
      "importe": 2500.00,
      "saldo_banco": 12500.00,
      "moneda": "PEN",
      "tipo_cambio": 1.0
    }
  ],
  "total": 45,
  "errores": []
}
```

---

## Nota sobre free tier de Render.com
El servicio gratuito se "duerme" tras 15 min de inactividad.
La primera solicitud después del sleep tarda ~30 segundos.
Las siguientes son instantáneas.

Para evitar el cold start, configurar un cron en Ferozo que haga ping cada 10 min:
```
*/10 * * * * curl -s https://smartcontrol-pdf-bancario.onrender.com > /dev/null
```
