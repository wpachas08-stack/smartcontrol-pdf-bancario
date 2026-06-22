"""
Parser BCP (Banco de Crédito del Perú) — PDF y Excel

Método de extracción PDF: agrupación de palabras por posición Y (coordenadas).
Cada línea de movimiento tiene fecha en x≈45, cargo/abono y saldo al final.

Formato por línea:
  DD-MM  DESCRIPCION  [MED AT]  [LUGAR]  [NRO OP]  [HORA]  [ORIGEN] [TIPO]  MONTO  SALDO
  Cargos: terminan en "-"  (ej: 50.00-)
  Abonos: sin signo        (ej: 500.00)
"""
import re
import pdfplumber
from openpyxl import load_workbook
from parsers.base import BaseParser


class BcpPdfParser(BaseParser):

    def parsear(self, ruta: str, password: str = None) -> dict:
        res         = self.respuesta_vacia()
        movimientos = []
        errores     = []

        try:
            open_kwargs = {"password": password} if password else {}
            with pdfplumber.open(ruta, **open_kwargs) as pdf:
                anio = "2026"
                texto_cabecera = ""

                for i, page in enumerate(pdf.pages):
                    words = page.extract_words()

                    if i == 0:
                        # Extraer cabecera de la primera página
                        texto_cabecera = " ".join(w["text"] for w in words)
                        cab = self._extraer_cabecera(texto_cabecera)
                        res["cabecera"] = cab
                        anio = cab.get("anio", "2026")

                    # Agrupar palabras por línea (coordenada Y redondeada)
                    lineas = {}
                    for w in words:
                        y = round(w["top"] / 2) * 2
                        if y not in lineas:
                            lineas[y] = []
                        lineas[y].append(w)

                    # Procesar cada línea
                    for y in sorted(lineas.keys()):
                        palabras = sorted(lineas[y], key=lambda w: w["x0"])
                        texto_linea = " ".join(w["text"] for w in palabras)
                        mov = self._parsear_linea(texto_linea, anio)
                        if mov:
                            movimientos.append(mov)

        except Exception as e:
            errores.append(f"Error al leer PDF BCP: {str(e)}")

        if not res["cabecera"].get("periodo") and movimientos:
            res["cabecera"]["periodo"] = self.periodo_desde_fechas(movimientos)

        res["movimientos"]  = movimientos
        res["errores"]      = errores
        res["total_leidos"] = len(movimientos)
        return res

    def _extraer_cabecera(self, texto: str) -> dict:
        cab = {
            "numero_cuenta": None,
            "periodo":       None,
            "saldo_inicial": None,
            "saldo_final":   None,
            "moneda":        "PEN",
            "anio":          "2026",
        }

        # Número de cuenta BCP: 191-7365047-0-12
        m = re.search(r'(\d{3}-\d{7}-\d-\d{2})', texto)
        if m:
            cab["numero_cuenta"] = m.group(1)

        # Período: DEL31/03/2026AL30/04/2026
        m = re.search(r'DEL\s*\d{2}/\d{2}/(\d{4})\s*AL\s*\d{2}/(\d{2})/(\d{4})', texto)
        if m:
            cab["anio"]    = m.group(3)
            cab["periodo"] = m.group(3) + m.group(2)  # AAAAMM

        # Saldo final del resumen
        m = re.search(r'(\d{2}/\d{2}/\d{4})\s+([\d,\.]+)\s+[\d,\.]+\s+[\d,\.]+\s+[\d,\.]+\s+[\d,\.]+\s+[\d,\.]+\s+[\d,\.]+\s+([\d,\.]+)', texto)
        if m:
            cab["saldo_inicial"] = self.parse_importe(m.group(2))
            cab["saldo_final"]   = self.parse_importe(m.group(3))

        if "DOLAR" in texto.upper() or "USD" in texto.upper():
            cab["moneda"] = "USD"

        return cab

    def _parsear_linea(self, linea: str, anio: str) -> dict | None:
        """
        Parsea una línea de movimiento BCP.
        Formato: DD-MM  DESCRIPCION ... MONTO  SALDO
        Cargo:   el monto termina en "-"
        Abono:   el monto no tiene signo
        """
        # La línea debe empezar con DD-MM
        m = re.match(r'^(\d{2})-(\d{2})\s+(.+)$', linea)
        if not m:
            return None

        dd   = m.group(1)
        mm   = m.group(2)
        rest = m.group(3).strip()

        fecha = f"{anio}-{mm}-{dd}"

        # Extraer los últimos 2 números de la línea (monto y saldo)
        # Patrón: número con coma/punto, opcionalmente con "-" al final
        nums = re.findall(r'((?:\d{1,3}(?:,\d{3})*|\d+)(?:\.\d{2})?|\.\d{2})(-?)', rest)

        if len(nums) < 1:
            return None

        # Último par = saldo contable
        saldo_val, _ = nums[-1]
        saldo = self.parse_importe(saldo_val)

        # Penúltimo par = monto del movimiento
        if len(nums) >= 2:
            monto_val, monto_sig = nums[-2]
        else:
            monto_val, monto_sig = nums[-1]
            saldo = None

        importe  = self.parse_importe(monto_val)
        es_cargo = monto_sig == '-'

        if importe <= 0:
            return None

        # Extraer descripción: todo antes de los números finales
        # Quitar números con guión del final
        desc = re.sub(r'\s+(?:(?:\d{1,3}(?:,\d{3})*|\d+)(?:\.\d{2})?|\.\d{2})-?\s*(?:(?:\d{1,3}(?:,\d{3})*|\d+)(?:\.\d{2})?|\.\d{2})-?\s*$', '', rest)
        desc = re.sub(r'\s+', ' ', desc).strip()

        # Quitar residuos numéricos del final de la descripción
        desc = re.sub(r'\s+\d{4,}\s*$', '', desc).strip()

        if not desc:
            return None

        # Extraer número de operación (6 dígitos aislados)
        nro_op = None
        m_nro  = re.search(r'\b(\d{6})\b', rest)
        if m_nro:
            nro_op = m_nro.group(1)

        return {
            "fecha_operacion": fecha,
            "fecha_valor":     fecha,
            "referencia":      nro_op,
            "descripcion":     desc[:300],
            "tipo":            "cargo" if es_cargo else "abono",
            "importe":         importe,
            "saldo_banco":     saldo,
            "moneda":          "PEN",
            "tipo_cambio":     1.0,
        }


class BcpExcelParser(BaseParser):
    """Parser BCP Excel. Columnas: Fecha|NroOp|Desc|Cargo|Abono|Saldo"""

    def parsear(self, ruta: str, password: str = None) -> dict:
        res = self.respuesta_vacia()
        movimientos = []
        errores = []
        try:
            wb = load_workbook(ruta, read_only=True, data_only=True)
            ws = wb.active
            texto_cab = ""
            for row in ws.iter_rows(min_row=1, max_row=7, values_only=True):
                texto_cab += " ".join(str(c or "") for c in row) + " "
            res["cabecera"] = self._extraer_cabecera_excel(texto_cab)
            fila_inicio = self._detectar_fila_inicio(ws)
            for row in ws.iter_rows(min_row=fila_inicio, values_only=True):
                fecha = self.parse_fecha(row[0] if row else None)
                if not fecha:
                    continue
                nro_op = str(row[1] or "").strip() if len(row) > 1 else ""
                desc   = str(row[2] or "").strip() if len(row) > 2 else ""
                cargo  = self.parse_importe(row[3] if len(row) > 3 else None)
                abono  = self.parse_importe(row[4] if len(row) > 4 else None)
                saldo  = self.parse_importe(row[5] if len(row) > 5 else None)
                if cargo <= 0 and abono <= 0 or not desc:
                    continue
                tipo = "cargo" if cargo > 0 else "abono"
                movimientos.append({
                    "fecha_operacion": fecha, "fecha_valor": fecha,
                    "referencia": nro_op or None, "descripcion": desc[:300],
                    "tipo": tipo, "importe": round(cargo if cargo > 0 else abono, 2),
                    "saldo_banco": saldo or None,
                    "moneda": res["cabecera"]["moneda"], "tipo_cambio": 1.0,
                })
            wb.close()
            if not res["cabecera"].get("periodo") and movimientos:
                res["cabecera"]["periodo"] = self.periodo_desde_fechas(movimientos)
            res["movimientos"] = movimientos
            res["total_leidos"] = len(movimientos)
        except Exception as e:
            errores.append(f"Error Excel BCP: {str(e)}")
        res["errores"] = errores
        return res

    def _extraer_cabecera_excel(self, texto: str) -> dict:
        t = texto.upper()
        cab = {"numero_cuenta": None, "periodo": None, "saldo_inicial": None, "saldo_final": None, "moneda": "PEN"}
        m = re.search(r'(\d{3}-\d{7}-\d-\d{2})', t)
        if m:
            cab["numero_cuenta"] = m.group(1)
        m = re.search(r'(\d{2})[/\-](\d{4})', t)
        if m:
            cab["periodo"] = m.group(2) + m.group(1)
        if "DOLAR" in t or "USD" in t:
            cab["moneda"] = "USD"
        return cab

    def _detectar_fila_inicio(self, ws) -> int:
        for i, row in enumerate(ws.iter_rows(min_row=5, max_row=20, values_only=True), start=5):
            if row and self.parse_fecha(row[0]):
                return i
        return 9
