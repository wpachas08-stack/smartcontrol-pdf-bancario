"""
BaseParser — clase base con utilidades compartidas por todos los parsers.
"""
import re
from datetime import datetime
from typing import Optional


class BaseParser:

    def parsear(self, ruta: str) -> dict:
        raise NotImplementedError

    # ------------------------------------------------------------------
    # Parseo de importes
    # ------------------------------------------------------------------
    def parse_importe(self, valor) -> float:
        """
        Convierte string de importe a float.
        Maneja: "1.234,56" → 1234.56  |  "1,234.56" → 1234.56  |  "1234.56" → 1234.56
        """
        if valor is None:
            return 0.0
        s = str(valor).strip().replace(" ", "").replace("\xa0", "")
        s = s.replace("S/", "").replace("$", "").replace("US$", "").strip()
        if not s or s in ("-", "—", ""):
            return 0.0
        # Detectar formato europeo: punto como miles, coma como decimal
        if re.search(r',\d{2}$', s):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
        try:
            return round(float(s), 2)
        except ValueError:
            return 0.0

    # ------------------------------------------------------------------
    # Parseo de fechas
    # ------------------------------------------------------------------
    def parse_fecha(self, valor) -> Optional[str]:
        """Retorna fecha en formato YYYY-MM-DD o None."""
        if valor is None:
            return None

        # Si ya es datetime/date de Python (viene de Excel)
        if hasattr(valor, 'strftime'):
            return valor.strftime('%Y-%m-%d')

        # Si es número serial de Excel
        if isinstance(valor, (int, float)) and 40000 < valor < 55000:
            try:
                from openpyxl.utils.datetime import from_excel
                dt = from_excel(valor)
                return dt.strftime('%Y-%m-%d')
            except Exception:
                pass

        s = str(valor).strip()

        formatos = [
            r'^(\d{2})/(\d{2})/(\d{4})$',  # DD/MM/YYYY
            r'^(\d{2})-(\d{2})-(\d{4})$',  # DD-MM-YYYY
            r'^(\d{4})-(\d{2})-(\d{2})$',  # YYYY-MM-DD
            r'^(\d{2})/(\d{2})/(\d{2})$',  # DD/MM/YY
        ]

        for pat in formatos:
            m = re.match(pat, s)
            if m:
                g = m.groups()
                if len(g[0]) == 4:              # YYYY-MM-DD
                    return f"{g[0]}-{g[1]}-{g[2]}"
                elif int(g[2]) > 31:             # DD/MM/YYYY
                    return f"{g[2]}-{g[1]}-{g[0]}"
                else:                            # DD/MM/YY
                    anio = f"20{g[2]}" if int(g[2]) < 50 else f"19{g[2]}"
                    return f"{anio}-{g[1]}-{g[0]}"

        return None

    # ------------------------------------------------------------------
    # Estructura de respuesta vacía
    # ------------------------------------------------------------------
    def respuesta_vacia(self) -> dict:
        return {
            "cabecera": {
                "numero_cuenta": None,
                "periodo":       None,
                "saldo_inicial": None,
                "saldo_final":   None,
                "moneda":        "PEN",
            },
            "movimientos":  [],
            "errores":      [],
            "total_leidos": 0,
        }

    # ------------------------------------------------------------------
    # Determinar período desde lista de fechas
    # ------------------------------------------------------------------
    def periodo_desde_fechas(self, movimientos: list) -> Optional[str]:
        fechas = [m['fecha_operacion'] for m in movimientos if m.get('fecha_operacion')]
        if not fechas:
            return None
        fechas.sort()
        return fechas[0][:7].replace('-', '')  # AAAAMM desde primera fecha
