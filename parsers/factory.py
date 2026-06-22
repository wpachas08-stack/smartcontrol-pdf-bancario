"""
ParserFactory — detecta el banco y retorna el parser correcto.
Desencripta el PDF antes de pasarlo al parser.
"""
import re
import tempfile
import os
from pathlib import Path

import pikepdf
import pdfplumber


# Palabras clave para detectar banco desde el texto del PDF/Excel
FIRMAS_BANCO = {
    "bcp":       ["BANCO DE CREDITO", "BANCO DE CRÉDITO", "CREDIBANCO", " BCP "],
    "bbva":      ["BBVA", "BANCO CONTINENTAL", "CONTINENTAL"],
    "scotiabank":["SCOTIABANK", "BANCO SCOTIABANK"],
    "interbank": ["INTERBANK", "BANCO INTERNACIONAL"],
    "nacion":    ["BANCO DE LA NACION", "BANCO DE LA NACIÓN", "BN "],
}


class ParserFactory:

    def procesar(
        self,
        ruta: str,
        extension: str,
        ruc: str,
        banco_forzado: str = "",
    ) -> dict:
        """
        Punto de entrada. Desencripta si es PDF, detecta banco, parsea.
        """
        ruta_trabajo  = ruta
        ruta_dec_tmp  = None

        try:
            if extension == "pdf":
                ruta_trabajo, ruta_dec_tmp = self._desencriptar(ruta, ruc)

            banco = banco_forzado.lower() if banco_forzado else \
                    self._detectar_banco(ruta_trabajo, extension)

            parser = self._get_parser(banco, extension)
            result = parser.parsear(ruta_trabajo)
            result["ok"]    = True
            result["banco"] = banco
            return result

        finally:
            if ruta_dec_tmp and os.path.exists(ruta_dec_tmp):
                os.unlink(ruta_dec_tmp)

    # ------------------------------------------------------------------
    # Desencriptar PDF con pikepdf (soporta AES-128/256 y RC4)
    # ------------------------------------------------------------------
    def _desencriptar(self, ruta: str, password: str):
        """
        Intenta abrir el PDF. Si está encriptado lo desencripta con pikepdf.
        Retorna (ruta_trabajo, ruta_temporal_o_None).
        """
        try:
            pdf = pikepdf.open(ruta)
            pdf.close()
            return ruta, None  # No está encriptado
        except pikepdf.PasswordError:
            pass  # Necesita contraseña
        except Exception:
            return ruta, None

        # Desencriptar con contraseña (RUC)
        fd, ruta_dec = tempfile.mkstemp(suffix=".pdf", prefix="dec_")
        os.close(fd)

        try:
            with pikepdf.open(ruta, password=password) as pdf:
                pdf.save(ruta_dec)
            return ruta_dec, ruta_dec
        except pikepdf.PasswordError:
            os.unlink(ruta_dec)
            raise ValueError(
                f"Contraseña incorrecta. El RUC '{password}' no abre este PDF. "
                "Verifique que el RUC de la empresa coincide con el del estado de cuenta."
            )

    # ------------------------------------------------------------------
    # Detección de banco
    # ------------------------------------------------------------------
    def _detectar_banco(self, ruta: str, extension: str) -> str:
        texto = ""
        try:
            if extension == "pdf":
                with pdfplumber.open(ruta) as pdf:
                    for page in pdf.pages[:2]:
                        texto += (page.extract_text() or "")
                        if len(texto) > 500:
                            break
            else:
                from openpyxl import load_workbook
                wb = load_workbook(ruta, read_only=True, data_only=True)
                ws = wb.active
                for row in ws.iter_rows(max_row=15, values_only=True):
                    for cell in row:
                        if cell:
                            texto += str(cell) + " "
                wb.close()
        except Exception:
            pass

        texto_up = texto.upper()
        for banco, firmas in FIRMAS_BANCO.items():
            for firma in firmas:
                if firma in texto_up:
                    return banco
        return "generico"

    # ------------------------------------------------------------------
    # Instanciar parser
    # ------------------------------------------------------------------
    def _get_parser(self, banco: str, extension: str):
        from parsers.bcp        import BcpPdfParser, BcpExcelParser
        from parsers.bbva       import BbvaPdfParser, BbvaExcelParser
        from parsers.scotiabank import ScotiabankPdfParser, ScotiabankExcelParser
        from parsers.interbank  import InterbankPdfParser, InterbankExcelParser
        from parsers.nacion     import NacionPdfParser, NacionExcelParser
        from parsers.generico   import GenericoPdfParser, GenericoExcelParser

        mapa = {
            ("bcp",        "pdf"):  BcpPdfParser,
            ("bcp",        "xlsx"): BcpExcelParser,
            ("bcp",        "xls"):  BcpExcelParser,
            ("bbva",       "pdf"):  BbvaPdfParser,
            ("bbva",       "xlsx"): BbvaExcelParser,
            ("bbva",       "xls"):  BbvaExcelParser,
            ("scotiabank", "pdf"):  ScotiabankPdfParser,
            ("scotiabank", "xlsx"): ScotiabankExcelParser,
            ("scotiabank", "xls"):  ScotiabankExcelParser,
            ("interbank",  "pdf"):  InterbankPdfParser,
            ("interbank",  "xlsx"): InterbankExcelParser,
            ("interbank",  "xls"):  InterbankExcelParser,
            ("nacion",     "pdf"):  NacionPdfParser,
            ("nacion",     "xlsx"): NacionExcelParser,
            ("nacion",     "xls"):  NacionExcelParser,
        }
        key   = (banco, extension)
        clase = mapa.get(key)
        if clase is None:
            clase = GenericoPdfParser if extension == "pdf" else GenericoExcelParser
        return clase()
