"""Testa a importação dos relatórios RELFUNCGERAL, inclusive .xls HTML vazios."""

import ast
import re
import tempfile
import unittest
import zipfile
from datetime import datetime
from html.parser import HTMLParser
from pathlib import Path
from typing import Any

APP_SOURCE = Path(__file__).resolve().parents[1] / "pgr_app" / "app.py"


def _load_receipt_reader():
    """Isola as funções de leitura para não exigir banco/Flask ao testar XLS."""
    relevant = {
        "_HtmlTableParser", "_receipt_norm_header", "RECEIPT_PDF_COLUMNS",
        "_parse_html_table_rows", "_html_text", "_parse_xlsx_table_rows",
        "_parse_binary_xls_table_rows", "_receipt_header_map",
        "_receipt_rows_from_raw_rows", "_read_receipt_rows_from_file",
        "_read_receipt_spreadsheet",
    }
    parsed = ast.parse(APP_SOURCE.read_text(encoding="utf-8"))
    nodes = [
        n for n in parsed.body
        if (isinstance(n, (ast.FunctionDef, ast.ClassDef)) and n.name in relevant)
        or (isinstance(n, ast.Assign) and any(
            isinstance(t, ast.Name) and t.id in relevant for t in n.targets
        ))
    ]
    namespace = {
        "Path": Path, "HTMLParser": HTMLParser, "re": re, "tempfile": tempfile,
        "zipfile": zipfile, "datetime": datetime, "Any": Any,
        "load_workbook": None, "xlrd": None,
    }
    exec(compile(ast.fix_missing_locations(ast.Module(nodes, [])), str(APP_SOURCE), "exec"), namespace)
    return namespace["_read_receipt_spreadsheet"]


class RecibosEsocialImportTests(unittest.TestCase):
    def setUp(self):
        self.reader = _load_receipt_reader()
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.folder = Path(self.temp.name)

    def test_html_aspnet_vazio_exibe_orientacao(self):
        xls = self.folder / "RELFUNCGERAL.xls"
        xls.write_text('''<!DOCTYPE html><html><body>
          <form action="./rela_excel.aspx?evento=s2220">
          <input type="hidden" name="__VIEWSTATE" value="123"/>
          </form></body></html>''', encoding="utf-8")
        with self.assertRaisesRegex(ValueError, "exportado sem dados") as exc:
            self.reader(xls)
        self.assertIn("exporte novamente", str(exc.exception))

    def test_html_xls_com_registros_e_aceito(self):
        xls = self.folder / "RELFUNCGERAL.xls"
        xls.write_text('''<!DOCTYPE html><html><body><table>
          <tr><th>EVENTO</th><th>empresa</th><th>NOME</th><th>CPF</th>
          <th>TIPO</th><th>STATUS</th><th>DATA</th>
          <th>Recibo eSocial</th><th>Recibo Sefaz</th></tr>
          <tr><td>S-2220</td><td>EMPRESA DE TESTE</td><td>COLABORADOR DE TESTE</td>
          <td>00011122233</td><td>ADMISSIONAL</td><td>PROCESSADO</td>
          <td>01/10/2026</td><td>NUMERO_TESTE</td><td>SEFAZ_TESTE</td></tr>
          </table></body></html>''', encoding="utf-8")
        rows = self.reader(xls)
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["nome"], "COLABORADOR DE TESTE")
        self.assertEqual(rows[0]["reciboesocial"], "NUMERO_TESTE")

        archive = self.folder / "planilhas.zip"
        with zipfile.ZipFile(archive, "w") as zf:
            zf.write(xls, arcname=xls.name)
        self.assertEqual(self.reader(archive), rows)


if __name__ == "__main__":
    unittest.main()
