"""Regressão isolada: leitura do arquivo envios e geração do PDF Recibo eSocial."""
from __future__ import annotations

import ast
import re
import tempfile
import unittest
import zipfile
from datetime import datetime
from html.parser import HTMLParser
from pathlib import Path
from typing import Any


APP_FILE = Path(__file__).resolve().parents[1] / 'pgr_app' / 'app.py'
SYMBOLS = {
    '_HtmlTableParser', '_receipt_norm_header', 'RECEIPT_PDF_COLUMNS',
    'LEGACY_RECEIPT_PDF_COLUMNS', '_format_receipt_cnpj', '_receipt_columns',
    '_parse_html_table_rows', '_html_text', '_parse_xlsx_table_rows',
    '_parse_binary_xls_table_rows', '_receipt_header_map',
    '_receipt_rows_from_raw_rows', '_read_receipt_rows_from_file',
    '_read_receipt_spreadsheet', '_wrap_pdf_text', '_draw_centered_cell',
    '_receipt_pdf_font_config', '_build_receipts_pdf',
}


def load_receipt_helpers():
    """Extrai as funções reais sem executar a inicialização do Flask e do banco."""
    module = ast.parse(APP_FILE.read_text(encoding='utf-8'))
    selected = [
        n for n in module.body if
        (isinstance(n, (ast.FunctionDef, ast.ClassDef)) and n.name in SYMBOLS)
        or (isinstance(n, ast.Assign) and any(isinstance(t, ast.Name) and t.id in SYMBOLS for t in n.targets))
    ]
    ns = {
        're': re, 'Path': Path, 'Any': Any,
        'HTMLParser': HTMLParser, 'datetime': datetime,
        'tempfile': tempfile, 'zipfile': zipfile,
        'load_workbook': None, 'xlrd': None,
    }
    exec('from __future__ import annotations', ns)
    exec(compile(ast.Module(body=selected, type_ignores=[]), str(APP_FILE), 'exec'), ns)
    return ns


class TestRecibosEsocialEnvios(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.api = load_receipt_helpers()

    def test_cnpj_com_zero_inicial_e_formatacao(self):
        formatter = self.api['_format_receipt_cnpj']
        self.assertEqual(formatter('04854015000178'), '04.854.015/0001-78')
        self.assertEqual(formatter(4854015000178), '04.854.015/0001-78')
        self.assertEqual(formatter('04.854.015/0001-78'), '04.854.015/0001-78')
        with self.assertRaises(ValueError):
            formatter('123456')

    def test_importacao_html_ignora_mensagem(self):
        html = '''<!doctype html><html><body><table>
        <tr><th>EMPRESA</th><th>CNPJ</th><th>FUNCIONARIO</th><th>CPF</th><th>Data</th><th>SEFAZ</th><th>esocial</th><th>RESULTADO</th><th>MENSAGEM</th></tr>
        <tr><td>EMPRESA TESTE</td><td>04854015000178</td><td>TRABALHADOR EXEMPLO</td><td>00123456789</td><td>08/08/2026</td><td>ABC</td><td>1.1.0000000</td><td>AUTORIZADO</td><td>texto não exportável</td></tr>
        </table></body></html>'''
        with tempfile.TemporaryDirectory() as td:
            path = Path(td)/'envios.xls'
            path.write_text(html, encoding='utf-8')
            records = self.api['_read_receipt_spreadsheet'](path)
            self.assertEqual(len(records), 1)
            self.assertEqual(len(records[0]), 8)
            self.assertEqual(records[0]['cnpj'], '04.854.015/0001-78')
            self.assertEqual(records[0]['cpf'], '00123456789')
            self.assertNotIn('mensagem', records[0])
            try:
                import fitz
            except ImportError:
                self.skipTest('PyMuPDF não instalado')
            out = Path(td)/'recibo.pdf'
            self.api['_build_receipts_pdf'](records, out)
            with fitz.open(out) as pdf:
                text = '\n'.join(p.get_text() for p in pdf)
                self.assertIn('04.854.015/0001-78', text)
                self.assertIn('00123456789', text)
                self.assertIn('RESULTADO', text)
                self.assertNotIn('MENSAGEM', text)
                self.assertNotIn('texto não exportável', text)
                self.assertLess(pdf[0].rect.height, pdf[0].rect.width)

    def test_formato_legado_preservado(self):
        rows = [
            ['EVENTO','empresa','NOME','CPF','TIPO','STATUS','DATA','Recibo eSocial','Recibo Sefaz'],
            ['S-2220','EMPRESA','NOME','123','ADM','AUTORIZADO','01/08/2026','1.1.123','abc']
        ]
        records = self.api['_receipt_rows_from_raw_rows'](rows)
        self.assertEqual(len(records), 1)
        self.assertEqual(records[0]['evento'], 'S-2220')
        self.assertEqual(len(self.api['_receipt_columns'](records)), 9)


if __name__ == '__main__':
    unittest.main()
