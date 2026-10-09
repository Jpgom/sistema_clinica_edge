"""Regressões do Recibo eSocial do portal EDGE (edge_app/application.py).

Teste isolado das funções sem inicializar Flask e seus bancos de dados.
"""
import ast
import os
import re
import tempfile
import unittest
import unicodedata
from unittest.mock import patch
from datetime import datetime, timedelta
from pathlib import Path

import pandas as pd

APP_PATH = Path(__file__).resolve().parents[1] / "edge_app" / "application.py"
NAMES = {
    "normalize_text", "normalize_company_name", "extract_cnpj", "format_cnpj",
    "extract_employer_document", "format_employer_document", "format_employer_document_filename",
    "_find_column_optional", "find_column", "_normalize_marker", "_row_has_ok_esocial",
    "_strip_cnpj_from_company", "read_esocial_base_rows", "select_esocial_rows_for_company",
    "_candidate_priority",
    "prepare_dataframe", "_normalize_person_name", "_safe_text", "_cpf_text",
    "_parse_excel_date", "_format_date", "_score_esocial_export_sheet",
    "_load_esocial_export_tables", "read_esocial_export_file", "make_paragraph",
    "build_esocial_receipt_pdf",
}


def receipt_functions():
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.units import mm
    from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
    from reportlab.platypus import Paragraph, SimpleDocTemplate, Table, TableStyle
    tree = ast.parse(APP_PATH.read_text(encoding="utf-8"))
    nodes = [n for n in tree.body if isinstance(n, ast.FunctionDef) and n.name in NAMES]
    if len(nodes) != len(NAMES):
        raise RuntimeError("Funções do módulo de recibos alteradas; confira os testes.")
    namespace = dict(
        os=os, re=re, pd=pd, unicodedata=unicodedata, Path=Path,
        datetime=datetime, timedelta=timedelta, colors=colors,
        A4=A4, landscape=landscape, mm=mm,
        ParagraphStyle=ParagraphStyle, getSampleStyleSheet=getSampleStyleSheet,
        Paragraph=Paragraph, SimpleDocTemplate=SimpleDocTemplate, Table=Table, TableStyle=TableStyle,
    )
    exec("from __future__ import annotations", namespace)
    exec(compile(ast.Module(body=nodes, type_ignores=[]), str(APP_PATH), "exec"), namespace)
    return namespace


class TestReciboEsocialPortal(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.fn = receipt_functions()

    def test_html_com_extensao_xls_pdf_colunas_a_h(self):
        html = """<!DOCTYPE html><html><body><table>
        <tr><th>EMPRESA</th><th>CNPJ</th><th>FUNCIONARIO</th><th>CPF</th><th>Data</th><th>SEFAZ</th><th>esocial</th><th>RESULTADO</th><th>MENSAGEM</th></tr>
        <tr><td>EMPRESA TESTE</td><td>04854015000178</td><td>TESTE FUNCIONARIO</td><td>00123456789</td><td>08/08/2026</td><td>UUID-123</td><td>1.1.0000020</td><td>AUTORIZADO</td><td>mensagem não deve aparecer</td></tr>
        </table></body></html>"""
        with tempfile.TemporaryDirectory() as td:
            planilha = Path(td) / "envios.xls"
            planilha.write_text(html, encoding="utf-8")
            df, sheet = self.fn["read_esocial_export_file"](str(planilha))
            self.assertEqual(len(df), 1)
            self.assertEqual(df.iloc[0]["CNPJ"], "04854015000178")
            self.assertEqual(df.iloc[0]["CPF"], "00123456789")
            self.assertEqual(df.iloc[0]["FORMATO_ENVIO"], "ENVIO")
            self.assertEqual(sheet, "TABELA HTML 1")

            import fitz
            pdf_path = Path(td) / "recibo.pdf"
            self.fn["build_esocial_receipt_pdf"](df, str(pdf_path), "EMPRESA TESTE", "04854015000178")
            with fitz.open(str(pdf_path)) as pdf:
                content = "\n".join(page.get_text() for page in pdf)
                self.assertIn("04.854.015/0001-78", content)
                self.assertIn("00123456789", content)
                for column in ("EMPRESA", "CNPJ", "FUNCIONARIO", "CPF", "Data", "SEFAZ", "esocial", "RESULTADO"):
                    self.assertIn(column, content)
                self.assertNotIn("MENSAGEM", content)
                self.assertNotIn("mensagem não deve aparecer", content)

    def test_cnpj_numerico_sem_zero(self):
        self.assertEqual(self.fn["extract_cnpj"](4854015000178), "04854015000178")
        self.assertEqual(self.fn["format_cnpj"]("04854015000178"), "04.854.015/0001-78")
        self.assertEqual(self.fn["extract_cnpj"]("123456789012"), "")

    def test_cpf_de_empregador_pessoa_fisica_em_envios_xls(self):
        html = """<html><body><table>
        <tr><th>EMPRESA</th><th>CNPJ</th><th>FUNCIONARIO</th><th>CPF</th><th>Data</th><th>SEFAZ</th><th>esocial</th><th>RESULTADO</th><th>MENSAGEM</th></tr>
        <tr><td>EMPREGADOR PESSOA FISICA</td><td>12345678909</td><td>TRABALHADOR A</td><td>98765432100</td><td>01/10/2026</td><td></td><td>1.1.000222</td><td>AUTORIZADO</td><td>não incluir</td></tr>
        </table></body></html>"""
        with tempfile.TemporaryDirectory() as td:
            xls = Path(td) / "envios.xls"
            xls.write_text(html, encoding="utf-8")
            df, _ = self.fn["read_esocial_export_file"](str(xls))
            self.assertEqual(len(df), 1)
            self.assertEqual(df.iloc[0]["CNPJ"], "12345678909")  # CPF do empregador
            self.assertEqual(df.iloc[0]["CPF"], "98765432100")  # CPF do funcionário
            pdf = Path(td) / "empregador.pdf"
            self.fn["build_esocial_receipt_pdf"](df, str(pdf), "EMPREGADOR PESSOA FISICA", "12345678909")
            import fitz
            with fitz.open(pdf) as result:
                txt = " ".join(page.get_text() for page in result)
            self.assertIn("123.456.789-09", txt)
            self.assertIn("CPF/CNPJ", txt)
            self.assertIn("98765432100", txt)
            self.assertNotIn("não incluir", txt)
            self.assertEqual(
                self.fn["format_employer_document_filename"]("12345678909"), "123.456.789-09"
            )

    def test_base_aceita_cpf_empregador_e_cruzamento_pelo_nome_funcionario(self):
        base_data = pd.DataFrame([
            {"SETOR": "EMPREGADOR PESSOA FISICA - CPF: 123.456.789-09", "FUNCIONÁRIO": "TRABALHADOR A", "OBSERVACAO": "OK E-SOCIAL"},
            {"SETOR": "EMPREGADOR PESSOA FISICA - CPF: 123.456.789-09", "FUNCIONÁRIO": "TRABALHADOR B", "OBSERVACAO": "OK E-SOCIAL"},
            {"SETOR": "OUTRO EMPREGADOR - CPF: 222.222.222-22", "FUNCIONÁRIO": "TRABALHADOR A", "OBSERVACAO": "OK E-SOCIAL"},
        ])
        with patch.object(pd, "read_excel", return_value=base_data):
            base_rows = self.fn["read_esocial_base_rows"]("planilha.xlsx", "OUTUBRO.2026 OK")
        self.assertEqual(len(base_rows), 3)
        filtered = base_rows[base_rows["CNPJ"] == "12345678909"]
        self.assertEqual(len(filtered), 2)
        self.assertEqual(filtered.iloc[0]["EMPRESA_NOME"], "EMPREGADOR PESSOA FISICA")
        exports = pd.DataFrame([
            {"CNPJ": "12345678909", "NOME_KEY": "TRABALHADOR A", "FUNCIONARIO": "TRABALHADOR A", "RECIBO": "1.1.223", "EVENTO": "S-2220"},
            {"CNPJ": "12345678909", "NOME_KEY": "OUTRO FUNCIONARIO", "FUNCIONARIO": "OUTRO FUNCIONARIO", "RECIBO": "1.1.225", "EVENTO": "S-2220"},
        ])
        selected, missing = self.fn["select_esocial_rows_for_company"](filtered, exports)
        self.assertEqual(len(selected), 1)
        self.assertEqual(missing, ["TRABALHADOR B"])

    def test_documento_do_empregador_preserva_zero_inicial(self):
        doc = self.fn["extract_employer_document"]
        self.assertEqual(doc("012.345.678-90"), "01234567890")
        self.assertEqual(doc(1234567890), "01234567890")
        self.assertEqual(self.fn["format_employer_document"]("01234567890"), "012.345.678-90")
        self.assertEqual(doc("04.854.015/0001-78"), "04854015000178")
        self.assertEqual(doc(4854015000178), "04854015000178")
        self.assertEqual(doc("documento incorreto"), "")
        self.assertEqual(doc("123456789012"), "")

    def test_base_sem_cpf_do_empregador_conserva_nome_para_correspondencia_segura(self):
        base_data = pd.DataFrame([
            {"SETOR": "EMPREGADOR PESSOA FISICA", "FUNCIONÁRIO": "TRABALHADOR A", "OBSERVACAO": "OK E-SOCIAL"},
        ])
        with patch.object(pd, "read_excel", return_value=base_data):
            base_rows = self.fn["read_esocial_base_rows"]("planilha.xlsx", "OUTUBRO.2026 OK")
        self.assertEqual(base_rows.iloc[0]["CNPJ"], "")
        self.assertEqual(base_rows.iloc[0]["EMPRESA_KEY"], "EMPREGADOR PESSOA FISICA")

    def test_extracao_legada_continua_compativel(self):
        html = """<html><body><table>
        <tr><th>EVENTO</th><th>EMPRESA</th><th>CNPJ</th><th>FUNCIONARIO</th><th>CPF</th><th>DATA REF.</th><th>STATUS</th><th>RECIBO</th></tr>
        <tr><td>S-2220</td><td>EMPRESA TESTE</td><td>04854015000178</td><td>EXEMPLO FUNCIONARIO</td><td>00123456789</td><td>08/08/2026</td><td>AUTORIZADO</td><td>1.1.1234</td></tr>
        </table></body></html>"""
        with tempfile.TemporaryDirectory() as td:
            planilha = Path(td) / "RELFUNCGERAL.xls"
            planilha.write_text(html, encoding="utf-8")
            df, _ = self.fn["read_esocial_export_file"](str(planilha))
            self.assertEqual(df.iloc[0]["FORMATO_ENVIO"], "LEGADO")
            self.assertEqual(df.iloc[0]["RECIBO"], "1.1.1234")

    def test_html_indice_excel_sem_dados_informa_solucao(self):
        with tempfile.TemporaryDirectory() as td:
            index = Path(td) / "RELFUNCGERAL.xls"
            index.write_text('<html><frameset><frame src="RELFUNCGERAL_arquivos/sheet001.htm"></frameset></html>')
            with self.assertRaisesRegex(ValueError, "página índice do Excel sem tabela"):
                self.fn["read_esocial_export_file"](str(index))


if __name__ == "__main__":
    unittest.main()
