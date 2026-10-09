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
    "_employer_documents_in_text", "extract_employer_document", "format_employer_document", "format_employer_document_filename",
    "_find_column_optional", "find_column", "_normalize_marker", "_row_has_ok_esocial",
    "_strip_cnpj_from_company", "read_esocial_base_rows", "select_esocial_rows_for_company",
    "_candidate_priority", "_is_valid_esocial_receipt",
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
    namespace["ESOCIAL_MONTHS"] = {
        "JANEIRO": "JANEIRO", "FEVEREIRO": "FEVEREIRO", "MARCO": "MARÇO",
        "ABRIL": "ABRIL", "MAIO": "MAIO", "JUNHO": "JUNHO",
        "JULHO": "JULHO", "AGOSTO": "AGOSTO", "SETEMBRO": "SETEMBRO",
        "OUTUBRO": "OUTUBRO", "NOVEMBRO": "NOVEMBRO", "DEZEMBRO": "DEZEMBRO",
    }
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

    def test_cnpj_duplicado_na_base_abril_e_numeros_no_nome(self):
        value = "B G SANTOS NAZARE EIRELI - ME 16.804.110/0001-46 - 16.804.110/0001-46"
        self.assertEqual(self.fn["extract_employer_document"](value), "16804110000146")
        self.assertEqual(
            self.fn["_strip_cnpj_from_company"](value, "16804110000146"),
            "B G SANTOS NAZARE EIRELI - ME",
        )
        self.assertEqual(self.fn["extract_employer_document"](
            "EMPRESA 123 - 16.804.110/0001-46 - 16.804.110/0001-46"
        ), "16804110000146")
        base = pd.DataFrame([
            {"SETOR": value, "FUNCIONÁRIO": "ROMULO HENRIQUE CORREA LIMA",
             "DATA": datetime(2026, 4, 23), "OBS": "OK E-SOCIAL"},
        ])
        with patch.object(pd, "read_excel", return_value=base):
            rows = self.fn["read_esocial_base_rows"]("planilha.xlsx", "ABRIL.2026")
        self.assertEqual(rows.iloc[0]["CNPJ"], "16804110000146")
        self.assertEqual(rows.iloc[0]["EMPRESA_NOME"], "B G SANTOS NAZARE EIRELI - ME")
        export = pd.DataFrame([
            {"CNPJ": "16804110000146", "NOME_KEY": "ROMULO HENRIQUE CORREA LIMA",
             "FUNCIONARIO": "ROMULO HENRIQUE CORREA LIMA", "CPF": "03907802276",
             "DATA_REF_DATE": datetime(2026, 4, 23).date(),
             "RECIBO": "1.1.000000012", "EVENTO": "", "STATUS": "AUTORIZADO"}
        ])
        selection, missing = self.fn["select_esocial_rows_for_company"](rows, export)
        self.assertEqual(len(selection), 1)
        self.assertEqual(missing, [])

    def test_documentos_conflitantes_nao_se_associam_ao_primeiro(self):
        doc = self.fn["extract_employer_document"]
        self.assertEqual(doc("EMPRESA 16.804.110/0001-46 / 11.222.333/0001-44"), "")
        self.assertEqual(doc("EMPRESA 1680411000014616804110000146"), "")
        self.assertEqual(doc("EMPRESA - 16.804.110/0001-46 - 16804110000146"), "16804110000146")
        base = pd.DataFrame([
            {"SETOR": "EMPRESA - 16.804.110/0001-46 - 11.222.333/0001-44",
             "FUNCIONÁRIO": "FUNCIONARIO A", "OBS": "OK E-SOCIAL"},
            {"SETOR": "EMPRESA - 16.804.110/0001-46",
             "FUNCIONÁRIO": "FUNCIONARIO B", "OBS": "OK E-SOCIAL"},
        ])
        with patch.object(pd, "read_excel", return_value=base):
            parsed = self.fn["read_esocial_base_rows"]("base.xlsx", "ABRIL.2026")
        self.assertEqual(parsed["FUNCIONARIO_BASE"].tolist(), ["FUNCIONARIO B"])
        self.assertEqual(len(parsed.attrs["identification_warnings"]), 1)

    def test_nao_utiliza_recibo_rejeitado_ou_de_exame_em_outra_data(self):
        company = pd.DataFrame([{
            "NOME_KEY": "FUNCIONARIO A", "FUNCIONARIO_BASE": "FUNCIONARIO A",
            "BASE_DATE": datetime(2026, 4, 23).date(), "BASE_ROW": 10,
        }])
        export = pd.DataFrame([
            {"CNPJ": "16804110000146", "NOME_KEY": "FUNCIONARIO A",
             "CPF": "11111111111", "DATA_REF_DATE": datetime(2026, 4, 23).date(),
             "STATUS": "REJEITADO", "RECIBO": "1.1.111", "EVENTO": ""},
            {"CNPJ": "16804110000146", "NOME_KEY": "FUNCIONARIO A",
             "CPF": "11111111111", "DATA_REF_DATE": datetime(2026, 3, 23).date(),
             "STATUS": "AUTORIZADO", "RECIBO": "1.1.222", "EVENTO": ""},
            {"CNPJ": "16804110000146", "NOME_KEY": "FUNCIONARIO A",
             "CPF": "11111111111", "DATA_REF_DATE": datetime(2026, 4, 23).date(),
             "STATUS": "AUTORIZADO", "RECIBO": "", "EVENTO": ""},
        ])
        selected, missing = self.fn["select_esocial_rows_for_company"](company, export)
        self.assertEqual(len(selected), 0)
        self.assertTrue(missing)

    def test_homonimos_cpfs_diferentes_sao_sinalizados(self):
        company = pd.DataFrame([{
            "NOME_KEY": "NOME COMUM", "FUNCIONARIO_BASE": "NOME COMUM",
            "BASE_DATE": None, "BASE_ROW": 2,
        }])
        export = pd.DataFrame([
            {"CNPJ": "16804110000146", "NOME_KEY": "NOME COMUM",
             "CPF": "11111111111", "STATUS": "AUTORIZADO", "RECIBO": "1.1.111",
             "EVENTO": ""},
            {"CNPJ": "16804110000146", "NOME_KEY": "NOME COMUM",
             "CPF": "22222222222", "STATUS": "AUTORIZADO", "RECIBO": "1.1.222",
             "EVENTO": ""},
        ])
        selected, missing = self.fn["select_esocial_rows_for_company"](company, export)
        self.assertTrue(selected.empty)
        self.assertIn("homônimos", missing[0])

    def test_sem_data_na_base_impede_cruzamento_com_ano_errado(self):
        base = pd.DataFrame([{
            "NOME_KEY": "FUNCIONARIO A", "FUNCIONARIO_BASE": "FUNCIONARIO A",
            "BASE_DATE": None, "BASE_ROW": 12,
        }])
        export = pd.DataFrame([{
            "CNPJ": "16804110000146", "NOME_KEY": "FUNCIONARIO A", "CPF": "11111111111",
            "DATA_REF_DATE": datetime(2025, 4, 23).date(),
            "RECIBO": "1.1.345", "STATUS": "AUTORIZADO", "EVENTO": "S-2220",
        }])
        selected, missing = self.fn["select_esocial_rows_for_company"](base, export, "ABRIL", 2026)
        self.assertTrue(selected.empty)
        self.assertIn("fora da competência", missing[0])

    def test_exportacao_outros_eventos_nao_pode_virar_recibo_asos(self):
        html = """<html><body><table>
        <tr><th>EVENTO</th><th>EMPRESA</th><th>CNPJ</th><th>FUNCIONARIO</th>
        <th>CPF</th><th>DATA</th><th>STATUS</th><th>RECIBO</th></tr>
        <tr><td>S-2240</td><td>EMPRESA TESTE</td><td>16804110000146</td>
        <td>FUNCIONARIO A</td><td>11111111111</td><td>23/04/2026</td>
        <td>AUTORIZADO</td><td>1.1.345</td></tr></table></body></html>"""
        with tempfile.TemporaryDirectory() as td:
            f = Path(td) / "outro_evento.xls"
            f.write_text(html, encoding="utf-8")
            with self.assertRaisesRegex(ValueError, "não contém eventos S-2220"):
                self.fn["read_esocial_export_file"](str(f))

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
