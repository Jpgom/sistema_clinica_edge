"""Regressões das posições e do PDF do recibo pré-impresso."""

from __future__ import annotations

import copy
import unittest

from recibo_app.core import (
    build_layout,
    load_default_config,
    normalize_config,
    render_pdf,
    render_preview_png,
)


class ReciboCoreTests(unittest.TestCase):
    def setUp(self) -> None:
        self.config = load_default_config()
        self.payload = {
            "cliente": "Cliente de teste",
            "data": "24/09/2026",
            "itens": [
                {
                    "quantidade": "2",
                    "unidade": "UN",
                    "descricao": "Consulta ocupacional",
                    "preco_unitario": "1.234,56",
                }
            ],
        }

    def test_normalizacao_preserva_calibracao_e_nao_muda_entrada(self) -> None:
        original = copy.deepcopy(self.config)
        original["campo_offsets_mm"]["cliente"]["x"] = 3.25
        original["linha_offsets_mm"] = original["linha_offsets_mm"][:2]
        original["celula_offsets_mm"] = original["celula_offsets_mm"][:1]
        normalized = normalize_config(original)

        self.assertEqual(normalized["campo_offsets_mm"]["cliente"]["x"], 3.25)
        self.assertEqual(len(normalized["linha_offsets_mm"]), 10)
        self.assertEqual(len(normalized["celula_offsets_mm"]), 10)
        self.assertEqual(len(original["linha_offsets_mm"]), 2)
        normalized["campo_offsets_mm"]["cliente"]["x"] = 99
        self.assertEqual(original["campo_offsets_mm"]["cliente"]["x"], 3.25)

    def test_layout_soma_offsets_mas_ignora_correcoes_gdi(self) -> None:
        config = copy.deepcopy(self.config)
        config["linha_offsets_mm"][0]["x"] = 1.5
        config["celula_offsets_mm"][0]["descricao"]["y"] = -0.5
        config["impressao_direta_offset_x_mm"] = 35
        config["impressao_direta_offset_y_mm"] = -2
        layout = build_layout(self.payload, config)
        fields = {field["id"]: field for field in layout["fields"]}

        self.assertEqual(layout["total"], "2.469,12")
        self.assertEqual(fields["cliente"]["x_mm"], 25.0)
        self.assertEqual(fields["cliente"]["y_mm"], 39.2)
        self.assertEqual(fields["descricao-1"]["x_mm"], 35.0)
        self.assertEqual(fields["descricao-1"]["y_mm"], 53.5)
        self.assertEqual(fields["descricao-1"]["line"], 1)
        self.assertEqual(fields["descricao-1"]["key"], "descricao")

    def test_linhas_vazias_com_un_padrao_nao_sao_impressas(self) -> None:
        self.payload["itens"].append(
            {"quantidade": "", "unidade": "UN", "descricao": "", "preco_unitario": ""}
        )
        layout = build_layout(self.payload, self.config)
        self.assertFalse(any(field["id"].endswith("-2") for field in layout["fields"]))

    def test_numeros_invalidos_sao_rejeitados(self) -> None:
        for invalid in ("1,2,3", "nan", "-1", True, "1.2.3"):
            with self.subTest(invalid=invalid):
                payload = copy.deepcopy(self.payload)
                payload["itens"][0]["preco_unitario"] = invalid
                with self.assertRaises(ValueError):
                    build_layout(payload, self.config)

    def test_texto_longo_gera_aviso_e_impede_impressao_sem_truncar(self) -> None:
        payload = copy.deepcopy(self.payload)
        payload["itens"][0]["descricao"] = "Descrição muito extensa " * 4 + "fim"
        layout = build_layout(payload, self.config)
        description = next(f for f in layout["fields"] if f["id"] == "descricao-1")
        self.assertEqual(description["text"], payload["itens"][0]["descricao"])
        self.assertTrue(layout["print_blockers"])
        with self.assertRaisesRegex(ValueError, "Ajuste o recibo"):
            render_pdf(payload, self.config)
        self.assertTrue(render_pdf(payload, self.config, with_model=True).startswith(b"%PDF"))

    def test_colunas_sobrepostas_sao_bloqueadas(self) -> None:
        payload = copy.deepcopy(self.payload)
        payload["itens"][0]["descricao"] = "M" * 27
        payload["itens"][0]["preco_unitario"] = "999.999.999,00"
        layout = build_layout(payload, self.config)
        self.assertTrue(any("sobrepostos" in message for message in layout["print_blockers"]))
        with self.assertRaisesRegex(ValueError, "sobrepostos"):
            render_pdf(payload, self.config)

    def test_texto_malformado_e_data_invalida_sao_rejeitados(self) -> None:
        for value in (23, ["Cliente"], "Linha 1\nLinha 2", "X" * 121):
            with self.subTest(cliente=value):
                payload = copy.deepcopy(self.payload)
                payload["cliente"] = value
                with self.assertRaises(ValueError):
                    build_layout(payload, self.config)
        for value in ("31/02/2026", "2026-09-24", "1/9/2026"):
            with self.subTest(data=value):
                payload = copy.deepcopy(self.payload)
                payload["data"] = value
                with self.assertRaises(ValueError):
                    build_layout(payload, self.config)

    def test_pdf_final_exige_conteudo(self) -> None:
        empty_layout = build_layout({"data": "24/09/2026", "itens": []}, self.config)
        self.assertEqual(empty_layout["total"], "0,00")
        self.assertFalse(any(field["key"] == "total_geral" for field in empty_layout["fields"]))
        with self.assertRaisesRegex(ValueError, "Preencha o cliente"):
            render_pdf({"data": "24/09/2026", "itens": []}, self.config)

    def test_pdf_e_previa_usam_mesmas_posicoes(self) -> None:
        try:
            import pymupdf as fitz
        except ImportError:
            try:
                import fitz
            except ImportError:
                self.skipTest("PyMuPDF não está instalado neste interpretador")

        plain = fitz.open(stream=render_pdf(self.payload, self.config), filetype="pdf")
        model = fitz.open(
            stream=render_pdf(self.payload, self.config, with_model=True), filetype="pdf"
        )
        try:
            plain_page, model_page = plain[0], model[0]
            expected_width = self.config["papel_largura_mm"] * 72 / 25.4
            expected_height = self.config["papel_altura_mm"] * 72 / 25.4
            self.assertAlmostEqual(plain_page.rect.width, expected_width, places=2)
            self.assertAlmostEqual(plain_page.rect.height, expected_height, places=2)
            self.assertEqual(plain_page.get_text("words"), model_page.get_text("words"))
            self.assertEqual(len(plain_page.get_images()), 0)
            self.assertGreaterEqual(len(model_page.get_images()), 1)
            preview = render_preview_png(self.payload, self.config)
            self.assertTrue(preview.startswith(b"\x89PNG\r\n\x1a\n"))
        finally:
            plain.close()
            model.close()


if __name__ == "__main__":
    unittest.main()
