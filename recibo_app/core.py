"""Núcleo do recibo para impressão sobre o formulário pré-impresso.

Todas as coordenadas são medidas em milímetros desde o canto superior esquerdo
do papel. A prévia é rasterizada a partir do próprio PDF com o modelo ao fundo;
o PDF entregue para impressão usa o mesmo desenho, sem o fundo.
"""

from __future__ import annotations

import copy
import io
import json
import math
import re
from datetime import datetime
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path
from typing import Any


BASE_DIR = Path(__file__).resolve().parent
CONFIG_PATH = BASE_DIR / "posicionamento.json"
MODEL_PATH = BASE_DIR / "modelo_preview.png"
ITEM_KEYS = ("quantidade", "unidade", "descricao", "preco_unitario", "total_item")
FIELD_KEYS = ("cliente", "data", *ITEM_KEYS, "total_geral")
MAX_LINES = 10


def _reference_config() -> dict:
    with CONFIG_PATH.open("r", encoding="utf-8") as source:
        config = json.load(source)
    if not isinstance(config, dict):
        raise ValueError("O arquivo de posicionamento deve conter um objeto JSON.")
    return config


def _merge(base: dict, changes: dict) -> dict:
    for key, value in changes.items():
        if isinstance(value, dict) and isinstance(base.get(key), dict):
            _merge(base[key], value)
        else:
            base[key] = copy.deepcopy(value)
    return base


def _number(value: Any, label: str, low: float = -500.0, high: float = 500.0) -> float:
    if isinstance(value, bool) or value is None:
        raise ValueError(f"{label}: informe um número válido.")
    try:
        number = float(value.replace(",", ".") if isinstance(value, str) else value)
    except (TypeError, ValueError, OverflowError) as exc:
        raise ValueError(f"{label}: informe um número válido.") from exc
    if not math.isfinite(number) or not low <= number <= high:
        raise ValueError(f"{label}: valor fora do intervalo permitido.")
    return number


def _pair(value: Any, label: str) -> dict[str, float]:
    if not isinstance(value, dict):
        raise ValueError(f"{label}: informe os ajustes X e Y.")
    return {
        "x": _number(value.get("x"), f"{label} X"),
        "y": _number(value.get("y"), f"{label} Y"),
    }


def normalize_config(config: dict) -> dict:
    """Devolve cópia validada do perfil v4, preservando os offsets existentes.

    As listas de ajustes são completadas ou limitadas a dez linhas. Os offsets
    ``impressao_direta_*`` podem permanecer no JSON legado, mas não participam
    de qualquer cálculo de posição neste módulo.
    """
    if not isinstance(config, dict):
        raise ValueError("A configuração deve ser um objeto.")
    defaults = _reference_config()
    merged = _merge(copy.deepcopy(defaults), config)
    merged["config_version"] = 4

    merged["papel_largura_mm"] = _number(
        merged.get("papel_largura_mm"), "Largura do papel", 50.0, 500.0
    )
    merged["papel_altura_mm"] = _number(
        merged.get("papel_altura_mm"), "Altura do papel", 50.0, 500.0
    )
    merged["offset_x_mm"] = _number(merged.get("offset_x_mm"), "Ajuste geral X")
    merged["offset_y_mm"] = _number(merged.get("offset_y_mm"), "Ajuste geral Y")
    merged["fonte_tamanho_pt"] = _number(
        merged.get("fonte_tamanho_pt"), "Tamanho da fonte", 4.0, 24.0
    )
    if not isinstance(merged.get("fonte"), str) or not merged["fonte"].strip():
        raise ValueError("O nome da fonte é inválido.")

    for key, coordinate in (
        ("cliente", "x_mm"),
        ("data", "x_centro_mm"),
        ("quantidade", "x_direita_mm"),
        ("unidade", "x_centro_mm"),
        ("descricao", "x_mm"),
        ("preco_unitario", "x_direita_mm"),
        ("total_item", "x_direita_mm"),
        ("total_geral", "x_direita_mm"),
    ):
        entry = merged.get(key)
        if not isinstance(entry, dict):
            raise ValueError(f"Posição de {key} inválida.")
        entry[coordinate] = _number(entry.get(coordinate), f"{key} X")
    for key in ("cliente", "data", "total_geral"):
        merged[key]["y_mm"] = _number(merged[key].get("y_mm"), f"{key} Y")

    def ten_values(value: Any, default: list, label: str) -> list:
        if not isinstance(value, list):
            raise ValueError(f"{label} deve ser uma lista.")
        return [copy.deepcopy(value[i] if i < len(value) else default[i]) for i in range(MAX_LINES)]

    merged["linhas_y_mm"] = [
        _number(value, f"Linha {i + 1} Y")
        for i, value in enumerate(
            ten_values(merged.get("linhas_y_mm"), defaults["linhas_y_mm"], "Linhas")
        )
    ]
    offsets = merged.get("campo_offsets_mm")
    if not isinstance(offsets, dict):
        raise ValueError("Os ajustes por campo são inválidos.")
    merged["campo_offsets_mm"] = {
        key: _pair(offsets.get(key, defaults["campo_offsets_mm"][key]), f"Campo {key}")
        for key in FIELD_KEYS
    }
    merged["linha_offsets_mm"] = [
        _pair(value, f"Linha {i + 1}")
        for i, value in enumerate(
            ten_values(
                merged.get("linha_offsets_mm"), defaults["linha_offsets_mm"], "Ajustes por linha"
            )
        )
    ]
    cells = ten_values(
        merged.get("celula_offsets_mm"), defaults["celula_offsets_mm"], "Ajustes por célula"
    )
    normalized_cells: list[dict] = []
    for i, row in enumerate(cells):
        if not isinstance(row, dict):
            raise ValueError(f"Ajustes da linha {i + 1} inválidos.")
        normalized_cells.append(
            {
                key: _pair(
                    row.get(key, defaults["celula_offsets_mm"][i][key]),
                    f"Linha {i + 1}, {key}",
                )
                for key in ITEM_KEYS
            }
        )
    merged["celula_offsets_mm"] = normalized_cells

    # Mantidos apenas para leitura de perfis antigos; impressão via PDF os ignora.
    for legacy_key in ("impressao_direta_offset_x_mm", "impressao_direta_offset_y_mm"):
        if legacy_key in merged:
            merged[legacy_key] = _number(merged[legacy_key], legacy_key)
    return merged


def load_default_config() -> dict:
    """Carrega o perfil físico v4 distribuído com o modelo, sem alterá-lo."""
    return normalize_config(_reference_config())


def _decimal(value: Any, label: str) -> Decimal:
    if value is None or value == "":
        return Decimal("0")
    if isinstance(value, bool):
        raise ValueError(f"{label}: informe um número válido.")
    if isinstance(value, Decimal):
        result = value
    elif isinstance(value, (int, float)):
        try:
            result = Decimal(str(value))
        except InvalidOperation as exc:
            raise ValueError(f"{label}: informe um número válido.") from exc
    elif isinstance(value, str):
        text = value.strip().replace("\u00a0", "").replace(" ", "")
        if text.startswith("R$"):
            text = text[2:]
        if not text:
            return Decimal("0")
        if "," in text:
            valid = re.fullmatch(r"\+?(?:\d{1,3}(?:\.\d{3})+|\d+)(?:,\d+)?", text)
            text = text.replace(".", "").replace(",", ".")
        else:
            valid = re.fullmatch(r"\+?\d+(?:\.\d+)?", text)
        if not valid:
            raise ValueError(f"{label}: informe um número válido.")
        try:
            result = Decimal(text)
        except InvalidOperation as exc:
            raise ValueError(f"{label}: informe um número válido.") from exc
    else:
        raise ValueError(f"{label}: informe um número válido.")
    if not result.is_finite() or result < 0:
        raise ValueError(f"{label}: informe um número não negativo.")
    if result > Decimal("999999999"):
        raise ValueError(f"{label}: valor acima do limite permitido.")
    return result


def _money(value: Decimal) -> str:
    rounded = value.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    grouped = f"{rounded:,.2f}"
    return grouped.replace(",", "_").replace(".", ",").replace("_", ".")


def _quantity(value: Decimal) -> str:
    if value == value.to_integral_value():
        return str(int(value))
    return format(value.normalize(), "f").replace(".", ",")


def _text(value: Any, label: str, max_length: int) -> str:
    if value is None:
        return ""
    if not isinstance(value, str):
        raise ValueError(f"{label}: informe um texto válido.")
    text = value.strip()
    if len(text) > max_length:
        raise ValueError(f"{label}: limite de {max_length} caracteres excedido.")
    if any(ord(char) < 32 for char in text):
        raise ValueError(f"{label}: use apenas uma linha de texto.")
    return text


def _font_name() -> str:
    """Registra uma fonte embutida e igual em qualquer servidor ou navegador."""
    from reportlab import pdfbase
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont

    name = "ReciboVera"
    if name not in pdfmetrics.getRegisteredFontNames():
        font_path = Path(pdfbase.__file__).resolve().parent / ".." / "fonts" / "Vera.ttf"
        pdfmetrics.registerFont(TTFont(name, str(font_path.resolve())))
    return name


def _xy(config: dict, key: str, x: float, y: float, line_index: int | None = None) -> tuple[float, float]:
    x += config["offset_x_mm"] + config["campo_offsets_mm"][key]["x"]
    y += config["offset_y_mm"] + config["campo_offsets_mm"][key]["y"]
    if line_index is not None:
        x += config["linha_offsets_mm"][line_index]["x"]
        y += config["linha_offsets_mm"][line_index]["y"]
        x += config["celula_offsets_mm"][line_index][key]["x"]
        y += config["celula_offsets_mm"][line_index][key]["y"]
    return x, y


def build_layout(payload: dict, config: dict) -> dict:
    """Calcula os textos e as posições finais do recibo.

    ``line`` é 1 a 10 nos itens e ``None`` em cliente, data e total geral.
    ``total`` usa formato monetário brasileiro sem prefixo R$.
    """
    from reportlab.lib.units import mm
    from reportlab.pdfbase import pdfmetrics

    if not isinstance(payload, dict):
        raise ValueError("Os dados do recibo devem ser um objeto.")
    cfg = normalize_config(config)
    items = payload.get("itens", [])
    if not isinstance(items, list) or len(items) > MAX_LINES:
        raise ValueError("O recibo aceita até 10 itens.")
    font_size = cfg["fonte_tamanho_pt"]
    font_name = _font_name()
    fields: list[dict] = []
    warnings: list[str] = []
    print_blockers: list[str] = []
    bounds: list[tuple[dict, float, float, float, float]] = []

    def add(key: str, text: str, x: float, y: float, align: str, line: int | None = None) -> None:
        field = {
            "id": f"{key}-{line}" if line is not None else key,
            "key": key,
            "line": line,
            "text": text,
            "x_mm": round(x, 3),
            "y_mm": round(y, 3),
            "align": align,
            "font_size": font_size,
        }
        fields.append(field)
        width_mm = pdfmetrics.stringWidth(text, font_name, font_size) / mm
        left = x if align == "left" else x - width_mm if align == "right" else x - width_mm / 2
        right = left + width_mm
        ascent, descent = pdfmetrics.getAscentDescent(font_name, font_size)
        top = y - ascent / mm
        bottom = y - descent / mm
        bounds.append((field, left, right, top, bottom))
        if left < 0 or right > cfg["papel_largura_mm"] or top < 0 or bottom > cfg["papel_altura_mm"]:
            message = f"{field['id']}: texto ultrapassa a área do papel."
            warnings.append(message)
            print_blockers.append(message)

    cliente = _text(payload.get("cliente"), "Cliente", 120)
    has_content = bool(cliente)
    if cliente:
        x, y = _xy(cfg, "cliente", cfg["cliente"]["x_mm"], cfg["cliente"]["y_mm"])
        add("cliente", cliente, x, y, "left")
    data = _text(payload.get("data"), "Data", 32)
    if data:
        if not re.fullmatch(r"\d{2}/\d{2}/\d{4}", data):
            raise ValueError("Data: use o formato DD/MM/AAAA.")
        try:
            datetime.strptime(data, "%d/%m/%Y")
        except ValueError as exc:
            raise ValueError("Data: informe uma data válida.") from exc
        x, y = _xy(cfg, "data", cfg["data"]["x_centro_mm"], cfg["data"]["y_mm"])
        add("data", data, x, y, "center")

    total = Decimal("0")
    for index, item in enumerate(items):
        if not isinstance(item, dict):
            raise ValueError(f"Item da linha {index + 1} inválido.")
        line = index + 1
        quantity = _decimal(item.get("quantidade"), f"Quantidade da linha {line}")
        price = _decimal(item.get("preco_unitario"), f"Preço unitário da linha {line}")
        unit = _text(item.get("unidade"), f"Unidade da linha {line}", 12)
        description = _text(item.get("descricao"), f"Descrição da linha {line}", 160)
        if quantity == 0 and price == 0 and not description:
            continue
        has_content = True
        if quantity == 0 and (price != 0 or description):
            warnings.append(f"Linha {line}: informe a quantidade para calcular o total do item.")
        if price == 0 and quantity != 0:
            warnings.append(f"Linha {line}: informe o preço unitário para calcular o total do item.")
        item_total = (quantity * price).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
        total += item_total
        base_y = cfg["linhas_y_mm"][index]
        for key, text, x_base, align in (
            ("quantidade", _quantity(quantity) if quantity != 0 else "", cfg["quantidade"]["x_direita_mm"], "right"),
            ("unidade", unit, cfg["unidade"]["x_centro_mm"], "center"),
            ("descricao", description, cfg["descricao"]["x_mm"], "left"),
            ("preco_unitario", _money(price) if price != 0 else "", cfg["preco_unitario"]["x_direita_mm"], "right"),
            ("total_item", _money(item_total) if price != 0 else "", cfg["total_item"]["x_direita_mm"], "right"),
        ):
            if text:
                x, y = _xy(cfg, key, x_base, base_y, index)
                add(key, text, x, y, align, line)
        if description:
            width_mm = pdfmetrics.stringWidth(description, font_name, font_size) / mm
            start_x = _xy(cfg, "descricao", cfg["descricao"]["x_mm"], base_y, index)[0]
            price_x = _xy(cfg, "preco_unitario", cfg["preco_unitario"]["x_direita_mm"], base_y, index)[0]
            if start_x + width_mm > price_x - 14:
                message = f"Linha {line}: descrição pode invadir a coluna de preço."
                warnings.append(message)
                print_blockers.append(message)

    if has_content:
        x, y = _xy(cfg, "total_geral", cfg["total_geral"]["x_direita_mm"], cfg["total_geral"]["y_mm"])
        add("total_geral", _money(total), x, y, "right")
    for index, first in enumerate(bounds):
        first_field, first_left, first_right, first_top, first_bottom = first
        if first_field["line"] is None:
            continue
        for second in bounds[index + 1 :]:
            second_field, second_left, second_right, second_top, second_bottom = second
            if first_field["line"] != second_field["line"]:
                continue
            horizontal_overlap = min(first_right, second_right) - max(first_left, second_left)
            vertical_overlap = min(first_bottom, second_bottom) - max(first_top, second_top)
            if horizontal_overlap > 0.1 and vertical_overlap > 0.1:
                message = (
                    f"Linha {first_field['line']}: "
                    f"{first_field['key']} e {second_field['key']} estão sobrepostos."
                )
                warnings.append(message)
                print_blockers.append(message)
    return {
        "fields": fields,
        "total": _money(total),
        "warnings": warnings,
        "print_blockers": print_blockers,
        "paper_width_mm": cfg["papel_largura_mm"],
        "paper_height_mm": cfg["papel_altura_mm"],
    }


def render_pdf(payload: dict, config: dict, with_model: bool = False) -> bytes:
    """Gera PDF de tamanho real. Use ``with_model=False`` no papel pré-impresso."""
    from reportlab.lib.units import mm
    from reportlab.pdfgen import canvas

    cfg = normalize_config(config)
    layout = build_layout(payload, cfg)
    if not with_model and not any(
        field["key"] == "cliente" or field["line"] is not None for field in layout["fields"]
    ):
        raise ValueError("Preencha o cliente ou ao menos um item antes de imprimir.")
    if not with_model and layout["print_blockers"]:
        raise ValueError("Ajuste o recibo antes de imprimir: " + " ".join(layout["print_blockers"]))
    width = cfg["papel_largura_mm"] * mm
    height = cfg["papel_altura_mm"] * mm
    output = io.BytesIO()
    pdf = canvas.Canvas(output, pagesize=(width, height), pageCompression=1)
    pdf.setTitle("Recibo para formulário pré-impresso")
    if with_model:
        pdf.drawImage(str(MODEL_PATH), 0, 0, width=width, height=height)
    pdf.setFillColorRGB(0, 0, 0)
    pdf.setFont(_font_name(), cfg["fonte_tamanho_pt"])
    for field in layout["fields"]:
        x = field["x_mm"] * mm
        y = height - field["y_mm"] * mm
        if field["align"] == "right":
            pdf.drawRightString(x, y, field["text"])
        elif field["align"] == "center":
            pdf.drawCentredString(x, y, field["text"])
        else:
            pdf.drawString(x, y, field["text"])
    pdf.showPage()
    pdf.save()
    return output.getvalue()


def render_preview_png(payload: dict, config: dict) -> bytes:
    """Rasteriza o PDF com modelo; coordenadas e fonte são as da impressão."""
    try:
        import pymupdf as fitz
    except ImportError:
        import fitz

    document = fitz.open(stream=render_pdf(payload, config, with_model=True), filetype="pdf")
    try:
        page = document[0]
        pixels = page.get_pixmap(matrix=fitz.Matrix(3, 3), colorspace=fitz.csRGB, alpha=False)
        return pixels.tobytes("png")
    finally:
        document.close()
