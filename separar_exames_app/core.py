# -*- coding: utf-8 -*-
from __future__ import annotations

import csv
import hashlib
import json
import math
import os
import re
import shutil
import unicodedata
import uuid
from collections import Counter, defaultdict
from dataclasses import asdict, dataclass, field
from datetime import datetime
from difflib import SequenceMatcher
from pathlib import Path
from typing import Callable, Iterable, Optional
from concurrent.futures import ThreadPoolExecutor, as_completed

try:
    import pymupdf as fitz
except ImportError:  # pragma: no cover
    import fitz  # type: ignore

from PIL import Image, ImageOps

try:
    import pytesseract
except Exception:  # pragma: no cover
    pytesseract = None

try:
    from openpyxl import Workbook, load_workbook
except Exception:  # pragma: no cover
    Workbook = None
    load_workbook = None

APP_DIR = Path(__file__).resolve().parent
DATA_DIR = APP_DIR / "dados"
MODELS_DIR = DATA_DIR / "modelos"
MODELS_JSON = DATA_DIR / "modelos.json"
TESSDATA_DIR = DATA_DIR / "tessdata"
CONFIG_JSON = DATA_DIR / "config.json"

DEFAULT_EXAM_TYPES = ("ASO", "AUDIOMETRIA", "ESPIROMETRIA", "ACUIDADE VISUAL", "LAUDO PCD")
EXAM_TYPES = DEFAULT_EXAM_TYPES
OUTPUT_DIRNAME = "ARQUIVOS SEPARADOS"
REPORT_FILENAME = "RELATORIO_CONFERENCIA.xlsx"
EXAM_ALIASES = {
    "ASO": "ASO",
    "ATESTADO DE SAUDE OCUPACIONAL": "ASO",
    "ATESTADO SAUDE OCUPACIONAL": "ASO",
    "AUDIOMETRIA": "AUDIOMETRIA",
    "AVALIACAO AUDIOLOGICA": "AUDIOMETRIA",
    "EXAME AUDIOMETRICO": "AUDIOMETRIA",
    "ESPIROMETRIA": "ESPIROMETRIA",
    "ESPIROMETRICO": "ESPIROMETRIA",
    "ACUIDADE": "ACUIDADE VISUAL",
    "ACUIDADE VISUAL": "ACUIDADE VISUAL",
    "LAUDO PCD": "LAUDO PCD",
    "PCD": "LAUDO PCD",
    "PESSOA COM DEFICIENCIA": "LAUDO PCD",
}

INVALID_WINDOWS_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
SPACE_RE = re.compile(r"[ \t\u00a0]+")
CNPJ_RE = re.compile(r"(?<!\d)(\d{2})\s*[.\-/]?\s*(\d{3})\s*[.\-/]?\s*(\d{3})\s*[/\-.]?\s*(\d{4})\s*[-.]?\s*(\d{2})(?!\d)")
CPF_RE = re.compile(r"(?<!\d)(\d{3})\s*[.\-]?\s*(\d{3})\s*[.\-]?\s*(\d{3})\s*[-.]?\s*(\d{2})(?!\d)")
DATE_RE = re.compile(r"\b([0-3]?\d)[/\-.]([01]?\d)[/\-.]((?:19|20)?\d{2})\b")

# Palavras muito comuns não ajudam a reconhecer o layout/modelo.
MODEL_STOPWORDS = {
    "para", "pela", "pelo", "pelos", "pelas", "com", "sem", "uma", "umas", "uns", "dos", "das", "do", "da", "de",
    "que", "este", "esta", "esse", "essa", "seu", "sua", "seus", "suas", "nao", "sim", "como", "mais", "menos", "entre",
    "trabalho", "trabalhador", "empresa", "funcionario", "colaborador", "nome", "data", "exame", "medico", "medicina", "saude",
    "ocupacional", "brasil", "assinatura", "observacoes", "observacao", "idade", "sexo", "cargo", "setor", "funcao", "cnpj", "cpf",
}

# Regras padrão. Os modelos cadastrados pelo usuário entram como uma segunda camada.
SIGNATURES: dict[str, list[tuple[str, float]]] = {
    "ASO": [
        ("atestado de saude ocupacional", 42), ("a.s.o", 18), ("dados da empresa colaborador", 15),
        ("riscos ocupacionais", 10), ("exames complementares", 8), ("tipo de exame", 7),
        ("apto p funcao", 8), ("medico examinador", 5),
    ],
    "AUDIOMETRIA": [
        # Formatos mais comuns encontrados nos laudos da EDGE e de clínicas parceiras.
        # "Avaliação Audiométrica" precisa ter peso alto: nas versões anteriores esse
        # título era reconhecido na estrutura da página, mas não pontuava o suficiente
        # na classificação final e o exame acabava marcado como NÃO ENCONTRADO.
        ("avaliacao audiometrica", 58), ("avaliacao audiologica", 42), ("audiometria tonal", 38),
        ("audiometro", 22), ("meatoscopia", 18), ("limiares auditivos", 18),
        ("orelha direita", 11), ("orelha esquerda", 11),
        ("ouvido direito", 12), ("ouvido esquerdo", 12),
        ("parecer audiologico", 12), ("repouso auditivo", 12), ("via aerea", 6),
    ],
    "ESPIROMETRIA": [
        ("espirometria", 45), ("curva fluxo volume", 25), ("vef1", 20), ("cvf", 18),
        ("fef", 8), ("disturbio ventilatorio", 18), ("prova broncodilatadora", 12), ("funcao pulmonar", 10),
    ],
    "ACUIDADE VISUAL": [
        ("acuidade visual", 48), ("avaliacao da acuidade", 25), ("visao de perto", 18), ("visao de longe", 18),
        ("olho direito", 10), ("olho esquerdo", 10), ("tabela de snellen", 15), ("snellen", 10),
    ],
    "LAUDO PCD": [
        ("laudo pcd", 48), ("laudo caracterizador", 42), ("pessoa com deficiencia", 32),
        ("caracterizacao da deficiencia", 28), ("tipo de deficiencia", 18), ("cid", 7),
        ("deficiencia fisica", 10), ("deficiencia auditiva", 10), ("deficiencia visual", 10),
    ],
}

IGNORE_SIGNATURES: list[tuple[str, float]] = [
    ("formulario de encaminhamento de exames", 100),
    ("responda as perguntas abaixo", 95),
    ("comprovante de transferencia", 100),
    ("comprovante de transacao", 100),
    ("nota de balcao", 100),
    ("solicitacao de exame medico", 90),
    ("solicitacao de exame", 55),
    ("encaminhamento", 30),
    ("dados do pagador", 80),
    ("dados do recebedor", 80),
]

@dataclass
class ExpectedEmployee:
    row_id: int
    name: str
    cpf: str = ""
    company: str = ""
    cnpj: str = ""
    expected_exams: set[str] = field(default_factory=set)
    # Mantém a quantidade de ocorrências do mesmo tipo de exame na planilha.
    # Ex.: duas linhas ASO para o mesmo colaborador significam DOIS ASOs esperados
    # (p.ex. PERIÓDICO + MUDANÇA DE RISCOS), não uma duplicidade.
    expected_exam_counts: dict[str, int] = field(default_factory=dict)
    # Recibo associado a cada ocorrência esperada do exame, na mesma ordem da planilha.
    # O valor pode ser um número/código ou a expressão "A PRAZO".
    expected_exam_receipts: dict[str, list[str]] = field(default_factory=dict)
    source_sheet: str = ""

    @property
    def key(self) -> str:
        cpf_d = digits_only(self.cpf)
        company_key = digits_only(self.cnpj) or normalize_for_match(self.company)
        if is_valid_cpf(cpf_d):
            return cpf_d + (("|" + company_key) if company_key else "")
        return normalize_for_match(self.name) + (("|" + company_key) if company_key else "")

    def add_expected_exam(self, exam: str, count: int = 1, receipt: str = "") -> None:
        exam = normalize_exam(exam) or clean_value(exam).upper()
        if not exam or count <= 0:
            return
        self.expected_exams.add(exam)
        self.expected_exam_counts[exam] = self.expected_exam_counts.get(exam, 0) + int(count)
        receipt_value = clean_value(receipt)
        if normalize_for_match(receipt_value) in {"A PRAZO", "PRAZO"}:
            receipt_value = "A PRAZO"
        bucket = self.expected_exam_receipts.setdefault(exam, [])
        bucket.extend([receipt_value] * int(count))

    def expected_receipt(self, exam: str, occurrence_index: int = 0) -> str:
        values = self.expected_exam_receipts.get(exam) or []
        if 0 <= int(occurrence_index) < len(values):
            return clean_value(values[int(occurrence_index)])
        return ""

    def expected_count(self, exam: str) -> int:
        # Compatibilidade com objetos antigos/testes que preencham apenas expected_exams.
        if exam in self.expected_exam_counts:
            return max(0, int(self.expected_exam_counts.get(exam, 0)))
        return 1 if exam in self.expected_exams else 0

    @property
    def expected_total(self) -> int:
        if self.expected_exam_counts:
            return sum(max(0, int(v)) for v in self.expected_exam_counts.values())
        return len(self.expected_exams)

@dataclass
class ModelSample:
    id: str
    exam_type: str
    label: str
    source_filename: str
    page_number: int
    created_at: str
    token_signature: list[str]
    visual_hash: str
    aspect_ratio: float
    stored_file: str = ""

@dataclass
class PageAnalysis:
    id: str
    source_pdf: str
    source_pdf_path: str
    page_index: int
    page_number: int
    raw_text: str
    text_source: str
    ocr_warning: str
    exam_type: str
    exam_confidence: float
    exam_score_breakdown: str
    employee_name: str
    employee_cpf: str
    company: str
    cnpj: str
    employee_match_confidence: float
    employee_row_id: int | None
    expected: bool
    status: str
    reason: str
    output_file: str = ""
    page_hash: str = ""
    exam_subtype: str = ""
    receipt: str = ""

@dataclass
class ProcessingSummary:
    root_dir: str
    total_pages: int
    auto_saved: int
    pending: int
    ignored: int
    unexpected: int
    duplicates: int
    missing_expected: int
    analyses: list[PageAnalysis]
    missing_rows: list[dict]
    quick_scanned: int = 0
    detailed_ocr_pages: int = 0
    skipped_after_complete: int = 0
    fast_rejected: int = 0
    rescued_pages: int = 0
    audit_gaps: int = 0


def strip_accents(value: str) -> str:
    return "".join(ch for ch in unicodedata.normalize("NFD", str(value or "")) if unicodedata.category(ch) != "Mn")


def normalize_for_match(value: str) -> str:
    value = strip_accents(value).upper()
    value = re.sub(r"[^A-Z0-9]+", " ", value)
    return SPACE_RE.sub(" ", value).strip()



def _exam_types_json() -> Path:
    return DATA_DIR / "tipos_exames.json"


def _exam_aliases_json() -> Path:
    return DATA_DIR / "apelidos_exames.json"


def _clean_exam_type_name(value: str) -> str:
    value = clean_value(value).upper() if 'clean_value' in globals() else str(value or "").upper().strip()
    value = INVALID_WINDOWS_CHARS.sub(" ", value)
    value = SPACE_RE.sub(" ", value).strip(" .")
    if len(value) > 80:
        value = value[:80].rstrip()
    return value


def refresh_exam_types() -> tuple[str, ...]:
    """Atualiza a lista de tipos de exames com os tipos cadastrados pelo usuário."""
    global EXAM_TYPES
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    custom: list[str] = []
    try:
        raw = json.loads(_exam_types_json().read_text(encoding="utf-8")) if _exam_types_json().exists() else []
        if isinstance(raw, dict):
            raw = raw.get("types") or []
        for item in raw:
            name = _clean_exam_type_name(item.get("name") if isinstance(item, dict) else item)
            if name and name not in DEFAULT_EXAM_TYPES and name not in custom:
                custom.append(name)
    except Exception:
        custom = []
    EXAM_TYPES = tuple(dict.fromkeys([*DEFAULT_EXAM_TYPES, *custom]))

    # Adiciona aliases automáticos para todos os tipos, inclusive os cadastrados.
    # Importante: tipos criados pelo usuário precisam pontuar alto mesmo quando
    # ainda não têm modelo visual cadastrado. Antes eles recebiam peso 18 e,
    # como o piso padrão é 28, páginas com o nome exato do exame podiam cair em
    # "NÃO ENCONTRADO".
    aliases_by_exam: dict[str, list[str]] = defaultdict(list)
    for name in EXAM_TYPES:
        name_norm = normalize_for_match(name)
        EXAM_ALIASES.setdefault(name_norm, name)
        aliases_by_exam[name].append(name)
    try:
        aliases = json.loads(_exam_aliases_json().read_text(encoding="utf-8")) if _exam_aliases_json().exists() else {}
        if isinstance(aliases, dict):
            for canonical, values in aliases.items():
                canonical_name = _clean_exam_type_name(canonical)
                if canonical_name not in EXAM_TYPES:
                    continue
                for alias in values or []:
                    alias_n = normalize_for_match(alias)
                    if alias_n:
                        EXAM_ALIASES[alias_n] = canonical_name
                        aliases_by_exam[canonical_name].append(str(alias))
    except Exception:
        pass

    for name in EXAM_TYPES:
        if name in DEFAULT_EXAM_TYPES:
            # Mantém as regras manuais dos tipos originais, apenas acrescentando
            # o próprio nome como assinatura forte quando não existir.
            SIGNATURES.setdefault(name, [])
            if not any(normalize_for_match(k) == normalize_for_match(name) for k, _ in SIGNATURES[name]):
                SIGNATURES[name].append((name.lower(), 45.0))
            continue
        sigs: list[tuple[str, float]] = []
        seen_sig: set[str] = set()
        for alias in aliases_by_exam.get(name, [name]):
            alias_clean = clean_value(alias) if 'clean_value' in globals() else str(alias or '').strip()
            alias_norm = normalize_for_match(alias_clean)
            if not alias_norm or alias_norm in seen_sig:
                continue
            seen_sig.add(alias_norm)
            # Peso 48 faz o tipo ser reconhecido pelo nome/alias exato.
            sigs.append((alias_clean.lower(), 48.0))
            # Tokens relevantes do tipo ajudam quando o OCR quebra o título.
            tokens = [t for t in alias_norm.split() if len(t) >= 4 and t not in MODEL_STOPWORDS]
            if len(tokens) >= 2:
                sigs.append((" ".join(tokens[:3]).lower(), 24.0))
            elif tokens:
                sigs.append((tokens[0].lower(), 16.0))
        SIGNATURES[name] = sigs or [(name.lower(), 48.0)]
    return EXAM_TYPES


def get_exam_types() -> tuple[str, ...]:
    return refresh_exam_types()


def get_custom_exam_types() -> list[dict[str, object]]:
    refresh_exam_types()
    try:
        model_counts = Counter(m.exam_type for m in load_models())
    except Exception:
        model_counts = Counter()
    return [
        {"name": name, "default": name in DEFAULT_EXAM_TYPES, "model_count": int(model_counts.get(name, 0))}
        for name in EXAM_TYPES
    ]


def add_custom_exam_type(name: str, aliases: list[str] | None = None) -> str:
    name = _clean_exam_type_name(name)
    if not name:
        raise ValueError("Informe o nome do tipo de exame.")
    if len(name) < 2:
        raise ValueError("O nome do tipo de exame está muito curto.")
    refresh_exam_types()
    if name not in EXAM_TYPES:
        items: list[str] = []
        try:
            raw = json.loads(_exam_types_json().read_text(encoding="utf-8")) if _exam_types_json().exists() else []
            if isinstance(raw, dict):
                raw = raw.get("types") or []
            for item in raw:
                item_name = _clean_exam_type_name(item.get("name") if isinstance(item, dict) else item)
                if item_name and item_name not in items and item_name not in DEFAULT_EXAM_TYPES:
                    items.append(item_name)
        except Exception:
            items = []
        items.append(name)
        _exam_types_json().write_text(json.dumps(items, ensure_ascii=False, indent=2), encoding="utf-8")
    if aliases:
        try:
            alias_data = json.loads(_exam_aliases_json().read_text(encoding="utf-8")) if _exam_aliases_json().exists() else {}
            if not isinstance(alias_data, dict):
                alias_data = {}
        except Exception:
            alias_data = {}
        current = [str(x).strip() for x in alias_data.get(name, []) if str(x).strip()]
        for alias in aliases:
            alias = str(alias or "").strip()
            if alias and alias not in current:
                current.append(alias)
        alias_data[name] = current
        _exam_aliases_json().write_text(json.dumps(alias_data, ensure_ascii=False, indent=2), encoding="utf-8")
    refresh_exam_types()
    return name


def delete_custom_exam_type(name: str) -> bool:
    name = _clean_exam_type_name(name)
    if not name or name in DEFAULT_EXAM_TYPES:
        return False
    refresh_exam_types()
    # Não apaga tipo que tem modelo cadastrado para evitar quebrar reconhecimento/arquivo.
    if any(m.exam_type == name for m in load_models()):
        raise ValueError("Este tipo possui modelos cadastrados. Exclua os modelos antes de remover o tipo.")
    try:
        raw = json.loads(_exam_types_json().read_text(encoding="utf-8")) if _exam_types_json().exists() else []
        if isinstance(raw, dict):
            raw = raw.get("types") or []
    except Exception:
        raw = []
    kept = []
    for item in raw:
        item_name = _clean_exam_type_name(item.get("name") if isinstance(item, dict) else item)
        if item_name and item_name != name and item_name not in DEFAULT_EXAM_TYPES and item_name not in kept:
            kept.append(item_name)
    _exam_types_json().write_text(json.dumps(kept, ensure_ascii=False, indent=2), encoding="utf-8")
    try:
        alias_data = json.loads(_exam_aliases_json().read_text(encoding="utf-8")) if _exam_aliases_json().exists() else {}
        if isinstance(alias_data, dict) and name in alias_data:
            alias_data.pop(name, None)
            _exam_aliases_json().write_text(json.dumps(alias_data, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass
    refresh_exam_types()
    return True


def clean_line(value: str) -> str:
    value = str(value or "").replace("\x00", " ").replace("\r", " ")
    return SPACE_RE.sub(" ", value).strip()


def clean_value(value: str) -> str:
    value = str(value or "").replace("\n", " ").replace("\r", " ")
    value = SPACE_RE.sub(" ", value).strip(" :;|-–—_.,")
    return value.strip()


def digits_only(value: str) -> str:
    return re.sub(r"\D", "", str(value or ""))


def is_valid_cpf(value: str) -> bool:
    d = digits_only(value)
    if len(d) != 11 or len(set(d)) == 1:
        return False
    nums = [int(x) for x in d]
    for size in (9, 10):
        total = sum(nums[i] * (size + 1 - i) for i in range(size))
        check = (total * 10) % 11
        if check == 10:
            check = 0
        if check != nums[size]:
            return False
    return True


def is_valid_cnpj(value: str) -> bool:
    d = digits_only(value)
    if len(d) != 14 or len(set(d)) == 1:
        return False
    nums = [int(x) for x in d]
    def calc(base: list[int], weights: list[int]) -> int:
        rem = sum(a * b for a, b in zip(base, weights)) % 11
        return 0 if rem < 2 else 11 - rem
    d1 = calc(nums[:12], [5,4,3,2,9,8,7,6,5,4,3,2])
    d2 = calc(nums[:12] + [d1], [6,5,4,3,2,9,8,7,6,5,4,3,2])
    return nums[12] == d1 and nums[13] == d2


def document_is_valid(value: str) -> bool:
    d = digits_only(value)
    return is_valid_cpf(d) if len(d) == 11 else (is_valid_cnpj(d) if len(d) == 14 else False)


def format_cnpj(value: str) -> str:
    d = digits_only(value)
    if len(d) != 14:
        return clean_value(value)
    return f"{d[:2]}.{d[2:5]}.{d[5:8]}/{d[8:12]}-{d[12:]}"


def format_cpf(value: str) -> str:
    d = digits_only(value)
    if len(d) != 11:
        return clean_value(value)
    return f"{d[:3]}.{d[3:6]}.{d[6:9]}-{d[9:]}"


def format_company_document(value: str) -> str:
    """Formata o documento da empresa, aceitando CNPJ (14) ou CPF (11)."""
    d = digits_only(value)
    if len(d) == 14:
        return format_cnpj(d)
    if len(d) == 11:
        return format_cpf(d)
    return clean_value(value)


def safe_component(value: str, fallback: str, max_length: int = 95) -> str:
    value = clean_value(value)
    value = INVALID_WINDOWS_CHARS.sub("_", value)
    value = SPACE_RE.sub(" ", value).strip().rstrip(".")
    if len(value) > max_length:
        value = value[:max_length].rstrip()
    return value or fallback


def normalize_exam(value: str) -> str:
    refresh_exam_types()
    raw = normalize_for_match(value)
    if raw in EXAM_ALIASES:
        return EXAM_ALIASES[raw]
    # Procura primeiro aliases maiores. Isso evita que um alias curto, como
    # "PCD", vença nomes mais específicos em células com vários exames.
    for alias, canonical in sorted(EXAM_ALIASES.items(), key=lambda kv: len(kv[0]), reverse=True):
        if alias and alias in raw:
            return canonical
    return ""


def parse_exam_cell(value: str) -> list[str]:
    """Extrai todos os tipos de exames de uma célula da planilha.

    A versão anterior dividia por barra (/), o que quebrava exames como
    "RAIO-X / TÓRAX". Agora primeiro procura todos os tipos/aliases cadastrados
    dentro da célula inteira e só depois tenta separar por delimitadores.
    """
    refresh_exam_types()
    raw_text = clean_value(value)
    raw = normalize_for_match(raw_text)
    if not raw:
        return []
    found: list[str] = []
    for alias, canonical in sorted(EXAM_ALIASES.items(), key=lambda kv: len(kv[0]), reverse=True):
        if alias and alias in raw and canonical not in found:
            found.append(canonical)
    if found:
        return found
    pieces = re.split(r"[;,|\n]+", raw_text)
    # Só divide por barra quando a célula parece conter exames distintos, não
    # quando a barra faz parte do nome do exame.
    if "/" in raw_text and not re.search(r"RAIO\s*-?\s*X\s*/", raw_text, re.I):
        pieces.extend(raw_text.split("/"))
    for piece in pieces:
        ex = normalize_exam(piece)
        if ex and ex not in found:
            found.append(ex)
    return found


ASO_SUBTYPE_PATTERNS: tuple[tuple[str, tuple[str, ...]], ...] = (
    ("MUDANÇA DE RISCOS OCUPACIONAIS", ("MUDANCA DE RISCOS OCUPACIONAIS", "MUDANCA DE RISCO OCUPACIONAL", "MUDANCA DE RISCOS", "MUDANCA DE FUNCAO")),
    ("RETORNO AO TRABALHO", ("RETORNO AO TRABALHO", "RETORNO TRABALHO")),
    ("DEMISSIONAL", ("DEMISSIONAL",)),
    ("PERIÓDICO", ("PERIODICO", "PENODICO", "PERIDICO")),
    ("ADMISSIONAL", ("ADMISSIONAL",)),
)

def detect_aso_subtype(text: str) -> str:
    """Identifica o motivo do ASO sem alterar a classificação principal (ASO)."""
    n = normalize_for_match(text)
    if not n:
        return ""
    # Dá preferência ao texto próximo do rótulo "Tipo de Exame" quando existir.
    local = n
    m = re.search(r"TIPO DE EXAME\s*[: -]?\s*(.{0,90})", n)
    if m:
        local = m.group(1) + " " + n
    for canonical, aliases in ASO_SUBTYPE_PATTERNS:
        if any(alias in local for alias in aliases):
            return canonical
    return ""


def find_cnpj(text: str) -> str:
    m = CNPJ_RE.search(str(text or ""))
    if not m:
        return ""
    return format_cnpj("".join(m.groups()))


def find_cpf(text: str) -> str:
    # Prioriza CPF identificado por rótulo para não confundir códigos avulsos.
    norm = str(text or "")
    label = re.search(r"\bCPF\s*[:#-]?\s*([\d.\-\s]{11,20})", norm, re.I)
    if label:
        d = digits_only(label.group(1))[:11]
        if len(d) == 11 and is_valid_cpf(d):
            return format_cpf(d)
    # Sem o rótulo CPF, é mais seguro não adivinhar: CNPJ/CAEPF e códigos de ficha
    # podem ter 11 dígitos e causar uma associação errada.
    return ""


def first_date(text: str) -> str:
    m = DATE_RE.search(text or "")
    if not m:
        return ""
    dd, mm, yy = m.groups()
    if len(yy) == 2:
        yy = "20" + yy
    return f"{int(dd):02d}/{int(mm):02d}/{yy}"


def _find_labeled_value(text: str, labels: Iterable[str], stop_labels: Iterable[str], max_lines: int = 2) -> str:
    lines = [clean_line(x) for x in str(text or "").splitlines() if clean_line(x)]
    labels_n = [normalize_for_match(x) for x in labels]
    stops_n = [normalize_for_match(x) for x in stop_labels]
    for i, line in enumerate(lines):
        nline = normalize_for_match(line)
        for lab in labels_n:
            # posição aproximada no texto normalizado; depois usa regex no original com tolerância.
            if not re.search(rf"(^|\s){re.escape(lab)}(\s|$)", nline):
                continue
            # tenta localizar o rótulo no original sem depender de acentos.
            # em geral o valor vem depois de ':' ou '-'.
            for sep in (":", "-", "–", "—"):
                pos = line.find(sep)
                if pos >= 0 and normalize_for_match(line[:pos]).endswith(lab):
                    val = clean_value(line[pos + 1:])
                    val_n = normalize_for_match(val)
                    cut = len(val)
                    for stop in stops_n:
                        mm = re.search(rf"\s{re.escape(stop)}\s*[:#-]", val_n)
                        if mm:
                            cut = min(cut, mm.start())
                    val = clean_value(val[:cut])
                    if val:
                        return val
            # também aceita "ROTULO VALOR" sem dois pontos em OCR.
            if nline.startswith(lab + " "):
                approx = line[len(line) - max(0, len(nline) - len(lab)) :]
                approx = clean_value(approx)
                if approx and normalize_for_match(approx) != lab:
                    return approx
            # valor pode estar na linha seguinte.
            collected: list[str] = []
            for j in range(i + 1, min(len(lines), i + 1 + max_lines)):
                cand = lines[j]
                nc = normalize_for_match(cand)
                if any(nc.startswith(s + " ") or nc == s for s in stops_n):
                    break
                collected.append(cand)
                if len(" ".join(collected)) > 160:
                    break
            if collected:
                return clean_value(" ".join(collected))
    return ""


def clean_person_name(value: str) -> str:
    value = clean_value(value)
    n = normalize_for_match(value)
    cuts = [" COD ", " CODIGO ", " MATRICULA ", " CARGO ", " SETOR ", " CPF ", " RG ", " DATA ", " NASC ", " IDADE "]
    pos = len(n)
    for marker in cuts:
        p = n.find(marker)
        if p >= 0:
            pos = min(pos, p)
    if pos < len(n):
        # Normalização preserva aproximadamente o comprimento só para ASCII; melhor refazer por regex.
        value = re.split(r"\s+(?:COD(?:IGO)?|MATRICULA|CARGO|SETOR|CPF|RG|DATA|NASC|IDADE)\b", value, maxsplit=1, flags=re.I)[0]
    value = re.sub(r"\b(?:COD|CODIGO)\s*[:#-]?\s*\d.*$", "", value, flags=re.I)
    return clean_value(value)


def clean_company(value: str) -> str:
    value = clean_value(value)
    value = re.split(r"\s+CNPJ\b", value, maxsplit=1, flags=re.I)[0]
    cnpj = find_cnpj(value)
    if cnpj:
        d = digits_only(cnpj)
        p = digits_only(value).find(d)
        if p >= 0:
            # fallback: remove o número reconhecido via regex em vez de índice de dígitos.
            value = CNPJ_RE.sub("", value)
    return clean_value(value)


def extract_fields(text: str, exam_type: str) -> tuple[str, str, str, str]:
    name_labels = ("FUNCIONARIO", "NOME DO FUNCIONARIO", "COLABORADOR", "NOME", "PACIENTE", "AVALIADO")
    name_stops = ("EMPRESA", "CONVENIO", "CNPJ", "CPF", "RG", "DATA", "NASCIMENTO", "IDADE", "SEXO", "MATRICULA", "FUNCAO", "CARGO", "SETOR", "DOC")
    company_stops = ("CNPJ", "CPF", "ENDERECO", "TELEFONE", "FUNCIONARIO", "NOME", "FUNCAO", "CARGO", "SETOR", "DATA", "DOC")

    # Alguns modelos do sistema atual usam 'Funcionário' no ASO e 'Nome' nos exames complementares.
    if exam_type == "ASO":
        labels = ("FUNCIONARIO", "NOME DO FUNCIONARIO", "COLABORADOR", "NOME")
    else:
        labels = name_labels

    name = _find_labeled_value(text, labels, name_stops, max_lines=2)
    company = _find_labeled_value(text, ("EMPRESA", "CONVENIO"), company_stops, max_lines=2)
    cnpj = ""
    # O CNPJ precisa estar ligado ao campo EMPRESA/CONVÊNIO. Isso evita usar, por engano,
    # o CNPJ da clínica que costuma aparecer no cabeçalho dos exames.
    raw_lines = [clean_line(x) for x in str(text or "").splitlines() if clean_line(x)]
    for i, line in enumerate(raw_lines):
        nline = normalize_for_match(line)
        if "EMPRESA" not in nline and "CONVENIO" not in nline:
            continue
        segment = " ".join(raw_lines[i:i+2])
        candidate = find_cnpj(segment)
        if candidate:
            cnpj = candidate
            break
    cpf = find_cpf(text)
    return clean_person_name(name), clean_company(company), cnpj, cpf


_TESSERACT_EXE_CACHE: str | None = None

def locate_tesseract() -> str:
    """Localiza o binário do Tesseract sem guardar resultado negativo.

    No Render, quando o build é ajustado e o serviço reinicia, o binário costuma
    ficar em /usr/bin/tesseract. Em versões anteriores o app guardava "" em cache
    quando não encontrava o OCR; se o ambiente fosse corrigido depois, a tela
    continuava exibindo OCR não localizado até reiniciar o processo. Agora só
    guardamos em cache quando o executável é realmente encontrado.
    """
    global _TESSERACT_EXE_CACHE
    if _TESSERACT_EXE_CACHE and Path(_TESSERACT_EXE_CACHE).exists():
        return _TESSERACT_EXE_CACHE

    env_cmd = (os.environ.get("TESSERACT_CMD") or "").strip()
    candidates: list[Path] = []
    if env_cmd:
        candidates.append(Path(env_cmd))

    found = shutil.which("tesseract")
    if found:
        candidates.append(Path(found))

    candidates.extend([
        Path("/usr/bin/tesseract"),
        Path("/usr/local/bin/tesseract"),
        Path("/opt/render/project/src/.local/bin/tesseract"),
        Path(os.environ.get("PROGRAMFILES", r"C:\Program Files")) / "Tesseract-OCR" / "tesseract.exe",
        Path(os.environ.get("PROGRAMFILES(X86)", r"C:\Program Files (x86)")) / "Tesseract-OCR" / "tesseract.exe",
        Path.home() / "AppData" / "Local" / "Programs" / "Tesseract-OCR" / "tesseract.exe",
    ])
    for c in candidates:
        try:
            if c and c.exists():
                _TESSERACT_EXE_CACHE = str(c)
                return _TESSERACT_EXE_CACHE
        except Exception:
            continue
    return ""


def ocr_image(image: Image.Image, psm: int = 6, timeout_seconds: float | None = None) -> tuple[str, str]:
    """Executa OCR com limite de tempo para nenhuma página travar o lote inteiro."""
    if pytesseract is None:
        return "", "pytesseract não instalado"
    exe = locate_tesseract()
    if not exe:
        return "", "Tesseract OCR não localizado"
    pytesseract.pytesseract.tesseract_cmd = exe
    image = ImageOps.autocontrast(image.convert("L")).convert("RGB")
    local_por = TESSDATA_DIR / "por.traineddata"
    base_cfg = f"--psm {int(psm)}"
    # PSM 11 é usado na triagem leve. Se uma página anormal fizer o Tesseract
    # demorar demais, ela não pode bloquear centenas de páginas seguintes.
    timeout_limit = float(timeout_seconds if timeout_seconds is not None else (8.0 if int(psm) == 11 else 18.0))

    def _run(lang: str, cfg: str) -> str:
        return pytesseract.image_to_string(image, lang=lang, config=cfg, timeout=timeout_limit)

    if local_por.exists():
        try:
            cfg = f'--tessdata-dir "{TESSDATA_DIR}" {base_cfg}'
            return _run("por", cfg), ""
        except RuntimeError as exc:
            if "timeout" in str(exc).lower():
                return "", f"OCR excedeu {timeout_limit:.0f}s e a página foi liberada para continuar o lote"
        except Exception:
            pass
    try:
        return _run("por", base_cfg), ""
    except RuntimeError as exc:
        if "timeout" in str(exc).lower():
            return "", f"OCR excedeu {timeout_limit:.0f}s e a página foi liberada para continuar o lote"
        # RuntimeError não relacionado a timeout pode ser falta do idioma.
        first_exc = exc
    except Exception as exc:
        first_exc = exc
    try:
        return _run("eng", base_cfg), "OCR português indisponível; usado inglês"
    except RuntimeError as exc:
        if "timeout" in str(exc).lower():
            return "", f"OCR excedeu {timeout_limit:.0f}s e a página foi liberada para continuar o lote"
        return "", f"Falha no OCR: {exc}"
    except Exception as exc:
        return "", f"Falha no OCR: {exc}"


def ocr_images_batch(images: list[Image.Image], psm: int = 11, timeout_seconds: float = 18.0) -> tuple[list[str], str]:
    """OCR de várias regiões de página em uma única chamada ao Tesseract.

    As imagens são empilhadas verticalmente e o retorno TSV é separado de volta
    pelas coordenadas Y. Isso reduz drasticamente o custo de iniciar o processo
    do Tesseract dezenas/centenas de vezes em um lote grande.
    """
    if not images:
        return [], ""
    if len(images) == 1:
        text, warning = ocr_image(images[0], psm=psm, timeout_seconds=min(timeout_seconds, 10.0))
        return [text], warning
    if pytesseract is None:
        return ["" for _ in images], "pytesseract não instalado"
    exe = locate_tesseract()
    if not exe:
        return ["" for _ in images], "Tesseract OCR não localizado"
    pytesseract.pytesseract.tesseract_cmd = exe

    prepared = [ImageOps.autocontrast(im.convert("L")) for im in images]
    width = max(im.width for im in prepared)
    gap = 28
    offsets: list[tuple[int, int]] = []
    y = 0
    for im in prepared:
        offsets.append((y, y + im.height))
        y += im.height + gap
    canvas = Image.new("L", (width, max(1, y - gap)), 255)
    for im, (y0, _y1) in zip(prepared, offsets):
        canvas.paste(im, (0, y0))
    canvas = canvas.convert("RGB")

    local_por = TESSDATA_DIR / "por.traineddata"
    base_cfg = f"--psm {int(psm)}"

    def _run(lang: str, cfg: str):
        return pytesseract.image_to_data(
            canvas, lang=lang, config=cfg, output_type=pytesseract.Output.DICT,
            timeout=float(timeout_seconds),
        )

    warning = ""
    data = None
    if local_por.exists():
        try:
            data = _run("por", f'--tessdata-dir "{TESSDATA_DIR}" {base_cfg}')
        except Exception as exc:
            if "timeout" in str(exc).lower():
                return ["" for _ in images], f"OCR em lote excedeu {timeout_seconds:.0f}s; páginas liberadas"
    if data is None:
        try:
            data = _run("por", base_cfg)
        except Exception:
            try:
                data = _run("eng", base_cfg)
                warning = "OCR português indisponível; usado inglês"
            except Exception as exc:
                if "timeout" in str(exc).lower():
                    return ["" for _ in images], f"OCR em lote excedeu {timeout_seconds:.0f}s; páginas liberadas"
                return ["" for _ in images], f"Falha no OCR em lote: {exc}"

    # Agrupa palavras por página e linha, usando as coordenadas retornadas pelo TSV.
    page_lines: list[dict[tuple[int, int, int], list[tuple[int, str]]]] = [defaultdict(list) for _ in images]
    count = len(data.get("text", []))
    for k in range(count):
        word = str(data["text"][k] or "").strip()
        if not word:
            continue
        try:
            cy = int(data["top"][k]) + max(1, int(data["height"][k])) // 2
            left = int(data["left"][k])
        except Exception:
            continue
        page_idx = None
        local_y = cy
        for pi, (y0, y1) in enumerate(offsets):
            if y0 <= cy <= y1:
                page_idx = pi
                local_y = cy - y0
                break
        if page_idx is None:
            continue
        # line_num costuma reiniciar entre blocos; o Y local evita colisões entre linhas.
        key = (
            int(data.get("block_num", [0] * count)[k] or 0),
            int(data.get("par_num", [0] * count)[k] or 0),
            int(round(local_y / 8.0)),
        )
        page_lines[page_idx][key].append((left, word))

    texts: list[str] = []
    for lines in page_lines:
        ordered = []
        for key in sorted(lines, key=lambda z: z[2]):
            words = " ".join(w for _x, w in sorted(lines[key], key=lambda x: x[0])).strip()
            if words:
                ordered.append(words)
        texts.append("\n".join(ordered))
    return texts, warning


def render_page_image(page: fitz.Page, scale: float = 1.8) -> Image.Image:
    pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
    return Image.frombytes("RGB", (pix.width, pix.height), pix.samples)


def render_page_top_image(page: fitz.Page, scale: float = 0.82, top_fraction: float = 0.42) -> Image.Image:
    """Renderiza somente o topo da página para a triagem OCR inicial.

    Títulos do exame, identificação do paciente e campos principais normalmente
    ficam no primeiro terço/quase metade da folha. Ler só essa região reduz
    drasticamente a quantidade de pixels enviados ao Tesseract.
    """
    frac = min(0.65, max(0.25, float(top_fraction)))
    rect = page.rect
    clip = fitz.Rect(rect.x0, rect.y0, rect.x1, rect.y0 + rect.height * frac)
    pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale), clip=clip, alpha=False)
    return Image.frombytes("RGB", (pix.width, pix.height), pix.samples)


def _embedded_text_is_good(embedded: str) -> bool:
    return len(embedded) >= 180 and len(re.findall(r"[A-Za-zÀ-ÿ]{4,}", embedded)) >= 20


def page_text(page: fitz.Page, use_ocr: bool = True) -> tuple[str, str, str]:
    """Leitura detalhada, usada apenas quando a triagem rápida não é suficiente."""
    embedded = page.get_text("text").strip()
    source = "Texto interno"
    warning = ""
    if use_ocr and not _embedded_text_is_good(embedded):
        img = render_page_image(page, scale=1.65)
        ocr, warning = ocr_image(img, psm=6)
        if ocr.strip():
            source = "Texto interno + OCR detalhado" if embedded else "OCR detalhado"
            return ((embedded + "\n" + ocr).strip() if embedded else ocr.strip()), source, warning
    return embedded, source, warning


def _page_has_large_raster(page: fitz.Page, min_coverage: float = 0.22) -> bool:
    """Retorna True quando a página parece ser uma digitalização/imagem grande.

    Isso permite confiar no texto interno de PDFs digitais sem confundir uma pequena
    camada de texto de uma página escaneada com o conteúdo real da página.
    """
    try:
        page_area = max(1.0, float(page.rect.width * page.rect.height))
        covered = 0.0
        seen_rects: set[tuple[int, int, int, int]] = set()
        for item in page.get_images(full=True):
            xref = int(item[0])
            try:
                rects = page.get_image_rects(xref)
            except Exception:
                rects = []
            for rect in rects:
                key = (round(rect.x0), round(rect.y0), round(rect.x1), round(rect.y1))
                if key in seen_rects:
                    continue
                seen_rects.add(key)
                area = max(0.0, float(rect.width * rect.height))
                if area / page_area >= min_coverage:
                    return True
                covered += area
        return (covered / page_area) >= min_coverage
    except Exception:
        # Na dúvida, não descarta: preserva a segurança da extração.
        return True


def quick_page_text(page: fitz.Page, use_ocr: bool = True, embedded: str | None = None) -> tuple[str, str, str, Image.Image | None]:
    """Triagem rápida: texto interno primeiro; OCR leve apenas quando necessário.

    O OCR preliminar usa resolução menor que o OCR detalhado. A imagem retornada é
    reutilizada na classificação visual para evitar renderizações duplicadas.
    """
    if embedded is None:
        embedded = page.get_text("text").strip()
    if _embedded_text_is_good(embedded) or not use_ocr:
        return embedded, "Texto interno (triagem)", "", None
    # 0.86 reduz bastante a quantidade de pixels sem perder títulos/campos grandes
    # que servem apenas para decidir se a página merece análise detalhada.
    img = render_page_image(page, scale=0.78)
    ocr, warning = ocr_image(img, psm=11)
    if ocr.strip():
        text = ((embedded + "\n" + ocr).strip() if embedded else ocr.strip())
        return text, "OCR rápido", warning, img
    return embedded, "Texto interno (triagem)", warning, img

def token_signature(text: str, limit: int = 180) -> list[str]:
    n = normalize_for_match(text)
    tokens = [t for t in n.split() if len(t) >= 4 and not t.isdigit() and t not in MODEL_STOPWORDS]
    counts = Counter(tokens)
    # Mantém as palavras mais relevantes, mas sem repetir.
    return [t for t, _ in counts.most_common(limit)]


def visual_hash(image: Image.Image, size: int = 16) -> str:
    # dHash 16x16 = 256 bits. Robusto para o mesmo formulário digitalizado em qualidade diferente.
    img = ImageOps.grayscale(image).resize((size + 1, size), Image.Resampling.LANCZOS)
    px = list(img.getdata())
    bits = []
    for y in range(size):
        row = y * (size + 1)
        for x in range(size):
            bits.append(1 if px[row + x] > px[row + x + 1] else 0)
    value = 0
    for bit in bits:
        value = (value << 1) | bit
    return f"{value:0{size * size // 4}x}"


def hash_similarity(h1: str, h2: str) -> float:
    try:
        a, b = int(h1, 16), int(h2, 16)
        bits = max(len(h1), len(h2)) * 4
        return 1.0 - ((a ^ b).bit_count() / bits)
    except Exception:
        return 0.0


def token_similarity(a: Iterable[str], b: Iterable[str]) -> float:
    sa, sb = set(a), set(b)
    if not sa or not sb:
        return 0.0
    overlap = len(sa & sb)
    # overlap coefficient favorece modelos iguais mesmo quando OCR adiciona ruído.
    return overlap / max(1, min(len(sa), len(sb)))


def load_models() -> list[ModelSample]:
    refresh_exam_types()
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    MODELS_DIR.mkdir(parents=True, exist_ok=True)
    if not MODELS_JSON.exists():
        return []
    try:
        raw = json.loads(MODELS_JSON.read_text(encoding="utf-8"))
        return [ModelSample(**x) for x in raw if x.get("exam_type") in EXAM_TYPES]
    except Exception:
        return []


def save_models(models: list[ModelSample]) -> None:
    refresh_exam_types()
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    MODELS_DIR.mkdir(parents=True, exist_ok=True)
    MODELS_JSON.write_text(json.dumps([asdict(m) for m in models], ensure_ascii=False, indent=2), encoding="utf-8")


def model_file_reference(path: Path) -> str:
    """Guarda o PDF modelo de forma compatível com Render Persistent Disk.

    Em versões anteriores o sistema tentava salvar o caminho relativo ao diretório
    do código (APP_DIR). Quando o Render usa /var/data, o arquivo fica fora de
    /opt/render/project/src e isso causava: "is not in the subpath".
    Por isso, agora gravamos caminho absoluto para arquivos persistentes e
    aceitamos também referências antigas relativas.
    """
    try:
        return str(Path(path).resolve())
    except Exception:
        return str(path)


def resolve_model_file(stored_file: str) -> Path:
    raw = str(stored_file or "").strip()
    if not raw:
        return Path()
    p = Path(raw)
    if p.is_absolute():
        return p
    # Compatibilidade com modelos antigos salvos relativos ao app.
    candidate = APP_DIR / p
    if candidate.exists():
        return candidate
    # Compatibilidade com modelos novos salvos apenas pelo nome.
    candidate = MODELS_DIR / p.name
    if candidate.exists():
        return candidate
    return candidate


def add_model_from_pdf(path: Path, exam_type: str, label: str, use_ocr: bool = True, progress: Optional[Callable[[int,int],None]] = None) -> list[ModelSample]:
    exam_type = normalize_exam(exam_type) or exam_type
    if exam_type not in EXAM_TYPES:
        raise ValueError("Tipo de exame inválido")
    models = load_models()
    new_models: list[ModelSample] = []
    doc = fitz.open(path)
    try:
        total = doc.page_count
        for idx in range(total):
            page = doc[idx]
            text, _, _ = page_text(page, use_ocr=use_ocr)
            img = render_page_image(page, scale=1.0)
            stored_name = f"{uuid.uuid4().hex}_{safe_component(path.stem, 'modelo', 60)}_p{idx+1}.pdf"
            stored = MODELS_DIR / stored_name
            single = fitz.open()
            try:
                single.insert_pdf(doc, from_page=idx, to_page=idx)
                single.save(stored, garbage=4, deflate=True)
            finally:
                single.close()
            model = ModelSample(
                id=uuid.uuid4().hex,
                exam_type=exam_type,
                label=f"{label} - pág. {idx+1}" if total > 1 else label,
                source_filename=path.name,
                page_number=idx + 1,
                created_at=datetime.now().isoformat(timespec="seconds"),
                token_signature=token_signature(text),
                visual_hash=visual_hash(img),
                aspect_ratio=round(img.width / max(1, img.height), 4),
                stored_file=model_file_reference(stored),
            )
            models.append(model)
            new_models.append(model)
            if progress:
                progress(idx + 1, total)
    finally:
        doc.close()
    save_models(models)
    return new_models




def add_models_from_pdfs(paths: list[Path], exam_type: str, label_prefix: str, use_ocr: bool = True,
                         progress: Optional[Callable[[int, int, str], None]] = None) -> list[ModelSample]:
    """Cadastra vários PDFs do MESMO tipo de exame em uma única operação."""
    exam_type = normalize_exam(exam_type) or exam_type
    if exam_type not in EXAM_TYPES:
        raise ValueError("Tipo de exame inválido")
    paths = [Path(x) for x in paths if Path(x).exists()]
    if not paths:
        raise ValueError("Nenhum PDF de modelo foi selecionado")
    total_pages = 0
    for path in paths:
        doc = fitz.open(path)
        try:
            total_pages += doc.page_count
        finally:
            doc.close()
    models = load_models()
    created: list[ModelSample] = []
    done = 0
    for path in paths:
        doc = fitz.open(path)
        try:
            for idx in range(doc.page_count):
                page = doc[idx]
                text, _, _ = page_text(page, use_ocr=use_ocr)
                img = render_page_image(page, scale=0.9)
                stored_name = f"{uuid.uuid4().hex}_{safe_component(path.stem, 'modelo', 60)}_p{idx+1}.pdf"
                stored = MODELS_DIR / stored_name
                single = fitz.open()
                try:
                    single.insert_pdf(doc, from_page=idx, to_page=idx)
                    single.save(stored, garbage=4, deflate=True)
                finally:
                    single.close()
                base_label = label_prefix.strip() or exam_type
                if len(paths) > 1:
                    base_label = f"{base_label} - {path.stem}"
                label = f"{base_label} - pág. {idx+1}" if doc.page_count > 1 else base_label
                model = ModelSample(
                    id=uuid.uuid4().hex,
                    exam_type=exam_type,
                    label=label,
                    source_filename=path.name,
                    page_number=idx + 1,
                    created_at=datetime.now().isoformat(timespec="seconds"),
                    token_signature=token_signature(text),
                    visual_hash=visual_hash(img),
                    aspect_ratio=round(img.width / max(1, img.height), 4),
                    stored_file=model_file_reference(stored),
                )
                models.append(model)
                created.append(model)
                done += 1
                if progress:
                    progress(done, total_pages, f"{path.name} - página {idx+1}")
        finally:
            doc.close()
    save_models(models)
    return created

def add_model_from_page(path: Path, page_index: int, exam_type: str, label: str, use_ocr: bool = True) -> ModelSample:
    """Cadastra somente uma página de um PDF como amostra de modelo."""
    exam_type = normalize_exam(exam_type) or exam_type
    if exam_type not in EXAM_TYPES:
        raise ValueError("Tipo de exame inválido")
    doc = fitz.open(path)
    try:
        if page_index < 0 or page_index >= doc.page_count:
            raise ValueError("Página de modelo inválida")
        page = doc[page_index]
        text, _, _ = page_text(page, use_ocr=use_ocr)
        img = render_page_image(page, scale=1.0)
        stored_name = f"{uuid.uuid4().hex}_{safe_component(path.stem, 'modelo', 60)}_p{page_index+1}.pdf"
        stored = MODELS_DIR / stored_name
        single = fitz.open()
        try:
            single.insert_pdf(doc, from_page=page_index, to_page=page_index)
            single.save(stored, garbage=4, deflate=True)
        finally:
            single.close()
        model = ModelSample(
            id=uuid.uuid4().hex, exam_type=exam_type, label=label, source_filename=path.name, page_number=page_index+1,
            created_at=datetime.now().isoformat(timespec="seconds"), token_signature=token_signature(text),
            visual_hash=visual_hash(img), aspect_ratio=round(img.width / max(1, img.height), 4),
            stored_file=model_file_reference(stored),
        )
    finally:
        doc.close()
    models = load_models(); models.append(model); save_models(models)
    return model

def add_model_from_image(path: Path, exam_type: str, label: str, use_ocr: bool = True) -> ModelSample:
    exam_type = normalize_exam(exam_type) or exam_type
    if exam_type not in EXAM_TYPES:
        raise ValueError("Tipo de exame inválido")
    image = Image.open(path).convert("RGB")
    text, _ = ocr_image(image) if use_ocr else ("", "")
    ext = path.suffix.lower() if path.suffix else ".png"
    stored_name = f"{uuid.uuid4().hex}_{safe_component(path.stem, 'modelo', 60)}{ext}"
    stored = MODELS_DIR / stored_name
    shutil.copy2(path, stored)
    model = ModelSample(
        id=uuid.uuid4().hex, exam_type=exam_type, label=label, source_filename=path.name, page_number=1,
        created_at=datetime.now().isoformat(timespec="seconds"), token_signature=token_signature(text),
        visual_hash=visual_hash(image), aspect_ratio=round(image.width / max(1, image.height), 4),
        stored_file=model_file_reference(stored),
    )
    models = load_models(); models.append(model); save_models(models)
    return model


def delete_model(model_id: str) -> bool:
    models = load_models()
    kept: list[ModelSample] = []
    deleted = False
    for m in models:
        if m.id == model_id:
            deleted = True
            if m.stored_file:
                try:
                    resolve_model_file(m.stored_file).unlink(missing_ok=True)
                except Exception:
                    pass
        else:
            kept.append(m)
    save_models(kept)
    return deleted


def classify_page(text: str, image: Image.Image | None, models: list[ModelSample],
                  expected_hint: set[str] | None = None, candidate_types: set[str] | None = None,
                  recognition_floor: float = 28.0) -> tuple[str, float, str, bool]:
    n = normalize_for_match(text)
    ignore_score = sum(weight for phrase, weight in IGNORE_SIGNATURES if normalize_for_match(phrase) in n)
    page_tokens = token_signature(text)
    p_hash = visual_hash(image) if image is not None else ""
    aspect = (image.width / max(1, image.height)) if image is not None else 1.0

    allowed = set(candidate_types or EXAM_TYPES)
    best_type = ""
    best_score = 0.0
    breakdown: dict[str, float] = {}
    for exam in EXAM_TYPES:
        if exam not in allowed:
            continue
        keyword = sum(weight for phrase, weight in SIGNATURES[exam] if normalize_for_match(phrase) in n)
        # Para tipos cadastrados, o nome/alias do exame é a principal assinatura.
        # Pontua também quando todos os tokens relevantes aparecem, mesmo com
        # separadores diferentes no OCR.
        if exam not in DEFAULT_EXAM_TYPES and keyword < 30:
            exam_tokens = [t for t in normalize_for_match(exam).split() if len(t) >= 4 and t not in MODEL_STOPWORDS]
            if exam_tokens and all(t in n.split() for t in exam_tokens):
                keyword = max(keyword, 42.0)

        # Alguns sistemas imprimem o titulo como "ASO - Atestado de Saude
        # Ocupacional". O OCR pode errar ATESTADO, mas conservar o token ASO.
        # Quando ASO aparece no cabecalho, ele deve pesar mais que nomes de
        # exames complementares listados dentro do proprio ASO.
        if exam == "ASO":
            top_text = "\n".join(str(text or "").splitlines()[:14])
            top_norm = normalize_for_match(top_text)
            if re.search(r"(?:^| )ASO(?: |$)", top_norm):
                keyword += 75.0

        model_bonus = 0.0
        for m in models:
            if m.exam_type != exam:
                continue
            ts = token_similarity(page_tokens, m.token_signature)
            hs = hash_similarity(p_hash, m.visual_hash) if p_hash else 0.0
            ar = min(aspect, m.aspect_ratio) / max(aspect, m.aspect_ratio) if m.aspect_ratio else 1.0
            text_part = min(45.0, max(0.0, (ts - 0.18) / 0.62) * 45.0)
            visual_part = min(30.0, max(0.0, (hs - 0.62) / 0.32) * 30.0) * ar if p_hash else 0.0
            joint_bonus = 10.0 if ts >= 0.55 and hs >= 0.70 else 0.0
            model_bonus = max(model_bonus, min(85.0, text_part + visual_part + joint_bonus))
        hint = 6.0 if expected_hint and exam in expected_hint else 0.0
        score = keyword + model_bonus + hint
        if ignore_score >= 70:
            score -= 75
        elif ignore_score:
            score -= min(25, ignore_score * 0.3)
        score = max(0.0, min(100.0, score))
        breakdown[exam] = score
        if score > best_score:
            best_type, best_score = exam, score

    if best_score < float(recognition_floor):
        best_type = ""
    ignore = ignore_score >= 70 and best_score < 65
    bd = "; ".join(f"{k}={v:.0f}" for k, v in sorted(breakdown.items(), key=lambda x: x[1], reverse=True)[:3])
    if ignore_score:
        bd += f"; ignorar={ignore_score:.0f}"
    return best_type, round(best_score, 1), bd, ignore



def _token_close(tokens: set[str], target: str, threshold: float = 0.74) -> bool:
    """Reconhece uma palavra mesmo quando o OCR troca 1-3 caracteres.

    A função é usada somente na triagem rápida. Ela é deliberadamente simples e
    barata: compara poucas palavras-âncora contra os tokens já extraídos do
    cabeçalho, sem renderizar novamente a página nem chamar outro OCR.
    """
    target_n = normalize_for_match(target)
    if not target_n:
        return False
    if target_n in tokens:
        return True
    for tok in tokens:
        if len(tok) < 4:
            continue
        # Evita comparar palavras de tamanhos completamente diferentes.
        if abs(len(tok) - len(target_n)) > max(3, len(target_n) // 3):
            continue
        if SequenceMatcher(None, tok, target_n).ratio() >= threshold:
            return True
    return False


def structural_exam_candidates(text: str, candidate_types: set[str] | None = None) -> set[str]:
    """Detecta estrutura típica de exame em OCR ruidoso.

    Esta camada NÃO salva nada sozinha. Ela apenas impede que uma página
    potencialmente válida seja descartada antes do OCR detalhado. Assim, um
    cabeçalho lido como ``ATESTADO DE SAbDE OCUPACTONA`` ainda recebe uma
    segunda leitura, enquanto recibos/encaminhamentos continuam sendo pulados.
    """
    allowed = set(candidate_types or EXAM_TYPES)
    n = normalize_for_match(text)
    tokens = set(n.split())
    found: set[str] = set()

    def t(word: str, threshold: float = 0.74) -> bool:
        return _token_close(tokens, word, threshold)

    if "ASO" in allowed:
        title = t("atestado", 0.72) and t("ocupacional", 0.68)
        risks = t("riscos", 0.70) and t("ocupacionais", 0.68)
        company_block = t("empresa", 0.70) and (
            any(x in tokens for x in {"CNP", "CNPJ"}) or t("colaborador", 0.70) or t("funcionario", 0.70)
        )
        complementary = t("exames", 0.72) and t("complementares", 0.70)
        examiner = t("medico", 0.74) and t("examinador", 0.70)
        # Um título aproximado já é forte. Sem título, exigimos combinação de
        # blocos estruturais característicos do ASO para evitar falsos positivos.
        if title or (risks and company_block) or (risks and complementary) or (risks and examiner):
            found.add("ASO")

    if "AUDIOMETRIA" in allowed:
        audio_title = t("audiometria", 0.72) or t("audiometrica", 0.72) or t("audiologica", 0.72)
        ears = t("ouvido", 0.74) and (t("direito", 0.76) or t("esquerdo", 0.76))
        if audio_title or ears:
            found.add("AUDIOMETRIA")

    if "ESPIROMETRIA" in allowed:
        spiro_title = t("espirometria", 0.72) or (t("funcao", 0.76) and t("pulmonar", 0.72))
        # Em digitalizações de espirometria o OCR rápido frequentemente perde o
        # título "ESPIROMETRIA", mas ainda consegue ler termos do bloco de conclusão.
        # Esses sinais servem apenas para mandar a página ao OCR detalhado; não salvam
        # o arquivo automaticamente sem confirmar funcionário e tipo.
        lung_values = (
            any(x in tokens for x in {"VEF1", "CVF", "FEF25", "FEF50", "FEF75"})
            or t("vef1", 0.62)
            or t("cvf", 0.64)
            or (t("volume", 0.70) and t("expiratorio", 0.64))
            or (t("capacidade", 0.66) and t("vital", 0.70))
            or (t("funcao", 0.72) and t("ventilatoria", 0.64))
        )
        if spiro_title or lung_values:
            found.add("ESPIROMETRIA")

    if "ACUIDADE VISUAL" in allowed:
        acuity_title = t("acuidade", 0.72) and t("visual", 0.76)
        eye_block = t("olho", 0.76) and (t("direito", 0.76) or t("esquerdo", 0.76))
        vision_block = t("visao", 0.72) and (t("longe", 0.76) or t("perto", 0.76))
        if acuity_title or (eye_block and vision_block):
            found.add("ACUIDADE VISUAL")

    if "LAUDO PCD" in allowed:
        pcd_title = (t("laudo", 0.76) and ("PCD" in tokens or t("deficiencia", 0.72)))
        disability = t("pessoa", 0.76) and t("deficiencia", 0.72)
        if pcd_title or disability:
            found.add("LAUDO PCD")

    # Tipos novos cadastrados pelo usuário. Serve apenas como triagem para não
    # descartar uma página antes da varredura/modelo visual.
    for exam in allowed:
        if exam in DEFAULT_EXAM_TYPES:
            continue
        exam_tokens = [tok for tok in normalize_for_match(exam).split() if len(tok) >= 4 and tok not in MODEL_STOPWORDS]
        if exam_tokens and all(_token_close(tokens, tok, 0.68) for tok in exam_tokens[:4]):
            found.add(exam)
            continue
        for alias, canonical in EXAM_ALIASES.items():
            if canonical != exam:
                continue
            alias_tokens = [tok for tok in alias.split() if len(tok) >= 4 and tok not in MODEL_STOPWORDS]
            if alias_tokens and all(_token_close(tokens, tok, 0.68) for tok in alias_tokens[:4]):
                found.add(exam)
                break

    return found

def _cell_str(value) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return clean_value(str(value))


def _header_key(value: str) -> str:
    return normalize_for_match(value).replace(" ", "_")



def split_company_and_cnpj(value: str) -> tuple[str, str]:
    """Aceita EMPRESA / CNPJ ou EMPRESA / CPF na mesma célula.

    O nome do campo histórico continua ``cnpj`` por compatibilidade interna,
    mas o valor pode ser um CNPJ de 14 dígitos ou um CPF de 11 dígitos.
    """
    raw = clean_value(value)
    document = find_cnpj(raw)
    company = raw
    if document:
        company = CNPJ_RE.sub("", company, count=1)
    else:
        m = CPF_RE.search(raw)
        if m:
            document = format_cpf("".join(m.groups()))
            company = CPF_RE.sub("", company, count=1)
    if document:
        company = re.sub(r"\s*[/|;]\s*$", "", company.strip()).strip()
        company = clean_value(company)
    return company, document

def spreadsheet_sheets(path: Path) -> list[str]:
    if path.suffix.lower() == ".csv":
        return ["CSV"]
    if load_workbook is None:
        raise RuntimeError("openpyxl não está instalado")
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        return list(wb.sheetnames)
    finally:
        wb.close()


def _rows_from_file(path: Path, sheet_name: str | None = None) -> list[list[str]]:
    if path.suffix.lower() == ".csv":
        data = path.read_text(encoding="utf-8-sig", errors="replace")
        sample = data[:3000]
        try:
            dialect = csv.Sniffer().sniff(sample, delimiters=";,\t,")
        except Exception:
            dialect = csv.excel; dialect.delimiter = ";"
        return [[_cell_str(v) for v in row] for row in csv.reader(data.splitlines(), dialect)]
    if load_workbook is None:
        raise RuntimeError("openpyxl não está instalado")
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        ws = wb[sheet_name] if sheet_name and sheet_name in wb.sheetnames else wb[wb.sheetnames[0]]
        return [[_cell_str(v) for v in row] for row in ws.iter_rows(values_only=True)]
    finally:
        wb.close()


def load_expected_list(path: Path, sheet_name: str | None = None) -> list[ExpectedEmployee]:
    rows = _rows_from_file(path, sheet_name)
    # Remove linhas totalmente vazias no início.
    while rows and not any(clean_value(x) for x in rows[0]):
        rows.pop(0)
    if not rows:
        raise ValueError("A lista está vazia")

    # Localiza cabeçalho nas primeiras 15 linhas, procurando coluna de nome.
    header_idx = 0
    headers: list[str] = []
    name_aliases = {"FUNCIONARIO", "NOME", "NOME_DO_FUNCIONARIO", "COLABORADOR", "NOME_COLABORADOR"}
    for idx, row in enumerate(rows[:15]):
        hk = [_header_key(x) for x in row]
        if any(h in name_aliases for h in hk):
            header_idx, headers = idx, hk
            break
    else:
        headers = [_header_key(x) for x in rows[0]]

    def find_col(aliases: set[str]) -> int | None:
        for i, h in enumerate(headers):
            if h in aliases:
                return i
        return None

    name_col = find_col(name_aliases)
    if name_col is None:
        raise ValueError("Não encontrei a coluna de funcionário. Use um cabeçalho como FUNCIONÁRIO, NOME ou COLABORADOR.")
    cpf_col = find_col({"CPF", "CPF_FUNCIONARIO", "CPF_COLABORADOR"})
    company_col = find_col({"EMPRESA", "CONVENIO", "RAZAO_SOCIAL", "EMPRESA_CONVENIO"})
    cnpj_col = find_col({"CNPJ", "CNPJ_EMPRESA", "CPF_CNPJ", "CPF_CNPJ_EMPRESA", "CNPJ_CPF", "DOCUMENTO_EMPRESA"})
    exam_col = find_col({"EXAME", "TIPO_DE_EXAME", "EXAMES", "EXAMES_ESPERADOS", "TIPO_EXAME"})
    receipt_col = find_col({"RECIBO", "NUMERO_RECIBO", "NUMERO_DE_RECIBO", "N_RECIBO", "N_DE_RECIBO", "NR_RECIBO", "RECIBO_EXAME"})

    wide_exam_cols: dict[int, str] = {}
    for i, h in enumerate(headers):
        ex = normalize_exam(h.replace("_", " "))
        if ex:
            wide_exam_cols[i] = ex

    grouped: dict[str, ExpectedEmployee] = {}
    # Índices auxiliares para unir linhas do MESMO funcionário quando uma ocorrência
    # vem sem CPF (ou com CPF inválido por erro de digitação/OCR) e outra traz o CPF correto.
    # A empresa faz parte da identidade para não misturar homônimos de clientes diferentes.
    by_name_company: dict[str, list[ExpectedEmployee]] = defaultdict(list)
    by_cpf_company: dict[str, ExpectedEmployee] = {}
    row_id = 0
    for row in rows[header_idx + 1:]:
        if name_col >= len(row):
            continue
        name = clean_person_name(row[name_col])
        if not name:
            continue
        row_id += 1
        cpf_raw = row[cpf_col] if cpf_col is not None and cpf_col < len(row) else ""
        cpf_digits = digits_only(cpf_raw)
        if cpf_digits and len(cpf_digits) < 11 and len(cpf_digits) >= 9:
            cpf_digits = cpf_digits.zfill(11)
        cpf = format_cpf(cpf_digits) if len(cpf_digits) == 11 else clean_value(cpf_raw)
        company_raw = row[company_col] if company_col is not None and company_col < len(row) else ""
        company, cnpj_from_company = split_company_and_cnpj(company_raw)
        company = clean_company(company)
        cnpj_raw = row[cnpj_col] if cnpj_col is not None and cnpj_col < len(row) else ""
        cnpj_digits = digits_only(cnpj_raw)
        if cnpj_digits and len(cnpj_digits) < 14 and len(cnpj_digits) >= 12:
            cnpj_digits = cnpj_digits.zfill(14)
        if len(cnpj_digits) == 14:
            cnpj = format_cnpj(cnpj_digits)
        elif len(cnpj_digits) == 11:
            cnpj = format_cpf(cnpj_digits)
        else:
            cnpj = clean_value(cnpj_raw)
        if not digits_only(cnpj) and cnpj_from_company:
            cnpj = cnpj_from_company
        cpf_key = digits_only(cpf)
        cpf_valid = is_valid_cpf(cpf_key)
        company_key = digits_only(cnpj) or normalize_for_match(company)
        name_company_key = normalize_for_match(name) + (("|" + company_key) if company_key else "")
        cpf_company_key = cpf_key + (("|" + company_key) if company_key else "") if cpf_valid else ""

        emp = by_cpf_company.get(cpf_company_key) if cpf_company_key else None
        if emp is None:
            same_name = by_name_company.get(name_company_key, [])
            # Une com linha homônima da mesma empresa somente quando não há conflito entre
            # dois CPFs válidos. Isso corrige planilhas em que um exame possui CPF vazio e
            # outro exame do mesmo colaborador possui CPF preenchido.
            compatible = []
            for cand in same_name:
                cand_cpf = digits_only(cand.cpf)
                if not cpf_valid or not is_valid_cpf(cand_cpf) or cand_cpf == cpf_key:
                    compatible.append(cand)
            if len(compatible) == 1:
                emp = compatible[0]
            elif len(compatible) > 1 and cpf_valid:
                exact = [c for c in compatible if digits_only(c.cpf) == cpf_key]
                if len(exact) == 1:
                    emp = exact[0]

        if emp is None:
            key = cpf_company_key or (name_company_key + f"|ROW{row_id}")
            emp = ExpectedEmployee(row_id=row_id, name=name, cpf=(cpf if cpf_valid else ""), company=company, cnpj=cnpj, source_sheet=sheet_name or "")
            grouped[key] = emp
            by_name_company[name_company_key].append(emp)
        # Atualiza dados vazios ou não confiáveis com dados válidos de linhas posteriores.
        emp.company = emp.company or company
        emp.cnpj = emp.cnpj or cnpj
        if cpf_valid and not is_valid_cpf(emp.cpf):
            emp.cpf = format_cpf(cpf_key)
        if cpf_company_key:
            by_cpf_company[cpf_company_key] = emp
        receipt = row[receipt_col] if receipt_col is not None and receipt_col < len(row) else ""
        receipt = clean_value(receipt)
        if normalize_for_match(receipt) in {"A PRAZO", "PRAZO"}:
            receipt = "A PRAZO"

        if exam_col is not None and exam_col < len(row):
            for ex in parse_exam_cell(row[exam_col]):
                emp.add_expected_exam(ex, receipt=receipt)
        for col, ex in wide_exam_cols.items():
            if col >= len(row):
                continue
            marker = normalize_for_match(row[col])
            if marker in {"X", "SIM", "S", "OK", "1", "TRUE", "VERDADEIRO", "REALIZAR", "PREVISTO"}:
                emp.add_expected_exam(ex, receipt=receipt)

    employees = list(grouped.values())
    if not employees:
        raise ValueError("Nenhum funcionário válido foi encontrado na lista")
    if not any(e.expected_exams for e in employees):
        raise ValueError("Encontrei os funcionários, mas não identifiquei os tipos de exames esperados.")
    return employees


def _name_similarity(a: str, b: str) -> float:
    na, nb = normalize_for_match(a), normalize_for_match(b)
    if not na or not nb:
        return 0.0
    if na == nb:
        return 1.0
    ta, tb = set(na.split()), set(nb.split())
    token = len(ta & tb) / max(1, len(ta | tb))
    seq = SequenceMatcher(None, na, nb).ratio()
    return max(seq, token)


@dataclass
class EmployeeIndex:
    by_cpf: dict[str, list[ExpectedEmployee]]
    by_token: dict[str, set[int]]
    by_row: dict[int, ExpectedEmployee]
    normalized_name: dict[int, str]


def build_employee_index(employees: list[ExpectedEmployee]) -> EmployeeIndex:
    by_cpf: dict[str, list[ExpectedEmployee]] = defaultdict(list)
    by_token: dict[str, set[int]] = defaultdict(set)
    by_row: dict[int, ExpectedEmployee] = {}
    normalized_name: dict[int, str] = {}
    weak = {"DE", "DA", "DO", "DAS", "DOS", "E"}
    for emp in employees:
        by_row[emp.row_id] = emp
        en = normalize_for_match(emp.name)
        normalized_name[emp.row_id] = en
        cpf = digits_only(emp.cpf)
        if is_valid_cpf(cpf):
            by_cpf[cpf].append(emp)
        for tok in set(en.split()):
            if len(tok) >= 4 and tok not in weak:
                by_token[tok].add(emp.row_id)
    return EmployeeIndex(dict(by_cpf), dict(by_token), by_row, normalized_name)


def _candidate_employees(text: str, extracted_name: str, employees: list[ExpectedEmployee], index: EmployeeIndex | None) -> list[ExpectedEmployee]:
    if index is None or len(employees) <= 12:
        return employees
    source = normalize_for_match(extracted_name or text)
    tokens = [t for t in set(source.split()) if len(t) >= 4]
    scores: Counter[int] = Counter()
    for tok in tokens:
        for row_id in index.by_token.get(tok, ()):
            scores[row_id] += 1
    if not scores:
        return employees
    # Mantém apenas candidatos com sobreposição real; limita fuzzy matching caro.
    ranked = [index.by_row[row] for row, _ in scores.most_common(16) if row in index.by_row]
    return ranked or employees


def match_employee(text: str, extracted_name: str, extracted_cpf: str, employees: list[ExpectedEmployee], employee_index: EmployeeIndex | None = None) -> tuple[ExpectedEmployee | None, float, str]:
    cpf_d = digits_only(extracted_cpf)
    ntext = normalize_for_match(text)
    text_digits = digits_only(text)

    # Empresa + nome é a âncora mais segura quando a mesma pessoa/CPF aparece em
    # mais de uma empresa ou quando uma linha da planilha tem CPF ausente/incorreto.
    # Isso corrige o caso clássico: mesmo nome e mesmo CPF histórico, mas o documento
    # atual traz claramente outro CNPJ de empresa.
    company_ranked: list[tuple[float, ExpectedEmployee]] = []
    for emp in employees:
        company_doc = digits_only(emp.cnpj)
        if len(company_doc) not in {11, 14} or company_doc not in text_digits:
            continue
        name_score = _name_similarity(extracted_name, emp.name) if extracted_name else (1.0 if normalize_for_match(emp.name) in ntext else 0.0)
        if name_score >= 0.82:
            company_ranked.append((name_score, emp))
    company_ranked.sort(key=lambda x: x[0], reverse=True)
    if company_ranked:
        top_score, top_emp = company_ranked[0]
        if len(company_ranked) == 1 or top_score > company_ranked[1][0] + 0.03:
            return top_emp, 100.0 if top_score >= 0.95 else 98.0, "nome + documento da empresa"

    if is_valid_cpf(cpf_d):
        cpf_matches = employee_index.by_cpf.get(cpf_d, []) if employee_index else [emp for emp in employees if digits_only(emp.cpf) == cpf_d]
        if len(cpf_matches) == 1:
            return cpf_matches[0], 100.0, "CPF"
        if len(cpf_matches) > 1:
            ranked: list[tuple[float, ExpectedEmployee, str]] = []
            for emp in cpf_matches:
                score, reason = 92.0, "CPF repetido na lista"
                cnpj_d = digits_only(emp.cnpj)
                comp = normalize_for_match(emp.company)
                if cnpj_d and cnpj_d in text_digits:
                    score, reason = 100.0, "CPF + CNPJ"
                elif comp and comp in ntext:
                    score, reason = 98.0, "CPF + empresa"
                ranked.append((score, emp, reason))
            ranked.sort(key=lambda x: x[0], reverse=True)
            if ranked[0][0] >= 98 and (len(ranked) == 1 or ranked[0][0] > ranked[1][0]):
                return ranked[0][1], ranked[0][0], ranked[0][2]
            return None, 75.0, "CPF aparece para mais de uma empresa na lista; conferir"

    best: ExpectedEmployee | None = None
    best_score = 0.0
    best_reason = ""
    ranked_candidates: list[tuple[float, ExpectedEmployee, str]] = []
    candidates = _candidate_employees(text, extracted_name, employees, employee_index)
    for emp in candidates:
        en = normalize_for_match(emp.name)
        score = 0.0
        reason = ""
        if en and en in ntext:
            score = 96.0
            reason = "nome completo no documento"
        if extracted_name:
            s = min(96.0, _name_similarity(extracted_name, emp.name) * 100)
            if s > score:
                score = s
                reason = "nome extraído"
        # Desempata pessoas com nomes iguais usando empresa/CNPJ da lista.
        cnpj_d = digits_only(emp.cnpj)
        text_digits = digits_only(text)
        if cnpj_d and cnpj_d in text_digits and score >= 70:
            score = min(100.0, score + 5.0)
            reason += " + CNPJ"
        elif emp.company and score >= 70:
            comp = normalize_for_match(emp.company)
            if comp and comp in ntext:
                score = min(100.0, score + 3.0)
                reason += " + empresa"
        # Tenta comparar com cada linha, útil quando o rótulo NOME/FUNCIONÁRIO falha no OCR.
        if score < 88 and en:
            for line in text.splitlines()[:90]:
                nl = normalize_for_match(line)
                if len(nl) < 5:
                    continue
                if en in nl:
                    score = max(score, 96.0); reason = "nome em linha"
                    break
                # Só compara linhas plausíveis para evitar custo e falsos positivos.
                if 2 <= len(nl.split()) <= 10:
                    sim = _name_similarity(nl, emp.name) * 100
                    if sim > score:
                        score = sim * 0.96
                        reason = "linha semelhante"
        ranked_candidates.append((score, emp, reason))
        if score > best_score:
            best, best_score, best_reason = emp, score, reason
    if best_score < 68:
        return None, round(best_score, 1), best_reason

    # Nunca escolhe arbitrariamente entre homônimos. Se dois registros obtiverem
    # praticamente a mesma pontuação e nenhum deles foi desempatado por empresa/CNPJ,
    # a página vai para conferência em vez de ser vinculada à pessoa errada.
    ranked_candidates.sort(key=lambda x: x[0], reverse=True)
    if len(ranked_candidates) > 1:
        top, second = ranked_candidates[0], ranked_candidates[1]
        top_company = digits_only(top[1].cnpj) or normalize_for_match(top[1].company)
        second_company = digits_only(second[1].cnpj) or normalize_for_match(second[1].company)
        if (
            top[0] >= 88 and second[0] >= 88 and abs(top[0] - second[0]) <= 1.0
            and top_company != second_company
            and "+ CNPJ" not in top[2] and "+ empresa" not in top[2]
        ):
            return None, round(top[0], 1), "nome homônimo sem empresa/CPF confiável para desempate"
    return best, round(best_score, 1), best_reason


def output_root(base_dir: Path) -> Path:
    """Usa exatamente a pasta escolhida pelo usuário como saída final, com proteção contra apagamento acidental."""
    root = Path(base_dir).expanduser()
    root.mkdir(parents=True, exist_ok=True)
    marker = root / ".edge_extrator_output"
    children = list(root.iterdir())

    # Só limpa automaticamente uma pasta que já tenha sido usada pelo extrator.
    # Isso evita que um caminho digitado por engano (ex.: Downloads) seja esvaziado.
    managed = marker.exists() or (root / REPORT_FILENAME).exists()
    meaningful_children = [c for c in children if c.name != marker.name]
    if meaningful_children and not managed:
        raise RuntimeError(
            "A pasta de saída escolhida já contém arquivos que não foram identificados como resultado anterior do extrator. "
            "Para proteger seus documentos, escolha outro nome de pasta ou uma pasta vazia."
        )

    for child in meaningful_children:
        try:
            if child.is_dir():
                shutil.rmtree(child)
            else:
                child.unlink()
        except OSError as exc:
            raise RuntimeError(f"Não foi possível limpar a saída anterior: {child.name}. Feche o arquivo e tente novamente. ({exc})") from exc
    try:
        marker.write_text("EDGE Extrator de Exames - pasta gerenciada pelo programa\n", encoding="utf-8")
    except OSError:
        pass
    return root


def unique_file(path: Path) -> Path:
    if not path.exists():
        return path
    n = 2
    while True:
        p = path.with_name(f"{path.stem} ({n}){path.suffix}")
        if not p.exists():
            return p
        n += 1


def build_filename(exam: str, name: str, company: str, cnpj: str, exam_subtype: str = "", include_aso_subtype: bool = False, receipt: str = "") -> str:
    """Monta o nome do PDF.

    O motivo do ASO (admissional, periódico, mudança de riscos etc.) só entra
    no nome quando realmente é necessário diferenciar duas ou mais ocorrências
    de ASO previstas para o mesmo funcionário.
    """
    exam_label = exam
    if normalize_exam(exam) == "ASO" and include_aso_subtype and clean_value(exam_subtype):
        exam_label = f"ASO - {clean_value(exam_subtype)}"
    e = safe_component(exam_label, "EXAME", 58)
    n = safe_component(name, "NOME NAO IDENTIFICADO", 65)
    r = safe_component(clean_value(receipt), "", 38) if clean_value(receipt) else ""
    c = safe_component(company, "EMPRESA NAO IDENTIFICADA", 85)
    j = safe_component((format_company_document(cnpj) or "CPF-CNPJ NAO IDENTIFICADO").replace("/", "-"), "CPF-CNPJ NAO IDENTIFICADO", 24)
    receipt_part = f" - {r}" if r else ""
    return f"{e} - {n}{receipt_part} - {c} {j}.pdf"


def save_page(source_pdf: Path, page_index: int, target: Path) -> None:
    target.parent.mkdir(parents=True, exist_ok=True)
    doc = fitz.open(source_pdf)
    try:
        out = fitz.open()
        try:
            out.insert_pdf(doc, from_page=page_index, to_page=page_index)
            out.save(target, garbage=4, deflate=True)
        finally:
            out.close()
    finally:
        doc.close()


def append_page_to_pdf(target: Path, source_doc: fitz.Document, page_index: int) -> None:
    """Anexa uma página a um PDF já salvo, usando arquivo temporário seguro."""
    existing = fitz.open(target)
    try:
        existing.insert_pdf(source_doc, from_page=page_index, to_page=page_index)
        temp = target.with_name(target.stem + ".__tmp__.pdf")
        existing.save(temp, garbage=4, deflate=True)
    finally:
        existing.close()
    os.replace(temp, target)


def file_page_hash(source_pdf: Path, page_index: int) -> str:
    doc = fitz.open(source_pdf)
    try:
        page = doc[page_index]
        pix = page.get_pixmap(matrix=fitz.Matrix(0.7, 0.7), alpha=False)
        return hashlib.sha1(pix.samples).hexdigest()
    finally:
        doc.close()


def _save_analysis_page(doc: fitz.Document, page_index: int, target: Path) -> None:
    target.parent.mkdir(parents=True, exist_ok=True)
    out = fitz.open()
    try:
        out.insert_pdf(doc, from_page=page_index, to_page=page_index)
        out.save(target, garbage=4, deflate=True)
    finally:
        out.close()



def sync_to_all_files(root: Path, exam_file: Path) -> Path:
    """Compatibilidade: a saída agora já é a própria pasta unificada."""
    return exam_file


def _expected_assignment_count(employees: list[ExpectedEmployee]) -> int:
    return sum(e.expected_total for e in employees)

def _missing_assignments(employees: list[ExpectedEmployee], counts: Counter[tuple[str, str]]) -> list[tuple[ExpectedEmployee, str]]:
    missing: list[tuple[ExpectedEmployee, str]] = []
    for emp in employees:
        for exam in sorted(emp.expected_exams):
            need = max(0, emp.expected_count(exam) - counts[(emp.key, exam)])
            missing.extend([(emp, exam)] * need)
    return missing


def _rescue_missing_pages(
    pdf_paths: list[Path], employees: list[ExpectedEmployee], models: list[ModelSample], root: Path,
    analyses: list[PageAnalysis], seen_assignment_counts: Counter[tuple[str, str]],
    seen_page_hashes: set[str], assignment_last: dict[tuple[str, str], tuple[str, int, Path]],
    employee_index: EmployeeIndex, use_ocr: bool, auto_threshold: float, employee_threshold: float,
    progress: Optional[Callable[[int,int,str],None]] = None,
) -> int:
    """Segunda varredura fail-safe executada SOMENTE se ainda houver exames faltando.

    Diferente da triagem rápida, esta etapa não descarta páginas por heurística de velocidade:
    lê a página inteira (texto nativo e, quando necessário, OCR), procura apenas os tipos ainda
    faltantes e reavalia funcionário + empresa/CNPJ. O objetivo é impedir falso "NÃO ENCONTRADO".
    """
    missing = _missing_assignments(employees, seen_assignment_counts)
    if not missing:
        return 0
    missing_types = {exam for _, exam in missing}
    saved_page_keys = {(a.source_pdf_path, a.page_index) for a in analyses if a.status in {"SALVO_AUTOMATICO", "SALVO_MANUAL", "ANEXADO_CONTINUACAO"}}
    by_page = {(a.source_pdf_path, a.page_index): a for a in analyses}
    rescued = 0
    total = sum(fitz.open(p).page_count for p in pdf_paths)
    step = 0

    for pdf in pdf_paths:
        doc = fitz.open(pdf)
        try:
            for idx in range(doc.page_count):
                step += 1
                if not _missing_assignments(employees, seen_assignment_counts):
                    return rescued
                key_page = (str(pdf), idx)
                if key_page in saved_page_keys:
                    continue
                if progress:
                    progress(step, total, f"Varredura de segurança: {pdf.name} - página {idx+1}")
                page = doc[idx]
                embedded = page.get_text("text").strip()
                # Texto nativo completo é preferido. Em página escaneada/ruim, OCR da folha inteira.
                if _embedded_text_is_good(embedded) or (embedded and not _page_has_large_raster(page)):
                    text, text_source, warning = embedded, "Texto interno (varredura de segurança)", ""
                else:
                    text, text_source, warning = page_text(page, use_ocr=use_ocr)
                if not text.strip():
                    continue

                current_missing = _missing_assignments(employees, seen_assignment_counts)
                current_types = {exam for _, exam in current_missing}
                # Primeiro identifica o tipo REAL da página entre todos os tipos conhecidos.
                # Isso impede confundir um ASO com Acuidade/Audiometria apenas porque o ASO
                # lista esses exames na seção "Exames complementares".
                global_exam, global_conf, global_breakdown, global_ignore = classify_page(
                    text, None, models, None, candidate_types=set(EXAM_TYPES), recognition_floor=18.0
                )
                if global_exam and global_exam not in current_types and global_conf >= 60:
                    continue
                structural = structural_exam_candidates(text, current_types)
                exam, exam_conf, breakdown, hard_ignore = classify_page(
                    text, None, models, None, candidate_types=current_types, recognition_floor=18.0
                )
                current_has_models = any(m.exam_type in current_types for m in models)
                # Na varredura de segurança sempre usa o modelo visual quando houver
                # modelo cadastrado para o tipo faltante. Isso corrige PDFs cujo OCR lê
                # pouco texto, mas o formulário é igual ao modelo enviado.
                if current_has_models and (not exam or exam_conf < max(50.0, auto_threshold - 25.0) or structural):
                    img = render_page_image(page, scale=0.85)
                    exam2, conf2, breakdown2, ignore2 = classify_page(
                        text, img, models, None, candidate_types=current_types, recognition_floor=14.0
                    )
                    if conf2 > exam_conf or not exam:
                        exam, exam_conf, breakdown, hard_ignore = exam2, conf2, breakdown2, ignore2
                if not exam and structural:
                    # Estrutura genérica sem pontuação suficiente: mantém como possível tipo
                    # para cruzar com funcionário e mandar para conferência, em vez de sumir.
                    exam = sorted(structural)[0]
                    exam_conf = max(exam_conf, 35.0)
                    breakdown = (breakdown + "; " if breakdown else "") + "estrutura=35"
                    hard_ignore = False
                if not exam:
                    continue

                extracted_name, extracted_company, extracted_cnpj, extracted_cpf = extract_fields(text, exam)
                employee, emp_conf, emp_reason = match_employee(text, extracted_name, extracted_cpf, employees, employee_index)
                if employee is None or exam not in employee.expected_exams:
                    continue
                assignment = (employee.key, exam)
                if seen_assignment_counts[assignment] >= employee.expected_count(exam):
                    continue

                # No resgate exigimos evidência combinada. Nome/CNPJ forte pode compensar pequenas
                # falhas de OCR, mas página auxiliar nunca é salva apenas por proximidade textual.
                ntext = normalize_for_match(text)
                text_digits = digits_only(text)
                exact_name = normalize_for_match(employee.name) in ntext
                exact_company_doc = bool(digits_only(employee.cnpj) and digits_only(employee.cnpj) in text_digits)
                valid_cpf_hit = bool(is_valid_cpf(employee.cpf) and digits_only(employee.cpf) in text_digits)
                strong_identity = (exact_name and exact_company_doc) or valid_cpf_hit or emp_conf >= max(88.0, employee_threshold)
                strong_exam = exam_conf >= max(45.0, auto_threshold - 35.0)
                if global_exam == exam and global_conf >= max(45.0, auto_threshold - 35.0):
                    strong_exam = True
                # Para tipos personalizados com modelo cadastrado, aceita confiança menor
                # quando a identidade do funcionário é forte. Assim exames novos não ficam
                # como "não encontrado" só porque não têm regra manual no código.
                if strong_identity and exam not in DEFAULT_EXAM_TYPES and any(m.exam_type == exam for m in models) and exam_conf >= 42.0:
                    strong_exam = True
                if hard_ignore and exam_conf < 80:
                    continue

                existing = by_page.get(key_page)
                if not (strong_identity and strong_exam):
                    # Não deixa a página sumir: transforma uma possível correspondência em PENDENTE.
                    if existing is not None and emp_conf >= 60 and exam_conf >= 25:
                        existing.exam_type = exam
                        existing.exam_confidence = exam_conf
                        existing.exam_score_breakdown = breakdown
                        existing.employee_name = employee.name
                        existing.employee_cpf = employee.cpf
                        existing.company = employee.company
                        existing.cnpj = employee.cnpj
                        existing.employee_match_confidence = emp_conf
                        existing.employee_row_id = employee.row_id
                        existing.expected = True
                        existing.status = "PENDENTE"
                        existing.reason = f"Varredura de segurança encontrou possível {exam}; requer conferência. Match: {emp_reason}"
                        existing.raw_text = text
                        existing.text_source = text_source
                        existing.ocr_warning = warning
                    continue

                hash_img = render_page_image(page, scale=0.40)
                page_digest = hashlib.sha1(hash_img.tobytes()).hexdigest()
                if page_digest in seen_page_hashes:
                    continue
                slot = seen_assignment_counts[assignment]
                receipt = employee.expected_receipt(exam, slot)
                subtype = detect_aso_subtype(text) if exam == "ASO" else ""
                target = unique_file(root / build_filename(
                    exam, employee.name, employee.company, employee.cnpj, subtype,
                    include_aso_subtype=(exam == "ASO" and employee.expected_count(exam) > 1), receipt=receipt,
                ))
                _save_analysis_page(doc, idx, target)
                sync_to_all_files(root, target)
                rel = str(target.relative_to(root))
                seen_assignment_counts[assignment] += 1
                seen_page_hashes.add(page_digest)
                assignment_last[assignment] = (str(pdf), idx, target)
                saved_page_keys.add(key_page)
                rescued += 1

                if existing is None:
                    existing = PageAnalysis(
                        id=uuid.uuid4().hex, source_pdf=pdf.name, source_pdf_path=str(pdf), page_index=idx, page_number=idx+1,
                        raw_text=text, text_source=text_source, ocr_warning=warning, exam_type=exam, exam_confidence=exam_conf,
                        exam_score_breakdown=breakdown, employee_name=employee.name, employee_cpf=employee.cpf,
                        company=employee.company, cnpj=employee.cnpj, employee_match_confidence=emp_conf,
                        employee_row_id=employee.row_id, expected=True, status="SALVO_AUTOMATICO",
                        reason=f"Recuperado pela varredura de segurança. Match: {emp_reason}", output_file=rel,
                        page_hash=page_digest, exam_subtype=subtype, receipt=receipt,
                    )
                    analyses.append(existing); by_page[key_page] = existing
                else:
                    existing.raw_text = text; existing.text_source = text_source; existing.ocr_warning = warning
                    existing.exam_type = exam; existing.exam_confidence = exam_conf; existing.exam_score_breakdown = breakdown
                    existing.employee_name = employee.name; existing.employee_cpf = employee.cpf
                    existing.company = employee.company; existing.cnpj = employee.cnpj
                    existing.employee_match_confidence = emp_conf; existing.employee_row_id = employee.row_id
                    existing.expected = True; existing.status = "SALVO_AUTOMATICO"
                    existing.reason = f"Recuperado pela varredura de segurança. Match: {emp_reason}"
                    existing.output_file = rel; existing.page_hash = page_digest
                    existing.exam_subtype = subtype; existing.receipt = receipt
        finally:
            doc.close()
    return rescued


def process_pdfs(
    pdf_paths: list[Path], employees: list[ExpectedEmployee], output_base: Path,
    use_ocr: bool = True, auto_threshold: float = 80.0, employee_threshold: float = 82.0,
    progress: Optional[Callable[[int,int,str],None]] = None,
    fast_mode: bool = True, stop_when_complete: bool = True,
) -> ProcessingSummary:
    if not pdf_paths:
        raise ValueError("Nenhum PDF selecionado")
    # Confiabilidade máxima: nunca encerra a leitura antes do fim do lote.
    # Também desativa a triagem rápida: ela era boa para velocidade, mas podia
    # descartar páginas de exames digitalizados antes da leitura completa.
    stop_when_complete = False
    fast_mode = False
    models = load_models()
    root = output_root(output_base)

    total_pages = 0
    pages_per_pdf: dict[str, int] = {}
    for p in pdf_paths:
        d = fitz.open(p)
        try:
            total_pages += d.page_count
            pages_per_pdf[str(p)] = d.page_count
        finally:
            d.close()

    globally_expected = {exam for emp in employees for exam in emp.expected_exams}
    expected_total = _expected_assignment_count(employees)
    pcd_expected = "LAUDO PCD" in globally_expected
    employee_index = build_employee_index(employees)

    analyses: list[PageAnalysis] = []
    seen_assignment_counts: Counter[tuple[str, str]] = Counter()
    seen_page_hashes: set[str] = set()
    assignment_last: dict[tuple[str, str], tuple[str, int, Path]] = {}
    current = 0
    quick_scanned = 0
    detailed_ocr_pages = 0
    skipped_after_complete = 0
    fast_rejected = 0
    stop_all = False

    def completed_count() -> int:
        return sum(seen_assignment_counts.values())

    def record_fast_ignore(
        pdf: Path, idx: int, text: str, text_source: str, ocr_warning: str,
        reason: str, exam: str = "", exam_conf: float = 0.0, breakdown: str = "",
        extracted_name: str = "", extracted_company: str = "", extracted_cnpj: str = "", extracted_cpf: str = "",
        employee: ExpectedEmployee | None = None, emp_conf: float = 0.0, emp_reason: str = "",
    ) -> None:
        """Registra o descarte sem hash, renderização extra ou extração pesada."""
        name = employee.name if employee else extracted_name
        cpf = employee.cpf if employee and employee.cpf else extracted_cpf
        company = employee.company if employee and employee.company else extracted_company
        cnpj = employee.cnpj if employee and employee.cnpj else extracted_cnpj
        analyses.append(PageAnalysis(
            id=uuid.uuid4().hex, source_pdf=pdf.name, source_pdf_path=str(pdf), page_index=idx, page_number=idx+1,
            raw_text=text, text_source=text_source, ocr_warning=ocr_warning, exam_type=exam, exam_confidence=exam_conf,
            exam_score_breakdown=breakdown, employee_name=name, employee_cpf=cpf, company=company, cnpj=cnpj,
            employee_match_confidence=emp_conf, employee_row_id=employee.row_id if employee else None,
            expected=False, status="IGNORADO", reason=f"{reason}. Match: {emp_reason}" if emp_reason else reason,
            output_file="", page_hash="",
        ))

    # O Tesseract já usa múltiplos recursos internamente. Várias instâncias em
    # paralelo ficaram mais lentas e, em alguns computadores, podiam travar em
    # páginas grandes. A triagem rápida agora é sequencial e limitada por timeout:
    # menos picos de CPU/RAM e tempo muito mais previsível.
    ocr_workers = 1
    batch_size = 4
    quick_executor = None

    def prepare_quick_batch(doc: fitz.Document, start_idx: int) -> dict[int, tuple[str, str, str, Image.Image | None, bool]]:
        cache: dict[int, tuple[str, str, str, Image.Image | None, bool]] = {}
        end_idx = min(doc.page_count, start_idx + batch_size)
        scanned_jobs: list[tuple[int, str, Image.Image]] = []
        for j in range(start_idx, end_idx):
            pg = doc[j]
            embedded = pg.get_text("text").strip()
            # PDFs digitais confiáveis não entram no OCR.
            if _embedded_text_is_good(embedded) or (embedded and not _page_has_large_raster(pg)) or not use_ocr:
                cache[j] = (embedded, "Texto interno (triagem)", "", None, True)
                continue
            img = render_page_top_image(pg, scale=0.95, top_fraction=0.45)
            scanned_jobs.append((j, embedded, img))

        if scanned_jobs:
            batch_texts, batch_warning = ocr_images_batch(
                [job[2] for job in scanned_jobs], psm=11,
                timeout_seconds=max(12.0, 5.0 * len(scanned_jobs)),
            )
            for (j, embedded, _img), ocr in zip(scanned_jobs, batch_texts):
                text = ((embedded + "\n" + ocr).strip() if embedded and ocr.strip() else (ocr.strip() or embedded))
                cache[j] = (
                    text,
                    "OCR rápido em lote (cabeçalho)" if ocr.strip() else "Texto interno (triagem)",
                    batch_warning,
                    None,
                    False,
                )
        return cache

    try:
        for pdf_index, pdf in enumerate(pdf_paths):
            if stop_all:
                skipped_after_complete += pages_per_pdf.get(str(pdf), 0)
                continue
            doc = fitz.open(pdf)
            quick_cache: dict[int, tuple[str, str, str, Image.Image | None, bool]] = {}
            try:
                for idx in range(doc.page_count):
                    if fast_mode and idx not in quick_cache:
                        quick_cache = prepare_quick_batch(doc, idx)
                    if stop_all:
                        skipped_after_complete += doc.page_count - idx
                        break
                    current += 1
                    page = doc[idx]
                    if progress:
                        progress(current, total_pages, f"Triagem rápida: {pdf.name} - página {idx+1}")

                    # -------- ETAPA 0: DESCARTE ULTRARRÁPIDO --------
                    # Lê primeiro somente a camada de texto do PDF. Quando ela é confiável
                    # e não há qualquer evidência de um exame solicitado, a página termina
                    # aqui: sem OCR, sem renderização e sem cálculo de hash.
                    cached_text = cached_source = cached_warning = ""
                    cached_img = None
                    cached_digital = False
                    if fast_mode:
                        cached_text, cached_source, cached_warning, cached_img, cached_digital = quick_cache.get(idx, ("", "", "", None, False))
                    embedded_fast = page.get_text("text").strip() if (fast_mode and cached_digital) else (page.get_text("text").strip() if not fast_mode else "")
                    prev_is_pcd = bool(
                        pcd_expected and analyses and analyses[-1].source_pdf_path == str(pdf)
                        and analyses[-1].page_index == idx - 1 and analyses[-1].exam_type == "LAUDO PCD"
                        and analyses[-1].employee_row_id
                    )

                    # -------- ETAPA 1: TRIAGEM RÁPIDA --------
                    if fast_mode:
                        probe_exam = ""; probe_conf = 0.0; probe_breakdown = ""; probe_ignore = False
                        can_trust_embedded = False
                        if embedded_fast:
                            probe_exam, probe_conf, probe_breakdown, probe_ignore = classify_page(
                                embedded_fast, None, models, None,
                                candidate_types=globally_expected, recognition_floor=14.0,
                            )
                            # Texto longo é confiável por si só. Texto curto também é confiável
                            # em PDF digital quando não existe uma grande imagem ocupando a página.
                            can_trust_embedded = _embedded_text_is_good(embedded_fast)
                            if not can_trust_embedded and not prev_is_pcd:
                                can_trust_embedded = not _page_has_large_raster(page)

                            if probe_ignore and not prev_is_pcd:
                                record_fast_ignore(
                                    pdf, idx, embedded_fast, "Texto interno (descarte instantâneo)", "",
                                    "Documento auxiliar/comprovante/encaminhamento descartado antes do OCR",
                                    probe_exam, probe_conf, probe_breakdown,
                                )
                                fast_rejected += 1
                                continue
                            if can_trust_embedded and not probe_exam and not prev_is_pcd:
                                record_fast_ignore(
                                    pdf, idx, embedded_fast, "Texto interno (descarte instantâneo)", "",
                                    "Página digital sem evidência de nenhum exame solicitado; análise pesada ignorada",
                                    "", probe_conf, probe_breakdown,
                                )
                                fast_rejected += 1
                                continue

                        # Se a camada interna já identificou um exame, não roda OCR preliminar.
                        # Para páginas escaneadas/ambíguas, usa OCR leve de baixa resolução.
                        if embedded_fast and probe_exam and can_trust_embedded:
                            text, text_source, ocr_warning, quick_img = embedded_fast, "Texto interno (triagem)", "", None
                        else:
                            if cached_text or cached_img is not None:
                                text, text_source, ocr_warning, quick_img = cached_text, cached_source, cached_warning, cached_img
                            else:
                                text, text_source, ocr_warning, quick_img = quick_page_text(
                                    page, use_ocr=use_ocr, embedded=embedded_fast
                                )
                        quick_scanned += 1
                        image_for_class = quick_img
                        extracted_name, extracted_company, extracted_cnpj, extracted_cpf = extract_fields(text, "ASO")
                        employee, emp_conf, emp_reason = match_employee(text, extracted_name, extracted_cpf, employees, employee_index)
                        hint = employee.expected_exams if employee else None

                        # Primeiro classifica só por texto/modelos textuais. Só renderiza imagem
                        # para comparação visual quando o texto não foi suficiente.
                        exam, exam_conf, breakdown, hard_ignore = classify_page(
                            text, image_for_class, models, hint,
                            candidate_types=globally_expected, recognition_floor=18.0,
                        )
                        if models and image_for_class is None and not hard_ignore and (not exam or exam_conf < auto_threshold):
                            image_for_class = render_page_image(page, scale=0.72)
                            exam, exam_conf, breakdown, hard_ignore = classify_page(
                                text, image_for_class, models, hint,
                                candidate_types=globally_expected, recognition_floor=18.0,
                            )

                        # Página com evidência suficiente para um exame realmente esperado pode
                        # ser concluída direto pela triagem. Antes de usar OCR da página inteira,
                        # fazemos uma SEGUNDA TRIAGEM apenas na região de identificação. Essa é a
                        # proteção contra o falso descarte da v3.0 sem voltar à lentidão antiga.
                        expected_quick = bool(employee and exam and exam in employee.expected_exams)
                        quick_secure = bool(
                            expected_quick and emp_conf >= employee_threshold and exam_conf >= auto_threshold and not hard_ignore
                        )
                        structural_candidates = structural_exam_candidates(text, globally_expected)
                        structural_rescue = bool(structural_candidates)

                        if structural_rescue and not quick_secure and use_ocr:
                            # OCR intermediário: lê somente ~45% do topo, em resolução suficiente
                            # para nomes/CPF/CNPJ/título. É muito mais barato que OCR completo.
                            id_img = render_page_top_image(page, scale=1.15, top_fraction=0.45)
                            id_ocr, id_warning = ocr_image(id_img, psm=6, timeout_seconds=10)
                            if id_ocr.strip():
                                text = ((embedded_fast + "\n" + id_ocr).strip() if embedded_fast else id_ocr.strip())
                                text_source = "OCR de identificação (triagem segura)"
                                if id_warning:
                                    ocr_warning = id_warning
                                extracted_name, extracted_company, extracted_cnpj, extracted_cpf = extract_fields(text, "ASO")
                                employee, emp_conf, emp_reason = match_employee(
                                    text, extracted_name, extracted_cpf, employees, employee_index
                                )
                                hint = employee.expected_exams if employee else None
                                exam, exam_conf, breakdown, hard_ignore = classify_page(
                                    text, None, models, hint,
                                    candidate_types=globally_expected, recognition_floor=18.0,
                                )
                                # A estrutura também é recalculada com a leitura mais nítida.
                                structural_candidates = structural_exam_candidates(text, globally_expected)
                                structural_rescue = bool(structural_candidates)
                                expected_quick = bool(employee and exam and exam in employee.expected_exams)
                                quick_secure = bool(
                                    expected_quick and emp_conf >= employee_threshold and exam_conf >= auto_threshold and not hard_ignore
                                )

                        need_detailed = False
                        if not quick_secure:
                            if not hard_ignore and employee and emp_conf >= 76 and (exam or structural_rescue):
                                # Só aprofunda quando também existe sinal do TIPO solicitado.
                                # Isso evita OCR completo em audiometria/acuidade de alguém cuja
                                # planilha pede apenas ASO, por exemplo.
                                need_detailed = True
                            elif not hard_ignore and exam and exam_conf >= 48 and emp_conf >= 45:
                                need_detailed = True
                            elif (
                                not hard_ignore and employee and emp_conf >= 76
                                and any(x != "ASO" for x in employee.expected_exams)
                            ):
                                # Se o nome do funcionário foi reconhecido com boa segurança e a
                                # planilha espera exame complementar, não descarte a página apenas
                                # porque o título pequeno sumiu no OCR rápido. É comum em laudos de
                                # espirometria o cabeçalho ser lido como apenas "Laudo Médico".
                                need_detailed = True
                            elif structural_rescue:
                                # A estrutura do documento já indica que pode ser um exame solicitado.
                                # Mesmo que o OCR rápido tenha lido mal o nome do funcionário, fazemos
                                # uma leitura detalhada antes de descartar. Isso corrige falsos "NÃO
                                # ENCONTRADO" em audiometrias/espirometrias escaneadas com texto pequeno.
                                # O arquivo só será salvo depois que a leitura detalhada confirmar também
                                # o funcionário e os níveis de confiança normais do sistema.
                                need_detailed = True

                        # Quando a triagem já prova que a página não interessa, encerra aqui.
                        # Isso evita inclusive o hash de duplicidade das páginas irrelevantes.
                        if hard_ignore and not prev_is_pcd and not structural_rescue:
                            record_fast_ignore(
                                pdf, idx, text, text_source, ocr_warning,
                                "Documento auxiliar/comprovante/encaminhamento identificado na triagem rápida",
                                exam, exam_conf, breakdown, extracted_name, extracted_company, extracted_cnpj, extracted_cpf,
                                employee, emp_conf, emp_reason,
                            )
                            fast_rejected += 1
                            continue
                        if not exam and not need_detailed and not prev_is_pcd:
                            record_fast_ignore(
                                pdf, idx, text, text_source, ocr_warning,
                                "Triagem rápida não encontrou evidência de um dos exames solicitados",
                                "", exam_conf, breakdown, extracted_name, extracted_company, extracted_cnpj, extracted_cpf,
                                employee, emp_conf, emp_reason,
                            )
                            fast_rejected += 1
                            continue

                        if need_detailed and use_ocr:
                            if progress:
                                progress(current, total_pages, f"OCR detalhado: {pdf.name} - página {idx+1}")
                            detailed_ocr_pages += 1
                            text, text_source, ocr_warning = page_text(page, use_ocr=True)
                            extracted_name, extracted_company, extracted_cnpj, extracted_cpf = extract_fields(text, "ASO")
                            employee, emp_conf, emp_reason = match_employee(text, extracted_name, extracted_cpf, employees, employee_index)
                            hint = employee.expected_exams if employee else None
                            # Reaproveita imagem rápida; modelos visuais não precisam de alta resolução.
                            if image_for_class is None and models:
                                image_for_class = render_page_image(page, scale=0.72)
                            exam, exam_conf, breakdown, hard_ignore = classify_page(
                                text, image_for_class, models, hint,
                                candidate_types=globally_expected, recognition_floor=28.0,
                            )
                            if hard_ignore and not prev_is_pcd:
                                record_fast_ignore(
                                    pdf, idx, text, text_source, ocr_warning,
                                    "Documento descartado após OCR de confirmação",
                                    exam, exam_conf, breakdown, extracted_name, extracted_company, extracted_cnpj, extracted_cpf,
                                    employee, emp_conf, emp_reason,
                                )
                                fast_rejected += 1
                                continue
                            if not exam and not prev_is_pcd:
                                record_fast_ignore(
                                    pdf, idx, text, text_source, ocr_warning,
                                    "OCR de confirmação não encontrou exame solicitado; página descartada sem processamento adicional",
                                    "", exam_conf, breakdown, extracted_name, extracted_company, extracted_cnpj, extracted_cpf,
                                    employee, emp_conf, emp_reason,
                                )
                                fast_rejected += 1
                                continue
                    else:
                        if progress:
                            progress(current, total_pages, f"OCR completo: {pdf.name} - página {idx+1}")
                        text, text_source, ocr_warning = page_text(page, use_ocr=use_ocr)
                        detailed_ocr_pages += 1 if use_ocr else 0
                        extracted_name, extracted_company, extracted_cnpj, extracted_cpf = extract_fields(text, "ASO")
                        employee, emp_conf, emp_reason = match_employee(text, extracted_name, extracted_cpf, employees, employee_index)
                        hint = employee.expected_exams if employee else None
                        img = render_page_image(page, scale=0.85) if models else None
                        exam, exam_conf, breakdown, hard_ignore = classify_page(
                            text, img, models, hint, candidate_types=globally_expected, recognition_floor=28.0
                        )

                    # Extrai novamente com o tipo reconhecido.
                    if exam:
                        extracted_name2, extracted_company2, extracted_cnpj2, extracted_cpf2 = extract_fields(text, exam)
                        employee2, emp_conf2, emp_reason2 = match_employee(
                            text, extracted_name2 or extracted_name, extracted_cpf2 or extracted_cpf, employees, employee_index
                        )
                        if employee2 and emp_conf2 >= emp_conf:
                            employee, emp_conf, emp_reason = employee2, emp_conf2, emp_reason2
                            extracted_name, extracted_company, extracted_cnpj, extracted_cpf = extracted_name2, extracted_company2, extracted_cnpj2, extracted_cpf2

                    # Continuação de Laudo PCD.
                    if exam == "LAUDO PCD" and (employee is None or emp_conf < employee_threshold) and analyses:
                        prev = analyses[-1]
                        if prev.source_pdf_path == str(pdf) and prev.page_index == idx - 1 and prev.exam_type == "LAUDO PCD" and prev.employee_row_id:
                            inherited = employees_by_row(employees, prev.employee_row_id)
                            if inherited:
                                employee = inherited
                                emp_conf = max(emp_conf, 90.0)
                                emp_reason = "continuação do Laudo PCD anterior"

                    name = employee.name if employee else extracted_name
                    cpf = employee.cpf if employee and employee.cpf else extracted_cpf
                    company = employee.company if employee and employee.company else extracted_company
                    cnpj = employee.cnpj if employee and employee.cnpj else extracted_cnpj
                    expected = bool(employee and exam and exam in employee.expected_exams)
                    exam_subtype = detect_aso_subtype(text) if exam == "ASO" else ""

                    # A impressão digital visual só é calculada quando a página realmente
                    # pode ser salva. Páginas irrelevantes não gastam renderização/hash.
                    page_digest = ""
                    status = "IGNORADO"; reason = "Documento não reconhecido como exame desejado"
                    output_file = ""
                    receipt = ""
                    if employee and exam and expected:
                        receipt = employee.expected_receipt(exam, seen_assignment_counts[(employee.key, exam)])
                    review_exam_floor = 35.0
                    if hard_ignore:
                        status = "IGNORADO"; reason = "Documento auxiliar/comprovante/encaminhamento identificado na triagem"
                    elif not exam:
                        status = "IGNORADO"; reason = "Triagem não encontrou evidência de um dos exames solicitados"
                    elif not employee:
                        status = "IGNORADO"; reason = f"{exam} identificado, mas não pertence a funcionário solicitado na lista"
                    elif not expected:
                        status = "IGNORADO"; reason = f"{exam} de {employee.name} não foi solicitado na lista"
                    elif emp_conf < 68.0:
                        status = "IGNORADO"; reason = f"Vínculo com funcionário muito fraco ({emp_conf:.0f}%); descartado para evitar pendência inútil"
                    elif exam_conf < review_exam_floor:
                        status = "PENDENTE"; reason = f"Possível {exam} solicitado para {employee.name}, mas a evidência do tipo ficou fraca ({exam_conf:.0f}%). Conferir manualmente."
                    elif emp_conf < employee_threshold:
                        status = "PENDENTE"; reason = f"Possível {exam} solicitado para {employee.name}, mas funcionário precisa de confirmação ({emp_conf:.0f}%)"
                    elif exam_conf < auto_threshold:
                        status = "PENDENTE"; reason = f"{exam} solicitado para {employee.name}, mas o tipo precisa de confirmação ({exam_conf:.0f}%)"
                    else:
                        # Hash determinístico da imagem completa. Não inclui o texto OCR, pois o
                        # OCR pode variar levemente entre duas leituras da MESMA página e gerar
                        # falsos "documentos diferentes".
                        hash_img = render_page_image(page, scale=0.40)
                        page_digest = hashlib.sha1(hash_img.tobytes()).hexdigest()
                        assignment = (employee.key, exam)
                        previous_assignment = assignment_last.get(assignment)
                        is_pcd_continuation = bool(
                            exam == "LAUDO PCD" and previous_assignment and previous_assignment[0] == str(pdf)
                            and previous_assignment[1] == idx - 1
                        )
                        expected_slots = employee.expected_count(exam)
                        saved_slots = seen_assignment_counts[assignment]
                        receipt = employee.expected_receipt(exam, saved_slots)
                        if page_digest in seen_page_hashes:
                            status = "IGNORADO"; reason = "Página idêntica já processada; duplicata real descartada automaticamente"
                        elif is_pcd_continuation:
                            target = previous_assignment[2]
                            append_page_to_pdf(target, doc, idx)
                            sync_to_all_files(root, target)
                            status = "ANEXADO_CONTINUACAO"; reason = "Página anexada ao Laudo PCD anterior"
                            output_file = str(target.relative_to(root))
                            assignment_last[assignment] = (str(pdf), idx, target)
                            seen_page_hashes.add(page_digest)
                        elif expected_slots > 0 and saved_slots >= expected_slots:
                            status = "DUPLICADO"
                            reason = f"Quantidade esperada de {exam} já atingida ({saved_slots}/{expected_slots}) para este funcionário"
                        else:
                            status = "SALVO_AUTOMATICO"
                            reason = "Correspondência segura com a lista esperada"
                            if expected_slots > 1:
                                reason += f"; ocorrência {saved_slots + 1} de {expected_slots}"
                            target = unique_file(root / build_filename(
                                exam, name, company, cnpj, exam_subtype,
                                include_aso_subtype=(exam == "ASO" and expected_slots > 1), receipt=receipt,
                            ))
                            _save_analysis_page(doc, idx, target)
                            sync_to_all_files(root, target)
                            output_file = str(target.relative_to(root))
                            seen_assignment_counts[assignment] += 1
                            seen_page_hashes.add(page_digest)
                            assignment_last[assignment] = (str(pdf), idx, target)

                    if status in {"PENDENTE", "DUPLICADO"}:
                        # Pendências permanecem disponíveis para revisão pela própria interface,
                        # mas não criam PDFs extras na pasta final.
                        output_file = ""

                    analyses.append(PageAnalysis(
                        id=uuid.uuid4().hex, source_pdf=pdf.name, source_pdf_path=str(pdf), page_index=idx, page_number=idx+1,
                        raw_text=text, text_source=text_source, ocr_warning=ocr_warning, exam_type=exam, exam_confidence=exam_conf,
                        exam_score_breakdown=breakdown, employee_name=name, employee_cpf=cpf, company=company, cnpj=cnpj,
                        employee_match_confidence=emp_conf, employee_row_id=employee.row_id if employee else None,
                        expected=expected, status=status, reason=f"{reason}. Match: {emp_reason}" if emp_reason else reason,
                        output_file=output_file, page_hash=page_digest, exam_subtype=exam_subtype, receipt=(receipt if employee and exam else ""),
                    ))

                    # Quando tudo pedido já foi encontrado, não continua gastando OCR em páginas
                    # que não podem mudar o resultado. Laudo PCD é exceção por poder continuar em
                    # páginas seguintes.
                    if stop_when_complete and expected_total and completed_count() >= expected_total and not pcd_expected:
                        stop_all = True
            finally:
                doc.close()

    finally:
        if quick_executor is not None:
            quick_executor.shutdown(wait=True, cancel_futures=False)

    # FAIL-SAFE: se qualquer exame continuar faltando, varre novamente TODAS as páginas
    # sem os atalhos do modo turbo. Essa etapa prioriza confiabilidade sobre velocidade.
    rescued_pages = _rescue_missing_pages(
        pdf_paths, employees, models, root, analyses, seen_assignment_counts, seen_page_hashes,
        assignment_last, employee_index, use_ocr, auto_threshold, employee_threshold, progress,
    )

    completed_counts: Counter[tuple[str, str]] = Counter()
    for a in analyses:
        if a.status not in {"SALVO_AUTOMATICO", "SALVO_MANUAL"} or not a.employee_row_id or not a.exam_type:
            continue
        emp = employees_by_row(employees, a.employee_row_id)
        if emp:
            completed_counts[(emp.key, a.exam_type)] += 1

    missing_rows: list[dict] = []
    for emp in employees:
        for exam in sorted(emp.expected_exams):
            expected_count = emp.expected_count(exam)
            done_count = completed_counts[(emp.key, exam)]
            missing_count = max(0, expected_count - done_count)
            if not missing_count:
                continue
            candidates = [a for a in analyses if a.employee_row_id == emp.row_id and a.exam_type == exam and a.status in {"PENDENTE", "DUPLICADO"}]
            for slot in range(missing_count):
                candidate = candidates[slot] if slot < len(candidates) else (candidates[0] if candidates else None)
                detail = candidate.reason if candidate else "Nenhuma página correspondente foi localizada"
                if expected_count > 1:
                    detail = f"Esperados {expected_count}; gerados {done_count}. {detail}"
                missing_rows.append({
                    "Funcionario": emp.name, "CPF": emp.cpf, "Empresa": emp.company, "CNPJ": emp.cnpj, "Exame esperado": exam,
                    "Situacao": "PENDENTE DE CONFERENCIA" if candidate else "NAO ENCONTRADO",
                    "Detalhe": detail,
                })

    audited_page_keys = {(a.source_pdf_path, a.page_index) for a in analyses}
    audit_gaps = max(0, total_pages - len(audited_page_keys))
    summary = ProcessingSummary(
        root_dir=str(root), total_pages=total_pages,
        auto_saved=sum(a.status == "SALVO_AUTOMATICO" for a in analyses),
        pending=sum(a.status == "PENDENTE" for a in analyses),
        ignored=sum(a.status == "IGNORADO" for a in analyses),
        unexpected=sum(a.status == "IGNORADO" and bool(a.exam_type) and not a.expected for a in analyses),
        duplicates=sum(a.status == "DUPLICADO" for a in analyses),
        missing_expected=len(missing_rows), analyses=analyses, missing_rows=missing_rows,
        quick_scanned=quick_scanned, detailed_ocr_pages=detailed_ocr_pages,
        skipped_after_complete=skipped_after_complete,
        fast_rejected=fast_rejected,
        rescued_pages=rescued_pages,
        audit_gaps=audit_gaps,
    )
    write_reports(summary)
    return summary

def employees_by_row(employees: list[ExpectedEmployee], row_id: int | None) -> ExpectedEmployee | None:
    if row_id is None:
        return None
    for e in employees:
        if e.row_id == row_id:
            return e
    return None



def recompute_summary(summary: ProcessingSummary, employees: list[ExpectedEmployee]) -> None:
    completed_counts: Counter[tuple[str, str]] = Counter()
    for a in summary.analyses:
        if a.status not in {"SALVO_AUTOMATICO", "SALVO_MANUAL"} or not a.employee_row_id or not a.exam_type:
            continue
        emp = employees_by_row(employees, a.employee_row_id)
        if emp:
            completed_counts[(emp.key, a.exam_type)] += 1

    missing_rows: list[dict] = []
    for emp in employees:
        for exam in sorted(emp.expected_exams):
            expected_count = emp.expected_count(exam)
            done_count = completed_counts[(emp.key, exam)]
            missing_count = max(0, expected_count - done_count)
            if not missing_count:
                continue
            candidates = [a for a in summary.analyses if a.employee_row_id == emp.row_id and a.exam_type == exam and a.status in {"PENDENTE", "DUPLICADO"}]
            for slot in range(missing_count):
                candidate = candidates[slot] if slot < len(candidates) else (candidates[0] if candidates else None)
                detail = candidate.reason if candidate else "Nenhuma página correspondente foi localizada"
                if expected_count > 1:
                    detail = f"Esperados {expected_count}; gerados {done_count}. {detail}"
                missing_rows.append({
                    "Funcionario": emp.name, "CPF": emp.cpf, "Empresa": emp.company, "CNPJ": emp.cnpj, "Exame esperado": exam,
                    "Situacao": "PENDENTE DE CONFERENCIA" if candidate else "NAO ENCONTRADO",
                    "Detalhe": detail,
                })
    summary.missing_rows = missing_rows
    summary.auto_saved = sum(a.status == "SALVO_AUTOMATICO" for a in summary.analyses)
    summary.pending = sum(a.status == "PENDENTE" for a in summary.analyses)
    summary.ignored = sum(a.status in {"IGNORADO","IGNORADO_MANUAL"} for a in summary.analyses)
    summary.unexpected = sum(a.status == "IGNORADO" and bool(a.exam_type) and not a.expected for a in summary.analyses)
    summary.duplicates = sum(a.status == "DUPLICADO" for a in summary.analyses)
    summary.missing_expected = len(missing_rows)


def write_reports(summary: ProcessingSummary) -> None:
    """Gera somente um Excel de conferência dentro da pasta final."""
    root = Path(summary.root_dir)
    if Workbook is None:
        raise RuntimeError("openpyxl não instalado; não foi possível gerar o Excel de conferência")

    saved_statuses = {"SALVO_AUTOMATICO", "SALVO_MANUAL"}

    # Cada PDF efetivamente salvo representa uma ocorrência esperada. Isso preserva
    # duas linhas ASO do mesmo funcionário quando a planilha pede dois ASOs.
    completed: list[PageAnalysis] = [
        a for a in summary.analyses
        if a.status in saved_statuses and a.expected and a.exam_type
    ]

    conference_rows: list[list] = []
    for a in completed:
        subtype_note = f" | Subtipo: {a.exam_subtype}" if a.exam_subtype else ""
        conference_rows.append([
            a.employee_name, a.employee_cpf, a.company, a.cnpj, a.exam_type,
            "SIM", "OK", a.output_file, a.source_pdf, a.page_number,
            f"Tipo {a.exam_confidence:.0f}% / funcionário {a.employee_match_confidence:.0f}%{subtype_note}",
        ])
    for r in summary.missing_rows:
        situation = str(r.get("Situacao", ""))
        result = "PENDENTE" if "PENDENTE" in situation.upper() else "NÃO ENCONTRADO"
        conference_rows.append([
            r.get("Funcionario", ""), r.get("CPF", ""), r.get("Empresa", ""), r.get("CNPJ", ""),
            r.get("Exame esperado", ""), "NÃO", result, "", "", "", r.get("Detalhe", ""),
        ])

    conference_rows.sort(key=lambda row: (normalize_for_match(row[0]), normalize_for_match(row[4])))
    expected_total = len(conference_rows)
    generated_total = sum(1 for row in conference_rows if row[5] == "SIM")
    pending_total = sum(1 for row in conference_rows if row[6] == "PENDENTE")
    missing_total = sum(1 for row in conference_rows if row[6] == "NÃO ENCONTRADO")

    occurrences: list[list] = []
    for a in summary.analyses:
        if a.status in {"PENDENTE", "DUPLICADO"}:
            severity = "ATENÇÃO" if a.status == "PENDENTE" else "AVISO"
            occurrences.append([severity, a.status, a.employee_name, a.employee_cpf, a.exam_type, a.source_pdf, a.page_number, a.reason])
    for r in summary.missing_rows:
        if "PENDENTE" not in str(r.get("Situacao", "")).upper():
            occurrences.append(["ERRO", "NÃO ENCONTRADO", r.get("Funcionario", ""), r.get("CPF", ""), r.get("Exame esperado", ""), "", "", r.get("Detalhe", "")])

    if summary.missing_expected == 0 and getattr(summary, "audit_gaps", 0) == 0:
        general_status = "TUDO CERTO" if not occurrences else "TUDO CERTO COM AVISOS"
        all_generated = "SIM"
    elif summary.missing_expected:
        general_status = "ATENÇÃO - EXISTEM EXAMES NÃO CONCLUÍDOS"
        all_generated = "NÃO"
    else:
        general_status = "ATENÇÃO - EXISTEM PÁGINAS SEM REGISTRO DE AUDITORIA"
        all_generated = "NÃO"

    wb = Workbook()
    resumo = wb.active
    resumo.title = "Resumo"
    resumo.append(["CONFERÊNCIA DA SEPARAÇÃO", ""])
    resumo.append(["Status geral", general_status])
    resumo.append(["Tudo que era esperado foi gerado?", all_generated])
    resumo.append(["Exames esperados", expected_total])
    resumo.append(["Exames gerados", generated_total])
    resumo.append(["Pendentes de conferência", pending_total])
    resumo.append(["Não encontrados", missing_total])
    resumo.append(["Duplicados identificados", summary.duplicates])
    resumo.append(["PDFs/páginas analisados", summary.total_pages])
    resumo.append(["Páginas descartadas sem análise pesada", summary.fast_rejected])
    resumo.append(["Páginas que exigiram OCR detalhado", summary.detailed_ocr_pages])
    resumo.append(["Exames recuperados pela varredura de segurança", getattr(summary, "rescued_pages", 0)])
    resumo.append(["Páginas sem registro de auditoria", getattr(summary, "audit_gaps", 0)])
    resumo.append(["Pasta de saída", str(root)])
    resumo.append(["Gerado em", datetime.now().strftime("%d/%m/%Y %H:%M:%S")])
    resumo.column_dimensions["A"].width = 38
    resumo.column_dimensions["B"].width = 72
    resumo.freeze_panes = "A2"

    conf = wb.create_sheet("Conferência")
    headers = ["Funcionário", "CPF", "Empresa", "CPF/CNPJ Empresa", "Exame esperado", "Gerado?", "Resultado", "Arquivo PDF", "PDF origem", "Página", "Detalhe"]
    conf.append(headers)
    for row in conference_rows:
        conf.append(row)
    conf.freeze_panes = "A2"
    conf.auto_filter.ref = conf.dimensions

    occ = wb.create_sheet("Ocorrências")
    occ_headers = ["Nível", "Ocorrência", "Funcionário", "CPF", "Exame", "PDF origem", "Página", "Detalhe"]
    occ.append(occ_headers)
    if occurrences:
        for row in occurrences:
            occ.append(row)
    else:
        occ.append(["OK", "Nenhuma ocorrência relevante", "", "", "", "", "", "Todos os exames esperados foram concluídos sem pendências."])
    occ.freeze_panes = "A2"
    occ.auto_filter.ref = occ.dimensions

    audit = wb.create_sheet("Auditoria_Paginas")
    audit_headers = ["PDF origem", "Página", "Fonte da leitura", "Status", "Exame identificado", "Funcionário", "CPF", "Empresa", "CPF/CNPJ Empresa", "Esperado?", "Conf. exame", "Conf. funcionário", "Motivo"]
    audit.append(audit_headers)
    for a in sorted(summary.analyses, key=lambda x: (normalize_for_match(x.source_pdf), x.page_number)):
        audit.append([
            a.source_pdf, a.page_number, a.text_source, a.status, a.exam_type, a.employee_name, a.employee_cpf,
            a.company, a.cnpj, "SIM" if a.expected else "NÃO", a.exam_confidence, a.employee_match_confidence, a.reason,
        ])
    audit.freeze_panes = "A2"
    audit.auto_filter.ref = audit.dimensions

    # Formatação simples e legível.
    from openpyxl.styles import Alignment, Font, PatternFill
    title_fill = PatternFill("solid", fgColor="1F4E78")
    header_fill = PatternFill("solid", fgColor="D9EAF7")
    ok_fill = PatternFill("solid", fgColor="E2F0D9")
    warn_fill = PatternFill("solid", fgColor="FFF2CC")
    error_fill = PatternFill("solid", fgColor="FCE4D6")

    resumo["A1"].font = Font(bold=True, color="FFFFFF", size=13)
    resumo["B1"].font = Font(bold=True, color="FFFFFF", size=13)
    resumo["A1"].fill = title_fill; resumo["B1"].fill = title_fill
    for row in range(2, resumo.max_row + 1):
        resumo.cell(row, 1).font = Font(bold=True)
    resumo["B2"].fill = ok_fill if summary.missing_expected == 0 and getattr(summary, "audit_gaps", 0) == 0 else error_fill
    resumo["B3"].fill = ok_fill if all_generated == "SIM" else error_fill

    for sheet in (conf, occ, audit):
        for cell in sheet[1]:
            cell.font = Font(bold=True)
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center")
        for row in sheet.iter_rows(min_row=2):
            for cell in row:
                cell.alignment = Alignment(vertical="top", wrap_text=True)
        for col in sheet.columns:
            values = [str(c.value or "") for c in list(col)[:150]]
            width = min(60, max(10, max((len(v) for v in values), default=10) + 2))
            sheet.column_dimensions[col[0].column_letter].width = width

    for row in range(2, conf.max_row + 1):
        result = str(conf.cell(row, 7).value or "")
        fill = ok_fill if result == "OK" else (warn_fill if result == "PENDENTE" else error_fill)
        conf.cell(row, 6).fill = fill
        conf.cell(row, 7).fill = fill

    for row in range(2, occ.max_row + 1):
        level = str(occ.cell(row, 1).value or "")
        fill = ok_fill if level == "OK" else (warn_fill if level in {"ATENÇÃO", "AVISO"} else error_fill)
        occ.cell(row, 1).fill = fill
        occ.cell(row, 2).fill = fill

    temp = root / (REPORT_FILENAME + ".tmp")
    final = root / REPORT_FILENAME
    wb.save(temp)
    os.replace(temp, final)


def resolve_pending(summary: ProcessingSummary, analysis_id: str, exam_type: str, employee: ExpectedEmployee, action: str = "SALVAR", learn: bool = False, employees: list[ExpectedEmployee] | None = None) -> Path | None:
    root = Path(summary.root_dir)
    item = next((a for a in summary.analyses if a.id == analysis_id), None)
    if not item:
        raise ValueError("Página não encontrada")
    if action == "IGNORAR":
        if item.output_file:
            old = root / item.output_file
            if old.exists() and old.parent == root:
                old.unlink(missing_ok=True)
                item.output_file = ""
        item.status = "IGNORADO_MANUAL"; item.reason = "Ignorado manualmente"
        if employees:
            recompute_summary(summary, employees)
        write_reports(summary); return None

    source = Path(item.source_pdf_path)
    # Laudos PCD podem ter várias páginas. Se o usuário confirmar uma página de
    # continuação imediatamente após outra já salva para o mesmo funcionário,
    # anexa ao mesmo arquivo em vez de criar PDFs separados.
    previous = next((a for a in summary.analyses
                     if a.source_pdf_path == item.source_pdf_path
                     and a.page_index == item.page_index - 1
                     and a.exam_type == "LAUDO PCD"
                     and a.employee_row_id == employee.row_id
                     and a.status in {"SALVO_AUTOMATICO", "SALVO_MANUAL", "ANEXADO_CONTINUACAO"}
                     and a.output_file), None) if exam_type == "LAUDO PCD" else None
    if previous:
        target = root / previous.output_file
        doc = fitz.open(source)
        try:
            append_page_to_pdf(target, doc, item.page_index)
        finally:
            doc.close()
        sync_to_all_files(root, target)
        continuation = True
    else:
        subtype = detect_aso_subtype(item.raw_text) if exam_type == "ASO" else ""
        saved_statuses = {"SALVO_AUTOMATICO", "SALVO_MANUAL", "ANEXADO_CONTINUACAO"}
        occurrence_index = sum(1 for a in summary.analyses
                               if a.id != item.id and a.employee_row_id == employee.row_id and a.exam_type == exam_type
                               and a.status in saved_statuses and a.output_file)
        receipt = employee.expected_receipt(exam_type, occurrence_index)
        target = unique_file(root / build_filename(
            exam_type, employee.name, employee.company or item.company, employee.cnpj or item.cnpj, subtype,
            include_aso_subtype=(exam_type == "ASO" and employee.expected_count("ASO") > 1), receipt=receipt,
        ))
        save_page(source, item.page_index, target)
        sync_to_all_files(root, target)
        continuation = False
    # Remove a cópia temporária de pendência, se existir.
    if item.output_file:
        old = root / item.output_file
        if old.exists() and old.parent == root:
            old.unlink(missing_ok=True)
    item.exam_type = exam_type; item.exam_subtype = detect_aso_subtype(item.raw_text) if exam_type == "ASO" else ""; item.employee_name = employee.name; item.employee_cpf = employee.cpf
    if not continuation:
        item.receipt = receipt
    item.company = employee.company or item.company; item.cnpj = employee.cnpj or item.cnpj
    item.employee_row_id = employee.row_id; item.expected = exam_type in employee.expected_exams
    item.status = "ANEXADO_CONTINUACAO" if continuation else "SALVO_MANUAL"
    item.reason = "Página anexada manualmente ao Laudo PCD anterior" if continuation else "Confirmado manualmente"
    item.output_file = str(target.relative_to(root))
    if learn:
        add_model_from_page(source, item.page_index, exam_type, f"Aprendido - {source.stem} pág {item.page_number}", use_ocr=True)
    if employees:
        recompute_summary(summary, employees)
    write_reports(summary)
    return target


def create_expected_list_example(path: Path) -> None:
    bundled = APP_DIR / "MODELO_FUNCIONARIOS.xlsx"
    try:
        if bundled.exists() and bundled.resolve() != path.resolve():
            path.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(bundled, path)
            return
    except Exception:
        pass
    if Workbook is None:
        raise RuntimeError("openpyxl não instalado")
    wb = Workbook(); ws = wb.active; ws.title = "Funcionarios"
    ws.append(["FUNCIONÁRIO", "EMPRESA", "EXAMES ESPERADOS", "RECIBO"])
    ws.freeze_panes = "A2"
    widths = [32, 62, 42, 24]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[chr(64+i)].width = w
    for c in ws[1]:
        c.font = c.font.copy(bold=True)
    wb.save(path)
