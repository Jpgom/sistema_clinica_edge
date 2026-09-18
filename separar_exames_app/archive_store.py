# -*- coding: utf-8 -*-
from __future__ import annotations

import hashlib
import io
import json
import re
import shutil
import threading
import unicodedata
import zipfile
from datetime import datetime
from pathlib import Path
from typing import Any, Iterable


def digits_only(value: str) -> str:
    return re.sub(r"\D+", "", str(value or ""))


def document_kind(value: str) -> str:
    d = digits_only(value)
    if len(d) == 11:
        return "CPF"
    if len(d) == 14:
        return "CNPJ"
    return "CPF/CNPJ"


def format_document(value: str) -> str:
    d = digits_only(value)
    if len(d) == 11:
        return f"{d[:3]}.{d[3:6]}.{d[6:9]}-{d[9:]}"
    if len(d) == 14:
        return f"{d[:2]}.{d[2:5]}.{d[5:8]}/{d[8:12]}-{d[12:]}"
    return d or str(value or "").strip()


def _ascii(value: str) -> str:
    value = unicodedata.normalize("NFKD", str(value or ""))
    return "".join(ch for ch in value if not unicodedata.combining(ch))


def norm(value: str) -> str:
    return re.sub(r"\s+", " ", _ascii(value).upper()).strip()


def safe_part(value: str, fallback: str = "SEM NOME", max_len: int = 90) -> str:
    value = norm(value)
    value = re.sub(r'[<>:"/\\|?*\x00-\x1f]+', " ", value)
    value = re.sub(r"\s+", " ", value).strip(" .")
    return value[:max_len] or fallback


def file_sha256(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


class ArchiveStore:
    """Arquivo persistente simples baseado em JSON + PDFs.

    O servidor atual usa um único processo com múltiplas threads; um RLock protege
    as gravações e o índice é salvo de forma atômica para evitar corrupção.
    """

    def __init__(self, root: Path):
        self.root = Path(root)
        self.files_root = self.root / "arquivos"
        self.index_path = self.root / "index.json"
        self.lock = threading.RLock()
        self.files_root.mkdir(parents=True, exist_ok=True)
        if not self.index_path.exists():
            self._write_index({"format": 1, "next_id": 1, "documents": []})

    def _read_index(self) -> dict[str, Any]:
        with self.lock:
            for candidate in (self.index_path, self.index_path.with_suffix(".bak")):
                try:
                    data = json.loads(candidate.read_text(encoding="utf-8"))
                    if not isinstance(data, dict) or not isinstance(data.get("documents"), list):
                        raise ValueError("Índice inválido")
                    data.setdefault("next_id", 1)
                    return data
                except Exception:
                    continue
            return {"format": 1, "next_id": 1, "documents": []}

    def _write_index(self, data: dict[str, Any]) -> None:
        self.root.mkdir(parents=True, exist_ok=True)
        temp = self.index_path.with_suffix(".tmp")
        backup = self.index_path.with_suffix(".bak")
        temp.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        if self.index_path.exists():
            try:
                shutil.copy2(self.index_path, backup)
            except Exception:
                pass
        temp.replace(self.index_path)

    def _company_key(self, company_name: str, company_document: str) -> str:
        doc = digits_only(company_document)
        return f"DOC:{doc}" if doc else "NOME:" + norm(company_name or "EMPRESA NAO IDENTIFICADA")

    def _unique_target(self, folder: Path, filename: str) -> Path:
        folder.mkdir(parents=True, exist_ok=True)
        filename = safe_part(Path(filename).stem, "DOCUMENTO", 160) + ".pdf"
        target = folder / filename
        if not target.exists():
            return target
        n = 2
        while True:
            candidate = folder / f"{target.stem} ({n}){target.suffix}"
            if not candidate.exists():
                return candidate
            n += 1

    def count_for_job(self, job_id: str) -> int:
        data = self._read_index()
        return len({d.get("sha256") for d in data["documents"] if d.get("source_job_id") == job_id and d.get("sha256")})

    def save_documents(self, documents: Iterable[dict[str, Any]], year: int, month: int, source_job_id: str = "") -> dict[str, Any]:
        year, month = int(year), int(month)
        if year < 2000 or year > 2100 or month < 1 or month > 12:
            raise ValueError("Competência inválida.")
        added: list[dict[str, Any]] = []
        skipped: list[dict[str, Any]] = []
        errors: list[dict[str, Any]] = []
        with self.lock:
            data = self._read_index()
            docs = data["documents"]
            existing_hashes = {(int(d.get("year", 0)), int(d.get("month", 0)), str(d.get("sha256", ""))): d for d in docs}
            changed = False
            for raw in documents:
                try:
                    src = Path(str(raw.get("source_path") or ""))
                    if not src.is_file():
                        raise FileNotFoundError(f"Arquivo não encontrado: {src.name}")
                    digest = file_sha256(src)
                    existing = existing_hashes.get((year, month, digest))
                    if existing:
                        skipped.append({"id": int(existing["id"]), "filename": existing.get("original_filename", src.name), "reason": "Já arquivado nesta competência"})
                        continue
                    company_name = str(raw.get("company_name") or "").strip()
                    company_document = digits_only(raw.get("company_document") or "")
                    company_key = self._company_key(company_name, company_document)
                    if company_document:
                        company_label = f"{document_kind(company_document)} {company_document}"
                    else:
                        company_label = safe_part(company_name, "EMPRESA NAO IDENTIFICADA", 90)
                    exam_type = str(raw.get("exam_type") or "OUTROS").strip().upper() or "OUTROS"
                    period_dir = self.files_root / str(year) / f"{month:02d}" / company_label / safe_part(exam_type, "OUTROS", 50)
                    target = self._unique_target(period_dir, src.name)
                    shutil.copy2(src, target)
                    doc_id = int(data.get("next_id", 1))
                    data["next_id"] = doc_id + 1
                    item = {
                        "id": doc_id, "year": year, "month": month,
                        "company_name": company_name, "company_document": company_document,
                        "company_document_kind": document_kind(company_document), "company_key": company_key,
                        "employee_name": str(raw.get("employee_name") or "").strip(),
                        "employee_cpf": digits_only(raw.get("employee_cpf") or ""),
                        "exam_type": exam_type, "exam_subtype": str(raw.get("exam_subtype") or "").strip().upper(),
                        "receipt": str(raw.get("receipt") or "").strip(),
                        "exam_date": str(raw.get("exam_date") or "").strip(),
                        "stored_path": target.relative_to(self.root).as_posix(), "original_filename": src.name,
                        "source_job_id": source_job_id, "source_analysis_id": str(raw.get("source_analysis_id") or ""),
                        "sha256": digest, "size_bytes": int(src.stat().st_size),
                        "created_at": datetime.now().isoformat(timespec="seconds"),
                    }
                    docs.append(item)
                    existing_hashes[(year, month, digest)] = item
                    added.append({"id": doc_id, "filename": src.name})
                    changed = True
                except Exception as exc:
                    errors.append({"filename": Path(str(raw.get("source_path") or "arquivo")).name, "error": str(exc)})
            if changed:
                self._write_index(data)
        return {"added": added, "skipped": skipped, "errors": errors, "added_count": len(added), "skipped_count": len(skipped), "error_count": len(errors)}

    def periods(self) -> list[dict[str, int]]:
        counts: dict[tuple[int, int], int] = {}
        for d in self._read_index()["documents"]:
            key = (int(d["year"]), int(d["month"]))
            counts[key] = counts.get(key, 0) + 1
        return [{"year": y, "month": m, "count": n} for (y, m), n in sorted(counts.items(), reverse=True)]

    def companies(self, year: int | None = None, month: int | None = None) -> list[dict[str, Any]]:
        grouped: dict[str, dict[str, Any]] = {}
        for d in self._read_index()["documents"]:
            if year and int(d["year"]) != int(year): continue
            if month and int(d["month"]) != int(month): continue
            key = str(d["company_key"])
            x = grouped.setdefault(key, {"key": key, "name": d.get("company_name") or "Empresa não identificada", "document_raw": d.get("company_document", ""), "document_kind": d.get("company_document_kind", "CPF/CNPJ"), "count": 0})
            x["count"] += 1
        items = []
        for x in grouped.values():
            x = dict(x); x["document"] = format_document(x["document_raw"]); items.append(x)
        return sorted(items, key=lambda x: norm(x["name"]))

    def exam_types(self) -> list[str]:
        return sorted({str(d.get("exam_type") or "") for d in self._read_index()["documents"] if d.get("exam_type")})

    def _matches(self, d: dict[str, Any], *, q: str = "", year: int | None = None, month: int | None = None,
                 company_key: str = "", exam_type: str = "", receipt_filter: str = "", ids: list[int] | None = None) -> bool:
        if year and int(d.get("year", 0)) != int(year): return False
        if month and int(d.get("month", 0)) != int(month): return False
        if company_key and d.get("company_key") != company_key: return False
        if exam_type and str(d.get("exam_type", "")).upper() != str(exam_type).upper(): return False
        if receipt_filter:
            rf = norm(receipt_filter)
            receipt = str(d.get("receipt") or "").strip()
            is_prazo = norm(receipt) in {"A PRAZO", "PRAZO"}
            if rf in {"A PRAZO", "PRAZO", "A PRAZOS"} and not is_prazo: return False
            if rf in {"RECIBO", "RECIBOS"} and (not receipt or is_prazo): return False
        if ids and int(d.get("id", 0)) not in ids: return False
        if q.strip():
            nq = norm(q)
            qdigits = digits_only(q)
            hay = " | ".join(norm(d.get(k, "")) for k in ("employee_name", "company_name", "original_filename", "exam_type", "exam_subtype", "receipt"))
            if nq not in hay:
                if not qdigits or (qdigits not in digits_only(d.get("employee_cpf", "")) and qdigits not in digits_only(d.get("company_document", "")) and qdigits not in digits_only(d.get("receipt", ""))):
                    return False
        return True

    def _public(self, d: dict[str, Any]) -> dict[str, Any]:
        return {
            "id": int(d["id"]), "year": int(d["year"]), "month": int(d["month"]),
            "competency": f"{int(d['month']):02d}/{int(d['year'])}",
            "company_name": d.get("company_name", ""), "company_document": format_document(d.get("company_document", "")),
            "company_document_raw": d.get("company_document", ""), "company_document_kind": d.get("company_document_kind", "CPF/CNPJ"),
            "company_key": d.get("company_key", ""), "employee_name": d.get("employee_name", ""),
            "employee_cpf": format_document(d.get("employee_cpf", "")), "employee_cpf_raw": d.get("employee_cpf", ""),
            "exam_type": d.get("exam_type", ""), "exam_subtype": d.get("exam_subtype", ""), "receipt": d.get("receipt", ""), "exam_date": d.get("exam_date", ""),
            "original_filename": d.get("original_filename", ""), "size_bytes": int(d.get("size_bytes", 0)), "created_at": d.get("created_at", ""),
        }

    def search(self, *, q: str = "", year: int | None = None, month: int | None = None,
               company_key: str = "", exam_type: str = "", receipt_filter: str = "", page: int = 1, page_size: int = 100) -> dict[str, Any]:
        page, page_size = max(1, int(page)), max(1, min(250, int(page_size)))
        matches = [d for d in self._read_index()["documents"] if self._matches(d, q=q, year=year, month=month, company_key=company_key, exam_type=exam_type, receipt_filter=receipt_filter)]
        matches.sort(key=lambda d: (-int(d["year"]), -int(d["month"]), norm(d.get("company_name", "")), norm(d.get("employee_name", "")), norm(d.get("exam_type", "")), int(d["id"])))
        total = len(matches); pages = max(1, (total + page_size - 1)//page_size); page=min(page,pages)
        part = matches[(page-1)*page_size:page*page_size]
        company_count = len({d.get("company_key") for d in matches})
        employee_count = len({(norm(d.get("employee_name", "")), digits_only(d.get("employee_cpf", ""))) for d in matches})
        return {"items":[self._public(d) for d in part], "total":total, "page":page, "page_size":page_size, "pages":pages, "companies":company_count, "employees":employee_count}

    def get(self, doc_id: int) -> dict[str, Any] | None:
        for d in self._read_index()["documents"]:
            if int(d.get("id", 0)) == int(doc_id): return self._public(d)
        return None

    def path_for(self, doc_id: int) -> Path | None:
        for d in self._read_index()["documents"]:
            if int(d.get("id", 0)) != int(doc_id): continue
            target = (self.root / d["stored_path"]).resolve(); root=self.root.resolve()
            if root not in target.parents or not target.is_file(): return None
            return target
        return None

    def filtered_rows(self, *, q: str = "", year: int | None = None, month: int | None = None,
                      company_key: str = "", exam_type: str = "", receipt_filter: str = "", ids: list[int] | None = None) -> list[dict[str, Any]]:
        wanted = [int(i) for i in (ids or [])]
        items = [d for d in self._read_index()["documents"] if self._matches(d, q=q, year=year, month=month, company_key=company_key, exam_type=exam_type, receipt_filter=receipt_filter, ids=wanted or None)]
        items.sort(key=lambda d: (norm(d.get("company_name", "")), norm(d.get("employee_name", "")), norm(d.get("exam_type", "")), int(d["id"])))
        return items

    def delete_documents(self, ids: Iterable[int]) -> dict[str, Any]:
        """Exclui documentos do arquivo permanente e remove pastas vazias.

        A remoção usa os IDs do índice, apaga somente arquivos que estejam dentro
        da raiz persistente e atualiza o índice de forma atômica.
        """
        wanted = {int(x) for x in ids if str(x).strip()}
        if not wanted:
            return {"deleted_count": 0, "deleted": [], "missing": []}

        deleted: list[dict[str, Any]] = []
        missing: list[int] = []
        with self.lock:
            data = self._read_index()
            docs = data.get("documents", [])
            by_id = {int(d.get("id", 0)): d for d in docs}
            missing = sorted(i for i in wanted if i not in by_id)
            targets = [by_id[i] for i in sorted(wanted) if i in by_id]
            root = self.root.resolve()
            files_root = self.files_root.resolve()

            for d in targets:
                stored = str(d.get("stored_path") or "")
                path = (self.root / stored).resolve()
                try:
                    if path.is_file() and root in path.parents:
                        path.unlink()
                except Exception:
                    # Não impede a remoção do registro do índice caso o PDF já
                    # tenha sido apagado manualmente do disco.
                    pass
                deleted.append({"id": int(d.get("id", 0)), "filename": d.get("original_filename", "")})

                # Limpa somente diretórios vazios dentro de arquivos/.
                parent = path.parent
                while parent != files_root and files_root in parent.parents:
                    try:
                        parent.rmdir()
                    except OSError:
                        break
                    parent = parent.parent

            if targets:
                data["documents"] = [d for d in docs if int(d.get("id", 0)) not in wanted]
                self._write_index(data)

        return {"deleted_count": len(deleted), "deleted": deleted, "missing": missing}

    def make_zip(self, rows: list[dict[str, Any]], label: str = "EXAMES_ARQUIVADOS") -> io.BytesIO:
        """Cria ZIP organizado por empresa e tipo de exame.

        - Se houver apenas uma empresa: PASTA RAIZ / TIPO / PDFs
        - Se houver mais de uma: PASTA RAIZ / EMPRESA / TIPO / PDFs
        - Só são criadas pastas de tipos que realmente tenham arquivos.
        """
        mem = io.BytesIO(); used: set[str] = set(); root = self.root.resolve()
        folder = safe_part(label, "EXAMES_ARQUIVADOS", 150)
        company_keys = {str(r.get("company_key") or "") for r in rows}
        multi_company = len(company_keys) > 1

        # Evita colisão de nomes de pasta entre empresas homônimas.
        company_folder_by_key: dict[str, str] = {}
        used_company_names: dict[str, str] = {}
        for r in rows:
            key = str(r.get("company_key") or "")
            if key in company_folder_by_key:
                continue
            name = safe_part(r.get("company_name") or "EMPRESA NAO IDENTIFICADA", "EMPRESA NAO IDENTIFICADA", 105)
            normalized = norm(name)
            if normalized in used_company_names and used_company_names[normalized] != key:
                doc = digits_only(r.get("company_document") or "")
                if doc:
                    name = safe_part(f"{name} - {doc}", name, 125)
                else:
                    name = safe_part(f"{name} - {len(company_folder_by_key)+1}", name, 125)
            used_company_names[norm(name)] = key
            company_folder_by_key[key] = name

        with zipfile.ZipFile(mem, "w", zipfile.ZIP_DEFLATED) as zf:
            for r in rows:
                src = (self.root / r["stored_path"]).resolve()
                if not src.is_file() or root not in src.parents:
                    continue
                filename = Path(str(r.get("original_filename") or src.name)).name
                filename = re.sub(r'[\\/:*?"<>|\x00-\x1f]+', " - ", filename).strip(" .") or "DOCUMENTO.pdf"
                if not filename.lower().endswith(".pdf"):
                    filename += ".pdf"

                exam_folder = safe_part(r.get("exam_type") or "OUTROS", "OUTROS", 70)
                parts = [folder]
                if multi_company:
                    parts.append(company_folder_by_key.get(str(r.get("company_key") or ""), "EMPRESA NAO IDENTIFICADA"))
                parts.append(exam_folder)
                arc = "/".join(parts + [filename])
                base = arc; n = 2
                while arc in used:
                    p = Path(base)
                    arc = str(p.with_name(f"{p.stem} ({n}){p.suffix}")).replace("\\", "/")
                    n += 1
                used.add(arc)
                zf.write(src, arc)
        mem.seek(0)
        return mem
