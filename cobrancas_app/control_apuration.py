"""Apuração do controle mensal; valores da planilha e origens vinculadas à responsável."""
from collections import Counter, defaultdict
import hashlib
import json


def process(api, conn, comp, ws, path, original_name, auto_register_missing=False):
    # Compatibilidade com clientes antigos: a opção nunca cria cadastros.
    auto_register_missing = False
    competency_id = comp["id"]
    locked = {r["company_id"]: r for r in conn.execute(
        "SELECT * FROM competency_companies WHERE competency_id=? AND (complementary_sent_at IS NOT NULL OR complementary_paid=1)", (competency_id,))}
    editable = "competency_id=? AND company_id NOT IN (SELECT company_id FROM competency_companies WHERE competency_id=? AND (complementary_sent_at IS NOT NULL OR complementary_paid=1))"
    if (ws.max_column or 0) > 200:
        raise ValueError("O controle deve ter até 200 colunas.")
    reported_rows = ws.max_row
    header_row, headers = api.control_header(ws)
    data_last_row, data_last_col = api.control_sheet_data_bounds(path, ws.title)
    if data_last_row < header_row:
        raise ValueError("A planilha de controle não possui linhas de dados.")

    def column(*names):
        return next((headers[api.norm_header(n)] - 1 for n in names if api.norm_header(n) in headers), None)

    employee_col = column("FUNCIONARIO", "COLABORADOR", "NOME DO FUNCIONARIO")
    exam_col = column("TIPO DE EXAME", "EXAME")
    receipt_col = column("RECIBO", "Nº RECIBO", "NO RECIBO", "RECEBIMENTO", "FORMA DE PAGAMENTO", "PAGAMENTO")
    company_col = column("SETOR", "EMPRESA", "RAZAO SOCIAL")
    document_col = column("CNPJ/CPF", "CNPJ CPF", "CNPJ", "CPF DA EMPRESA", "CPF EMPRESA")
    # No modelo de controle, CPF ao lado de SETOR pertence ao paciente.
    # CPF genérico só identifica empresa nos layouts sem SETOR/EMPRESA.
    if document_col is None and company_col is None:
        document_col = column("CPF")
    date_col, job_col, value_col = column("DATA"), column("FUNCAO"), column("VALOR")
    if value_col is None:
        raise ValueError("A planilha precisa da coluna VALOR para apurar complementares.")
    observations_col, status_col = column("OBS", "OBSERVACAO", "OBSERVACOES"), column("STATUS", "SITUACAO")

    def cell(row, index):
        return row[index] if index is not None and index < len(row) else None

    all_companies = conn.execute("SELECT c.*,u.name unit_name FROM companies c JOIN units u ON u.id=c.unit_id").fetchall()
    selected = {api.digits(r["cnpj"]): r for r in all_companies if r["unit_id"] == comp["unit_id"]}
    aliases = {r["document"]:r for r in conn.execute("SELECT * FROM complementary_sources WHERE unit_id=?",(comp["unit_id"],))}
    companies_by_id = {r["id"]:r for r in all_companies}
    active_exam_catalog = {}
    active_exams = set()
    for r in conn.execute("SELECT name,exam_key FROM complementary_exam_types WHERE active=1"):
        normalized_key = api.normalize_text(r["exam_key"])
        active_exam_catalog[normalized_key] = r["name"]
        active_exam_catalog.setdefault(api.norm_header(r["exam_key"]), r["name"])
        active_exams.add(normalized_key)
    overrides = {r["fingerprint"]: r for r in conn.execute("SELECT * FROM exam_items WHERE competency_id=? AND duplicate_override=1", (competency_id,)) if r["fingerprint"]}
    prior_by_fingerprint = {}
    for r in conn.execute("""SELECT e.*,cp.month,cp.year FROM exam_items e
        JOIN competencies cp ON cp.id=e.competency_id
        JOIN competency_companies cc ON cc.competency_id=e.competency_id AND cc.company_id=e.company_id
        WHERE e.competency_id<>? AND e.status='OK'
        AND (cc.complementary_sent_at IS NOT NULL OR cc.complementary_paid=1 OR COALESCE(cp.status,'')='FECHADA')""", (competency_id,)):
        fingerprint = r["fingerprint"] or api.exam_fingerprint(r["company_id"], r["employee"], r["exam_key"], r["exam_date"], r["job_title"] or "")
        prior_by_fingerprint.setdefault(fingerprint, r)

    digest = hashlib.sha256(path.read_bytes()).hexdigest()
    run_id = conn.execute("INSERT INTO apuration_runs(competency_id,original_name,file_hash,stats_json,created_at) VALUES(?,?,?,?,?)", (competency_id, original_name, digest, "{}", api.now_iso())).lastrowid
    conn.execute("DELETE FROM exam_items WHERE " + editable, (competency_id, competency_id))
    api.sync_competency_companies(conn, competency_id)
    conn.execute("UPDATE competency_companies SET complementary_amount=0,updated_at=? WHERE " + editable, (api.now_iso(), competency_id, competency_id))
    conn.execute("UPDATE competency_companies SET complementary_amount_manual=0, complementary_value_pending=COALESCE((SELECT complementary_value_required FROM companies WHERE id=company_id),0) WHERE " + editable, (competency_id, competency_id))
    stats = {key: 0 for key in ("rows", "not_a_prazo", "a_prazo", "candidate", "charged", "unpriced", "unregistered", "paid_at_exam", "other_unit", "duplicates", "historical_duplicates", "ignored_exam_type", "invalid_rows", "cancelled", "registered_companies", "price_differences")}
    stats.update(sheet_name=ws.title, reported_rows=reported_rows, effective_last_row=data_last_row,
        price_source="planilha", price_source_label="Coluna VALOR da planilha de controle", auto_register_missing=bool(auto_register_missing), run_id=run_id)
    stats.update(preserved_rows=0, preserved_companies=len(locked),
        preserved_total=sum(int(round((r["complementary_amount"] or 0) * 100)) for r in locked.values()) / 100)
    seen, unmatched_documents = set(), set()
    totals, source_total = defaultdict(int), 0
    source_values, exam_counts = Counter(), Counter()
    detected_exam_counts, unregistered_exam_counts = Counter(), Counter()
    unpriced_names = defaultdict(set)
    unrecognized, summaries = {}, {}
    registered_ids = set()

    def audit(row_number, decision, reason, document="", company=None, source_company="", employee="", source_exam="", canonical_exam="", exam_date=None, unit_price=None, amount=None, fingerprint=""):
        conn.execute("""INSERT INTO apuration_audit(run_id,competency_id,source_row,decision,reason,document,company_id,company_name,employee,source_exam,canonical_exam,exam_date,unit_price,amount,fingerprint,created_at)
            VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""", (run_id, competency_id, row_number, decision, reason, document,
            company["id"] if company else None, company["name"] if company else source_company, employee, source_exam, canonical_exam,
            exam_date.isoformat() if exam_date else "", unit_price, amount, fingerprint, api.now_iso()))

    max_control_col = max(data_last_col, max(headers.values()) if headers else 1)
    for row_number, row in enumerate(ws.iter_rows(min_row=header_row + 1, max_row=data_last_row, max_col=max_control_col, values_only=True), header_row + 1):
        if not any(value is not None and str(value).strip() for value in row):
            continue
        stats["rows"] += 1
        if stats["rows"] > 100000:
            raise ValueError("O controle deve ter até 100.000 linhas com dados. A apuração anterior foi preservada.")
        stats["last_data_row"] = row_number
        employee, source_exam = str(cell(row, employee_col) or "").strip(), str(cell(row, exam_col) or "").strip()
        company_cell = cell(row, company_col)
        source_company = api.company_name_from_cell(company_cell)
        documents = set(api.control_company_documents(company_cell))
        explicit_document = cell(row, document_col)
        if explicit_document is not None and str(explicit_document).strip():
            documents.update(api.control_company_documents(explicit_document))
        document = next(iter(documents)) if len(documents) == 1 else ""
        exam_date = api.parse_excel_date(cell(row, date_col))
        context = dict(document=document, source_company=source_company, employee=employee, source_exam=source_exam, exam_date=exam_date)
        receipt_value = cell(row, receipt_col) if receipt_col is not None else cell(row, 2)
        if api.normalize_text(receipt_value) != "A PRAZO":
            stats["not_a_prazo"] += 1
            audit(row_number, "IGNORADO", "RECIBO/RECEBIMENTO NÃO É A PRAZO", **context)
            continue
        stats["a_prazo"] += 1
        # DEPOSITANTE é irrelevante para cobrança neste modelo; somente
        # STATUS/OBS podem cancelar um atendimento.
        if api.row_cancelled([cell(row, observations_col), cell(row, status_col)]):
            stats["cancelled"] += 1
            audit(row_number, "IGNORADO", "ATENDIMENTO CANCELADO OU NÃO COBRAR", **context)
            continue
        detected_exam_name = api.normalize_text(source_exam)
        if detected_exam_name:
            detected_exam_counts[detected_exam_name] += 1
        exam_name = api.canonical_control_exam(source_exam, active_exam_catalog)
        if not exam_name:
            stats["ignored_exam_type"] += 1
            if detected_exam_name:
                unregistered_exam_counts[detected_exam_name] += 1
            audit(row_number, "IGNORADO", "TIPO DE EXAME NÃO ESTÁ CADASTRADO/ATIVO COMO COMPLEMENTAR", **context)
            continue
        context["canonical_exam"] = exam_name
        exam_key = api.normalize_text(exam_name)
        stats["candidate"] += 1
        exam_counts[exam_name] += 1
        if not employee:
            stats["invalid_rows"] += 1
            audit(row_number, "IGNORADO", "FUNCIONÁRIO AUSENTE", **context)
            continue
        raw_date = cell(row, date_col)
        if raw_date is not None and str(raw_date).strip() and exam_date is None:
            raise ValueError(f"Linha {row_number}: DATA inválida. A apuração anterior foi preservada.")
        try:
            source_value = api.parse_money_strict(cell(row, value_col))
        except ValueError as exc:
            raise ValueError(f"Linha {row_number}: VALOR inválido. A apuração anterior foi preservada.") from exc
        source_values["missing" if source_value is None else "zero" if source_value == 0 else "positive"] += 1
        if source_value is not None:
            source_total += int(round(source_value * 100))

        alias = aliases.get(document)
        company = companies_by_id.get(alias["mother_id"]) if alias else (selected.get(document) if document else None)
        if alias and alias["name"]: source_company = alias["name"]
        reason = None
        if not document:
            reason = "MAIS DE UM CNPJ/CPF INFORMADO" if len(documents) > 1 else "SETOR/EMPRESA SEM CNPJ/CPF VÁLIDO"
        elif company and company["unit_id"] != comp["unit_id"]:
            reason = "EMPRESA RESPONSÁVEL PERTENCE A OUTRA UNIDADE: CORRIJA O VÍNCULO"
        elif company and not company["active"]:
            reason = "EMPRESA INATIVA NESTA UNIDADE"
        elif not company:
            if any(api.digits(r["cnpj"]) == document for r in all_companies):
                reason = "CNPJ/CPF CADASTRADO SOMENTE EM OUTRA UNIDADE"
            else:
                reason = "CNPJ/CPF NÃO CADASTRADO NESTA UNIDADE"
        if reason:
            stats["other_unit"] += 1
            if document:
                unmatched_documents.add(document)
            key = (document, source_company, reason)
            item = unrecognized.setdefault(key, {"document": document, "name": source_company or "Empresa não informada", "reason": reason, "rows": 0, "exams": Counter()})
            item["rows"] += 1
            item["exams"][exam_name] += 1
            audit(row_number, "IGNORADO", reason, **context)
            continue
        context["company"] = company
        summary = summaries.setdefault(company["id"], {"company_id": company["id"], "document": company["cnpj"], "name": company["name"], "registered_now": company["id"] in registered_ids,
            "missing_email": not api.valid_email(company["email"]), "rows": 0, "charged": 0, "unpriced": 0, "paid_at_exam": 0, "duplicates": 0, "historical_duplicates": 0, "total_cents": 0, "source_reference_cents": 0, "exams": Counter(), "unpriced_exams": Counter()})
        summary["rows"] += 1
        summary["exams"][exam_name] += 1
        if source_value is not None:
            summary["source_reference_cents"] += int(round(source_value * 100))
        if company["id"] in locked:
            stats["preserved_rows"] += 1
            audit(row_number, "PRESERVADO", "COBRANÇA JÁ ENVIADA OU PAGA: HISTÓRICO MANTIDO; LINHA NÃO REAPURADA", **context)
            continue
        if exam_key not in active_exams:
            stats["unregistered"] += 1
            audit(row_number, "IGNORADO", "TIPO DE EXAME NÃO ESTÁ ATIVO NO CADASTRO", **context)
            continue
        if not company["bill_complementaries"]:
            stats["paid_at_exam"] += 1
            summary["paid_at_exam"] += 1
            audit(row_number, "IGNORADO", "CADASTRO NÃO HABILITA COBRANÇA MENSAL DE COMPLEMENTARES", **context)
            continue
        job_title = str(cell(row, job_col) or "").strip()
        fingerprint = api.exam_fingerprint("origem:"+str(comp["unit_id"])+":"+document if alias else company["id"], employee, exam_key, exam_date, job_title)
        if fingerprint in seen:
            stats["duplicates"] += 1
            summary["duplicates"] += 1
            audit(row_number, "IGNORADO", "DUPLICADO NA MESMA PLANILHA", fingerprint=fingerprint, **context)
            continue
        seen.add(fingerprint)
        unit_price = source_value
        prior, override = prior_by_fingerprint.get(fingerprint), overrides.get(fingerprint)
        if prior and not override:
            stats["historical_duplicates"] += 1
            summary["historical_duplicates"] += 1
            status, amount = "JA_COBRADO", 0
            note = f"Já faturado em {api.MONTHS.get(prior['month'], prior['month'])}/{prior['year']}. Libere manualmente somente se for um atendimento distinto."
            decision, audit_reason = "IGNORADO", "JÁ COBRADO EM COMPETÊNCIA ANTERIOR"
        elif unit_price is None:
            stats["unpriced"] += 1
            summary["unpriced"] += 1
            summary["unpriced_exams"][exam_name] += 1
            unpriced_names[company["id"]].add(exam_name)
            status, amount = "SEM_PRECO", None
            note = "Preencha VALOR neste atendimento da planilha e reprocesse, ou confirme um total manual para esta competência."
            decision, audit_reason = "PENDENTE", "VALOR AUSENTE NA PLANILHA"
        else:
            stats["charged"] += 1
            summary["charged"] += 1
            cents = int(round(unit_price * 100))
            totals[company["id"]] += cents
            summary["total_cents"] += cents
            status, amount = "OK", unit_price
            note = "Cobrança duplicada liberada manualmente pelo usuário." if override else None
            decision, audit_reason = "INCLUÍDO", "COBRADO PELA COLUNA VALOR DA PLANILHA"
            if source_value is not None and int(round(source_value * 100)) != cents:
                stats["price_differences"] += 1
        conn.execute("""INSERT INTO exam_items(competency_id,company_id,employee,exam_name,exam_key,exam_date,job_title,receipt,source_value,unit_price,total,source_row,status,note,created_at,fingerprint,duplicate_of_item_id,duplicate_override,source_document,source_company)
            VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""", (competency_id, company["id"], employee, exam_name, exam_key, exam_date.isoformat() if exam_date else None,
            job_title, "A PRAZO", source_value, unit_price, amount, row_number, status, note, api.now_iso(), fingerprint, prior["id"] if prior else None, int(bool(override)), document, source_company or company["name"]))
        audit(row_number, decision, audit_reason, unit_price=unit_price, amount=amount, fingerprint=fingerprint, **context)

    if not stats["rows"]:
        raise ValueError("A planilha de controle está vazia. A apuração anterior foi preservada.")
    for company_id, cents in totals.items():
        conn.execute("UPDATE competency_companies SET complementary_amount=?,updated_at=? WHERE competency_id=? AND company_id=?", (cents / 100, api.now_iso(), competency_id, company_id))
    stats.update(unmatched_documents=sorted(unmatched_documents), unrecognized_companies=list(unrecognized.values()), exam_counts=dict(exam_counts),
        detected_exam_counts=dict(detected_exam_counts), unregistered_exam_counts=dict(unregistered_exam_counts), source_values=dict(source_values),
        source_total_reference=source_total / 100, charged_total=sum(totals.values()) / 100, identified_companies=len(summaries), companies_without_email=sum(item["missing_email"] for item in summaries.values()),
        companies_without_price=sum(bool(item["unpriced"]) for item in summaries.values()), company_summary=[dict(item, total=item["total_cents"] / 100, source_total_reference=item["source_reference_cents"] / 100) for item in summaries.values()])
    stored_name = f"competencia_{competency_id}_{digest[:12]}.xlsx"
    stored_path = api.CONTROL_DIR / stored_name
    if path.resolve() != stored_path.resolve():
        stored_path.write_bytes(path.read_bytes())
    conn.execute("UPDATE apuration_runs SET stats_json=? WHERE id=?", (json.dumps(stats, ensure_ascii=False), run_id))
    conn.execute("UPDATE competencies SET control_filename=?,control_hash=?,control_stored_name=?,processed_at=?,updated_at=? WHERE id=?", (original_name, digest, stored_name, api.now_iso(), api.now_iso(), competency_id))
    return stats, unpriced_names
