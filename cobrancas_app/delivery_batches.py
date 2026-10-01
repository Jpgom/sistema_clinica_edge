"""Pacotes de transporte por destinatário, preservando cada cobrança e grupo explícito."""
import billing_policy
from decimal import Decimal
from pathlib import Path
import html
import re
import sqlite3

import billing_groups
from mail_transport import normalize_cc


def normalize_email(value):
    return str(value or "").strip().lower()


def _candidate(row, billing, email_type, force):
    prefix = "fixed" if billing == "FIXED" else "complementary"
    if not row["active"] or (billing == "COMPLEMENTARY" and not row["bill_complementaries"]):
        return False
    if row[f"{prefix}_paid"]:
        return False
    if email_type.endswith("_REMINDER"):
        if not row[f"{prefix}_sent_at"]:
            return False
    elif force:
        # Reenvio explícito: somente cobranças que já foram enviadas e ainda não
        # foram pagas. Evita que uma empresa nova/pendente entre por acidente no
        # mesmo lote apenas por compartilhar o destinatário.
        if not row[f"{prefix}_sent_at"]:
            return False
    elif row[f"{prefix}_sent_at"]:
        return False
    return billing_policy.pending(row, billing) or Decimal(str(row[f"{prefix}_amount"] or 0)) > 0 or (billing == "COMPLEMENTARY" and billing_policy.unpriced_blocks(row))


def package_context(conn, competency_id, company_id, email_type, force=False):
    """Os candidatos financeiros são definidos antes de validar preço e anexos.

    Assim, uma empresa sem documentação ou preço bloqueia o pacote inteiro do
    destinatário. Um grupo cadastrado nunca é absorvido pelo lote automático.
    """
    group, owner, own_rows, billing = billing_groups.context(conn, competency_id, company_id, email_type)
    if group:
        rows = billing_groups.outstanding_rows(conn, competency_id, company_id, email_type, force)[2]
        return dict(group=group, representative_id=owner, rows=rows, all_rows=own_rows,
                    billing=billing, delivery_kind="explicit_group")
    company = conn.execute("SELECT * FROM companies WHERE id=?", (company_id,)).fetchone()
    comp = conn.execute("SELECT * FROM competencies WHERE id=?", (competency_id,)).fetchone()
    registered = conn.execute("SELECT group_id FROM billing_group_members WHERE company_id=?", (company_id,)).fetchone()
    all_rows = own_rows
    if company and comp and company["unit_id"] == comp["unit_id"] and not registered:
        # Usa a mesma normalização da prévia/SMTP, inclusive tabulações e
        # espaços Unicode vindos de planilhas, em vez do trim ASCII do SQLite.
        try:
            conn.execute("SELECT edge_email_key(NULL)").fetchone()
        except sqlite3.OperationalError:
            conn.create_function("edge_email_key", 1, normalize_email, deterministic=True)
        all_rows = conn.execute("""SELECT cc.*, c.name,c.cnpj,c.active,c.email,c.email_cc,
            c.bill_complementaries,c.unit_id,
            (SELECT COUNT(*) FROM exam_items e WHERE e.competency_id=cc.competency_id
             AND e.company_id=cc.company_id AND e.status='SEM_PRECO') unpriced
            FROM competency_companies cc JOIN companies c ON c.id=cc.company_id
            WHERE cc.competency_id=? AND c.unit_id=? AND edge_email_key(c.email)=?
            AND NOT EXISTS (SELECT 1 FROM billing_group_members m WHERE m.company_id=c.id)
            ORDER BY c.id""", (competency_id, comp["unit_id"], normalize_email(company["email"]))).fetchall()
    rows = [row for row in all_rows if _candidate(row, billing, email_type, force)]
    owner = min((row["company_id"] for row in rows), default=company_id)
    return dict(group=None, representative_id=owner, rows=rows, all_rows=all_rows,
                billing=billing, delivery_kind="recipient_batch" if len(rows) > 1 else "individual")


def representative_id(conn, competency_id, company_id, email_type, force=False):
    return package_context(conn, competency_id, company_id, email_type, force)["representative_id"]


def delivery_ids(conn, competency_id, company_id, email_type, force=False):
    ctx = package_context(conn, competency_id, company_id, email_type, force)
    if ctx["group"]:
        return billing_groups.delivery_ids(conn, competency_id, company_id, email_type, force)
    return [row["company_id"] for row in ctx["rows"]]


def send_reason(api, conn, competency_id, company_id, email_type, force=False):
    ctx = package_context(conn, competency_id, company_id, email_type, force)
    if ctx["group"]:
        reason = billing_groups.send_reason(api, conn, competency_id, company_id, email_type, force)
        if reason:
            return reason
        owner = conn.execute("SELECT email,email_cc FROM companies WHERE id=?", (ctx["representative_id"],)).fetchone()
        try:
            normalize_cc(api.valid_email, owner["email"], [owner["email_cc"]])
        except ValueError as exc:
            return str(exc)
        return None
    comp = conn.execute("SELECT * FROM competencies WHERE id=?", (competency_id,)).fetchone()
    if not comp:
        return "Competência não encontrada."
    if api.competency_is_closed(comp) and not email_type.endswith("_REMINDER"):
        return "A competência está fechada."
    if not ctx["rows"]:
        return "Nenhuma cobrança não paga está elegível para este envio."
    for row in ctx["rows"]:
        label = row["name"]
        if row["unit_id"] != comp["unit_id"]:
            return f"{label}: a empresa pertence a outra unidade."
        if not api.valid_email(normalize_email(row["email"])):
            return f"{label}: e-mail principal ausente ou inválido."
        if billing_policy.pending(row, ctx["billing"]):
            return f"{label}: confira e confirme o valor na prévia antes de enviar."
        if ctx["billing"] == "COMPLEMENTARY" and billing_policy.unpriced_blocks(row):
            return f"{label}: há exames sem preço. Resolva antes de enviar o pacote deste destinatário."
        try:
            normalize_cc(api.valid_email, row["email"], [row["email_cc"]])
        except ValueError as exc:
            return f"{label}: {exc}"
        docs = conn.execute("SELECT stored_name FROM documents WHERE competency_id=? AND company_id=? AND billing_type=?",
                            (competency_id, row["company_id"], ctx["billing"])).fetchall()
        if not docs:
            return f"{label}: documentos não importados. Complete o pacote deste destinatário antes de enviar."
        if any(not api.stored_file_path(doc["stored_name"]).is_file() for doc in docs):
            return f"{label}: um documento está ausente no disco. Reimporte antes de enviar o pacote deste destinatário."
    return None


def eligible_ids(api, conn, competency_id, email_type):
    if email_type not in billing_groups.EMAIL_TYPES:
        raise ValueError("Tipo de envio inválido.")
    seen = set()
    ids = []
    flag = "consolidate_fixed" if email_type.startswith("FIXED") else "consolidate_complementary"
    for row in conn.execute(f"""SELECT cc.company_id,c.unit_id,c.email,m.group_id,g.{flag} consolidate
        FROM competency_companies cc JOIN companies c ON c.id=cc.company_id
        LEFT JOIN billing_group_members m ON m.company_id=c.id
        LEFT JOIN billing_groups g ON g.id=m.group_id
        WHERE cc.competency_id=? ORDER BY c.name,c.id""", (competency_id,)):
        key = ("explicit", row["group_id"]) if row["group_id"] and row["consolidate"] else (
            ("registered_individual", row["company_id"]) if row["group_id"] else
            ("recipient", row["unit_id"], normalize_email(row["email"])))
        if key in seen:
            continue
        seen.add(key)
        owner = representative_id(conn, competency_id, row["company_id"], email_type)
        if send_reason(api, conn, competency_id, owner, email_type) is None:
            ids.append(owner)
    return ids


def _attachments(api, conn, competency_id, member, billing, batched):
    attachments = []
    used_names = set()
    for doc in conn.execute("SELECT * FROM documents WHERE competency_id=? AND company_id=? AND billing_type=? ORDER BY id",
                            (competency_id, member["company_id"], billing)):
        path = api.stored_file_path(doc["stored_name"])
        if not path.is_file():
            continue
        name = doc["original_name"]
        if batched:
            company_label = re.sub(r'[\\/:*?"<>|\r\n]+', "_", member["name"]).strip()[:80]
            name = f"{member['company_id']} - {company_label} - {name}"
            if name.casefold() in used_names:
                # Dois documentos da mesma empresa podem ter o mesmo nome
                # original (por exemplo, dois boletos importados de pastas).
                original = Path(name)
                counter = 2
                while name.casefold() in used_names:
                    name = f"{original.stem} ({counter}){original.suffix}"
                    counter += 1
            used_names.add(name.casefold())
        attachments.append(dict(path=str(path), filename=name, company_id=member["company_id"], document_id=doc["id"]))
    return attachments


def build_payload(api, conn, competency_id, company_id, email_type, company_ids=None):
    ctx = package_context(conn, competency_id, company_id, email_type)
    if ctx["group"]:
        payload = billing_groups.build_payload(api, conn, competency_id, company_id, email_type, company_ids)
        if payload:
            payload.update(batched=False, recipient_batched=False, delivery_kind="explicit_group",
                           attachment_company_ids=[payload["primary_company_id"]])
        return payload
    by_id = {row["company_id"]: row for row in ctx["all_rows"]}
    if company_ids is None:
        ids = [row["company_id"] for row in ctx["rows"]]
        if not ids:
            # Preserva a consulta da prévia depois do envio/baixa; send_reason
            # continua exigindo cobranças pendentes para permitir novo envio.
            field = "fixed_amount" if ctx["billing"] == "FIXED" else "complementary_amount"
            ids = [row["company_id"] for row in ctx["all_rows"] if row["active"] and float(row[field] or 0) > 0]
    else:
        ids = sorted(set(int(cid) for cid in company_ids))
    if any(cid not in by_id for cid in ids):
        raise ValueError("Os participantes ou destinatários do pacote mudaram. Revise e crie uma nova fila.")
    if not ids:
        payload = api._build_individual_email(conn, competency_id, company_id, email_type)
        if payload:
            payload.update(company_ids=[], members=[], primary_company_id=company_id, group_name=None,
                           grouped=False, batched=False, recipient_batched=False, delivery_kind="individual",
                           attachment_company_ids=[company_id])
        return payload
    owner = min(ids)
    payload = api._build_individual_email(conn, competency_id, owner, email_type)
    if not payload:
        return None
    prefix = "fixed" if ctx["billing"] == "FIXED" else "complementary"
    members = [dict(company_id=cid, name=by_id[cid]["name"], cnpj=by_id[cid]["cnpj"],
                    amount=float(by_id[cid][f"{prefix}_amount"] or 0)) for cid in ids]
    batched = len(ids) > 1
    row = dict(payload["row"])
    row["email"] = normalize_email(row["email"])
    # Prévia exibe também cópias inválidas; send_reason bloqueia o envio e
    # identifica a empresa que precisa de correção, sem descartá-las em silêncio.
    row["email_cc"] = "; ".join(normalize_cc(api.valid_email, row["email"],
                                 [by_id[cid]["email_cc"] for cid in ids], strict=False))
    payload.update(row=row, company_ids=ids, members=members, group_name=None,
                   primary_company_id=owner, grouped=False, batched=batched, recipient_batched=batched,
                   delivery_kind="recipient_batch" if batched else "individual", attachment_company_ids=ids,
                   attachments=[attachment for member in members for attachment in
                                _attachments(api, conn, competency_id, member, ctx["billing"], batched)])
    payload["delivery_label"] = f"{len(members)} empresas → {row['email']}" if batched else row["name"]
    if not batched:
        return payload
    total = sum((Decimal(str(member["amount"])) for member in members), Decimal("0")).quantize(Decimal("0.01"))
    payload["amount"] = float(total)
    label = api.month_label(row["month"], row["year"])
    reminder = email_type.endswith("_REMINDER")
    title = "MENSALIDADES" if ctx["billing"] == "FIXED" else "EXAMES COMPLEMENTARES"
    payload["subject"] = f"{'LEMBRETE DE PAGAMENTO - ' if reminder else ''}{title} - {label} - {len(members)} EMPRESAS"
    lines = "".join(f"<tr><td style='padding:8px'>{html.escape(member['name'])}<br><small>{html.escape(api.format_document(member['cnpj']))}</small></td>"
                    f"<td style='padding:8px;text-align:right'>{html.escape(api.money(member['amount']))}</td></tr>" for member in members)
    details = []
    text_details = []
    if ctx["billing"] == "COMPLEMENTARY":
        for member in members:
            summary = api.complementary_summary(conn, competency_id, member["company_id"])
            details.append(f"<h3>{html.escape(member['name'])}</h3><ul>" + "".join(
                f"<li>{html.escape(item['exam_name'])}: {item['qty']} × {html.escape(api.money(item['unit_price']))} = {html.escape(api.money(item['total']))}</li>"
                for item in summary) + "</ul>")
            text_details.extend(f"{member['name']} - {item['exam_name']}: {item['qty']} x {api.money(item['unit_price'])} = {api.money(item['total'])}" for item in summary)
    action = "Lembramos as cobranças pendentes" if reminder else "Encaminhamos as cobranças"
    signature = api.setting_get("email_signature", "EDGE Saúde Ocupacional")
    payload["html"] = ("<div style='font-family:Arial,sans-serif;color:#243447;line-height:1.5'><p>Prezados,</p>"
        f"<p>{action} de {title.lower()}, competência <strong>{html.escape(label)}</strong>, das empresas abaixo.</p>"
        f"<table style='width:100%;border-collapse:collapse'><thead><tr><th style='text-align:left'>Empresa</th><th style='text-align:right'>Valor</th></tr></thead><tbody>{lines}</tbody></table>"
        f"<p style='font-size:18px'><strong>Total das cobranças: {html.escape(api.money(total))}</strong></p>{''.join(details)}"
        "<p>Seguem anexos os documentos próprios de cada empresa, identificados pelo nome da empresa.</p>"
        f"<p>Atenciosamente,<br>{html.escape(signature).replace(chr(10), '<br>')}</p></div>")
    payload["text"] = (f"Prezados,\n\n{action} de {title.lower()} - {label}\n\n" +
        "\n".join(f"{member['name']} ({api.format_document(member['cnpj'])}): {api.money(member['amount'])}" for member in members) +
        f"\nTotal das cobranças: {api.money(total)}\n\n" + "\n".join(text_details) +
        f"\n\nSeguem os documentos próprios de cada empresa, identificados pelo nome da empresa.\n{signature}")
    return payload
