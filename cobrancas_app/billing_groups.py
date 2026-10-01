"""Agrupamento explícito de cobranças, sem inferir grupos pelo destinatário."""
import billing_policy
from decimal import Decimal
from pathlib import Path
import html
import json

EMAIL_TYPES = {"FIXED", "COMPLEMENTARY", "FIXED_REMINDER", "COMPLEMENTARY_REMINDER"}


def migrate_billing_groups(conn):
    conn.executescript("""
        CREATE TABLE IF NOT EXISTS billing_groups (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            unit_id INTEGER NOT NULL REFERENCES units(id) ON DELETE CASCADE,
            name TEXT NOT NULL,
            primary_company_id INTEGER NOT NULL REFERENCES companies(id) ON DELETE CASCADE,
            consolidate_fixed INTEGER NOT NULL DEFAULT 1,
            consolidate_complementary INTEGER NOT NULL DEFAULT 1,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL,
            UNIQUE(unit_id, name)
        );
        CREATE TABLE IF NOT EXISTS billing_group_members (
            group_id INTEGER NOT NULL REFERENCES billing_groups(id) ON DELETE CASCADE,
            company_id INTEGER NOT NULL UNIQUE REFERENCES companies(id) ON DELETE CASCADE,
            PRIMARY KEY(group_id, company_id)
        );
        CREATE INDEX IF NOT EXISTS idx_billing_groups_unit ON billing_groups(unit_id);
    """)


def context(conn, competency_id, company_id, email_type):
    if email_type not in EMAIL_TYPES:
        raise ValueError("Tipo de envio inválido.")
    billing = "FIXED" if email_type.startswith("FIXED") else "COMPLEMENTARY"
    comp = conn.execute("SELECT * FROM competencies WHERE id=?", (competency_id,)).fetchone()
    if not comp:
        return None, company_id, [], billing
    flag = "consolidate_fixed" if billing == "FIXED" else "consolidate_complementary"
    group = conn.execute(f"""SELECT g.* FROM billing_groups g
        JOIN billing_group_members m ON m.group_id=g.id
        WHERE m.company_id=? AND g.unit_id=? AND g.{flag}=1""", (company_id, comp["unit_id"])).fetchone()
    owner = group["primary_company_id"] if group else company_id
    if group:
        rows = conn.execute("""SELECT cc.*, c.name, c.cnpj, c.active, c.email, c.email_cc,
            c.bill_complementaries, c.unit_id,
            (SELECT COUNT(*) FROM exam_items e WHERE e.competency_id=cc.competency_id
             AND e.company_id=cc.company_id AND e.status='SEM_PRECO') unpriced
            FROM competency_companies cc JOIN companies c ON c.id=cc.company_id
            JOIN billing_group_members m ON m.company_id=c.id
            WHERE cc.competency_id=? AND m.group_id=? ORDER BY c.id""", (competency_id, group["id"])).fetchall()
    else:
        rows = conn.execute("""SELECT cc.*, c.name, c.cnpj, c.active, c.email, c.email_cc,
            c.bill_complementaries, c.unit_id,
            (SELECT COUNT(*) FROM exam_items e WHERE e.competency_id=cc.competency_id
             AND e.company_id=cc.company_id AND e.status='SEM_PRECO') unpriced
            FROM competency_companies cc JOIN companies c ON c.id=cc.company_id
            WHERE cc.competency_id=? AND cc.company_id=?""", (competency_id, company_id)).fetchall()
    return group, owner, rows, billing


def outstanding_rows(conn, competency_id, company_id, email_type, force=False):
    group, owner, rows, billing = context(conn, competency_id, company_id, email_type)
    prefix = "fixed" if billing == "FIXED" else "complementary"
    reminder = email_type.endswith("_REMINDER")
    candidates = []
    for r in rows:
        if not r["active"] or (billing == "COMPLEMENTARY" and not r["bill_complementaries"]):
            continue
        if r[f"{prefix}_paid"]:
            continue
        if reminder and not r[f"{prefix}_sent_at"]:
            continue
        if not reminder and force and not r[f"{prefix}_sent_at"]:
            continue
        if not reminder and not force and r[f"{prefix}_sent_at"]:
            continue
        # SEM_PRECO também bloqueia quando o subtotal ainda é zero.
        if Decimal(str(r[f"{prefix}_amount"] or 0)) <= 0 and not billing_policy.pending(r, billing) and not (billing == "COMPLEMENTARY" and billing_policy.unpriced_blocks(r)):
            continue
        candidates.append(r)
    return group, owner, candidates, billing


def delivery_ids(conn, competency_id, company_id, email_type, force=False):
    group, owner, candidates, billing = outstanding_rows(conn, competency_id, company_id, email_type, force)
    ids = [r["company_id"] for r in candidates]
    if not group or not candidates:
        return ids

    # Em uma cobrança agrupada, participantes com R$ 0,00 também precisam constar
    # no mesmo recibo/e-mail. Eles não alteram o total, mas registram quais empresas
    # estão abrangidas pela cobrança única.
    _, _, all_rows, _ = context(conn, competency_id, company_id, email_type)
    prefix = "fixed" if billing == "FIXED" else "complementary"
    reminder = email_type.endswith("_REMINDER")
    for row in all_rows:
        if row["company_id"] in ids or not row["active"]:
            continue
        if billing == "COMPLEMENTARY" and not row["bill_complementaries"]:
            continue
        if row[f"{prefix}_paid"]:
            continue
        if Decimal(str(row[f"{prefix}_amount"] or 0)) > 0:
            continue
        if reminder and not row[f"{prefix}_sent_at"]:
            continue
        if not reminder and force and not row[f"{prefix}_sent_at"]:
            continue
        if not reminder and not force and row[f"{prefix}_sent_at"]:
            continue
        ids.append(row["company_id"])
    return ids


def send_reason(api, conn, competency_id, company_id, email_type, force=False):
    group, owner_id, rows, billing = outstanding_rows(conn, competency_id, company_id, email_type, force)
    comp = conn.execute("SELECT * FROM competencies WHERE id=?", (competency_id,)).fetchone()
    if not comp:
        return "Competência não encontrada."
    if api.competency_is_closed(comp) and not email_type.endswith("_REMINDER"):
        return "A competência está fechada."
    if not rows:
        return "Nenhuma cobrança não paga está elegível para este envio."
    owner = conn.execute("SELECT * FROM companies WHERE id=? AND unit_id=?", (owner_id, comp["unit_id"])).fetchone()
    if not owner or not owner["active"]:
        return "A empresa responsável está ausente ou inativa."
    if not api.valid_email(owner["email"]):
        return "E-mail da empresa responsável ausente ou inválido."
    if any(r["unit_id"] != comp["unit_id"] for r in rows):
        return "O grupo possui uma empresa em outra unidade. Corrija o cadastro do grupo."
    if any(billing_policy.pending(r, billing) for r in rows):
        return "Conferir e confirmar o valor de todas as participantes antes do envio."
    if billing == "COMPLEMENTARY" and any(billing_policy.unpriced_blocks(r) for r in rows):
        return "Há exames sem preço em uma das empresas participantes."
    docs = conn.execute("SELECT stored_name FROM documents WHERE competency_id=? AND company_id=? AND billing_type=?",
                        (competency_id, owner_id, billing)).fetchall()
    if not docs:
        return "Documentos da empresa responsável não importados."
    if any(not api.stored_file_path(d["stored_name"]).is_file() for d in docs):
        return "Um documento da empresa responsável está ausente no disco. Reimporte os documentos."
    return None


def eligible_ids(api, conn, competency_id, email_type):
    if email_type not in EMAIL_TYPES:
        raise ValueError("Tipo de envio inválido.")
    company_ids = conn.execute("""SELECT cc.company_id FROM competency_companies cc
        JOIN companies c ON c.id=cc.company_id WHERE cc.competency_id=? ORDER BY c.name""", (competency_id,)).fetchall()
    owners = dict.fromkeys(context(conn, competency_id, r["company_id"], email_type)[1] for r in company_ids)
    return [cid for cid in owners if send_reason(api, conn, competency_id, cid, email_type) is None]


def build_payload(api, conn, competency_id, company_id, email_type, company_ids=None):
    group, owner_id, rows, billing = context(conn, competency_id, company_id, email_type)
    if company_ids is None:
        ids = delivery_ids(conn, competency_id, company_id, email_type)
        # A prévia permanece consultável após o envio ou a baixa.
        if not ids:
            ids = [r["company_id"] for r in rows if r["active"] and (billing != "COMPLEMENTARY" or r["bill_complementaries"])]
    else:
        ids = list(dict.fromkeys(int(cid) for cid in company_ids))
    members_by_id = {r["company_id"]: r for r in rows}
    if any(cid not in members_by_id for cid in ids):
        raise ValueError("Os participantes do grupo mudaram. Crie uma nova fila após revisar a cobrança.")
    payload = api._build_individual_email(conn, competency_id, owner_id, email_type)
    if not payload:
        return None
    prefix = "fixed" if billing == "FIXED" else "complementary"
    members = [{"company_id": cid, "name": members_by_id[cid]["name"], "cnpj": members_by_id[cid]["cnpj"],
                "amount": float(members_by_id[cid][f"{prefix}_amount"] or 0)} for cid in ids]
    payload.update(company_ids=ids, members=members, group_name=group["name"] if group else None,
                   primary_company_id=owner_id, grouped=bool(group))
    if not group:
        return payload
    total = float(sum((Decimal(str(m["amount"])) for m in members), Decimal("0")).quantize(Decimal("0.01")))
    owner = payload["row"]
    label = api.month_label(owner["month"], owner["year"])
    title = "MENSALIDADE" if billing == "FIXED" else "EXAMES COMPLEMENTARES"
    reminder = email_type.endswith("_REMINDER")
    payload["amount"] = total
    payload["subject"] = f"{'LEMBRETE DE PAGAMENTO - ' if reminder else ''}{title} - {label} - {owner['name']}"
    lines = "".join(f"<tr><td style='padding:8px'>{html.escape(m['name'])}<br><small>{html.escape(api.format_document(m['cnpj']))}</small></td>"
                    f"<td style='padding:8px;text-align:right'>{html.escape(api.money(m['amount']))}</td></tr>" for m in members)
    details = ""
    text_details = []
    if billing == "COMPLEMENTARY":
        for member in members:
            summary = api.complementary_summary(conn, competency_id, member["company_id"])
            details += f"<h3>{html.escape(member['name'])}</h3><ul>" + "".join(
                f"<li>{html.escape(r['exam_name'])}: {r['qty']} × {html.escape(api.money(r['unit_price']))} = {html.escape(api.money(r['total']))}</li>" for r in summary) + "</ul>"
            text_details.extend(f"{member['name']} - {r['exam_name']}: {r['qty']} x {api.money(r['unit_price'])} = {api.money(r['total'])}" for r in summary)
    action = "Lembramos que permanece pendente" if reminder else "Encaminhamos"
    payload["html"] = ("<div style='font-family:Arial,sans-serif;color:#243447;line-height:1.5'><p>Prezados,</p>"
        f"<p><strong>Empresa responsável: {html.escape(owner['name'])}</strong><br>CNPJ/CPF: {html.escape(api.format_document(owner['cnpj']))}</p>"
        f"<p>{action} a cobrança agrupada de {title.lower()}, competência <strong>{label}</strong>.</p>"
        f"<p>Esta é uma <strong>cobrança única</strong> que contempla {len(members)} empresa(s). Empresas com valor individual de R$ 0,00 estão incluídas no mesmo recibo, pois o valor correspondente já está concentrado na empresa responsável ou em outra participante do grupo.</p>"
        f"<table style='width:100%;border-collapse:collapse'><thead><tr><th style='text-align:left'>Empresa contemplada</th><th style='text-align:right'>Valor considerado</th></tr></thead><tbody>{lines}</tbody></table>"
        f"<p style='font-size:18px'><strong>Total da cobrança: {html.escape(api.money(total))}</strong></p>{details}"
        f"<p>Seguem anexos os documentos emitidos em nome de {html.escape(owner['name'])}, referentes ao total acima.</p>"
        f"<p>Atenciosamente,<br>{html.escape(api.smtp_config()['email_signature'])}</p></div>")
    payload["text"] = (f"Prezados,\n\nEmpresa responsável: {owner['name']}\nCNPJ/CPF: {api.format_document(owner['cnpj'])}\n"
        f"Cobrança agrupada de {title.lower()} - {label}\n"
        f"Esta cobrança única contempla {len(members)} empresa(s). Participantes com R$ 0,00 permanecem incluídas no mesmo recibo.\n\n" + "\n".join(f"{m['name']} ({api.format_document(m['cnpj'])}): {api.money(m['amount'])}" for m in members)
        + f"\nTotal: {api.money(total)}\n\n" + "\n".join(text_details)
        + f"\n\nSeguem os documentos emitidos em nome de {owner['name']}.\n{api.smtp_config()['email_signature']}")
    # Os anexos continuam exclusivamente os da responsável, sem agregação de arquivos.
    return payload


def register_routes(api):
    from flask import abort, flash, redirect, render_template, request, url_for
    import sqlite3

    def form_context(group_id=None):
        conn = api.db()
        try:
            group = conn.execute("SELECT * FROM billing_groups WHERE id=?", (group_id,)).fetchone() if group_id else None
            if group_id and not group:
                abort(404)
            members = {r["company_id"] for r in conn.execute("SELECT company_id FROM billing_group_members WHERE group_id=?", (group_id,))} if group else set()
            return dict(group=group, selected_members=members,
                units=conn.execute("SELECT * FROM units WHERE active=1 ORDER BY name").fetchall(),
                companies=conn.execute("""SELECT c.*,u.name unit_name,g.name billing_group_name,m.group_id
                    FROM companies c JOIN units u ON u.id=c.unit_id
                    LEFT JOIN billing_group_members m ON m.company_id=c.id
                    LEFT JOIN billing_groups g ON g.id=m.group_id WHERE c.active=1 ORDER BY u.name,c.name""").fetchall())
        finally:
            conn.close()

    @api.app.route("/billing-groups")
    def billing_groups():
        conn = api.db()
        try:
            groups = conn.execute("""SELECT g.*,u.name unit_name,c.name primary_name,c.email,
                (SELECT COUNT(*) FROM billing_group_members m WHERE m.group_id=g.id) member_count
                FROM billing_groups g JOIN units u ON u.id=g.unit_id
                JOIN companies c ON c.id=g.primary_company_id ORDER BY u.name,g.name""").fetchall()
            return render_template("billing_groups.html", groups=groups)
        finally:
            conn.close()

    @api.app.route("/billing-groups/new", methods=["GET", "POST"])
    @api.app.route("/billing-groups/<int:group_id>/edit", methods=["GET", "POST"])
    def billing_group_edit(group_id=None):
        ctx = form_context(group_id)
        if request.method == "POST":
            conn = api.db()
            try:
                conn.execute("BEGIN IMMEDIATE")
                name = str(request.form.get("name") or "").strip()
                unit_id = int(request.form.get("unit_id") or 0)
                primary = int(request.form.get("primary_company_id") or 0)
                ids = sorted({int(cid) for cid in request.form.getlist("company_ids")} | {primary})
                fixed = int(bool(request.form.get("consolidate_fixed")))
                complementary = int(bool(request.form.get("consolidate_complementary")))
                if not name or len(name) > 120:
                    raise ValueError("Informe um nome para o grupo (até 120 caracteres).")
                if not fixed and not complementary:
                    raise ValueError("Escolha mensalidade, complementares ou ambas as modalidades.")
                if len(ids) < 2:
                    raise ValueError("Selecione ao menos duas empresas, incluindo a responsável.")
                marks = ",".join("?" for _ in ids)
                companies = conn.execute(f"SELECT * FROM companies WHERE id IN ({marks})", ids).fetchall()
                if len(companies) != len(ids) or any(c["unit_id"] != unit_id or not c["active"] for c in companies):
                    raise ValueError("Todas as empresas precisam estar ativas e pertencer à unidade escolhida.")
                owner = next(c for c in companies if c["id"] == primary)
                if not api.valid_email(owner["email"]):
                    raise ValueError("Cadastre um e-mail válido na empresa responsável antes de criar o grupo.")
                existing = conn.execute(f"SELECT company_id FROM billing_group_members WHERE company_id IN ({marks}) AND group_id != ?", (*ids, group_id or 0)).fetchall()
                if existing:
                    raise ValueError("Uma empresa selecionada já participa de outro grupo. Edite o grupo atual primeiro.")
                reason=api.delivery_mutation_reason(conn,unit_id=unit_id)
                if reason:raise ValueError(reason)
                stamp = api.now_iso()
                if group_id:
                    if ctx["group"]["unit_id"] != unit_id:
                        raise ValueError("A unidade de um grupo existente não pode ser alterada.")
                    conn.execute("UPDATE billing_groups SET name=?,primary_company_id=?,consolidate_fixed=?,consolidate_complementary=?,updated_at=? WHERE id=?", (name,primary,fixed,complementary,stamp,group_id))
                    conn.execute("DELETE FROM billing_group_members WHERE group_id=?", (group_id,))
                else:
                    group_id = conn.execute("INSERT INTO billing_groups(unit_id,name,primary_company_id,consolidate_fixed,consolidate_complementary,created_at,updated_at) VALUES(?,?,?,?,?,?,?)", (unit_id,name,primary,fixed,complementary,stamp,stamp)).lastrowid
                conn.executemany("INSERT INTO billing_group_members(group_id,company_id) VALUES(?,?)", [(group_id,cid) for cid in ids])
                conn.commit()
                api.audit_event("SALVAR_GRUPO_COBRANCA", "billing_group", group_id, name, {"company_ids":ids,"primary_company_id":primary,"fixed":fixed,"complementary":complementary})
                flash("Grupo salvo. Os documentos e o destinatário serão os da empresa responsável; os valores das participantes serão somados.", "success")
                return redirect(url_for("billing_groups"))
            except (ValueError, sqlite3.IntegrityError) as exc:
                conn.rollback()
                flash("Já existe um grupo com esse nome nesta unidade." if isinstance(exc, sqlite3.IntegrityError) else str(exc), "error")
                ctx.update(submitted=request.form)
            finally:
                conn.close()
        return render_template("billing_group_edit.html", **ctx)

    @api.app.route("/billing-groups/<int:group_id>/delete", methods=["POST"])
    def billing_group_delete(group_id):
        conn = api.db()
        try:
            conn.execute("BEGIN IMMEDIATE")
            group = conn.execute("SELECT * FROM billing_groups WHERE id=?", (group_id,)).fetchone()
            if not group:
                abort(404)
            reason=api.delivery_mutation_reason(conn,unit_id=group["unit_id"])
            if reason:
                flash(reason, "error")
            else:
                conn.execute("DELETE FROM billing_groups WHERE id=?", (group_id,))
                conn.commit()
                api.audit_event("DESFAZER_GRUPO_COBRANCA", "billing_group", group_id, group["name"])
                flash("Grupo desfeito. Os próximos envios serão individuais. O histórico e os pagamentos foram preservados.", "success")
        finally:
            conn.close()
        return redirect(url_for("billing_groups"))
