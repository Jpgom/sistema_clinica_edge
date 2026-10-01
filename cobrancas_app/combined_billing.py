"""Um envelope reúne as modalidades pendentes, preservando seus lançamentos."""
from decimal import Decimal
import html

import delivery_batches
from mail_transport import normalize_cc


def envelopes(conn, competency_id):
    companies = conn.execute('''SELECT cc.company_id,c.email,m.group_id FROM competency_companies cc
        JOIN companies c ON c.id=cc.company_id LEFT JOIN billing_group_members m ON m.company_id=c.id
        WHERE cc.competency_id=? ORDER BY c.id''', (competency_id,)).fetchall()
    memberships = {r['company_id']: r['group_id'] for r in companies}
    packets = {}
    result = {}
    for billing in ('FIXED', 'COMPLEMENTARY'):
        for company in companies:
            ctx = delivery_batches.package_context(conn, competency_id, company['company_id'], billing)
            if not ctx['rows']:
                continue
            owner = ctx['representative_id']
            packet_key = (billing, owner)
            if packet_key in packets:
                continue
            packets[packet_key] = True
            ids = sorted(delivery_batches.delivery_ids(conn, competency_id, owner, billing))
            owner_row = conn.execute('SELECT * FROM companies WHERE id=?', (owner,)).fetchone()
            key = (memberships.get(owner), delivery_batches.normalize_email(owner_row['email']))
            envelope = result.setdefault(key, {'owner': owner, 'packets': [], 'ids': set(), 'email': key[1], 'conflict': None})
            envelope['owner'] = min(envelope['owner'], owner)
            envelope['ids'].update(ids)
            envelope['packets'].append(dict(ctx, ids=ids))
    # Modalidades configuradas para destinatários diferentes exigem correção
    # explícita: não enviar uma parte isoladamente nem trocar o destinatário.
    destinations = {}
    for key, envelope in result.items():
        for cid in envelope['ids']:
            destinations.setdefault(cid, set()).add(key)
    for envelope in result.values():
        envelope['ids'] = sorted(envelope['ids'])
        if any(len(destinations[cid]) > 1 for cid in envelope['ids']):
            envelope['conflict'] = ('Mensalidade e complementares estão configurados para destinatários diferentes no grupo. '
                'Ajuste os e-mails ou configure ambas as modalidades para a mesma responsável antes do envio conjunto.')
    return sorted(result.values(), key=lambda e: e['owner'])


def context(conn, competency_id, company_id):
    items = envelopes(conn, competency_id)
    return next((e for e in items if e['owner'] == company_id), None) or next((e for e in items if company_id in e['ids']), None)


def reason(api, conn, competency_id, envelope):
    if not envelope:
        return 'Nenhuma cobrança pendente para este envio.'
    if envelope['conflict']:
        return envelope['conflict']
    for packet in envelope['packets']:
        problem = delivery_batches.send_reason(api, conn, competency_id, packet['representative_id'], packet['billing'])
        if problem:
            label = 'Mensalidade' if packet['billing'] == 'FIXED' else 'Complementares'
            return label + ': ' + problem
    return None


def send_reason(api, conn, competency_id, company_id):
    return reason(api, conn, competency_id, context(conn, competency_id, company_id))


def eligible_ids(api, conn, competency_id):
    return [e['owner'] for e in envelopes(conn, competency_id) if reason(api, conn, competency_id, e) is None]


def build_payload(api, conn, competency_id, company_id, company_ids=None, envelope=None):
    envelope = envelope or context(conn, competency_id, company_id)
    if not envelope:
        return None
    if company_ids is not None and set(company_ids) != set(envelope['ids']):
        raise ValueError('As empresas deste e-mail mudaram. Revise e crie uma nova fila.')
    owner = envelope['owner']
    payload = api._build_individual_email(conn, competency_id, owner, 'FIXED')
    if not payload:
        return None
    row = dict(payload['row'])
    row['email'] = envelope['email']
    components, by_id, refs, cc_values = {}, {}, set(), []
    groups = set()
    for packet in envelope['packets']:
        billing = packet['billing']
        prefix = 'fixed' if billing == 'FIXED' else 'complementary'
        all_rows = {r['company_id']: r for r in packet['all_rows']}
        for cid in packet['ids']:
            member = all_rows[cid]
            by_id[cid] = member
            components.setdefault(str(cid), {})[billing] = float(member[prefix+'_amount'] or 0)
        attachment_ids = [packet['representative_id']] if packet['group'] else packet['ids']
        refs.update((cid, billing) for cid in attachment_ids)
        if packet['group']:
            groups.add(packet['group']['name'])
            cc_values.append(conn.execute('SELECT email_cc FROM companies WHERE id=?', (packet['representative_id'],)).fetchone()[0])
        else:
            cc_values.extend(all_rows[cid]['email_cc'] for cid in packet['ids'])
    row['email_cc'] = '; '.join(normalize_cc(api.valid_email, row['email'], cc_values, strict=False))
    members = []
    for cid in envelope['ids']:
        values = components[str(cid)]
        fixed = values.get('FIXED', 0)
        complementary = values.get('COMPLEMENTARY', 0)
        members.append(dict(company_id=cid, name=by_id[cid]['name'], cnpj=by_id[cid]['cnpj'],
            fixed_amount=fixed, complementary_amount=complementary,
            amount=float(Decimal(str(fixed))+Decimal(str(complementary)))))
    attachments = []
    for cid, billing in sorted(refs):
        company = conn.execute('SELECT id company_id,name FROM companies WHERE id=?', (cid,)).fetchone()
        docs = delivery_batches._attachments(api, conn, competency_id, company, billing, True)
        for doc in docs:
            doc['filename'] = ('MENSALIDADE - ' if billing == 'FIXED' else 'COMPLEMENTARES - ') + doc['filename']
        attachments.extend(docs)
    total = float(sum((Decimal(str(m['amount'])) for m in members), Decimal('0')))
    label = api.month_label(row['month'], row['year'])
    signature = api.smtp_config()['email_signature']
    rows, lines = [], []
    for m in members:
        name = m['name'] + ' — ' + api.format_document(m['cnpj'])
        values = components[str(m['company_id'])]
        fixed = api.money(m['fixed_amount']) if 'FIXED' in values else 'Não incluída neste envio'
        complementary = api.money(m['complementary_amount']) if 'COMPLEMENTARY' in values else 'Não incluídos neste envio'
        rows.append('<tr><td style="padding:10px">'+html.escape(name)+'</td><td style="padding:10px">'+html.escape(fixed)+
            '</td><td style="padding:10px">'+html.escape(complementary)+'</td><td style="padding:10px">'+html.escape(api.money(m['amount']))+'</td></tr>')
        lines.append(name+'\nMensalidade comum: '+fixed+'\nComplementares: '+complementary+'\nTotal neste envio: '+api.money(m['amount']))
    note = 'Somente os valores incluídos neste envio compõem o total a pagar. Cobranças já enviadas ou pagas permanecem no histórico.'
    group_note = ('Nas modalidades agrupadas, os documentos são emitidos em nome da responsável. Participantes com R$ 0,00 já estão abrangidas pelo valor do grupo.' if groups else
        'Seguem os documentos de cada empresa, identificados por empresa e modalidade.')
    payload.update(row=row, billing='COMBINED', title='COBRANÇA DA COMPETÊNCIA', combined=True,
        subject=f"COBRANÇA - {label} - {row['name'] if len(members)==1 else str(len(members))+' EMPRESAS'}",
        amount=total, company_ids=envelope['ids'], members=members, financial_components=components,
        document_refs=[list(ref) for ref in sorted(refs)], attachment_company_ids=sorted({cid for cid, _ in refs}),
        primary_company_id=owner, group_name=' · '.join(sorted(groups)) or None, grouped=bool(groups),
        batched=len(members)>1, recipient_batched=len(members)>1 and not groups,
        delivery_kind='combined', delivery_label=f"{len(members)} empresa(s) → {row['email']}", attachments=attachments)
    payload['html'] = ('<div style="font-family:Arial,sans-serif;color:#243447;line-height:1.5"><p>Prezados,</p>'
        f'<p>Encaminhamos em um único e-mail as cobranças pendentes da competência <strong>{html.escape(label)}</strong>.</p>'
        '<table style="width:100%;border-collapse:collapse"><thead><tr style="background:#e7eff5"><th>Empresa / CNPJ</th>'
        '<th>Mensalidade comum</th><th>Complementares</th><th>Total neste envio</th></tr></thead><tbody>'+''.join(rows)+'</tbody></table>'
        f'<p><strong>Total a pagar neste envio: {html.escape(api.money(total))}</strong></p><p>{note}</p><p>{group_note}</p>'
        f"<p>Atenciosamente,<br>{html.escape(signature).replace(chr(10), '<br>')}</p></div>")
    payload['text'] = f'Prezados,\n\nCobranças pendentes da competência {label}, reunidas em um único e-mail.\n\n'+'\n\n'.join(lines)+f'\n\nTotal a pagar neste envio: {api.money(total)}\n{note}\n{group_note}\n\n{signature}'
    return payload
