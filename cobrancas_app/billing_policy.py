"""Origens de complementares e definição manual de valores por competência."""
import html
import re


def migrate(conn):
    additions = {
        'companies': {'fixed_value_required': 'INTEGER NOT NULL DEFAULT 0', 'complementary_value_required': 'INTEGER NOT NULL DEFAULT 0'},
        'competency_companies': {'fixed_value_pending': 'INTEGER NOT NULL DEFAULT 0', 'complementary_value_pending': 'INTEGER NOT NULL DEFAULT 0', 'complementary_amount_manual': 'INTEGER NOT NULL DEFAULT 0'},
        'exam_items': {'source_document': 'TEXT', 'source_company': 'TEXT'},
    }
    for table, columns in additions.items():
        existing = {r[1] for r in conn.execute('PRAGMA table_info('+table+')')}
        for name, declaration in columns.items():
            if name not in existing:conn.execute(f'ALTER TABLE {table} ADD COLUMN {name} {declaration}')
    conn.execute('''CREATE TABLE IF NOT EXISTS complementary_sources (
        unit_id INTEGER NOT NULL REFERENCES units(id),
        document TEXT NOT NULL,
        mother_id INTEGER NOT NULL REFERENCES companies(id) ON DELETE CASCADE,
        name TEXT NOT NULL DEFAULT '',
        PRIMARY KEY(unit_id,document))''')


def save_company(api, conn, company_id, form):
    company = conn.execute('SELECT * FROM companies WHERE id=?',(company_id,)).fetchone()
    if 'billing_policy_present' in form:
        reason = api.delivery_mutation_reason(conn, unit_id=company['unit_id'])
        if reason:raise ValueError(reason)
        conn.execute('UPDATE companies SET fixed_value_required=?,complementary_value_required=? WHERE id=?',
            (int(bool(form.get('fixed_value_required'))), int(bool(form.get('complementary_value_required'))), company_id))
        parsed = {}
        for line in str(form.get('complementary_sources') or '').splitlines():
            if not line.strip():continue
            doc_text, _, name = line.partition('|')
            document = api.digits(doc_text)
            if len(document) not in (11,14):raise ValueError('Informe um CNPJ/CPF por linha em Origens dos complementares. Use CNPJ | nome opcional.')
            if document == api.digits(company['cnpj']):raise ValueError('O CNPJ da própria empresa já está incluído. Informe somente outros CNPJs.')
            if document in parsed:raise ValueError('CNPJ/CPF repetido na lista de origens: '+document)
            other = conn.execute('SELECT mother_id FROM complementary_sources WHERE unit_id=? AND document=?',(company['unit_id'],document)).fetchone()
            if other and other['mother_id'] != company_id:raise ValueError('Este CNPJ/CPF já está vinculado a outra empresa responsável nesta unidade: '+document)
            if conn.execute('SELECT 1 FROM companies WHERE unit_id=? AND cnpj=? AND id<>?',(company['unit_id'],document,company_id)).fetchone():
                raise ValueError('Este CNPJ/CPF já tem cadastro próprio. Para empresas cadastradas, use Cobranças agrupadas: '+document)
            parsed[document] = name.strip()[:200]
        if parsed and not company['bill_complementaries']:raise ValueError('Habilite complementares na empresa responsável para vincular outros CNPJs.')
        conn.execute('DELETE FROM complementary_sources WHERE mother_id=?',(company_id,))
        conn.executemany('INSERT INTO complementary_sources(unit_id,document,mother_id,name) VALUES(?,?,?,?)',
            [(company['unit_id'],doc,company_id,name) for doc,name in parsed.items()])
        # A modalidade passa a ficar pendente apenas nos lançamentos ainda abertos.
        for prefix in ('fixed','complementary'):
            required = int(bool(form.get(prefix+'_value_required')))
            conn.execute(f'''UPDATE competency_companies SET {prefix}_value_pending=CASE WHEN {prefix}_amount_manual=1 THEN 0 ELSE ? END
                WHERE company_id=? AND {prefix}_sent_at IS NULL AND {prefix}_paid=0
                AND competency_id IN (SELECT id FROM competencies WHERE status<>'FECHADA')''',(required,company_id))


def sources_text(conn, company_id):
    return '\n'.join(r['document'] + (' | '+r['name'] if r['name'] else '') for r in conn.execute(
        'SELECT document,name FROM complementary_sources WHERE mother_id=? ORDER BY document',(company_id,)))


def pending(row, billing):
    prefix='fixed' if billing=='FIXED' else 'complementary'
    return bool(dict(row).get(prefix+'_value_pending'))


def unpriced_blocks(row):
    return bool(row['unpriced']) and not bool(dict(row).get('complementary_amount_manual'))


def enrich_email(api, conn, competency_id, payload):
    if not payload or payload['billing'] not in {'COMPLEMENTARY','COMBINED'}:return payload
    ids=payload.get('company_ids') or [payload['row']['company_id']]
    sections=[];texts=[]
    for cid in ids:
        if payload['billing']=='COMBINED' and 'COMPLEMENTARY' not in payload['financial_components'].get(str(cid),{}):continue
        cc=conn.execute('SELECT cc.*,c.name,c.cnpj FROM competency_companies cc JOIN companies c ON c.id=cc.company_id WHERE competency_id=? AND company_id=?',(competency_id,cid)).fetchone()
        rows=conn.execute('''SELECT COALESCE(NULLIF(source_document,''),?) document,
            MAX(COALESCE(NULLIF(source_company,''),?)) name,COUNT(*) qty,SUM(COALESCE(total,0)) total
            FROM exam_items WHERE competency_id=? AND company_id=? AND status IN ('OK','SEM_PRECO') GROUP BY 1 ORDER BY 1''',
            (cc['cnpj'],cc['name'],competency_id,cid)).fetchall()
        if not rows and not cc['complementary_amount_manual']:continue
        title='Complementares sob responsabilidade de '+cc['name']
        text=title+'\n'
        detail='<h3>'+html.escape(title)+'</h3><p>Os atendimentos dos CNPJs abaixo estão incluídos na cobrança da empresa responsável.</p><table style="width:100%;border-collapse:collapse"><tr><th style="text-align:left">Origem dos atendimentos</th><th>Qtd.</th><th>Valor da planilha</th></tr>'
        for row in rows:
            label=(row['name'] or '')+' — '+api.format_document(row['document'])
            detail+='<tr><td>'+html.escape(label)+'</td><td>'+str(row['qty'])+'</td><td>'+html.escape(api.money(row['total']))+'</td></tr>'
            text+=f"{label}: {row['qty']} atendimento(s), {api.money(row['total'])}\n"
        detail+='</table>'
        if cc['complementary_amount_manual']:
            note='Valor final definido manualmente para esta competência: '+api.money(cc['complementary_amount'])+'. Os valores dos atendimentos acima servem para discriminar a origem, sem acréscimo ao total final.'
            detail+='<p><strong>'+html.escape(note)+'</strong></p>';text+=note
        sections.append(detail);texts.append(text)
    if sections:
        block=''.join(sections)
        if '<p>Atenciosamente,' in payload['html']:
            payload['html']=payload['html'].replace('<p>Atenciosamente,',block+'<p>Atenciosamente,',1)
        else:payload['html']+=block
        if '\n\nSeguem' in payload['text']:
            payload['text']=payload['text'].replace('\n\nSeguem','\n\n'+'\n\n'.join(texts)+'\n\nSeguem',1)
        else:payload['text']+='\n\n'+'\n\n'.join(texts)
    return payload
