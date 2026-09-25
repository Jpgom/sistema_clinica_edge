import fs from 'node:fs/promises';
import { FileBlob, SpreadsheetFile } from '@oai/artifact-tool';

const dir = 'C:/sites/sistema_clinica_edge/outputs/01a0d409-c7cb-7882-8379-5af976f25090';
const modelPath = 'C:/Users/Usuário/Downloads/MODELO_CADASTRO_EMPRESAS (5).xlsx';
const sourcePath = 'C:/Users/Usuário/Downloads/EMPRESAS BELEM (1) (1).xlsx';
const outputPath = `${dir}/MODELO_CADASTRO_EMPRESAS_preenchido.xlsx`;

const digits = value => String(value ?? '').replace(/\D/g, '');
const clean = value => String(value ?? '').trim();
const requested = (await fs.readFile(`${dir}/requested_ids.txt`, 'utf8'))
  .split(/\r?\n/)
  .map(s => s.trim())
  .filter(Boolean);

const sourceBook = await SpreadsheetFile.importXlsx(await FileBlob.load(sourcePath));
const source = sourceBook.worksheets.getItem('EMPRESAS');
const values = source.getRange('A1:G108').values;
const expectedHeaders = ['UNIDADE', 'CNPJ/CPF', 'EMPRESA', 'EMAIL', 'EMAIL_CC', 'VALOR_FIXO_MENSAL', 'ATIVA'];
if (expectedHeaders.some((h, i) => values[0][i] !== h)) {
  throw new Error('Cabeçalhos da planilha de origem não correspondem aos esperados.');
}
const records = values.slice(1);
if (requested.length !== 107 || records.length !== requested.length) {
  throw new Error(`Quantidade de linhas inesperada: pedido ${requested.length}, origem ${records.length}.`);
}
for (let i = 0; i < records.length; i++) {
  if (digits(requested[i]) !== digits(records[i][1])) {
    throw new Error(`Identificador diverge na posição ${i + 1}: ${requested[i]} vs. ${records[i][1]}`);
  }
}

const namesById = new Map();
for (const record of records) {
  const key = digits(record[1]);
  const name = clean(record[2]);
  if (!name) continue;
  if (!namesById.has(key)) namesById.set(key, new Set());
  namesById.get(key).add(name);
}

const emailPattern = /[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}/g;
function emailsIn(value) {
  const raw = clean(value);
  if (!raw) return [];
  const matches = [...raw.matchAll(emailPattern)].map(match => match[0]);
  const separators = raw.replace(emailPattern, '').replace(/[\s/;-]/g, '');
  if (separators) throw new Error(`Texto de e-mail não reconhecido: ${raw}`);
  return matches;
}
function splitEmails(main, cc) {
  const all = [...emailsIn(main), ...emailsIn(cc)];
  const seen = new Set();
  const unique = all.filter(email => {
    const key = email.toLowerCase();
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
  return [unique[0] ?? null, unique.length > 1 ? unique.slice(1).join(';') : null];
}

let filledNames = 0;
let missingEmails = 0;
const outputRows = records.map((record, i) => {
  const key = digits(requested[i]);
  let name = clean(record[2]);
  if (!name) {
    const choices = namesById.get(key);
    if (!choices || choices.size !== 1) throw new Error(`EMPRESA ausente e ambígua: ${requested[i]}`);
    name = [...choices][0];
    filledNames++;
  }
  const [email, cc] = splitEmails(record[3], record[4]);
  if (!email) missingEmails++;
  const active = clean(record[6]);
  if (!active) throw new Error(`ATIVA ausente na linha ${i + 2} da origem.`);
  return [requested[i], name, email, cc, active];
});

const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(modelPath));
const sheet = workbook.worksheets.getItem('EMPRESAS');
const headers = sheet.getRange('A1:E1').values[0];
if (JSON.stringify(headers) !== JSON.stringify(['CNPJ', 'EMPRESA', 'EMAIL', 'EMAIL_CC', 'ATIVO'])) {
  throw new Error('Cabeçalhos do modelo inesperados.');
}
if (sheet.tables.items.length !== 0) throw new Error('Tabela nativa inesperada no modelo.');
sheet.getRange('A2:E108').values = outputRows;
sheet.getRange('B1:B108').format.columnWidth = 82;
sheet.getRange('C1:C108').format.columnWidth = 46;
sheet.getRange('D1:D108').format.columnWidth = 68;
workbook.recalculate();

const actual = sheet.getRange('A2:E108').values;
for (let i = 0; i < outputRows.length; i++) {
  for (let j = 0; j < 5; j++) {
    if ((actual[i]?.[j] ?? null) !== (outputRows[i][j] ?? null)) {
      throw new Error(`Falha ao gravar célula ${i + 2},${j + 1}: ${actual[i]?.[j]} vs. ${outputRows[i][j]}`);
    }
  }
}
const top = await workbook.inspect({ kind: 'table', range: 'EMPRESAS!A1:E6', include: 'values,formulas', tableMaxRows: 6, tableMaxCols: 5, maxChars: 3500 });
console.log(top.ndjson);
const errors = await workbook.inspect({
  kind: 'match',
  searchTerm: '#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!',
  options: { useRegex: true, maxResults: 300 },
  summary: 'erros finais',
  maxChars: 1500,
});
console.log(errors.ndjson);

for (const [range, file] of [
  ['A1:E8', 'preview-top.png'],
  ['A63:E70', 'preview-empty-email.png'],
]) {
  const blob = await workbook.render({ sheetName: 'EMPRESAS', range, scale: 1.5, format: 'png' });
  await fs.writeFile(`${dir}/${file}`, new Uint8Array(await blob.arrayBuffer()));
}

const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);
console.log(JSON.stringify({
  outputPath, rows: outputRows.length, filledNames, missingEmails,
  withEmailCc: outputRows.filter(r => r[3]).length,
  uniqueIds: new Set(requested.map(digits)).size,
}));
