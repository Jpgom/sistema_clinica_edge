import fs from 'node:fs/promises';
import { FileBlob, SpreadsheetFile } from '@oai/artifact-tool';

const file = 'C:/Users/Usuário/Downloads/MODELO_CADASTRO_EMPRESAS (5).xlsx';
const dir = 'C:/sites/sistema_clinica_edge/outputs/01a0d409-c7cb-7882-8379-5af976f25090';
const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(file));
const summary = await workbook.inspect({ kind: 'workbook,sheet,table', maxChars: 3500, tableMaxRows: 4, tableMaxCols: 6 });
console.log(summary.ndjson);
const image = await workbook.render({ sheetName: 'EMPRESAS', range: 'A1:E3', scale: 2, format: 'png' });
await fs.writeFile(`${dir}/baseline.png`, new Uint8Array(await image.arrayBuffer()));
