import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
// Import default (não namespace): sob tsx/Node, o pacote 'xlsx' (CJS) não
// expõe corretamente XLSX.readFile como named export — só via .default.
import XLSX from 'xlsx';
import { namesMatch } from '../src/shared/nameMatch.ts';
import type { SalariosStore } from '../src/shared/types.ts';
import { parseSalariosSheet, parseExcecoesSheet, isExcecoesSheetName } from './xlsxParsing.ts';

const __dirname = path.dirname(fileURLToPath(import.meta.url));

// Caminho absoluto, resolvido a partir da localização deste arquivo — não
// depende de qual diretório o Windows usou para abrir o programa.
export const DADOS_DIR = process.env.DADOS_DIR
  ? path.resolve(process.env.DADOS_DIR)
  : path.resolve(__dirname, '..', 'dados');

export const SALARIOS_PATH = path.join(DADOS_DIR, 'salarios.xlsx');

const SHEET_SALARIOS = 'Planilha1';
const SHEET_EXCECOES = 'excessão extra';

// Serializa toda leitura/escrita nesse arquivo único para evitar que duas
// operações concorrentes (ex: cadastrar + excluir em sequência rápida)
// leiam o mesmo estado antigo e uma sobrescreva a outra.
let queue: Promise<unknown> = Promise.resolve();
const enqueue = <T>(fn: () => T | Promise<T>): Promise<T> => {
  const result = queue.then(fn);
  queue = result.catch(() => {});
  return result;
};

const ensureFileExists = () => {
  fs.mkdirSync(DADOS_DIR, { recursive: true });
  if (!fs.existsSync(SALARIOS_PATH)) {
    const wb = XLSX.utils.book_new();
    const wsSal = XLSX.utils.aoa_to_sheet([['FUNCIONÁRIOS', 'SALÁRIO']]);
    XLSX.utils.book_append_sheet(wb, wsSal, SHEET_SALARIOS);
    const wsExc = XLSX.utils.aoa_to_sheet([[]]);
    XLSX.utils.book_append_sheet(wb, wsExc, SHEET_EXCECOES);
    XLSX.writeFile(wb, SALARIOS_PATH);
  }
};

const loadWorkbook = (): XLSX.WorkBook => {
  ensureFileExists();
  return XLSX.readFile(SALARIOS_PATH);
};

const findExcecoesSheetName = (wb: XLSX.WorkBook): string | null => {
  return wb.SheetNames.find(isExcecoesSheetName) ?? null;
};

const parseWorkbook = (wb: XLSX.WorkBook): SalariosStore => {
  const salSheetName = wb.SheetNames[0];
  const salRows: any[][] = salSheetName
    ? XLSX.utils.sheet_to_json(wb.Sheets[salSheetName], { header: 1 })
    : [];
  const salarios = parseSalariosSheet(salRows);

  const excSheetName = findExcecoesSheetName(wb);
  const excRows: any[][] = excSheetName
    ? XLSX.utils.sheet_to_json(wb.Sheets[excSheetName], { header: 1 })
    : [];
  const excecoes = parseExcecoesSheet(excRows);

  return { salarios, excecoes };
};

const saveWorkbook = (wb: XLSX.WorkBook) => {
  fs.mkdirSync(DADOS_DIR, { recursive: true });
  XLSX.writeFile(wb, SALARIOS_PATH);
};

export const getStore = (): Promise<SalariosStore> =>
  enqueue(() => parseWorkbook(loadWorkbook()));

export const upsertSalario = (nome: string, valorHora: number): Promise<SalariosStore> =>
  enqueue(() => {
    const wb = loadWorkbook();
    const salSheetName = wb.SheetNames[0];
    const ws = wb.Sheets[salSheetName];
    const rows: any[][] = XLSX.utils.sheet_to_json(ws, { header: 1 });
    if (rows.length === 0) rows.push(['FUNCIONÁRIOS', 'SALÁRIO']);

    let updated = false;
    for (let r = 1; r < rows.length; r++) {
      const existing = String(rows[r]?.[0] || '');
      if (existing && namesMatch(existing, nome)) {
        rows[r][1] = valorHora;
        updated = true;
        break;
      }
    }
    if (!updated) {
      rows.push([nome, valorHora]);
    }

    wb.Sheets[salSheetName] = XLSX.utils.aoa_to_sheet(rows);
    saveWorkbook(wb);
    return parseWorkbook(wb);
  });

export const deleteSalario = (nome: string): Promise<SalariosStore> =>
  enqueue(() => {
    const wb = loadWorkbook();
    const salSheetName = wb.SheetNames[0];
    const ws = wb.Sheets[salSheetName];
    const rows: any[][] = XLSX.utils.sheet_to_json(ws, { header: 1 });
    const header = rows[0] || ['FUNCIONÁRIOS', 'SALÁRIO'];
    const kept = rows.slice(1).filter(row => !(row?.[0] && namesMatch(String(row[0]), nome)));

    wb.Sheets[salSheetName] = XLSX.utils.aoa_to_sheet([header, ...kept]);
    saveWorkbook(wb);
    return parseWorkbook(wb);
  });

export const addExcecao = (nome: string): Promise<SalariosStore> =>
  enqueue(() => {
    const wb = loadWorkbook();
    let excSheetName = findExcecoesSheetName(wb);
    if (!excSheetName) {
      excSheetName = SHEET_EXCECOES;
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([[]]), excSheetName);
    }
    const ws = wb.Sheets[excSheetName];
    const rows: any[][] = XLSX.utils.sheet_to_json(ws, { header: 1 });
    const already = rows.some(row => row?.[0] && namesMatch(String(row[0]), nome));
    if (!already) {
      rows.push([nome]);
    }
    wb.Sheets[excSheetName] = XLSX.utils.aoa_to_sheet(rows);
    saveWorkbook(wb);
    return parseWorkbook(wb);
  });

export const removeExcecao = (nome: string): Promise<SalariosStore> =>
  enqueue(() => {
    const wb = loadWorkbook();
    const excSheetName = findExcecoesSheetName(wb);
    if (!excSheetName) return parseWorkbook(wb);

    const ws = wb.Sheets[excSheetName];
    const rows: any[][] = XLSX.utils.sheet_to_json(ws, { header: 1 });
    const kept = rows.filter(row => !(row?.[0] && namesMatch(String(row[0]), nome)));

    wb.Sheets[excSheetName] = XLSX.utils.aoa_to_sheet(kept);
    saveWorkbook(wb);
    return parseWorkbook(wb);
  });
