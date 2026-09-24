// Parsing tolerante das duas abas do salarios.xlsx, portado da lógica
// original que já foi validada contra o arquivo real do usuário.

const HEADER_KEYWORDS = ['NOME', 'FUNCIONARIO', 'FUNCIONÁRIO', 'COLABORADOR', 'EMPREGADO', 'NOMES'];

const BLACKLIST = new Set([
  'NOME', 'NOMES', 'NOME COMPLETO', 'FUNCIONARIO', 'FUNCIONÁRIO', 'COLABORADOR',
  'COLABORADORES', 'EXCEÇÃO', 'EXCESSÃO', 'CARÊNCIA', 'CARENCIA', 'SALARIO',
  'SALÁRIO', 'VALOR', 'CARGO', 'FUNÇÃO', 'FUNCAO', 'SETOR', 'DEPARTAMENTO',
  'CIDADE', 'ESTADO', 'OBS', 'OBSERVAÇÃO', 'OBSERVACOES', 'STATUS', 'MOTIVO',
  'TOTAL', 'TOTAIS', 'DATA', 'MATRICULA', 'MATRÍCULA', 'CPF'
]);

export const parseValorHora = (raw: any): number | null => {
  let valorStr = String(raw ?? '').replace('R$', '').trim();
  if (valorStr.includes(',') && valorStr.includes('.')) {
    valorStr = valorStr.replace(/\./g, '').replace(',', '.');
  } else if (valorStr.includes(',')) {
    valorStr = valorStr.replace(',', '.');
  }
  const num = parseFloat(valorStr);
  return Number.isFinite(num) ? num : null;
};

// Aba de salários: coluna A = nome, coluna B = valor-hora, a partir da linha 2.
export const parseSalariosSheet = (rows: any[][]): Record<string, number> => {
  const map: Record<string, number> = {};
  for (let r = 1; r < rows.length; r++) {
    const row = rows[r];
    if (!row || row.length < 2) continue;
    const nome = String(row[0] || '').trim().toUpperCase();
    const num = parseValorHora(row[1]);
    if (nome && num !== null) {
      map[nome] = num;
    }
  }
  return map;
};

// Aba de exceção: pode ou não ter cabeçalho; localiza a coluna de nomes
// evitando ler colunas de cargo/cidade/etc.
export const parseExcecoesSheet = (rows: any[][]): string[] => {
  const detected: string[] = [];
  if (!rows || rows.length === 0) return detected;

  let nameColIdx = 0;
  let foundHeaderCol = -1;

  for (let r = 0; r < Math.min(rows.length, 4); r++) {
    const rData = rows[r] || [];
    for (let c = 0; c < rData.length; c++) {
      const cellText = String(rData[c] || '').trim().toUpperCase();
      if (HEADER_KEYWORDS.some(kw => cellText.includes(kw))) {
        foundHeaderCol = c;
        break;
      }
    }
    if (foundHeaderCol !== -1) break;
  }

  if (foundHeaderCol !== -1) {
    nameColIdx = foundHeaderCol;
  } else {
    const sample0 = rows.slice(0, 5).map(r => String(r[0] || '').trim());
    const sample1 = rows.slice(0, 5).map(r => String(r[1] || '').trim());
    if (sample0.some(s => /^\d+$/.test(s)) && sample1.some(s => s.length >= 3 && !/^\d+$/.test(s))) {
      nameColIdx = 1;
    }
  }

  for (let r = 0; r < rows.length; r++) {
    const val = rows[r]?.[nameColIdx];
    if (val === undefined || val === null) continue;
    const nomeExc = String(val).trim().toUpperCase();
    if (!nomeExc || BLACKLIST.has(nomeExc)) continue;
    if (/^\d+(\.\d+)?$/.test(nomeExc) || nomeExc.length < 3) continue;
    if (!detected.includes(nomeExc)) detected.push(nomeExc);
  }

  return detected;
};

const normalizeSheetName = (sName: string): string =>
  sName.trim().toLowerCase().replace(/ç/g, 'c').replace(/ã/g, 'a').replace(/õ/g, 'o').replace(/\s+/g, '');

export const isExcecoesSheetName = (sName: string): boolean => {
  const sNorm = normalizeSheetName(sName);
  return sNorm.includes('excess') || sNorm.includes('excec') || sNorm.includes('carencia');
};
