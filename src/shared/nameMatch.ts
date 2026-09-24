// Comparação estrita de NOME COMPLETO, tolerante a diferenças de formatação
// (acentos, preposições, truncamento de 31 caracteres nas abas do Excel).
// Usado tanto pelo servidor (checagem de duplicidade) quanto pela tela
// (bater nome da aba do ponto contra a lista de salários/exceções).

export const cleanCompareName = (text: string): string => {
  if (!text) return '';
  return text
    .normalize('NFD')
    .replace(/[̀-ͯ]/g, '') // remove acentos
    .toUpperCase()
    // Remove prefixos como '01 - ', '0023 ', 'MATR. 123 - '
    .replace(/^(?:MATR(?:ICULA)?\.?\s*)?[0-9\-_./#:\s]+/, '')
    // Remove sufixos como ' - 123', ' (0012)', ' [99]'
    .replace(/[-_/#:\s]*[([]?\d+[)\]]?\s*$/, '')
    .replace(/[^A-Z\s]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
};

const PREPOSITIONS = new Set(['DE', 'DA', 'DO', 'DOS', 'DAS', 'E', 'D']);

export const getCoreTokens = (cleanedName: string): string[] => {
  return cleanedName
    .split(' ')
    .filter(t => t.length > 0 && !PREPOSITIONS.has(t));
};

// Compara dois nomes com a mesma regra: igualdade exata normalizada,
// igualdade ignorando preposições, ou truncamento nativo do Excel (31 chars).
export const namesMatch = (a: string, b: string): boolean => {
  const normA = cleanCompareName(a);
  const normB = cleanCompareName(b);
  if (!normA || !normB) return false;

  if (normA === normB) return true;

  const coreA = getCoreTokens(normA).join(' ');
  const coreB = getCoreTokens(normB).join(' ');
  if (coreA && coreB && coreA === coreB) return true;

  if (
    (normA.length >= 28 && normB.startsWith(normA)) ||
    (normB.length >= 28 && normA.startsWith(normB))
  ) {
    return true;
  }

  return false;
};

export const isNameInList = (name: string, list: string[]): boolean => {
  return list.some(item => namesMatch(name, item));
};

// Retorna a taxa horária correspondente, ou undefined se o nome não bate
// com nenhuma chave da lista de salários (em vez de cair numa taxa padrão
// silenciosa — quem chama deve tratar undefined como "funcionário novo").
export const findRateForName = (
  name: string,
  salariosMap: Record<string, number>
): number | undefined => {
  for (const key of Object.keys(salariosMap)) {
    if (namesMatch(name, key)) {
      return salariosMap[key];
    }
  }
  return undefined;
};
