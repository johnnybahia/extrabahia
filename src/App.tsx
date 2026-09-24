import React, { useState, useRef, useEffect, useCallback } from 'react';
import {
  Upload,
  CheckCircle2,
  AlertTriangle,
  Download,
  Check,
  Sparkles,
  Eye,
  Plus,
  X,
  Trash2,
  AlertCircle,
  Clock,
  DollarSign,
  RefreshCw,
  UserPlus,
} from 'lucide-react';
import * as XLSX from 'xlsx';
import { isNameInList, findRateForName } from './shared/nameMatch';
import type { SalariosStore } from './shared/types';

// Regras fixas do cálculo (antes ajustáveis na aba "Simulador", removida).
// A tolerância (minutos de carência) agora é editável na tela e fica salva
// na aba "Configuração" do dados/salarios.xlsx — DEFAULT só vale antes do
// primeiro carregamento.
const TOLERANCIA_MINUTOS_DEFAULT = 19;
const START_ROW = 3;
const INCLUDE_EXCESS_ONLY = false;
const APPLY_100_PERCENT = false;

interface DetalheDia {
  linha: number;
  labelDia: string;
  exU50Str: string;
  sec50: number;
  ex100Str: string;
  status50: 'aprovado' | 'descartado' | 'zerado';
  motivo: string;
}

interface ResultadoFuncionario {
  nome: string;
  ehExcecao: boolean;
  extras50: string;
  total50: number;
  extras100: string;
  total100: number;
  totalGeral: number;
  diasAprovados: number;
  diasDescartados: number;
  detalhesDias: DetalheDia[];
}

const parseCellSeconds = (val: any): number => {
  if (val === null || val === undefined || val === '') return 0;
  if (typeof val === 'number') {
    return Math.round(val * 24 * 3600);
  }
  const str = String(val).trim();
  const match = str.match(/(\d{1,3}):(\d{2})(?::(\d{2}))?$/);
  if (match) {
    const h = parseInt(match[1], 10);
    const m = parseInt(match[2], 10);
    const s = parseInt(match[3] || '0', 10);
    return h * 3600 + m * 60 + s;
  }
  return 0;
};

const formatSecondsToHHMM = (totalSec: number) => {
  if (totalSec <= 0) return '00:00';
  const h = Math.floor(totalSec / 3600);
  const m = Math.floor((totalSec % 3600) / 60);
  return `${h.toString().padStart(2, '0')}:${m.toString().padStart(2, '0')}`;
};

const formatBRL = (n: number) =>
  n.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });

// Processa o workbook de ponto já carregado: uma aba por funcionário.
// Abas cujo nome não bate com nenhuma chave de salariosMap voltam em
// "pendentes" (nome exato da aba) em vez de usar uma taxa padrão.
const processPontoWorkbook = (
  wb: XLSX.WorkBook,
  salariosMap: Record<string, number>,
  excecoesList: string[],
  toleranciaMinutos: number
): { resultados: ResultadoFuncionario[]; pendentes: string[] } => {
  const resultados: ResultadoFuncionario[] = [];
  const pendentes: string[] = [];
  const limitSec = toleranciaMinutos * 60;

  wb.SheetNames.forEach(sheetName => {
    const ws = wb.Sheets[sheetName];
    const rows: any[][] = XLSX.utils.sheet_to_json(ws, { header: 1 });
    if (!rows || rows.length < 2) return;

    const header = rows[0] || [];
    let colU = 20;
    let colV = 21;
    let colW = 22;
    let colX = 23;

    header.forEach((val, idx) => {
      const h = String(val || '').trim().toUpperCase();
      if (h === 'EXU50') colU = idx;
      if (h === 'EXS100' || h === 'EXS5100') colV = idx;
      if (h === 'EXD100') colW = idx;
      if (h === 'EXF100') colX = idx;
    });

    let employeeFullName = sheetName;
    for (let r = 0; r < Math.min(rows.length, 5); r++) {
      const rData = rows[r] || [];
      for (let c = 0; c < rData.length; c++) {
        const val = String(rData[c] || '');
        const match = val.match(/funcion[aá]rio\s*:\s*(.+)/i);
        if (match && match[1]?.trim()) {
          employeeFullName = match[1].trim();
          break;
        }
      }
    }

    const isExcecao =
      isNameInList(sheetName, excecoesList) ||
      (employeeFullName !== sheetName && isNameInList(employeeFullName, excecoesList));

    let totalSec50 = 0;
    let totalSec100 = 0;
    let diasAprovados = 0;
    let diasDescartados = 0;
    const detalhesDias: DetalheDia[] = [];

    const startIdx = Math.max(1, START_ROW - 1);
    for (let r = startIdx; r < rows.length; r++) {
      if (r === 1) continue; // linha 2 = Totais do sistema

      const row = rows[r];
      if (!row) continue;

      const firstCell = String(row[0] || '').toLowerCase();
      const uCellStr = String(row[colU] || '').toLowerCase();
      if (firstCell.includes('total') || uCellStr.includes('total')) continue;

      const sec50 = parseCellSeconds(row[colU]);
      const exU50Str = formatSecondsToHHMM(sec50);

      const sec100 = parseCellSeconds(row[colV]) + parseCellSeconds(row[colW]) + parseCellSeconds(row[colX]);
      const ex100Str = formatSecondsToHHMM(sec100);

      const linhaExcel = r + 1;
      const labelDia = String(row[0] || `Linha ${linhaExcel}`);

      if (isExcecao) {
        if (sec50 > 0) {
          diasAprovados++;
          totalSec50 += sec50;
          detalhesDias.push({
            linha: linhaExcel,
            labelDia,
            exU50Str,
            sec50,
            ex100Str,
            status50: 'aprovado',
            motivo: 'Exceção: cada minuto é calculado integralmente',
          });
        } else {
          detalhesDias.push({
            linha: linhaExcel,
            labelDia,
            exU50Str,
            sec50: 0,
            ex100Str,
            status50: 'zerado',
            motivo: 'Sem extras 50% neste dia',
          });
        }
      } else if (sec50 > limitSec) {
        diasAprovados++;
        totalSec50 += INCLUDE_EXCESS_ONLY ? Math.max(0, sec50 - limitSec) : sec50;
        detalhesDias.push({
          linha: linhaExcel,
          labelDia,
          exU50Str,
          sec50,
          ex100Str,
          status50: 'aprovado',
          motivo: `Ultrapassou ${toleranciaMinutos} min (soma integral: ${exU50Str})`,
        });
      } else if (sec50 > 0) {
        diasDescartados++;
        detalhesDias.push({
          linha: linhaExcel,
          labelDia,
          exU50Str,
          sec50,
          ex100Str,
          status50: 'descartado',
          motivo: `Feito ${exU50Str} (<= ${toleranciaMinutos} min carência descartado)`,
        });
      } else {
        detalhesDias.push({
          linha: linhaExcel,
          labelDia,
          exU50Str,
          sec50: 0,
          ex100Str,
          status50: 'zerado',
          motivo: 'Sem extras 50% neste dia',
        });
      }

      if (APPLY_100_PERCENT && !isExcecao) {
        if (sec100 > limitSec) totalSec100 += sec100;
      } else {
        totalSec100 += sec100;
      }
    }

    const nomeFinal = employeeFullName || sheetName;
    const rate = findRateForName(sheetName, salariosMap) ?? findRateForName(employeeFullName, salariosMap);

    if (rate === undefined) {
      pendentes.push(sheetName);
      return;
    }

    const h50Dec = totalSec50 / 3600;
    const h100Dec = totalSec100 / 3600;
    const total50 = h50Dec * rate * 1.5;
    const total100 = h100Dec * rate * 2.0;

    resultados.push({
      nome: nomeFinal,
      ehExcecao: isExcecao,
      extras50: formatSecondsToHHMM(totalSec50),
      total50,
      extras100: formatSecondsToHHMM(totalSec100),
      total100,
      totalGeral: total50 + total100,
      diasAprovados,
      diasDescartados,
      detalhesDias,
    });
  });

  return { resultados, pendentes };
};

export default function App() {
  const [salariosMap, setSalariosMap] = useState<Record<string, number>>({});
  const [excecoesList, setExcecoesList] = useState<string[]>([]);
  const [salariosLoading, setSalariosLoading] = useState<boolean>(true);
  const [salariosError, setSalariosError] = useState<string>('');

  const [pontoFileName, setPontoFileName] = useState<string>('');
  const [uploadError, setUploadError] = useState<string>('');
  const [processedResults, setProcessedResults] = useState<ResultadoFuncionario[]>([]);
  const [selectedEmployeeDetail, setSelectedEmployeeDetail] = useState<ResultadoFuncionario | null>(null);
  const pontoWbRef = useRef<XLSX.WorkBook | null>(null);
  const pendentesAnterioresRef = useRef<string[]>([]);

  const [pendentesNovos, setPendentesNovos] = useState<string[]>([]);
  const [showNovosModal, setShowNovosModal] = useState<boolean>(false);
  const [pendingRates, setPendingRates] = useState<Record<string, string>>({});
  const [savingNome, setSavingNome] = useState<string | null>(null);
  const [registroErro, setRegistroErro] = useState<string>('');

  const [salarioParaExcluir, setSalarioParaExcluir] = useState<string | null>(null);
  const [excluindo, setExcluindo] = useState<boolean>(false);
  const [excluirErro, setExcluirErro] = useState<string>('');

  const [newExcecaoInput, setNewExcecaoInput] = useState<string>('');
  const [excecaoSalvando, setExcecaoSalvando] = useState<boolean>(false);
  const [excecaoErro, setExcecaoErro] = useState<string>('');

  const [toleranciaMinutos, setToleranciaMinutos] = useState<number>(TOLERANCIA_MINUTOS_DEFAULT);
  const [toleranciaInput, setToleranciaInput] = useState<string>(String(TOLERANCIA_MINUTOS_DEFAULT));
  const [toleranciaSalvando, setToleranciaSalvando] = useState<boolean>(false);
  const [toleranciaErro, setToleranciaErro] = useState<string>('');

  const reprocessPonto = useCallback((map: Record<string, number>, list: string[], tolerancia: number) => {
    const wb = pontoWbRef.current;
    if (!wb) return;
    const { resultados, pendentes } = processPontoWorkbook(wb, map, list, tolerancia);
    // Só reabre o modal se surgiu um nome pendente que não estava lá antes —
    // evita reinterromper o usuário toda vez que ele mexe em outra coisa
    // (ex: adicionar uma exceção) enquanto um pendente já visto continua em aberto.
    const anteriores = pendentesAnterioresRef.current;
    const surgiuPendenteNovo = pendentes.some(nome => !anteriores.includes(nome));
    pendentesAnterioresRef.current = pendentes;

    setProcessedResults(resultados);
    setPendentesNovos(pendentes);
    setSelectedEmployeeDetail(null);
    if (surgiuPendenteNovo) setShowNovosModal(true);
  }, []);

  const loadSalarios = useCallback(async () => {
    setSalariosLoading(true);
    setSalariosError('');
    try {
      const resp = await fetch('/api/salarios');
      if (!resp.ok) throw new Error(`Servidor respondeu ${resp.status}`);
      const store: SalariosStore = await resp.json();
      setSalariosMap(store.salarios);
      setExcecoesList(store.excecoes);
      setToleranciaMinutos(store.toleranciaMinutos);
      setToleranciaInput(String(store.toleranciaMinutos));
      reprocessPonto(store.salarios, store.excecoes, store.toleranciaMinutos);
    } catch (err: any) {
      setSalariosError(
        `Não consegui ler a lista de salários (${err?.message || err}). Verifique se o programa "iniciar.bat" está rodando e se o arquivo dados/salarios.xlsx não está aberto no Excel.`
      );
    } finally {
      setSalariosLoading(false);
    }
  }, [reprocessPonto]);

  useEffect(() => {
    loadSalarios();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const handlePontoFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    setUploadError('');
    const file = e.target.files?.[0];
    if (!file) return;
    setPontoFileName(file.name);

    const reader = new FileReader();
    reader.onload = evt => {
      try {
        const data = new Uint8Array(evt.target?.result as ArrayBuffer);
        const wb = XLSX.read(data, { type: 'array' });
        if (!wb.SheetNames || wb.SheetNames.length === 0) {
          setUploadError('Nenhuma aba válida encontrada no arquivo selecionado.');
          return;
        }
        pontoWbRef.current = wb;
        reprocessPonto(salariosMap, excecoesList, toleranciaMinutos);
      } catch (err: any) {
        setUploadError(`Erro ao processar arquivo de ponto: ${err?.message || err}`);
      }
    };
    reader.onerror = () => setUploadError('Erro ao ler o arquivo selecionado.');
    reader.readAsArrayBuffer(file);
  };

  const handleSalvarNovoFuncionario = async (nome: string) => {
    const raw = (pendingRates[nome] || '').replace(',', '.');
    const valorHora = parseFloat(raw);
    if (!Number.isFinite(valorHora) || valorHora <= 0) {
      setRegistroErro(`Informe um valor de hora válido (maior que zero) para ${nome}.`);
      return;
    }
    setRegistroErro('');
    setSavingNome(nome);
    try {
      const resp = await fetch('/api/salarios', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ nome, valorHora }),
      });
      const body = await resp.json();
      if (!resp.ok) throw new Error(body?.error || `Servidor respondeu ${resp.status}`);
      const store: SalariosStore = body;
      setSalariosMap(store.salarios);
      setExcecoesList(store.excecoes);
      setPendingRates(prev => {
        const next = { ...prev };
        delete next[nome];
        return next;
      });
      reprocessPonto(store.salarios, store.excecoes, toleranciaMinutos);
    } catch (err: any) {
      setRegistroErro(`Erro ao cadastrar ${nome}: ${err?.message || err}`);
    } finally {
      setSavingNome(null);
    }
  };

  const handleConfirmarExclusao = async () => {
    if (!salarioParaExcluir) return;
    setExcluindo(true);
    setExcluirErro('');
    try {
      const resp = await fetch(`/api/salarios/${encodeURIComponent(salarioParaExcluir)}`, { method: 'DELETE' });
      const body = await resp.json();
      if (!resp.ok) throw new Error(body?.error || `Servidor respondeu ${resp.status}`);
      const store: SalariosStore = body;
      setSalariosMap(store.salarios);
      setExcecoesList(store.excecoes);
      reprocessPonto(store.salarios, store.excecoes, toleranciaMinutos);
      setSalarioParaExcluir(null);
    } catch (err: any) {
      setExcluirErro(`Erro ao excluir: ${err?.message || err}`);
    } finally {
      setExcluindo(false);
    }
  };

  const handleAdicionarExcecao = async () => {
    const nome = newExcecaoInput.trim().toUpperCase();
    if (!nome) return;
    setExcecaoSalvando(true);
    setExcecaoErro('');
    try {
      const resp = await fetch('/api/excecoes', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ nome }),
      });
      const body = await resp.json();
      if (!resp.ok) throw new Error(body?.error || `Servidor respondeu ${resp.status}`);
      const store: SalariosStore = body;
      setSalariosMap(store.salarios);
      setExcecoesList(store.excecoes);
      setNewExcecaoInput('');
      reprocessPonto(store.salarios, store.excecoes, toleranciaMinutos);
    } catch (err: any) {
      setExcecaoErro(`Erro ao adicionar exceção: ${err?.message || err}`);
    } finally {
      setExcecaoSalvando(false);
    }
  };

  const handleRemoverExcecao = async (nome: string) => {
    setExcecaoErro('');
    try {
      const resp = await fetch(`/api/excecoes/${encodeURIComponent(nome)}`, { method: 'DELETE' });
      const body = await resp.json();
      if (!resp.ok) throw new Error(body?.error || `Servidor respondeu ${resp.status}`);
      const store: SalariosStore = body;
      setSalariosMap(store.salarios);
      setExcecoesList(store.excecoes);
      reprocessPonto(store.salarios, store.excecoes, toleranciaMinutos);
    } catch (err: any) {
      setExcecaoErro(`Erro ao remover exceção: ${err?.message || err}`);
    }
  };

  const handleSalvarTolerancia = async () => {
    const minutos = parseInt(toleranciaInput, 10);
    if (!Number.isInteger(minutos) || minutos < 0) {
      setToleranciaErro('Informe um número inteiro de minutos (0 ou mais).');
      return;
    }
    setToleranciaErro('');
    setToleranciaSalvando(true);
    try {
      const resp = await fetch('/api/configuracao', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ toleranciaMinutos: minutos }),
      });
      const body = await resp.json();
      if (!resp.ok) throw new Error(body?.error || `Servidor respondeu ${resp.status}`);
      const store: SalariosStore = body;
      setSalariosMap(store.salarios);
      setExcecoesList(store.excecoes);
      setToleranciaMinutos(store.toleranciaMinutos);
      setToleranciaInput(String(store.toleranciaMinutos));
      reprocessPonto(store.salarios, store.excecoes, store.toleranciaMinutos);
    } catch (err: any) {
      setToleranciaErro(`Erro ao salvar tolerância: ${err?.message || err}`);
    } finally {
      setToleranciaSalvando(false);
    }
  };

  const handleExportCalculatedExcel = () => {
    if (processedResults.length === 0) return;
    const wb = XLSX.utils.book_new();
    const dataRows = processedResults.map(r => ({
      'NOME DO FUNCIONÁRIO': r.nome,
      'REGRA APLICADA': r.ehExcecao ? 'EXCEÇÃO (CADA MINUTO)' : `GERAL (> ${toleranciaMinutos}m)`,
      'EXTRAS 50%': r.extras50,
      'TOTAL 50%': Number(r.total50.toFixed(2)),
      'EXTRAS 100%': r.extras100,
      'TOTAL 100%': Number(r.total100.toFixed(2)),
      TOTAL: Number(r.totalGeral.toFixed(2)),
      [`DIAS 50% APROV. (> ${toleranciaMinutos}m)`]: r.diasAprovados,
      [`DIAS 50% DESC. (<= ${toleranciaMinutos}m)`]: r.diasDescartados,
    }));
    const ws = XLSX.utils.json_to_sheet(dataRows);
    XLSX.utils.book_append_sheet(wb, ws, 'Relatorio');
    XLSX.writeFile(wb, 'EXTRAS CALCULADAS.xlsx');
  };

  const salariosOrdenados = Object.entries(salariosMap).sort((a, b) => a[0].localeCompare(b[0], 'pt-BR'));
  const valorParaExcluir = salarioParaExcluir ? salariosMap[salarioParaExcluir] : undefined;

  return (
    <div className="min-h-screen bg-slate-950 text-slate-100 flex flex-col font-sans">
      <header className="border-b border-slate-800 bg-slate-900/80 backdrop-blur sticky top-0 z-20 px-6 py-4">
        <div className="max-w-5xl mx-auto flex flex-wrap items-center justify-between gap-4">
          <div className="flex items-center gap-3">
            <div className="p-2.5 rounded-xl bg-gradient-to-tr from-amber-500 to-orange-600 text-white shadow-lg shadow-amber-500/20">
              <Clock className="w-6 h-6" />
            </div>
            <div>
              <h1 className="text-lg font-bold text-white leading-tight">Calculadora de Horas Extras</h1>
              <p className="text-xs text-slate-400">Regra de carência de {toleranciaMinutos} minutos</p>
            </div>
          </div>
        </div>
      </header>

      <main className="flex-1 max-w-5xl w-full mx-auto p-6 space-y-6">
        {/* Lista de Salários (pasta fixa) */}
        <div className="bg-slate-900 border border-slate-800 rounded-2xl p-6 shadow-sm">
          <div className="flex items-center justify-between gap-4 pb-4 border-b border-slate-800">
            <h3 className="text-base font-bold text-white flex items-center gap-2">
              <DollarSign className="w-5 h-5 text-emerald-400" />
              Lista de Salários (dados/salarios.xlsx)
            </h3>
            <button
              onClick={loadSalarios}
              disabled={salariosLoading}
              className="px-3 py-2 bg-slate-800 hover:bg-slate-700 disabled:opacity-50 text-slate-200 text-xs font-semibold rounded-xl border border-slate-700 flex items-center gap-1.5 transition-colors"
            >
              <RefreshCw className={`w-3.5 h-3.5 ${salariosLoading ? 'animate-spin' : ''}`} />
              Recarregar
            </button>
          </div>

          {salariosError && (
            <div className="mt-4 p-3.5 rounded-xl bg-rose-950/40 border border-rose-800/50 text-rose-300 text-xs flex items-center gap-2">
              <AlertTriangle className="w-4 h-4 text-rose-400 shrink-0" />
              <span>{salariosError}</span>
            </div>
          )}

          {!salariosError && salariosLoading && (
            <p className="text-xs text-slate-400 mt-4">Carregando lista de salários...</p>
          )}

          {!salariosError && !salariosLoading && (
            <div className="mt-4 max-h-72 overflow-y-auto">
              {salariosOrdenados.length === 0 ? (
                <p className="text-xs text-slate-500 italic">
                  Nenhum funcionário cadastrado ainda. Cadastros aparecem aqui automaticamente ao enviar um arquivo de ponto.
                </p>
              ) : (
                <table className="w-full text-xs text-left">
                  <thead>
                    <tr className="border-b border-slate-800 text-slate-400 font-semibold uppercase tracking-wider text-[11px]">
                      <th className="py-2 px-3">Funcionário</th>
                      <th className="py-2 px-3">Valor Hora</th>
                      <th className="py-2 px-3 w-10"></th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-800/60">
                    {salariosOrdenados.map(([nome, valor]) => (
                      <tr key={nome} className="hover:bg-slate-800/40">
                        <td className="py-2 px-3 font-sans text-slate-200">{nome}</td>
                        <td className="py-2 px-3 font-mono text-emerald-400">R$ {formatBRL(valor)}</td>
                        <td className="py-2 px-3 text-right">
                          <button
                            onClick={() => {
                              setExcluirErro('');
                              setSalarioParaExcluir(nome);
                            }}
                            title="Excluir da lista de salários"
                            className="p-1.5 rounded-lg text-slate-500 hover:text-rose-400 hover:bg-rose-500/10 transition-colors"
                          >
                            <Trash2 className="w-3.5 h-3.5" />
                          </button>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              )}
            </div>
          )}

          {/* Exceções extra */}
          <div className="mt-5 p-4 rounded-xl bg-purple-950/40 border border-purple-800/50 text-purple-200 text-xs space-y-3">
            <div className="flex items-start justify-between gap-2">
              <div className="flex items-center gap-2">
                <Sparkles className="w-4 h-4 text-purple-400 shrink-0" />
                <span className="font-bold text-purple-300">Funcionários em Exceção (sem carência de {toleranciaMinutos} min):</span>
              </div>
              <span className="text-[11px] text-purple-400/80 font-mono">{excecoesList.length} cadastrado(s)</span>
            </div>

            <div className="flex flex-wrap gap-1.5 items-center">
              {excecoesList.length === 0 ? (
                <span className="text-slate-500 italic text-xs">Nenhum funcionário na lista de exceção.</span>
              ) : (
                excecoesList.map((exc, idx) => (
                  <span
                    key={idx}
                    className="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-lg bg-purple-500/20 text-purple-200 font-mono text-xs border border-purple-500/30 shadow-sm"
                  >
                    <Check className="w-3 h-3 text-purple-400" />
                    {exc}
                    <button
                      type="button"
                      onClick={() => handleRemoverExcecao(exc)}
                      className="hover:text-rose-400 transition-colors p-0.5"
                      title="Remover exceção"
                    >
                      <X className="w-3 h-3" />
                    </button>
                  </span>
                ))
              )}
            </div>

            {excecaoErro && <p className="text-rose-300 text-[11px]">{excecaoErro}</p>}

            <div className="flex items-center gap-2 pt-2 border-t border-purple-900/40">
              <input
                type="text"
                value={newExcecaoInput}
                onChange={e => setNewExcecaoInput(e.target.value)}
                onKeyDown={e => {
                  if (e.key === 'Enter') {
                    e.preventDefault();
                    handleAdicionarExcecao();
                  }
                }}
                placeholder="Nome completo exatamente como na aba do ponto..."
                className="flex-1 bg-slate-950 border border-purple-800/60 rounded-lg px-3 py-1.5 text-xs text-white placeholder:text-slate-500 focus:outline-none focus:border-purple-400"
              />
              <button
                type="button"
                onClick={handleAdicionarExcecao}
                disabled={excecaoSalvando || !newExcecaoInput.trim()}
                className="px-3 py-1.5 bg-purple-600 hover:bg-purple-500 disabled:opacity-50 text-white text-xs font-semibold rounded-lg flex items-center gap-1 transition-colors"
              >
                <Plus className="w-3.5 h-3.5" />
                Adicionar Exceção
              </button>
            </div>
          </div>
        </div>

        {/* Upload do ponto */}
        <div className="bg-slate-900 border border-slate-800 rounded-2xl p-6 shadow-sm">
          <h3 className="text-base font-bold text-white flex items-center gap-2">
            <Upload className="w-5 h-5 text-amber-400" />
            Enviar Arquivo de Ponto
          </h3>
          <p className="text-xs text-slate-400 mt-1">
            Envie o arquivo de ponto (.xlsx) para calcular as horas extras automaticamente.
          </p>

          <div className="mt-4 p-3 rounded-xl bg-slate-950 border border-slate-800 flex flex-wrap items-center gap-2">
            <label htmlFor="tolerancia-minutos" className="text-xs text-slate-300 font-semibold">
              Tolerância (minutos de carência):
            </label>
            <input
              id="tolerancia-minutos"
              type="text"
              inputMode="numeric"
              value={toleranciaInput}
              onChange={e => setToleranciaInput(e.target.value)}
              onKeyDown={e => {
                if (e.key === 'Enter') handleSalvarTolerancia();
              }}
              className="w-16 bg-slate-900 border border-slate-700 rounded-lg px-2 py-1.5 text-xs text-white text-center focus:outline-none focus:border-amber-400"
            />
            <button
              onClick={handleSalvarTolerancia}
              disabled={toleranciaSalvando || toleranciaInput === String(toleranciaMinutos)}
              className="px-3 py-1.5 bg-amber-500 hover:bg-amber-400 disabled:opacity-50 disabled:cursor-not-allowed text-slate-950 text-xs font-bold rounded-lg transition-colors"
            >
              {toleranciaSalvando ? 'Salvando...' : 'Salvar'}
            </button>
            <span className="text-[11px] text-slate-500">
              Dias com extra 50% ≤ {toleranciaMinutos} min são descartados na regra geral. Salvo em dados/salarios.xlsx.
            </span>
            {toleranciaErro && <span className="w-full text-[11px] text-rose-300">{toleranciaErro}</span>}
          </div>

          <div className="mt-4">
            <div
              className={`p-4 rounded-xl border-2 border-dashed bg-slate-950 flex flex-col items-center justify-center text-center transition-colors ${
                salariosLoading ? 'border-slate-800 opacity-50' : 'border-slate-700 hover:border-amber-500/60'
              }`}
            >
              <input
                type="file"
                accept=".xlsx, .xls"
                id="upload-ponto"
                className="hidden"
                disabled={salariosLoading}
                onChange={handlePontoFileUpload}
              />
              <label
                htmlFor="upload-ponto"
                className={`flex flex-col items-center p-3 w-full ${salariosLoading ? 'cursor-not-allowed' : 'cursor-pointer'}`}
              >
                <div className="p-3 rounded-full bg-amber-500/10 text-amber-400 mb-2">
                  <Upload className="w-6 h-6" />
                </div>
                <span className="text-xs font-bold text-white">
                  {pontoFileName || 'Selecionar Arquivo de PONTO (.xlsx)'}
                </span>
                <span className="text-[11px] text-slate-400 mt-1">
                  {salariosLoading
                    ? 'Aguardando carregar a lista de salários...'
                    : pontoFileName
                    ? 'Clique para trocar'
                    : 'Deve conter as abas dos funcionários com colunas U, V, W, X'}
                </span>
              </label>
            </div>
          </div>

          {!showNovosModal && pendentesNovos.length > 0 && (
            <button
              onClick={() => setShowNovosModal(true)}
              className="mt-4 w-full p-3 rounded-xl bg-amber-950/40 border border-amber-700/50 text-amber-200 text-xs flex items-center gap-2 hover:bg-amber-950/60 transition-colors"
            >
              <UserPlus className="w-4 h-4 text-amber-400 shrink-0" />
              {pendentesNovos.length} funcionário(s) sem valor-hora cadastrado — clique para cadastrar
            </button>
          )}

          {uploadError && (
            <div className="mt-4 p-3.5 rounded-xl bg-rose-950/40 border border-rose-800/50 text-rose-300 text-xs flex items-center gap-2">
              <AlertTriangle className="w-4 h-4 text-rose-400 shrink-0" />
              <span>{uploadError}</span>
            </div>
          )}
        </div>

        {/* Resultado */}
        {processedResults.length > 0 && (
          <div className="bg-slate-900 border border-slate-800 rounded-2xl p-6 shadow-sm space-y-4">
            <div className="flex flex-col sm:flex-row items-start sm:items-center justify-between gap-4 pb-3 border-b border-slate-800">
              <div>
                <h4 className="text-sm font-bold text-white flex items-center gap-2">
                  <CheckCircle2 className="w-4 h-4 text-emerald-400" />
                  Resultado do Processamento ({processedResults.length} funcionários)
                </h4>
                <p className="text-xs text-slate-400 mt-0.5">
                  Regra Geral: dias com extras 50% ≤ {toleranciaMinutos} min descartados. Exceções: cada minuto calculado.
                </p>
              </div>
              <button
                onClick={handleExportCalculatedExcel}
                className="px-4 py-2 bg-emerald-600 hover:bg-emerald-500 text-white text-xs font-bold rounded-xl flex items-center gap-2 transition-colors shadow-lg shadow-emerald-600/20"
              >
                <Download className="w-3.5 h-3.5" />
                Exportar EXTRAS CALCULADAS.xlsx
              </button>
            </div>

            <div className="overflow-x-auto">
              <table className="w-full text-xs text-left">
                <thead>
                  <tr className="border-b border-slate-800 text-slate-400 font-semibold uppercase tracking-wider text-[11px]">
                    <th className="py-2.5 px-3">Funcionário (Auditoria)</th>
                    <th className="py-2.5 px-3">Regra</th>
                    <th className="py-2.5 px-3">Extras 50%</th>
                    <th className="py-2.5 px-3">Total 50% (R$)</th>
                    <th className="py-2.5 px-3">Extras 100%</th>
                    <th className="py-2.5 px-3">Total 100% (R$)</th>
                    <th className="py-2.5 px-3">Total Geral</th>
                    <th className="py-2.5 px-3 text-center">Dias Aprovados</th>
                    <th className="py-2.5 px-3 text-center">Dias Descartados</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-800/60 font-mono">
                  {processedResults.map((r, i) => (
                    <tr
                      key={i}
                      onClick={() => setSelectedEmployeeDetail(r)}
                      className="hover:bg-slate-800/50 cursor-pointer transition-colors group"
                      title="Clique para abrir auditoria dia a dia deste funcionário"
                    >
                      <td className="py-2.5 px-3 font-sans font-semibold text-slate-200">
                        <div className="flex items-center gap-2">
                          <span className="group-hover:text-amber-400 transition-colors">{r.nome}</span>
                          <span className="inline-flex items-center gap-1 text-[10px] px-2 py-0.5 rounded bg-indigo-500/20 text-indigo-300 border border-indigo-500/30 group-hover:bg-indigo-500/40">
                            <Eye className="w-3 h-3" /> Ver Dias
                          </span>
                        </div>
                      </td>
                      <td className="py-2.5 px-3 font-sans">
                        {r.ehExcecao ? (
                          <span className="inline-flex items-center px-2 py-0.5 rounded-full text-[10px] font-bold bg-purple-500/20 text-purple-300 border border-purple-500/30">
                            Exceção (Cada Minuto)
                          </span>
                        ) : (
                          <span className="inline-flex items-center px-2 py-0.5 rounded-full text-[10px] font-bold bg-slate-800 text-slate-400 border border-slate-700">
                            Geral (&gt; {toleranciaMinutos}m)
                          </span>
                        )}
                      </td>
                      <td className="py-2.5 px-3 text-amber-400 font-bold">{r.extras50}</td>
                      <td className="py-2.5 px-3 text-slate-200">R$ {formatBRL(r.total50)}</td>
                      <td className="py-2.5 px-3 text-blue-400 font-bold">{r.extras100}</td>
                      <td className="py-2.5 px-3 text-slate-200">R$ {formatBRL(r.total100)}</td>
                      <td className="py-2.5 px-3 text-emerald-400 font-bold">R$ {formatBRL(r.totalGeral)}</td>
                      <td className="py-2.5 px-3 text-center text-emerald-400 font-sans">{r.diasAprovados}</td>
                      <td className="py-2.5 px-3 text-center font-sans">
                        {r.diasDescartados > 0 ? (
                          <button
                            onClick={e => {
                              e.stopPropagation();
                              setSelectedEmployeeDetail(r);
                            }}
                            className="inline-flex items-center gap-1 font-bold text-rose-300 bg-rose-500/20 hover:bg-rose-500/30 px-2.5 py-1 rounded-md border border-rose-500/40 text-xs shadow-sm transition-all"
                            title="Clique para ver o dia descartado"
                          >
                            <AlertCircle className="w-3.5 h-3.5 text-rose-400" />
                            {r.diasDescartados} descartado{r.diasDescartados > 1 ? 's' : ''} (Ver)
                          </button>
                        ) : (
                          <span className="text-slate-500">0</span>
                        )}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {/* Modal: novos funcionários detectados */}
        {showNovosModal && pendentesNovos.length > 0 && (
          <div className="fixed inset-0 z-50 bg-black/75 backdrop-blur-sm flex items-center justify-center p-4">
            <div className="bg-slate-900 border border-slate-700 rounded-2xl w-full max-w-lg max-h-[85vh] flex flex-col shadow-2xl overflow-hidden">
              <div className="px-6 py-4 border-b border-slate-800 bg-slate-950 flex items-center justify-between">
                <div className="flex items-center gap-2">
                  <UserPlus className="w-5 h-5 text-amber-400" />
                  <h3 className="text-base font-bold text-white">Funcionário(s) Novo(s) Detectado(s)</h3>
                </div>
                <button
                  onClick={() => setShowNovosModal(false)}
                  className="p-1.5 rounded-lg bg-slate-800 hover:bg-slate-700 text-slate-400 hover:text-white transition-colors"
                >
                  <X className="w-4 h-4" />
                </button>
              </div>

              <div className="p-5 overflow-y-auto flex-1 space-y-3">
                <p className="text-xs text-slate-400">
                  Estas abas do arquivo de ponto não têm valor-hora cadastrado. O nome abaixo é exatamente o nome da aba —
                  informe o valor da hora (R$) de cada um para incluí-los na lista de salários.
                </p>

                {registroErro && (
                  <div className="p-3 rounded-xl bg-rose-950/40 border border-rose-800/50 text-rose-300 text-xs flex items-center gap-2">
                    <AlertTriangle className="w-4 h-4 text-rose-400 shrink-0" />
                    <span>{registroErro}</span>
                  </div>
                )}

                {pendentesNovos.map(nome => (
                  <div key={nome} className="p-3 rounded-xl bg-slate-950 border border-slate-800 flex items-center gap-2">
                    <span className="flex-1 text-xs font-semibold text-slate-200 font-mono">{nome}</span>
                    <span className="text-xs text-slate-500">R$</span>
                    <input
                      type="text"
                      inputMode="decimal"
                      value={pendingRates[nome] || ''}
                      onChange={e => setPendingRates(prev => ({ ...prev, [nome]: e.target.value }))}
                      onKeyDown={e => {
                        if (e.key === 'Enter') handleSalvarNovoFuncionario(nome);
                      }}
                      placeholder="0,00"
                      className="w-20 bg-slate-900 border border-slate-700 rounded-lg px-2 py-1.5 text-xs text-white text-right focus:outline-none focus:border-amber-400"
                    />
                    <button
                      onClick={() => handleSalvarNovoFuncionario(nome)}
                      disabled={savingNome === nome}
                      className="px-3 py-1.5 bg-amber-500 hover:bg-amber-400 disabled:opacity-50 text-slate-950 text-xs font-bold rounded-lg transition-colors"
                    >
                      {savingNome === nome ? '...' : 'Salvar'}
                    </button>
                  </div>
                ))}
              </div>

              <div className="px-6 py-3 border-t border-slate-800 bg-slate-950 flex items-center justify-end">
                <button
                  onClick={() => setShowNovosModal(false)}
                  className="px-4 py-1.5 rounded-lg bg-slate-800 hover:bg-slate-700 text-white text-xs font-semibold transition-colors"
                >
                  Cadastrar Depois
                </button>
              </div>
            </div>
          </div>
        )}

        {/* Modal: confirmar exclusão de salário */}
        {salarioParaExcluir && (
          <div className="fixed inset-0 z-50 bg-black/75 backdrop-blur-sm flex items-center justify-center p-4">
            <div className="bg-slate-900 border border-slate-700 rounded-2xl w-full max-w-sm shadow-2xl overflow-hidden">
              <div className="px-6 py-4 border-b border-slate-800 bg-slate-950 flex items-center gap-2">
                <AlertTriangle className="w-5 h-5 text-rose-400" />
                <h3 className="text-base font-bold text-white">Confirmar Exclusão</h3>
              </div>
              <div className="p-5 space-y-3">
                <p className="text-sm text-slate-300">
                  Excluir <strong className="text-white">{salarioParaExcluir}</strong> da lista de salários?
                </p>
                <p className="text-xs text-slate-500">
                  Isso remove o nome e o valor-hora
                  {valorParaExcluir !== undefined && <> (R$ {formatBRL(valorParaExcluir)})</>} da lista. Se ele aparecer
                  novamente em um arquivo de ponto, será tratado como funcionário novo.
                </p>
                {excluirErro && <p className="text-rose-300 text-xs">{excluirErro}</p>}
              </div>
              <div className="px-6 py-3 border-t border-slate-800 bg-slate-950 flex items-center justify-end gap-2">
                <button
                  onClick={() => setSalarioParaExcluir(null)}
                  disabled={excluindo}
                  className="px-4 py-1.5 rounded-lg bg-slate-800 hover:bg-slate-700 text-white text-xs font-semibold transition-colors disabled:opacity-50"
                >
                  Cancelar
                </button>
                <button
                  onClick={handleConfirmarExclusao}
                  disabled={excluindo}
                  className="px-4 py-1.5 rounded-lg bg-rose-600 hover:bg-rose-500 text-white text-xs font-semibold transition-colors disabled:opacity-50"
                >
                  {excluindo ? 'Excluindo...' : 'Excluir'}
                </button>
              </div>
            </div>
          </div>
        )}

        {/* Modal de Detalhamento Diário do Funcionário */}
        {selectedEmployeeDetail && (
          <div className="fixed inset-0 z-50 bg-black/75 backdrop-blur-sm flex items-center justify-center p-4">
            <div className="bg-slate-900 border border-slate-700 rounded-2xl w-full max-w-3xl max-h-[85vh] flex flex-col shadow-2xl overflow-hidden">
              <div className="px-6 py-4 border-b border-slate-800 bg-slate-950 flex items-center justify-between">
                <div>
                  <div className="flex items-center gap-2">
                    <h3 className="text-base font-bold text-white">{selectedEmployeeDetail.nome}</h3>
                    {selectedEmployeeDetail.ehExcecao ? (
                      <span className="px-2 py-0.5 rounded-full text-[10px] font-bold bg-purple-500/20 text-purple-300 border border-purple-500/30">
                        Exceção (Cada Minuto)
                      </span>
                    ) : (
                      <span className="px-2 py-0.5 rounded-full text-[10px] font-bold bg-slate-800 text-slate-400 border border-slate-700">
                        Regra Geral (&gt; {toleranciaMinutos} min)
                      </span>
                    )}
                  </div>
                  <p className="text-xs text-slate-400 mt-1">
                    Auditoria de cada dia trabalhado • Extras 50% (ExU50) e 100% (ExS100, ExD100, ExF100)
                  </p>
                </div>
                <button
                  onClick={() => setSelectedEmployeeDetail(null)}
                  className="p-1.5 rounded-lg bg-slate-800 hover:bg-slate-700 text-slate-400 hover:text-white transition-colors"
                >
                  <X className="w-4 h-4" />
                </button>
              </div>

              <div className="grid grid-cols-4 gap-2 p-4 bg-slate-950/60 border-b border-slate-800/80 text-center text-xs">
                <div className="p-2 rounded-lg bg-slate-900 border border-slate-800">
                  <span className="text-slate-400 block text-[10px]">Total 50%</span>
                  <span className="font-bold text-amber-400 font-mono text-sm">{selectedEmployeeDetail.extras50}</span>
                </div>
                <div className="p-2 rounded-lg bg-slate-900 border border-slate-800">
                  <span className="text-slate-400 block text-[10px]">Total 100%</span>
                  <span className="font-bold text-blue-400 font-mono text-sm">{selectedEmployeeDetail.extras100}</span>
                </div>
                <div className="p-2 rounded-lg bg-slate-900 border border-slate-800">
                  <span className="text-slate-400 block text-[10px]">Dias Aprovados 50%</span>
                  <span className="font-bold text-emerald-400 font-mono text-sm">{selectedEmployeeDetail.diasAprovados}</span>
                </div>
                <div className="p-2 rounded-lg bg-slate-900 border border-slate-800">
                  <span className="text-slate-400 block text-[10px]">Dias Descartados 50%</span>
                  <span
                    className={`font-bold font-mono text-sm ${
                      selectedEmployeeDetail.diasDescartados > 0 ? 'text-rose-400' : 'text-slate-400'
                    }`}
                  >
                    {selectedEmployeeDetail.diasDescartados}
                  </span>
                </div>
              </div>

              <div className="p-4 overflow-y-auto flex-1 space-y-2">
                {selectedEmployeeDetail.diasDescartados > 0 && (
                  <div className="p-3 rounded-xl bg-rose-950/30 border border-rose-800/40 text-rose-200 text-xs flex items-start gap-2 mb-3">
                    <AlertCircle className="w-4 h-4 text-rose-400 shrink-0 mt-0.5" />
                    <div>
                      <span className="font-bold block text-rose-300">
                        Por que houve {selectedEmployeeDetail.diasDescartados} dia(s) descartado(s)?
                      </span>
                      <span>
                        Na regra geral, dias com <strong>≤ {toleranciaMinutos} minutos</strong> de extra 50% (coluna
                        ExU50) não são computados para pagamento. Veja abaixo exatamente qual dia teve minutos abaixo do
                        limite:
                      </span>
                    </div>
                  </div>
                )}

                <table className="w-full text-xs text-left">
                  <thead>
                    <tr className="border-b border-slate-800 text-slate-400 uppercase tracking-wider text-[10px]">
                      <th className="py-2 px-2.5">Linha / Dia</th>
                      <th className="py-2 px-2.5">Extra 50% (ExU50)</th>
                      <th className="py-2 px-2.5">Extra 100%</th>
                      <th className="py-2 px-2.5">Status 50%</th>
                      <th className="py-2 px-2.5">Motivo / Explicação</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-800/60 font-mono text-xs">
                    {selectedEmployeeDetail.detalhesDias?.map((dia, idx) => {
                      const isDescartado = dia.status50 === 'descartado';
                      const isAprovado = dia.status50 === 'aprovado';
                      return (
                        <tr key={idx} className={isDescartado ? 'bg-rose-950/20 font-semibold' : 'hover:bg-slate-800/20'}>
                          <td className="py-2 px-2.5 font-sans text-slate-300">
                            {dia.labelDia} <span className="text-[10px] text-slate-500 font-mono">({dia.linha})</span>
                          </td>
                          <td
                            className={`py-2 px-2.5 font-bold ${
                              isDescartado ? 'text-rose-400' : isAprovado ? 'text-amber-400' : 'text-slate-500'
                            }`}
                          >
                            {dia.exU50Str}
                          </td>
                          <td className="py-2 px-2.5 text-blue-400">{dia.ex100Str}</td>
                          <td className="py-2 px-2.5 font-sans">
                            {isAprovado && (
                              <span className="px-2 py-0.5 rounded bg-emerald-500/20 text-emerald-300 text-[10px] font-bold border border-emerald-500/30">
                                Aprovado
                              </span>
                            )}
                            {isDescartado && (
                              <span className="px-2 py-0.5 rounded bg-rose-500/20 text-rose-300 text-[10px] font-bold border border-rose-500/30">
                                Descartado (&le; {toleranciaMinutos}m)
                              </span>
                            )}
                            {dia.status50 === 'zerado' && <span className="text-slate-600 text-[11px]">-</span>}
                          </td>
                          <td className="py-2 px-2.5 font-sans text-slate-400 text-[11px]">{dia.motivo}</td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>

              <div className="px-6 py-3 border-t border-slate-800 bg-slate-950 flex items-center justify-between text-xs text-slate-400">
                <span>
                  Dica: caso <strong>{selectedEmployeeDetail.nome}</strong> deva receber cada minuto sem tolerância de{' '}
                  {toleranciaMinutos}m, adicione o nome dele na lista de exceção acima.
                </span>
                <button
                  onClick={() => setSelectedEmployeeDetail(null)}
                  className="px-4 py-1.5 rounded-lg bg-slate-800 hover:bg-slate-700 text-white font-semibold transition-colors shrink-0 ml-3"
                >
                  Fechar
                </button>
              </div>
            </div>
          </div>
        )}
      </main>

      <footer className="border-t border-slate-800/80 bg-slate-950 py-4 px-6 text-center text-xs text-slate-500">
        Calculadora de Horas Extras e Processador de Ponto • Regra de {toleranciaMinutos} Minutos (CLT / Acordo)
      </footer>
    </div>
  );
}
