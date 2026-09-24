import 'dotenv/config';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import express from 'express';
import {
  getStore,
  upsertSalario,
  deleteSalario,
  addExcecao,
  removeExcecao,
  setToleranciaMinutos,
  DADOS_DIR,
} from './salariosStore.ts';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const PORT = Number(process.env.PORT) || 5175;

const app = express();
app.use(express.json());

const asyncRoute =
  (fn: (req: express.Request, res: express.Response) => Promise<void>) =>
  (req: express.Request, res: express.Response) => {
    fn(req, res).catch(err => {
      console.error(err);
      res.status(500).json({ error: err?.message || 'Erro interno no servidor.' });
    });
  };

const parseNomeParam = (raw: string): string => decodeURIComponent(raw).trim();

app.get(
  '/api/salarios',
  asyncRoute(async (_req, res) => {
    res.json(await getStore());
  })
);

app.post(
  '/api/salarios',
  asyncRoute(async (req, res) => {
    const nome = String(req.body?.nome || '').trim();
    const valorHora = Number(req.body?.valorHora);
    if (!nome) {
      res.status(400).json({ error: 'Nome é obrigatório.' });
      return;
    }
    if (!Number.isFinite(valorHora) || valorHora <= 0) {
      res.status(400).json({ error: 'Valor da hora precisa ser um número maior que zero.' });
      return;
    }
    res.json(await upsertSalario(nome, valorHora));
  })
);

app.delete(
  '/api/salarios/:nome',
  asyncRoute(async (req, res) => {
    const nome = parseNomeParam(req.params.nome);
    if (!nome) {
      res.status(400).json({ error: 'Nome é obrigatório.' });
      return;
    }
    res.json(await deleteSalario(nome));
  })
);

app.post(
  '/api/excecoes',
  asyncRoute(async (req, res) => {
    const nome = String(req.body?.nome || '').trim();
    if (!nome) {
      res.status(400).json({ error: 'Nome é obrigatório.' });
      return;
    }
    res.json(await addExcecao(nome));
  })
);

app.delete(
  '/api/excecoes/:nome',
  asyncRoute(async (req, res) => {
    const nome = parseNomeParam(req.params.nome);
    if (!nome) {
      res.status(400).json({ error: 'Nome é obrigatório.' });
      return;
    }
    res.json(await removeExcecao(nome));
  })
);

app.post(
  '/api/configuracao',
  asyncRoute(async (req, res) => {
    const toleranciaMinutos = Number(req.body?.toleranciaMinutos);
    if (!Number.isInteger(toleranciaMinutos) || toleranciaMinutos < 0) {
      res.status(400).json({ error: 'Tolerância precisa ser um número inteiro maior ou igual a zero.' });
      return;
    }
    res.json(await setToleranciaMinutos(toleranciaMinutos));
  })
);

app.post('/api/desligar', (_req, res) => {
  res.json({ ok: true });
  // Dá tempo da resposta chegar no navegador antes do processo (e a janela
  // preta do iniciar.bat) encerrar.
  setTimeout(() => process.exit(0), 200);
});

// Em produção (depois de `npm run build`), o próprio servidor entrega a
// interface — um processo só, uma porta só, mais simples pro atalho do Windows.
// App de tela única (sem rotas de cliente), então servir estático já basta:
// express.static entrega index.html em "/" e os assets em suas próprias rotas.
const distDir = path.resolve(__dirname, '..', 'dist');
if (fs.existsSync(path.join(distDir, 'index.html'))) {
  app.use(express.static(distDir));
}

app.listen(PORT, '127.0.0.1', () => {
  console.log(`Calculadora de Horas Extras rodando em http://localhost:${PORT}`);
  console.log(`Pasta de dados: ${DADOS_DIR}`);
  if (!fs.existsSync(path.join(distDir, 'index.html'))) {
    console.log('Aviso: build de produção não encontrada (rode "npm run build"). Servindo apenas a API.');
  }
});
