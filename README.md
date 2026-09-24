# Calculadora de Horas Extras

Processa o arquivo de ponto (uma aba por funcionário), calcula horas extras
de 50% e 100% com a regra de carência de 19 minutos, e mantém a lista de
salários automaticamente numa pasta fixa (`dados/salarios.xlsx`).

## Uso no Windows

**Primeira vez:**
1. Instale o [Node.js](https://nodejs.org) (versão LTS).
2. Dê duplo clique em `instalar.bat`.

**Uso do dia a dia:**
- Dê duplo clique em `iniciar.bat`. O navegador abre sozinho em
  `http://localhost:5175`. Para encerrar, feche a janela preta que abriu
  junto.

## Como funciona

- **Lista de salários**: fica em `dados/salarios.xlsx`, criada automaticamente
  na primeira execução. O programa lê esse arquivo sozinho — não precisa
  enviar manualmente.
- **Funcionário novo**: se o arquivo de ponto tiver uma aba com um nome que
  não está na lista de salários, o programa pede o valor da hora (R$) e
  cadastra automaticamente, usando exatamente o nome da aba.
- **Excluir da lista**: cada funcionário na lista de salários tem um botão
  de lixeira, com confirmação antes de excluir.
- **Lista de exceção**: funcionários nessa lista não têm carência de 19
  minutos — cada minuto extra é somado integralmente.
- **Arquivo de ponto**: continua sendo enviado manualmente a cada uso (botão
  "Selecionar Arquivo de PONTO").

Evite manter `dados/salarios.xlsx` aberto no Excel enquanto usa o programa —
como os dois tentam controlar o mesmo arquivo, um pode sobrescrever o outro.

## Desenvolvimento

```bash
npm install
npm run dev          # interface (http://localhost:3000)
npm run dev:server   # API local (http://localhost:5175), em outro terminal
```

`npm run build` gera a versão de produção em `dist/`; `npm start` sobe o
servidor único (API + interface) usado pelo `iniciar.bat`.
