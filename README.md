# Cadastro de Alunos — Mundo Z

Automação em **Python + Selenium** para cadastrar alunos em uma plataforma online (Mundo Z), inserir cada um no código da sua turma e, no fim, gerar **cartões de acesso em PDF** (usuário, senha e link) — um arquivo por turma.

O progresso é gravado diretamente na planilha (colunas `REGISTRADO` e `INSERIDO`), então a execução pode ser **pausada e retomada** sem refazer o que já foi feito, e é possível rodar **vários navegadores Chrome em paralelo**.

---

## 📋 Índice

- [Funcionalidades](#-funcionalidades)
- [Pré-requisitos](#-pré-requisitos)
- [Instalação](#-instalação)
- [Configuração (.env)](#-configuração-env)
- [Planilha de dados](#-planilha-de-dados)
- [Execução](#-execução)
  - [1. Cadastro e inserção de turma](#1-student_registrationpy)
  - [2. Cartões de acesso em PDF](#2-generate_login_cardspy)
- [Estrutura do projeto](#-estrutura-do-projeto)
- [Como funciona por dentro](#-como-funciona-por-dentro)
- [Segurança e boas práticas](#-segurança-e-boas-práticas)
- [Observações](#-observações)

---

## ✅ Funcionalidades

**`student_registration.py`**
- Lê a planilha `dados.xlsx` e processa apenas os alunos **pendentes** (`INSERIDO` ≠ `OK`).
- **Etapa 1 — Cadastro:** abre a página de cadastro, preenche nome, nickname (= RA), e-mail e senha, aceita os termos e clica em "Cadastrar-se".
- **Etapa 2 — Turma:** faz login com RA + senha, dispensa o tutorial (se aparecer), clica em "Inserir Código", digita o código da turma e confirma; depois encerra a sessão ("Sair").
- Grava `OK`/`FALHA` em `REGISTRADO` e `INSERIDO` a cada aluno, **salvando o Excel na hora**.
- Suporta **N navegadores em paralelo** (um Chrome real por thread), com divisão intercalada das pendências e um `threading.Lock` para proteger as gravações no Excel.
- Uma falha nunca derruba o roda: marca `FALHA`, salva e segue para o próximo aluno.

**`generate_login_cards.py`**
- Agrupa os alunos por `TURMA` e gera **um PDF por turma**.
- Grade de **4 cartões por página** (2 colunas × 2 linhas) em A4 retrato.
- Cada cartão mostra: cabeçalho "ACESSO MUNDO Z", nome (sem acentos, maiúsculas), turma, **usuário (RA)**, **senha** e **link de acesso**.
- Saída em `DIR_DATA/pdf/<turma>.pdf` (ex.: `2º ANO A` → `2_ano_a.pdf`).

---

## 🧰 Pré-requisitos

- **Python 3.10+**
- **Google Chrome** instalado (o Selenium abre o Chrome local)
- ChromeDriver gerenciado automaticamente pelo Selenium 4 (não precisa baixar manualmente)
- Uma planilha `dados.xlsx` com os alunos (veja [Planilha de dados](#-planilha-de-dados))

---

## ⚙️ Instalação

```bash
# 1. clone o repositório
git clone <url-do-repositorio>
cd register_students

# 2. crie e ative um ambiente virtual
python -m venv .venv
source .venv/bin/activate      # macOS / Linux
# .venv\Scripts\activate       # Windows

# 3. instale as dependências
pip install pandas openpyxl selenium python-dotenv fpdf unidecode
```

| Pacote | Para que serve |
|---|---|
| `pandas` + `openpyxl` | ler e salvar a planilha `.xlsx` |
| `selenium` | controlar o navegador Chrome |
| `python-dotenv` | carregar variáveis do arquivo `.env` |
| `fpdf` | gerar os PDFs dos cartões |
| `unidecode` | remover acentos dos nomes no PDF |

> Se preferir, gere um `requirements.txt` com `pip freeze > requirements.txt` e use `pip install -r requirements.txt`.

---

## 🔧 Configuração (`.env`)

Crie um arquivo **`.env`** na raiz do projeto (junto dos scripts):

```ini
URL_CADASTRAR=https://exemplo.com/cadastrar
URL_LOGIN=https://exemplo.com/
PASS_FIXED=SuaSenhaPadrao
DIR_DATA=C:/caminho/absoluto/para/register_students
```

| Variável | Obrigatória | Descrição |
|---|---|---|
| `URL_CADASTRAR` | Sim | Página de cadastro da plataforma |
| `URL_LOGIN` | Sim | Raiz do site (tela de login) |
| `PASS_FIXED` | Sim | Senha padrão usada no cadastro, no login e impressa nos cartões |
| `DIR_DATA` | Não | Pasta onde ficam `dados.xlsx` e os PDFs. **Padrão:** a própria pasta do script |

> ⚠️ **Nunca suba o `.env` para o Git.** Ele contém senha. O `.gitignore` da raiz já ignora arquivos `.env`.

---

## 📊 Planilha de dados

Coloque um arquivo chamado **`dados.xlsx`** em `DIR_DATA` (por padrão, na pasta do projeto).

| Coluna | Conteúdo | Exemplo |
|---|---|---|
| `NOME` | Nome completo do aluno | `Maria Silva` |
| `EMAIL` | E-mail de contato | `maria@email.com` |
| `RA` | Registro do aluno (vira o **usuário**) | `9024611` |
| `COD. TURMA` | Código da turma a ser inserido | `ABC123` |
| `TURMA` | Turma (usada só no PDF dos cartões) | `2º ANO A` |
| `REGISTRADO` | Controle: `OK` / `FALHA` / vazio | `OK` |
| `INSERIDO` | Controle: `OK` / `FALHA` / vazio | `OK` |

- **Deixe `REGISTRADO`/`INSERIDO` vazios** para os alunos que ainda precisam ser processados.
- O script remove o `.0` do RA automaticamente (caso o Excel tenha convertido para número).
- Aluno com `INSERIDO = OK` é **ignorado** — é assim que a retomada funciona.
- Aluno com `REGISTRADO = OK` pula direto para a etapa de login/inserção.

> ⚠️ A planilha contém **dados pessoais** (nome, e-mail, RA). Não versione: o `.gitignore` da raiz já cobre `**/dados.xlsx`.

---

## ▶️ Execução

### 1. `student_registration.py`

```bash
# de dentro da pasta do projeto
python student_registration.py
```

O script pergunta no terminal:

```text
Quantos navegadores em paralelo? (padrão 1):
```

Ou passe a quantidade direto como argumento:

```bash
python student_registration.py 1   # sequencial (1 navegador)
python student_registration.py 3   # 3 navegadores em paralelo
```

**Saída no console** (a cada aluno):

```text
Total pendentes: 24
Navegadores: 3
CADASTRO OK - Maria Silva (RA 9024611)
INSERIDO OK - Maria Silva (RA 9024611)
CADASTRO FALHA - João Souza (RA 9024612): ...
```

Para **retomar**: basta rodar de novo — só os pendentes serão processados.

**Encerramento:** o script finaliza sozinho quando não há mais pendências (fecha todos os navegadores automaticamente).

### 2. `generate_login_cards.py`

```bash
python generate_login_cards.py
```

Cada turma gera um aviso:

```text
Gerado: 2_ano_a.pdf
Gerado: 3_ano_b.pdf
```

Os PDFs ficam em `DIR_DATA/pdf/`.

---

## 📁 Estrutura do projeto

```text
register_students/
├── student_registration.py    # cadastra alunos e insere na turma
├── generate_login_cards.py    # gera os PDFs de cartões de acesso
├── README.md
├── .env                       # configurações (NÃO versionar)
├── dados.xlsx                 # planilha de alunos (NÃO versionar)
└── pdf/                       # PDFs gerados (NÃO versionar)
    ├── 2_ano_a.pdf
    └── ...
```

---

## 🧠 Como funciona por dentro

### Fluxo do `student_registration.py`

```text
para cada aluno pendente (INSERIDO ≠ OK)
        │
        ▼
┌─ ETAPA 1: CADASTRO ────────────────────────────────┐
│ só roda se REGISTRADO ≠ OK                         │
│ preenche Nome / Nickname(RA) / E-mail / Senha      │
│ aceita termos → "Cadastrar-se"                     │
│ sucesso → REGISTRADO = OK   |  erro → FALHA        │
└───────────────────────┬────────────────────────────┘
                        ▼
┌─ ETAPA 2: LOGIN + TURMA ───────────────────────────┐
│ só roda se REGISTRADO = OK                         │
│ login (RA + senha) → "Inserir Código"              │
│ digita COD. TURMA → "Continuar" (2x) → "Sair"      │
│ sucesso → INSERIDO = OK     |  erro → FALHA        │
└───────────────────────┬────────────────────────────┘
                        ▼
              salva o Excel e segue
```

- **Cadastro falhou** → não tenta inserir na turma (sem conta não há o que fazer).
- **Já registrado mas não inserido** → pula o cadastro e faz só login + inserção.
- **Erro nunca derruba o script** → grava `FALHA`, salva e continua.

### Paralelismo e proteção do Excel

```bash
python student_registration.py 4   # 4 navegadores simultâneos
```

- Os índices pendentes são **intercalados** entre as fatias (trabalho equilibrado).
- Cada fatia roda numa thread com o **seu próprio Chrome** (`ThreadPoolExecutor`).
- Fatias vazias são descartadas (ex.: 2 pendentes + 4 navegadores → só 2 abrem).
- Um `threading.Lock()` garante que **só um navegador escreve no Excel por vez** — a leitura dos dados e a gravação ficam dentro do bloco `with lock:`, evitando corrupção do arquivo. A navegação em si acontece fora do lock, cada um no seu navegador.

### Estimativa de tempo

Cada aluno leva em torno de **75–90 s** (as esperas fixas do script somam ~60 s só de `time.sleep`). Tempo total ≈ `pendentes × 90s ÷ N navegadores`.

---

## 🔐 Segurança e boas práticas

- **`.env` nunca vai para o Git** — contém a senha padrão dos alunos.
- **`dados.xlsx` e `pdf/` nunca vão para o Git** — contêm dados pessoais (LGPD).
- Rode primeiro com **1 navegador** para validar se os XPath ainda batem com a plataforma; mude a quantidade só depois.
- Se a plataforma mudar de layout, os seletores (`By.XPATH`, placeholders como `Nome`, `Usuário`, `Digite o código`) precisarão ser atualizados.

---

## 📌 Observações

- Os XPath e os textos dos botões (`Cadastrar-se`, `Entrar`, `Inserir Código`, `Continuar`, `Sair`, `Dispensar Tutorial`) foram escritos para a interface atual da plataforma — qualquer alteração de layout exige ajuste no script.
- O `time.sleep()` é usado em conjunto com `WebDriverWait`; aumente os tempos se a conexão for lenta.
- O RA é sempre tratado como **texto** (sem `.0`) para ser digitado corretamente nos campos.
- Nomes e turmas no PDF passam por `unidecode(...).upper()` — tudo sai em ASCII/maiúsculas, evitando problemas de fonte.
- Para trocar a pasta de dados sem editar código, altere `DIR_DATA` no `.env`.
