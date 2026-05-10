# 📚 Sistema Dom Mota — Gerenciador de Biblioteca Escolar

> Sistema desktop para gerenciamento de empréstimos, acervo e ranking de leitura da Escola Dom Mota.

---

## 👨‍💻 Desenvolvedor

**Davi Lucas Galdino Braz**  
Técnico em Desenvolvimento de Sistemas  
Projeto orientado pela professora **Gislene Laurenço**

---

## 🖥️ Requisitos

| Requisito | Versão mínima |
|-----------|--------------|
| Python | 3.10+ |
| customtkinter | 5.2+ |
| Pillow | 9.0+ |
| openpyxl | 3.1+ |

---

## ⚙️ Instalação

**1. Clone o repositório ou copie os arquivos para a pasta do projeto:**
```
/home/davi/Documentos/Projeto/App_Biblioteca-main/
```

**2. Crie e ative o ambiente virtual:**
```bash
python -m venv venv
source venv/bin/activate        # Linux/macOS
venv\Scripts\activate           # Windows
```

**3. Instale as dependências:**
```bash
pip install customtkinter pillow openpyxl
```

**4. Coloque o logo da escola na pasta do projeto:**
```
erasebg-transformed.png
```
> Se o arquivo não existir, o sistema exibe o texto "DOM MOTA" no lugar.

---

## 🗄️ Estrutura de Arquivos

```
App_Biblioteca-main/
│
├── main.py               # Sistema principal (interface gráfica)
├── migrar_db.py          # Script de migração — cria o turmas.db
│
├── biblioteca.db         # Banco principal (criado automaticamente)
├── turmas.db             # Banco de turmas (criado pelo migrar_db.py)
│
├── erasebg-transformed.png  # Logo da escola
│
└── Relatorios/           # Pasta gerada automaticamente
    ├── Controle_*.xlsx
    ├── RankingMensal_*.xlsx
    ├── Ranking_*.xlsx
    └── Export_*.xlsx
```

---

## 🚀 Como usar

### Primeiro uso

**Passo 1 — Popule o banco de alunos**

Se você já tem os alunos cadastrados no `biblioteca.db` (tabela `alunos`), pule para o Passo 2.

Caso contrário, cadastre os alunos manualmente pelo sistema em **Gerenciar Alunos**, ou use um script de importação.

**Passo 2 — Execute a migração para criar o `turmas.db`:**
```bash
python migrar_db.py
```

Esse script lê todos os alunos do `biblioteca.db` e cria o `turmas.db` com **uma tabela por turma**, no formato:

```
ano1_turma1_MANHA     →  30 alunos
ano1_turma2_MANHA     →  30 alunos
ano2_turma1_MANHA     →  30 alunos
ano2_turma2_TARDE     →  30 alunos
...
ano9_turma3_TARDE     →  35 alunos
```

**Passo 3 — Execute o sistema:**
```bash
python main.py
```

---

## 📋 Funcionalidades

### 📋 Empréstimos
Registra a retirada de livros por alunos.

- **Buscar por Nome** — digita parte do nome e o aluno aparece em tempo real
- **Selecionar por Turma** — painel lateral com todas as turmas; clica na turma e lista os alunos dela
- Ao registrar, informa: **código do livro** (4 dígitos), **título** e **data de entrega**

---

### 👤 Gerenciar Alunos
Cadastro completo de alunos.

- Campos: Nome, Série, Turma (1/2/3), Turno
- Busca em tempo real por nome
- Exportação para Excel (`Export_alunos_*.xlsx`)
- Exclusão com confirmação

---

### 📚 Acervo
Catálogo de livros da biblioteca.

- Campos: Título, Autor, Gênero
- **Código** — 4 dígitos numéricos únicos, definidos pela bibliotecária
- **Estante** — letras de A a H (seleção por botões)
- **Fileira** — números de 1 a 5 (seleção por botões)
- Pesquisa simultânea por qualquer combinação de campos
- Exportação para Excel (`Export_acervo_*.xlsx`)

---

### ⏳ Pendências
Lista todos os livros ainda não devolvidos.

- Mostra: livro, aluno, série/turma e data prevista de entrega
- Botão **Devolver** remove o empréstimo e adiciona +1 ao ranking do aluno automaticamente

---

### 🏆 Ranking
Acompanhamento de leitura por aluno e por turma.

- **Aba Alunos** — lista completa ordenada por livros lidos, com 🥇🥈🥉 para os 3 primeiros
- **Aba Turmas** — todas as turmas cadastradas, mesmo as com 0 leituras
- Botão **Exportar Excel** — gera planilha com as duas abas
- Botão **Zerar Ranking** — reseta todas as contagens (com confirmação)

---

### 📊 Relatórios

#### 📋 Controle de Leitura
Planilha por turma com:
- Cabeçalho amarelo com série, turma, turno e mês/ano
- Todos os alunos em ordem alfabética
- Colunas de 1ª a 4ª semana com "X" marcado se o aluno leu
- Coluna TOTAL de livros lidos
- Bordas no padrão escolar

#### 🏆 Ranking Mensal
Planilha no formato usado pela escola (igual à ficha impressa):
- Nome da escola no topo
- Mês selecionável (Janeiro a Dezembro)
- Turma e Professor(a)
- Alunos na vertical, semanas na horizontal
- "X" em azul nas semanas com leitura
- Total destacado em amarelo

#### 🏫 Ranking Completo
Exporta ranking geral com duas abas:
- **Ranking Alunos** — todos os alunos que leram, ordenados por quantidade
- **Ranking Turmas** — todas as turmas, incluindo as com 0 leituras

---

## 🗃️ Estrutura do Banco de Dados

### `biblioteca.db`

| Tabela | Colunas principais |
|--------|-------------------|
| `alunos` | nome, serie, turma, turno |
| `acervo` | titulo, autor, genero, codigo, estante, fileira |
| `emprestimos` | aluno_nome, livro_titulo, data_entrega, serie, turma, turno |
| `ranking` | nome, lidos, serie, turma, turno |
| `turmas` | serie, turma, turno (índice único) |

### `turmas.db`

Uma tabela por turma, nomeadas no formato `ano{serie}_turma{n}_{TURNO}`.

Cada tabela contém: `id`, `nome`, `serie`, `turma`, `turno`.

---

## 🔄 Fluxo de Uso Diário

```
Aluno pega livro
       ↓
  Empréstimos → Seleciona aluno → Informa código e título
       ↓
  [livro fica em Pendências]
       ↓
Aluno devolve livro
       ↓
  Pendências → Devolver
       ↓
  Ranking atualiza automaticamente (+1 para o aluno)
       ↓
Fim do mês → Relatórios → Ranking Mensal por turma
```

---

## ⚠️ Observações

- O campo **Turno** aceita texto livre (ex: `MANHÃ`, `TARDE`, `MANHA`). Use um padrão consistente no cadastro para que os filtros funcionem corretamente.
- O **código do livro** deve ter exatamente 4 dígitos numéricos. O sistema valida e bloqueia códigos inválidos.
- Execute `migrar_db.py` novamente sempre que novos alunos forem cadastrados e você quiser atualizar o `turmas.db`.
- Os relatórios são salvos automaticamente na pasta `Relatorios/` dentro do diretório do projeto.
- O banco `biblioteca.db` é criado automaticamente na primeira execução do `main.py`.
