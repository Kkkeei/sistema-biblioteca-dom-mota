📚 Sistema Dom Mota — Gerenciador de Biblioteca Escolar

    Sistema desktop para gerenciamento de empréstimos, acervo e ranking de leitura da Escola Dom Mota.

👨‍💻 Desenvolvedor

Davi Lucas Galdino Braz
Técnico em Desenvolvimento de Sistemas
Projeto orientado pela professora Gislene Laurenço
🖥️ Requisitos
Requisito 	Versão mínima
Python 	3.10+
customtkinter 	5.2+
Pillow 	9.0+
openpyxl 	3.1+
⚙️ Instalação

1. Clone o repositório ou copie os arquivos para a pasta do projeto:

/home/davi/Documentos/Projeto/App_Biblioteca-main/

2. Crie e ative o ambiente virtual:

python -m venv venv
source venv/bin/activate        # Linux/macOS
venv\Scripts\activate           # Windows

3. Instale as dependências:

pip install customtkinter pillow openpyxl

4. Coloque o logo da escola na pasta do projeto:

erasebg-transformed.png

    Se o arquivo não existir, o sistema exibe o texto "DOM MOTA" no lugar.

🗄️ Estrutura de Arquivos

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

🚀 Como usar
Primeiro uso

Passo 1 — Popule o banco de alunos

Se você já tem os alunos cadastrados no biblioteca.db (tabela alunos), pule para o Passo 2.

Caso contrário, cadastre os alunos manualmente pelo sistema em Gerenciar Alunos, ou use um script de importação.

Passo 2 — Execute a migração para criar o turmas.db:

python migrar_db.py

Esse script lê todos os alunos do biblioteca.db e cria o turmas.db com uma tabela por turma, no formato:

ano1_turma1_MANHA     →  30 alunos
ano1_turma2_MANHA     →  30 alunos
ano2_turma1_MANHA     →  30 alunos
ano2_turma2_TARDE     →  30 alunos
...
ano9_turma3_TARDE  
