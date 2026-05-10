import customtkinter as ctk
import sqlite3
from pathlib import Path
from tkinter import messagebox
from PIL import Image
import sys
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
from datetime import datetime
import locale

try:
    locale.setlocale(locale.LC_TIME, "pt_BR.UTF-8")
except Exception:
    try:
        locale.setlocale(locale.LC_TIME, "Portuguese_Brazil.1252")
    except Exception:
        pass

# ─────────────────────────────────────────────
#  CONFIGURAÇÕES GLOBAIS
# ─────────────────────────────────────────────
BASE_DIR   = Path(__file__).parent.absolute()
DB_PATH    = BASE_DIR / "biblioteca.db"
LOGO_PATH  = BASE_DIR / "erasebg-transformed.png"
EXPORT_DIR = BASE_DIR / "Relatorios"
EXPORT_DIR.mkdir(exist_ok=True)

ctk.set_appearance_mode("light")

TURMAS_OPCOES  = ["1", "2", "3"]
ESTANTES_LETRAS = ["A", "B", "C", "D", "E", "F", "G", "H"]
FILEIRAS_NUM    = ["1", "2", "3", "4", "5"]

MESES_PT = {
    1:"JANEIRO", 2:"FEVEREIRO", 3:"MARÇO",    4:"ABRIL",
    5:"MAIO",    6:"JUNHO",     7:"JULHO",     8:"AGOSTO",
    9:"SETEMBRO",10:"OUTUBRO", 11:"NOVEMBRO", 12:"DEZEMBRO"
}

# ── Paleta ───────────────────────────────────
C = {
    "amarelo":      "#FFCC00",
    "vermelho":     "#A31C1C",
    "vermelho_esc": "#7A1515",
    "verde":        "#2E7D32",
    "verde_esc":    "#1B5E20",
    "fundo":        "#F4F4F6",
    "card":         "#FFFFFF",
    "borda":        "#E8E8EC",
    "texto":        "#1A1A2E",
    "subtexto":     "#6B7280",
    "cinza_btn":    "#6B7280",
    "cinza_hover":  "#4B5563",
    "perigo":       "#B71C1C",
    "perigo_hover": "#7F0000",
    "inativo":      "#E5E7EB",
    "inativo_txt":  "#374151",
}

FONTE_TITULO  = ("Segoe UI", 20, "bold")
FONTE_SECAO   = ("Segoe UI", 14, "bold")
FONTE_CORPO   = ("Segoe UI", 13)
FONTE_PEQUENA = ("Segoe UI", 11)
FONTE_BOLD    = ("Segoe UI", 13, "bold")

# ── Helpers de estilo ─────────────────────────
def estilo_entry(width=180, **kw):
    return dict(width=width, height=38, corner_radius=10,
                border_color=C["borda"], border_width=1,
                fg_color="white", text_color=C["texto"],
                font=FONTE_CORPO, **kw)

def estilo_segmented(width=200):
    return dict(width=width, height=34,
                selected_color=C["vermelho"],
                selected_hover_color=C["vermelho_esc"],
                unselected_color=C["inativo"],
                unselected_hover_color="#D1D5DB",
                text_color="white", font=FONTE_BOLD, corner_radius=8)

def estilo_btn_primario(width=140, height=38, **kw):
    return dict(width=width, height=height, corner_radius=10,
                fg_color=C["vermelho"], hover_color=C["vermelho_esc"],
                text_color="white", font=FONTE_BOLD, **kw)

def estilo_btn_verde(width=140, height=38, **kw):
    return dict(width=width, height=height, corner_radius=10,
                fg_color=C["verde"], hover_color=C["verde_esc"],
                text_color="white", font=FONTE_BOLD, **kw)

def estilo_btn_cinza(width=120, height=38, **kw):
    return dict(width=width, height=height, corner_radius=10,
                fg_color=C["cinza_btn"], hover_color=C["cinza_hover"],
                text_color="white", font=FONTE_CORPO, **kw)

def cabecalho_tela(container, titulo, voltar_cmd):
    topo = ctk.CTkFrame(container, fg_color=C["vermelho"], corner_radius=0, height=58)
    topo.pack(fill="x"); topo.pack_propagate(False)
    ctk.CTkButton(topo, text="← Voltar", width=110, height=36,
                  fg_color="transparent", hover_color=C["vermelho_esc"],
                  text_color="white", font=("Segoe UI", 13),
                  command=voltar_cmd).pack(side="left", padx=18, pady=11)
    ctk.CTkLabel(topo, text=titulo, font=("Segoe UI", 16, "bold"),
                 text_color="white").pack(side="left", padx=6)

def separador(parent, pady=(0, 0)):
    ctk.CTkFrame(parent, height=1, fg_color=C["borda"]).pack(fill="x", pady=pady)

def lbl_sub(parent, texto):
    ctk.CTkLabel(parent, text=texto, font=FONTE_PEQUENA,
                 text_color=C["subtexto"]).pack(side="left", padx=(8, 2))

def thin_border():
    s = Side(style="thin")
    return Border(left=s, right=s, top=s, bottom=s)

def medium_border():
    s = Side(style="medium")
    return Border(left=s, right=s, top=s, bottom=s)

# ─────────────────────────────────────────────
#  BANCO DE DADOS
# ─────────────────────────────────────────────
def iniciar_db():
    conn   = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()

    cursor.execute("CREATE TABLE IF NOT EXISTS acervo (titulo TEXT, autor TEXT, genero TEXT)")
    cursor.execute("CREATE TABLE IF NOT EXISTS alunos (nome TEXT, serie TEXT, turma TEXT, turno TEXT)")
    cursor.execute("CREATE TABLE IF NOT EXISTS emprestimos (aluno_nome TEXT, livro_titulo TEXT, data_entrega TEXT, serie TEXT, turma TEXT, turno TEXT)")
    cursor.execute("CREATE TABLE IF NOT EXISTS ranking (nome TEXT PRIMARY KEY, lidos INTEGER DEFAULT 0, serie TEXT, turma TEXT, turno TEXT)")
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS turmas (
            id    INTEGER PRIMARY KEY AUTOINCREMENT,
            serie TEXT NOT NULL,
            turma TEXT NOT NULL,
            turno TEXT NOT NULL,
            UNIQUE(serie, turma, turno)
        )
    """)

    cursor.execute("CREATE INDEX IF NOT EXISTS idx_alunos_nome ON alunos (nome)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_acervo_titulo ON acervo (titulo)")

    def garantir_col(tabela, coluna, tipo="TEXT"):
        cursor.execute(f"PRAGMA table_info({tabela})")
        if coluna not in [r[1] for r in cursor.fetchall()]:
            cursor.execute(f"ALTER TABLE {tabela} ADD COLUMN {coluna} {tipo}")

    for col in ["turma", "turno"]:          garantir_col("alunos", col)
    for col in ["serie", "turma", "turno"]: garantir_col("emprestimos", col)
    for col in ["serie", "turma", "turno"]: garantir_col("ranking", col)
    for col in ["estante", "fileira", "codigo"]: garantir_col("acervo", col)

    # Sincroniza tabela turmas a partir dos alunos cadastrados
    cursor.execute("""
        INSERT OR IGNORE INTO turmas (serie, turma, turno)
        SELECT DISTINCT serie, turma, turno FROM alunos
        WHERE serie != '' AND turma != '' AND turno != ''
    """)

    conn.commit()
    conn.close()

def listar_turmas_db():
    """Retorna lista de (serie, turma, turno) ordenada."""
    conn = sqlite3.connect(DB_PATH); c = conn.cursor()
    c.execute("SELECT serie, turma, turno FROM turmas ORDER BY CAST(serie AS INTEGER), turma, turno")
    rows = c.fetchall(); conn.close()
    return rows

def turmas_da_serie(serie):
    conn = sqlite3.connect(DB_PATH); c = conn.cursor()
    c.execute("SELECT DISTINCT turma FROM alunos WHERE serie=? AND turma IS NOT NULL AND turma!='' ORDER BY turma",
              (serie.strip(),))
    rows = [r[0] for r in c.fetchall()]; conn.close()
    return rows if rows else TURMAS_OPCOES


# ─────────────────────────────────────────────
#  APLICAÇÃO
# ─────────────────────────────────────────────
class BibliotecaApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Sistema Dom Mota - v8.5")
        if sys.platform.startswith("linux"):
            self.attributes("-zoomed", True)
        else:
            self.state("zoomed")

        iniciar_db()
        self.configure(fg_color=C["fundo"])
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        self.topo = ctk.CTkFrame(self, height=130, fg_color=C["amarelo"], corner_radius=0)
        self.topo.grid(row=0, column=0, sticky="ew")
        try:
            img_pil   = Image.open(LOGO_PATH)
            self.logo = ctk.CTkImage(light_image=img_pil, size=(110, 110))
            ctk.CTkLabel(self.topo, image=self.logo, text="").place(relx=0.5, rely=0.5, anchor="center")
        except Exception:
            ctk.CTkLabel(self.topo, text="DOM MOTA", font=("Segoe UI", 35, "bold"),
                         text_color=C["vermelho"]).place(relx=0.5, rely=0.5, anchor="center")

        self.container = ctk.CTkFrame(self, fg_color=C["fundo"], corner_radius=0)
        self.container.grid(row=1, column=0, sticky="nsew")

        ctk.CTkButton(self, text="ℹ Créditos", width=100, height=28,
                      fg_color="#555", hover_color="#333", corner_radius=8,
                      font=("Segoe UI", 11), command=self.mostrar_creditos
                      ).place(relx=1.0, rely=1.0, anchor="se", x=-16, y=-16)

        self.mostrar_tela_principal()

    def limpar_container(self):
        for w in self.container.winfo_children(): w.destroy()

    # ─────────────────────────────────────────
    #  MENU PRINCIPAL
    # ─────────────────────────────────────────
    def mostrar_tela_principal(self):
        self.limpar_container()
        wrapper = ctk.CTkFrame(self.container, fg_color="transparent")
        wrapper.place(relx=0.5, rely=0.5, anchor="center")
        ctk.CTkLabel(wrapper, text="O que você deseja fazer?",
                     font=("Segoe UI", 15), text_color=C["subtexto"]).pack(pady=(0, 18))
        grid = ctk.CTkFrame(wrapper, fg_color="transparent"); grid.pack()
        icones = ["📋", "👤", "📚", "⏳", "🏆", "📊"]
        opcoes = [
            ("Empréstimos",      self.tela_emprestimo),
            ("Gerenciar Alunos", self.tela_alunos),
            ("Acervo",           self.tela_acervo),
            ("Pendências",       self.tela_pendentes),
            ("Ranking",          self.tela_ranking_unificado),
            ("Relatórios",       self.tela_relatorios),
        ]
        for i, ((txt, cmd), ico) in enumerate(zip(opcoes, icones)):
            card = ctk.CTkFrame(grid, fg_color=C["card"], corner_radius=16,
                                border_width=1, border_color=C["borda"],
                                width=200, height=120)
            card.grid(row=i//3, column=i%3, padx=12, pady=12)
            card.grid_propagate(False)
            inner = ctk.CTkFrame(card, fg_color="transparent")
            inner.place(relx=0.5, rely=0.5, anchor="center")
            ctk.CTkLabel(inner, text=ico, font=("Segoe UI", 28)).pack()
            ctk.CTkLabel(inner, text=txt, font=("Segoe UI", 13, "bold"),
                         text_color=C["texto"]).pack(pady=(4, 8))
            ctk.CTkButton(inner, text="Abrir",
                          **estilo_btn_primario(width=100, height=30),
                          command=cmd).pack()

    # ─────────────────────────────────────────
    #  CRÉDITOS
    # ─────────────────────────────────────────
    def mostrar_creditos(self):
        pop = ctk.CTkToplevel(self)
        pop.title("Créditos"); pop.geometry("420x280")
        pop.resizable(False, False); pop.attributes("-topmost", True)
        pop.configure(fg_color=C["card"])
        ctk.CTkLabel(pop, text="Sobre o Desenvolvedor",
                     font=FONTE_TITULO, text_color=C["vermelho"]).pack(pady=(24, 8))
        ctk.CTkFrame(pop, height=1, fg_color=C["borda"]).pack(fill="x", padx=30)
        ctk.CTkLabel(pop,
                     text="Davi Lucas Galdino Braz\nTécnico em Desenvolvimento de Sistemas\n\nProjeto orientado pela professora\nGislene Laurenço",
                     justify="center", font=FONTE_CORPO, text_color=C["texto"]).pack(pady=16)
        ctk.CTkButton(pop, text="Fechar", **estilo_btn_cinza(width=120),
                      command=pop.destroy).pack(pady=8)

    # ─────────────────────────────────────────
    #  EMPRÉSTIMOS  — com seleção de turma
    # ─────────────────────────────────────────
    def tela_emprestimo(self):
        self.limpar_container()
        cabecalho_tela(self.container, "📋  Empréstimos", self.mostrar_tela_principal)
        centro = ctk.CTkFrame(self.container, fg_color=C["fundo"], corner_radius=0)
        centro.pack(fill="both", expand=True)

        # ── Abas ────────────────────────────────────────────────────
        abas_frame = ctk.CTkFrame(centro, fg_color="transparent")
        abas_frame.pack(fill="x", padx=30, pady=(16, 0))

        btn_nome  = ctk.CTkButton(abas_frame, text="🔍 Buscar por Nome",
                                   width=190, height=36, corner_radius=10,
                                   fg_color=C["vermelho"], hover_color=C["vermelho_esc"],
                                   text_color="white", font=FONTE_BOLD)
        btn_nome.pack(side="left", padx=(0, 8))
        btn_turma = ctk.CTkButton(abas_frame, text="🏫 Selecionar por Turma",
                                   width=210, height=36, corner_radius=10,
                                   fg_color=C["inativo"], hover_color="#D1D5DB",
                                   text_color=C["inativo_txt"], font=FONTE_CORPO)
        btn_turma.pack(side="left")

        # Container principal que vai ser trocado entre abas
        area = ctk.CTkFrame(centro, fg_color="transparent")
        area.pack(fill="both", expand=True, padx=30, pady=12)

        def _montar_lista_alunos(parent, rows):
            for n, s, tr, tn in rows:
                card = ctk.CTkFrame(parent, fg_color=C["card"], corner_radius=10,
                                    border_width=1, border_color=C["borda"])
                card.pack(fill="x", pady=3)
                ctk.CTkLabel(card, text="👤", font=("Segoe UI", 18), width=44
                             ).pack(side="left", padx=10)
                info = ctk.CTkFrame(card, fg_color="transparent")
                info.pack(side="left", fill="x", expand=True, pady=10)
                ctk.CTkLabel(info, text=n, font=FONTE_BOLD,
                             text_color=C["texto"], anchor="w").pack(anchor="w", padx=4)
                meta = f"Série: {s or '—'}  ·  Turma: {tr or '—'}  ·  Turno: {tn or '—'}"
                ctk.CTkLabel(info, text=meta, font=FONTE_PEQUENA,
                             text_color=C["subtexto"], anchor="w").pack(anchor="w", padx=4)
                ctk.CTkButton(card, text="Registrar Empréstimo",
                              **estilo_btn_verde(width=180),
                              command=lambda x=n, y=s, z=tr, w=tn: self.pop_livro(x, y, z, w)
                              ).pack(side="right", padx=12, pady=10)

        # ── Aba: Busca por nome ──────────────────────────────────────
        def aba_nome():
            for w in area.winfo_children(): w.destroy()
            btn_nome.configure(fg_color=C["vermelho"], text_color="white", font=FONTE_BOLD)
            btn_turma.configure(fg_color=C["inativo"], text_color=C["inativo_txt"], font=FONTE_CORPO)

            card_busca = ctk.CTkFrame(area, fg_color=C["card"], corner_radius=14,
                                       border_width=1, border_color=C["borda"])
            card_busca.pack(fill="x")
            ctk.CTkLabel(card_busca, text="Buscar aluno por nome",
                         font=FONTE_SECAO, text_color=C["vermelho"]
                         ).pack(anchor="w", padx=20, pady=(16, 8))
            linha = ctk.CTkFrame(card_busca, fg_color="transparent")
            linha.pack(fill="x", padx=20, pady=(0, 16))
            busc = ctk.CTkEntry(linha, placeholder_text="🔍 Digite o nome do aluno...",
                                **estilo_entry(width=500))
            busc.pack(side="left"); busc.focus_set()

            sc = ctk.CTkScrollableFrame(area, fg_color="transparent", corner_radius=0)
            sc.pack(fill="both", expand=True, pady=(8, 0))

            def query(e=None):
                for w in sc.winfo_children(): w.destroy()
                t = busc.get().upper().strip()
                if len(t) < 2: return
                conn = sqlite3.connect(DB_PATH); c = conn.cursor()
                c.execute("SELECT nome, serie, turma, turno FROM alunos WHERE nome LIKE ? LIMIT 15",
                          (f"%{t}%",))
                rows = c.fetchall(); conn.close()
                if not rows:
                    ctk.CTkLabel(sc, text="Nenhum aluno encontrado.",
                                 font=FONTE_CORPO, text_color=C["subtexto"]).pack(pady=20)
                    return
                _montar_lista_alunos(sc, rows)

            busc.bind("<KeyRelease>", query)

        # ── Aba: Selecionar por turma ────────────────────────────────
        def aba_turma():
            for w in area.winfo_children(): w.destroy()
            btn_turma.configure(fg_color=C["vermelho"], text_color="white", font=FONTE_BOLD)
            btn_nome.configure(fg_color=C["inativo"], text_color=C["inativo_txt"], font=FONTE_CORPO)

            # Layout: esquerda = lista de turmas | direita = alunos da turma
            frame_split = ctk.CTkFrame(area, fg_color="transparent")
            frame_split.pack(fill="both", expand=True)
            frame_split.columnconfigure(0, weight=0)
            frame_split.columnconfigure(1, weight=1)
            frame_split.rowconfigure(0, weight=1)

            # ── Painel esquerdo: botões de turma rolável ─────────────
            panel_esq = ctk.CTkFrame(frame_split, fg_color=C["card"],
                                      corner_radius=14, border_width=1,
                                      border_color=C["borda"], width=180)
            panel_esq.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
            panel_esq.grid_propagate(False)

            ctk.CTkLabel(panel_esq, text="Turmas", font=FONTE_SECAO,
                         text_color=C["vermelho"]).pack(anchor="w", padx=14, pady=(14, 6))
            separador(panel_esq)

            sc_turmas = ctk.CTkScrollableFrame(panel_esq, fg_color="transparent",
                                                corner_radius=0, width=160)
            sc_turmas.pack(fill="both", expand=True, padx=4, pady=4)

            # ── Painel direito: lista de alunos ──────────────────────
            panel_dir = ctk.CTkFrame(frame_split, fg_color="transparent")
            panel_dir.grid(row=0, column=1, sticky="nsew")

            lbl_turma_sel = ctk.CTkLabel(panel_dir, text="← Selecione uma turma",
                                          font=FONTE_CORPO, text_color=C["subtexto"])
            lbl_turma_sel.pack(anchor="w", pady=(6, 4))

            sc_alunos = ctk.CTkScrollableFrame(panel_dir, fg_color="transparent", corner_radius=0)
            sc_alunos.pack(fill="both", expand=True)

            # Busca turmas únicas do banco (sem duplicatas)
            conn = sqlite3.connect(DB_PATH); c = conn.cursor()
            c.execute("""
                SELECT DISTINCT serie, turma, turno
                FROM alunos
                WHERE serie != '' AND turma != '' AND turno != ''
                ORDER BY CAST(serie AS INTEGER), turma, turno
            """)
            turmas_unicas = c.fetchall(); conn.close()

            btn_turmas_lista = []

            def mostrar_alunos_turma(s, t, tn, btn_clicado):
                # Reseta visual de todos os botões
                for b in btn_turmas_lista:
                    b.configure(fg_color=C["fundo"], text_color=C["texto"])
                btn_clicado.configure(fg_color=C["vermelho"], text_color="white")

                for w in sc_alunos.winfo_children(): w.destroy()
                lbl_turma_sel.configure(
                    text=f"Alunos — {s}º ano  Turma {t}  {tn}",
                    text_color=C["vermelho"], font=FONTE_BOLD)

                conn2 = sqlite3.connect(DB_PATH); c2 = conn2.cursor()
                c2.execute("""SELECT nome, serie, turma, turno FROM alunos
                              WHERE serie=? AND turma=? AND turno=? ORDER BY nome""",
                           (s, t, tn))
                rows = c2.fetchall(); conn2.close()
                _montar_lista_alunos(sc_alunos, rows)

            for s, t, tn in turmas_unicas:
                turno_curto = "M" if "MAN" in tn.upper() else "T" if "TAR" in tn.upper() else tn[:1]
                label = f"{s}º  T{t}  {turno_curto}"
                btn = ctk.CTkButton(sc_turmas, text=label,
                                     width=148, height=40, corner_radius=8,
                                     fg_color=C["fundo"], hover_color=C["borda"],
                                     text_color=C["texto"], font=FONTE_BOLD,
                                     border_width=1, border_color=C["borda"],
                                     anchor="w")
                btn.configure(command=lambda ss=s, tt=t, ttn=tn, b=btn: mostrar_alunos_turma(ss, tt, ttn, b))
                btn.pack(fill="x", pady=2)
                btn_turmas_lista.append(btn)

        btn_nome.configure(command=aba_nome)
        btn_turma.configure(command=aba_turma)
        aba_nome()

    def pop_livro(self, aluno, serie, turma, turno):
        pop = ctk.CTkToplevel(self)
        pop.geometry("480x380"); pop.resizable(False, False)
        pop.attributes("-topmost", True); pop.title("Registrar Empréstimo")
        pop.configure(fg_color=C["card"])
        ctk.CTkLabel(pop, text="📋  Novo Empréstimo",
                     font=FONTE_TITULO, text_color=C["vermelho"]).pack(pady=(24, 4))
        ctk.CTkLabel(pop, text=f"Aluno: {aluno}",
                     font=FONTE_BOLD, text_color=C["texto"]).pack()
        ctk.CTkFrame(pop, height=1, fg_color=C["borda"]).pack(fill="x", padx=30, pady=16)

        # Código do livro + título
        f_cod = ctk.CTkFrame(pop, fg_color="transparent"); f_cod.pack()
        lbl_sub(f_cod, "Código:")
        cod = ctk.CTkEntry(f_cod, placeholder_text="0000", **estilo_entry(80))
        cod.pack(side="left", padx=(0, 16))

        l = ctk.CTkEntry(pop, placeholder_text="Título do Livro", **estilo_entry(width=380))
        d = ctk.CTkEntry(pop, placeholder_text="Data de Entrega  (ex: 15/08/2025)",
                         **estilo_entry(width=380))
        l.pack(pady=6); d.pack(pady=6)

        def salvar():
            if l.get().strip():
                titulo_completo = f"[{cod.get().strip()}] {l.get().upper().strip()}" if cod.get().strip() else l.get().upper().strip()
                conn = sqlite3.connect(DB_PATH); c = conn.cursor()
                c.execute("INSERT INTO emprestimos (aluno_nome, livro_titulo, data_entrega, serie, turma, turno) VALUES (?,?,?,?,?,?)",
                          (aluno, titulo_completo, d.get().strip(), serie, turma, turno))
                conn.commit(); conn.close()
                pop.destroy()
                messagebox.showinfo("Sucesso", f"Empréstimo registrado para {aluno}!")

        ctk.CTkButton(pop, text="✓  Confirmar Empréstimo",
                      **estilo_btn_verde(width=240, height=42),
                      command=salvar).pack(pady=20)

    # ─────────────────────────────────────────
    #  GERENCIAR ALUNOS
    # ─────────────────────────────────────────
    def tela_alunos(self):
        self.limpar_container()
        cabecalho_tela(self.container, "👤  Gerenciar Alunos", self.mostrar_tela_principal)
        centro = ctk.CTkFrame(self.container, fg_color=C["fundo"], corner_radius=0)
        centro.pack(fill="both", expand=True)

        card_add = ctk.CTkFrame(centro, fg_color=C["card"], corner_radius=14,
                                border_width=1, border_color=C["borda"])
        card_add.pack(fill="x", padx=30, pady=(20, 0))
        ctk.CTkLabel(card_add, text="Cadastrar aluno", font=FONTE_SECAO,
                     text_color=C["vermelho"]).pack(anchor="w", padx=20, pady=(14, 8))

        linha = ctk.CTkFrame(card_add, fg_color="transparent")
        linha.pack(fill="x", padx=20, pady=(0, 16))
        en = ctk.CTkEntry(linha, placeholder_text="Nome completo", **estilo_entry(220))
        en.pack(side="left", padx=(0, 8))
        lbl_sub(linha, "Série:")
        es = ctk.CTkEntry(linha, placeholder_text="Ex: 6", **estilo_entry(70))
        es.pack(side="left", padx=(0, 8))
        lbl_sub(linha, "Turma:")
        seg_turma = ctk.CTkSegmentedButton(linha, values=TURMAS_OPCOES, **estilo_segmented(160))
        seg_turma.set("1"); seg_turma.pack(side="left", padx=(0, 8))

        def atualizar_turmas(e=None):
            s = es.get().strip().upper()
            opcoes = turmas_da_serie(s) if s else TURMAS_OPCOES
            seg_turma.configure(values=opcoes); seg_turma.set(opcoes[0])
        es.bind("<FocusOut>", atualizar_turmas); es.bind("<KeyRelease>", atualizar_turmas)

        lbl_sub(linha, "Turno:")
        etu = ctk.CTkEntry(linha, placeholder_text="Manhã / Tarde", **estilo_entry(130))
        etu.pack(side="left", padx=(0, 12))

        def salvar():
            if en.get().strip():
                conn = sqlite3.connect(DB_PATH); c = conn.cursor()
                nome_up = en.get().upper().strip()
                serie   = es.get().upper().strip()
                turma   = seg_turma.get().strip()
                turno   = etu.get().upper().strip()
                c.execute("INSERT INTO alunos (nome, serie, turma, turno) VALUES (?,?,?,?)",
                          (nome_up, serie, turma, turno))
                # sincroniza tabela turmas
                if serie and turma and turno:
                    c.execute("INSERT OR IGNORE INTO turmas (serie, turma, turno) VALUES (?,?,?)",
                              (serie, turma, turno))
                conn.commit(); conn.close()
                en.delete(0, "end"); listar()

        ctk.CTkButton(linha, text="+ Cadastrar", **estilo_btn_verde(140), command=salvar).pack(side="left")

        barra = ctk.CTkFrame(centro, fg_color="transparent")
        barra.pack(fill="x", padx=30, pady=(10, 0))
        ctk.CTkButton(barra, text="📥 Exportar Excel", **estilo_btn_cinza(160),
                      command=lambda: self.exportar_dados("alunos")).pack(side="left")
        busc = ctk.CTkEntry(barra, placeholder_text="🔍 Pesquisar aluno...", **estilo_entry(340))
        busc.pack(side="right")

        lbl_count = ctk.CTkLabel(centro, text="", font=FONTE_PEQUENA, text_color=C["subtexto"])
        lbl_count.pack(anchor="w", padx=34, pady=(6, 2))
        sc = ctk.CTkScrollableFrame(centro, fg_color="transparent", corner_radius=0)
        sc.pack(fill="both", expand=True, padx=30, pady=(4, 20))

        def listar(e=None):
            for w in sc.winfo_children(): w.destroy()
            q = busc.get().upper().strip()
            conn = sqlite3.connect(DB_PATH); c = conn.cursor()
            if not q:
                c.execute("SELECT nome, serie, turma, turno FROM alunos ORDER BY nome LIMIT 30")
            else:
                c.execute("SELECT nome, serie, turma, turno FROM alunos WHERE nome LIKE ? ORDER BY nome LIMIT 80",
                          (f"%{q}%",))
            rows = c.fetchall(); conn.close()
            lbl_count.configure(text=f"{len(rows)} aluno(s) encontrado(s)")
            for n, s, tr, tn in rows:
                card = ctk.CTkFrame(sc, fg_color=C["card"], corner_radius=10,
                                    border_width=1, border_color=C["borda"])
                card.pack(fill="x", pady=3)
                ctk.CTkLabel(card, text="👤", font=("Segoe UI", 16), width=40).pack(side="left", padx=10)
                info = ctk.CTkFrame(card, fg_color="transparent")
                info.pack(side="left", fill="x", expand=True, pady=8)
                ctk.CTkLabel(info, text=n, font=FONTE_BOLD, text_color=C["texto"], anchor="w").pack(anchor="w", padx=4)
                ctk.CTkLabel(info, text=f"Série: {s or '—'}  ·  Turma: {tr or '—'}  ·  Turno: {tn or '—'}",
                             font=FONTE_PEQUENA, text_color=C["subtexto"], anchor="w").pack(anchor="w", padx=4)
                ctk.CTkButton(card, text="✕", width=32, height=32,
                              fg_color="#FFEEEE", hover_color="#FFCCCC",
                              text_color=C["vermelho"], corner_radius=8, font=FONTE_BOLD,
                              command=lambda x=n: deletar(x)).pack(side="right", padx=12)

        def deletar(n):
            if messagebox.askyesno("Confirmar", f"Excluir o aluno '{n}'?"):
                conn = sqlite3.connect(DB_PATH); c = conn.cursor()
                c.execute("DELETE FROM alunos WHERE nome=?", (n,))
                conn.commit(); conn.close(); listar()

        busc.bind("<KeyRelease>", listar); listar()

    # ─────────────────────────────────────────
    #  ACERVO  — com estante (A-H), fileira (1-5) e código (4 dígitos)
    # ─────────────────────────────────────────
    def tela_acervo(self):
        self.limpar_container()
        cabecalho_tela(self.container, "📚  Acervo de Livros", self.mostrar_tela_principal)
        centro = ctk.CTkFrame(self.container, fg_color=C["fundo"], corner_radius=0)
        centro.pack(fill="both", expand=True)

        # ── Card adicionar ───────────────────────────────────────────
        card_add = ctk.CTkFrame(centro, fg_color=C["card"], corner_radius=14,
                                border_width=1, border_color=C["borda"])
        card_add.pack(fill="x", padx=30, pady=(20, 0))
        ctk.CTkLabel(card_add, text="Adicionar livro", font=FONTE_SECAO,
                     text_color=C["vermelho"]).pack(anchor="w", padx=20, pady=(14, 6))

        l1 = ctk.CTkFrame(card_add, fg_color="transparent")
        l1.pack(fill="x", padx=20, pady=(0, 4))
        et  = ctk.CTkEntry(l1, placeholder_text="Título",  **estilo_entry(260))
        ea  = ctk.CTkEntry(l1, placeholder_text="Autor",   **estilo_entry(200))
        eg  = ctk.CTkEntry(l1, placeholder_text="Gênero",  **estilo_entry(140))
        for w in [et, ea, eg]: w.pack(side="left", padx=(0, 8))

        l2 = ctk.CTkFrame(card_add, fg_color="transparent")
        l2.pack(fill="x", padx=20, pady=(0, 14))

        lbl_sub(l2, "Código (4 dígitos):")
        ecod = ctk.CTkEntry(l2, placeholder_text="0000", **estilo_entry(90))
        ecod.pack(side="left", padx=(0, 16))

        lbl_sub(l2, "Estante:")
        seg_est = ctk.CTkSegmentedButton(l2, values=ESTANTES_LETRAS, **estilo_segmented(400))
        seg_est.set("A"); seg_est.pack(side="left", padx=(0, 16))

        lbl_sub(l2, "Fileira:")
        seg_fil = ctk.CTkSegmentedButton(l2, values=FILEIRAS_NUM, **estilo_segmented(200))
        seg_fil.set("1"); seg_fil.pack(side="left", padx=(0, 16))

        ctk.CTkButton(l2, text="+ Adicionar", **estilo_btn_verde(130),
                      command=lambda: add()).pack(side="left")

        # ── Card pesquisar ───────────────────────────────────────────
        card_busc = ctk.CTkFrame(centro, fg_color=C["card"], corner_radius=14,
                                  border_width=1, border_color=C["borda"])
        card_busc.pack(fill="x", padx=30, pady=(10, 0))
        ctk.CTkLabel(card_busc, text="Pesquisar livros", font=FONTE_SECAO,
                     text_color=C["vermelho"]).pack(anchor="w", padx=20, pady=(14, 6))

        lb = ctk.CTkFrame(card_busc, fg_color="transparent")
        lb.pack(fill="x", padx=20, pady=(0, 14))
        bt  = ctk.CTkEntry(lb, placeholder_text="🔍 Título",         **estilo_entry(220))
        ba  = ctk.CTkEntry(lb, placeholder_text="👤 Autor",          **estilo_entry(160))
        bg  = ctk.CTkEntry(lb, placeholder_text="🏷 Gênero",         **estilo_entry(120))
        bcd = ctk.CTkEntry(lb, placeholder_text="# Código",          **estilo_entry(90))
        be  = ctk.CTkEntry(lb, placeholder_text="📦 Estante (A-H)",  **estilo_entry(120))
        bf  = ctk.CTkEntry(lb, placeholder_text="📍 Fileira (1-5)",  **estilo_entry(110))
        for w in [bt, ba, bg, bcd, be, bf]: w.pack(side="left", padx=(0, 8))
        ctk.CTkButton(lb, text="Limpar", **estilo_btn_cinza(90),
                      command=lambda: limpar_busca()).pack(side="left")

        lbl_count = ctk.CTkLabel(centro, text="", font=FONTE_PEQUENA, text_color=C["subtexto"])
        lbl_count.pack(anchor="w", padx=34, pady=(8, 2))
        sc = ctk.CTkScrollableFrame(centro, fg_color="transparent", corner_radius=0)
        sc.pack(fill="both", expand=True, padx=30, pady=(4, 20))

        def listar(e=None):
            for w in sc.winfo_children(): w.destroy()
            qt  = bt.get().upper().strip()
            qa  = ba.get().upper().strip()
            qg  = bg.get().upper().strip()
            qcd = bcd.get().strip()
            qe  = be.get().upper().strip()
            qf  = bf.get().strip()

            sql = "SELECT titulo, autor, genero, estante, fileira, codigo FROM acervo WHERE 1=1"
            params = []
            if qt:  sql += " AND titulo  LIKE ?"; params.append(f"%{qt}%")
            if qa:  sql += " AND autor   LIKE ?"; params.append(f"%{qa}%")
            if qg:  sql += " AND genero  LIKE ?"; params.append(f"%{qg}%")
            if qcd: sql += " AND codigo  LIKE ?"; params.append(f"%{qcd}%")
            if qe:  sql += " AND estante LIKE ?"; params.append(f"%{qe}%")
            if qf:  sql += " AND fileira LIKE ?"; params.append(f"%{qf}%")

            sem_filtro = not any([qt, qa, qg, qcd, qe, qf])
            sql += " ORDER BY titulo" + (" LIMIT 20" if sem_filtro else " LIMIT 100")

            conn = sqlite3.connect(DB_PATH); c = conn.cursor()
            c.execute(sql, params); rows = c.fetchall(); conn.close()

            sufixo = " — primeiros 20" if sem_filtro and len(rows) == 20 else ""
            lbl_count.configure(text=f"{len(rows)} livro(s) encontrado(s){sufixo}")

            for t, a, g, est, fil, cod in rows:
                card = ctk.CTkFrame(sc, fg_color=C["card"], corner_radius=10,
                                    border_width=1, border_color=C["borda"])
                card.pack(fill="x", pady=3)

                # Badge código
                cod_txt = cod if cod and cod.strip() else "----"
                ctk.CTkLabel(card, text=cod_txt, font=("Courier", 13, "bold"),
                             text_color=C["vermelho"], width=52
                             ).pack(side="left", padx=(10, 0))

                info = ctk.CTkFrame(card, fg_color="transparent")
                info.pack(side="left", fill="x", expand=True, pady=8)
                ctk.CTkLabel(info, text=t if len(t) <= 60 else t[:57]+"...",
                             font=FONTE_BOLD, text_color=C["texto"], anchor="w").pack(anchor="w", padx=4)
                partes = []
                if a and a.strip():   partes.append(f"Autor: {a}")
                if g and g.strip():   partes.append(f"Gênero: {g}")
                if est and est.strip(): partes.append(f"Estante: {est}")
                if fil and fil.strip(): partes.append(f"Fileira: {fil}")
                if partes:
                    ctk.CTkLabel(info, text="  ·  ".join(partes), font=FONTE_PEQUENA,
                                 text_color=C["subtexto"], anchor="w").pack(anchor="w", padx=4)

                ctk.CTkButton(card, text="✕", width=32, height=32,
                              fg_color="#FFEEEE", hover_color="#FFCCCC",
                              text_color=C["vermelho"], corner_radius=8, font=FONTE_BOLD,
                              command=lambda x=t: deletar_l(x)).pack(side="right", padx=12)

        def add():
            if et.get().strip():
                cod_val = ecod.get().strip()
                if cod_val and (not cod_val.isdigit() or len(cod_val) != 4):
                    messagebox.showwarning("Atenção", "O código deve ter exatamente 4 dígitos numéricos.")
                    return
                conn = sqlite3.connect(DB_PATH); c = conn.cursor()
                c.execute("INSERT INTO acervo (titulo, autor, genero, estante, fileira, codigo) VALUES (?,?,?,?,?,?)",
                          (et.get().upper().strip(), ea.get().upper().strip(),
                           eg.get().upper().strip(), seg_est.get().strip(),
                           seg_fil.get().strip(), cod_val))
                conn.commit(); conn.close()
                for w in [et, ea, eg, ecod]: w.delete(0, "end")
                listar()

        def deletar_l(t):
            if messagebox.askyesno("Confirmar", f"Excluir '{t}' do acervo?"):
                conn = sqlite3.connect(DB_PATH); c = conn.cursor()
                c.execute("DELETE FROM acervo WHERE titulo=?", (t,))
                conn.commit(); conn.close(); listar()

        def limpar_busca():
            for w in [bt, ba, bg, bcd, be, bf]: w.delete(0, "end")
            listar()

        for w in [bt, ba, bg, bcd, be, bf]: w.bind("<KeyRelease>", listar)
        listar()

    # ─────────────────────────────────────────
    #  PENDÊNCIAS
    # ─────────────────────────────────────────
    def tela_pendentes(self):
        self.limpar_container()
        cabecalho_tela(self.container, "⏳  Pendências de Devolução", self.mostrar_tela_principal)
        centro = ctk.CTkFrame(self.container, fg_color=C["fundo"], corner_radius=0)
        centro.pack(fill="both", expand=True)
        lbl_count = ctk.CTkLabel(centro, text="", font=FONTE_PEQUENA, text_color=C["subtexto"])
        lbl_count.pack(anchor="w", padx=34, pady=(12, 4))
        sc = ctk.CTkScrollableFrame(centro, fg_color="transparent", corner_radius=0)
        sc.pack(fill="both", expand=True, padx=30, pady=(4, 20))

        def carregar():
            for w in sc.winfo_children(): w.destroy()
            conn = sqlite3.connect(DB_PATH); c = conn.cursor()
            c.execute("SELECT rowid, aluno_nome, livro_titulo, data_entrega, serie, turma, turno FROM emprestimos ORDER BY aluno_nome")
            rows = c.fetchall(); conn.close()
            lbl_count.configure(text=f"{len(rows)} empréstimo(s) pendente(s)")
            if not rows:
                ctk.CTkLabel(sc, text="✅  Nenhuma pendência no momento.",
                             font=FONTE_CORPO, text_color=C["subtexto"]).pack(pady=30)
                return
            for rid, n, l, dt, s, tr, tn in rows:
                card = ctk.CTkFrame(sc, fg_color=C["card"], corner_radius=10,
                                    border_width=1, border_color=C["borda"])
                card.pack(fill="x", pady=3)
                ctk.CTkLabel(card, text="📖", font=("Segoe UI", 18), width=42).pack(side="left", padx=10)
                info = ctk.CTkFrame(card, fg_color="transparent")
                info.pack(side="left", fill="x", expand=True, pady=8)
                ctk.CTkLabel(info, text=l if len(l) <= 55 else l[:52]+"...",
                             font=FONTE_BOLD, text_color=C["texto"], anchor="w").pack(anchor="w", padx=4)
                meta = f"Aluno: {n}  ·  Série: {s or '—'}  ·  Turma: {tr or '—'}"
                if dt: meta += f"  ·  Entrega: {dt}"
                ctk.CTkLabel(info, text=meta, font=FONTE_PEQUENA,
                             text_color=C["subtexto"], anchor="w").pack(anchor="w", padx=4)
                ctk.CTkButton(card, text="✓ Devolver", **estilo_btn_verde(120),
                              command=lambda x=rid, y=n, sa=s, ta=tr, tu=tn: self.finalizar(x, y, sa, ta, tu, carregar)
                              ).pack(side="right", padx=12, pady=10)

        carregar()

    def finalizar(self, rid, nome, serie, turma, turno, cb):
        conn = sqlite3.connect(DB_PATH); c = conn.cursor()
        c.execute("INSERT INTO ranking (nome, lidos, serie, turma, turno) VALUES (?,1,?,?,?) ON CONFLICT(nome) DO UPDATE SET lidos=lidos+1",
                  (nome, serie, turma, turno))
        c.execute("DELETE FROM emprestimos WHERE rowid=?", (rid,))
        conn.commit(); conn.close(); cb()

    # ─────────────────────────────────────────
    #  RANKING
    # ─────────────────────────────────────────
    def tela_ranking_unificado(self):
        self.limpar_container()
        cabecalho_tela(self.container, "🏆  Ranking de Leitura", self.mostrar_tela_principal)
        centro = ctk.CTkFrame(self.container, fg_color=C["fundo"], corner_radius=0)
        centro.pack(fill="both", expand=True)

        barra = ctk.CTkFrame(centro, fg_color=C["card"], corner_radius=0, height=56)
        barra.pack(fill="x"); barra.pack_propagate(False)
        ctk.CTkButton(barra, text="📥 Exportar Excel", **estilo_btn_verde(160),
                      command=self.exportar_ranking).pack(side="left", padx=20, pady=9)
        ctk.CTkButton(barra, text="🗑 Zerar Ranking",
                      width=150, height=38, corner_radius=10,
                      fg_color=C["perigo"], hover_color=C["perigo_hover"],
                      text_color="white", font=FONTE_BOLD,
                      command=self.zerar_ranking).pack(side="left", pady=9)

        abas = ctk.CTkFrame(centro, fg_color="transparent")
        abas.pack(fill="x", padx=30, pady=(16, 0))
        btn_alunos = ctk.CTkButton(abas, text="🏆 Alunos", width=160, height=38,
                                   corner_radius=10, fg_color=C["vermelho"],
                                   hover_color=C["vermelho_esc"], text_color="white", font=FONTE_BOLD)
        btn_alunos.pack(side="left", padx=(0, 8))
        btn_turmas = ctk.CTkButton(abas, text="🏫 Turmas", width=160, height=38,
                                   corner_radius=10, fg_color=C["inativo"],
                                   hover_color="#D1D5DB", text_color=C["inativo_txt"], font=FONTE_CORPO)
        btn_turmas.pack(side="left")

        sc = ctk.CTkScrollableFrame(centro, fg_color="transparent", corner_radius=0)
        sc.pack(fill="both", expand=True, padx=30, pady=12)
        medalhas = {1:"🥇", 2:"🥈", 3:"🥉"}

        def renderizar_alunos():
            for w in sc.winfo_children(): w.destroy()
            btn_alunos.configure(fg_color=C["vermelho"], text_color="white", font=FONTE_BOLD)
            btn_turmas.configure(fg_color=C["inativo"], text_color=C["inativo_txt"], font=FONTE_CORPO)
            conn = sqlite3.connect(DB_PATH); c = conn.cursor()
            c.execute("SELECT nome, lidos, serie, turma, turno FROM ranking WHERE lidos > 0 ORDER BY lidos DESC")
            rows = c.fetchall(); conn.close()
            if not rows:
                ctk.CTkLabel(sc, text="Nenhum aluno no ranking ainda.",
                             font=FONTE_CORPO, text_color=C["subtexto"]).pack(pady=40)
                return
            for i, (n, l, s, tr, tn) in enumerate(rows, 1):
                card = ctk.CTkFrame(sc, fg_color=C["card"], corner_radius=10,
                                    border_width=1, border_color=C["borda"])
                card.pack(fill="x", pady=3)
                ctk.CTkLabel(card, text=medalhas.get(i, f"{i}º"),
                             font=("Segoe UI", 20), width=52).pack(side="left", padx=8)
                info = ctk.CTkFrame(card, fg_color="transparent")
                info.pack(side="left", fill="x", expand=True, pady=10)
                ctk.CTkLabel(info, text=n, font=FONTE_BOLD, text_color=C["texto"],
                             anchor="w").pack(anchor="w", padx=4)
                ctk.CTkLabel(info, text=f"Série: {s or '—'}  ·  Turma: {tr or '—'}  ·  Turno: {tn or '—'}",
                             font=FONTE_PEQUENA, text_color=C["subtexto"], anchor="w").pack(anchor="w", padx=4)
                ctk.CTkLabel(card, text=f"{l} livro{'s' if l != 1 else ''}",
                             font=FONTE_BOLD, text_color=C["verde"], width=90).pack(side="right", padx=16)

        def renderizar_turmas():
            for w in sc.winfo_children(): w.destroy()
            btn_turmas.configure(fg_color=C["vermelho"], text_color="white", font=FONTE_BOLD)
            btn_alunos.configure(fg_color=C["inativo"], text_color=C["inativo_txt"], font=FONTE_CORPO)
            conn = sqlite3.connect(DB_PATH); c = conn.cursor()
            c.execute("SELECT DISTINCT serie, turma, turno FROM alunos ORDER BY serie, turma, turno")
            turmas_alunos = c.fetchall()
            c.execute("SELECT serie, turma, turno, COALESCE(SUM(lidos),0) FROM ranking GROUP BY serie, turma, turno")
            dict_rank = {(r[0], r[1], r[2]): r[3] for r in c.fetchall()}
            conn.close()
            if not turmas_alunos:
                ctk.CTkLabel(sc, text="Nenhuma turma cadastrada ainda.",
                             font=FONTE_CORPO, text_color=C["subtexto"]).pack(pady=40)
                return
            dados = sorted([(s, tr, tn, dict_rank.get((s, tr, tn), 0))
                            for s, tr, tn in turmas_alunos], key=lambda x: -x[3])
            for i, (s, tr, tn, t) in enumerate(dados, 1):
                card = ctk.CTkFrame(sc, fg_color=C["card"], corner_radius=10,
                                    border_width=1, border_color=C["borda"])
                card.pack(fill="x", pady=3)
                pos = medalhas.get(i, f"{i}º") if t > 0 else "—"
                ctk.CTkLabel(card, text=pos, font=("Segoe UI", 20), width=52).pack(side="left", padx=8)
                info = ctk.CTkFrame(card, fg_color="transparent")
                info.pack(side="left", fill="x", expand=True, pady=10)
                ctk.CTkLabel(info, text=f"Série {s or '—'}  —  Turma {tr or '—'}",
                             font=FONTE_BOLD, text_color=C["texto"], anchor="w").pack(anchor="w", padx=4)
                if tn:
                    ctk.CTkLabel(info, text=f"Turno: {tn}", font=FONTE_PEQUENA,
                                 text_color=C["subtexto"], anchor="w").pack(anchor="w", padx=4)
                cor = C["verde"] if t > 0 else C["subtexto"]
                ctk.CTkLabel(card, text=f"{t} livro{'s' if t != 1 else ''}",
                             font=FONTE_BOLD, text_color=cor, width=90).pack(side="right", padx=16)

        btn_alunos.configure(command=renderizar_alunos)
        btn_turmas.configure(command=renderizar_turmas)
        renderizar_alunos()

    def exportar_ranking(self):
        wb  = Workbook(); ws1 = wb.active; ws1.title = "Ranking Alunos"
        ws1.append(["Posição","Aluno","Livros Lidos","Série","Turma","Turno"])
        conn = sqlite3.connect(DB_PATH); c = conn.cursor()
        c.execute("SELECT nome, lidos, serie, turma, turno FROM ranking WHERE lidos > 0 ORDER BY lidos DESC")
        for i, row in enumerate(c.fetchall(), 1):
            ws1.append([i, row[0], row[1], row[2] or "", row[3] or "", row[4] or ""])
        ws2 = wb.create_sheet("Ranking Turmas")
        ws2.append(["Posição","Série","Turma","Turno","Total Lidos"])
        c.execute("SELECT DISTINCT serie, turma, turno FROM alunos ORDER BY serie, turma, turno")
        turmas = c.fetchall()
        c.execute("SELECT serie, turma, turno, COALESCE(SUM(lidos),0) FROM ranking GROUP BY serie, turma, turno")
        dict_rank = {(r[0],r[1],r[2]):r[3] for r in c.fetchall()}
        conn.close()
        dados = sorted([(s, tr, tn, dict_rank.get((s,tr,tn),0)) for s,tr,tn in turmas], key=lambda x:-x[3])
        for i,(s,tr,tn,t) in enumerate(dados,1):
            ws2.append([i, s or "", tr or "", tn or "", t])
        caminho = EXPORT_DIR / f"Ranking_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        wb.save(caminho)
        messagebox.showinfo("Exportado", f"Ranking salvo em:\n{caminho}")

    def zerar_ranking(self):
        if messagebox.askyesno("Atenção", "Zerar a contagem de livros lidos de todos os alunos?\nEssa ação não pode ser desfeita."):
            conn = sqlite3.connect(DB_PATH); c = conn.cursor()
            c.execute("UPDATE ranking SET lidos = 0")
            conn.commit(); conn.close()
            messagebox.showinfo("Concluído", "Ranking zerado com sucesso!")
            self.tela_ranking_unificado()

    # ─────────────────────────────────────────
    #  RELATÓRIOS
    # ─────────────────────────────────────────
    def tela_relatorios(self):
        self.limpar_container()
        cabecalho_tela(self.container, "📊  Relatórios", self.mostrar_tela_principal)
        centro = ctk.CTkFrame(self.container, fg_color=C["fundo"], corner_radius=0)
        centro.pack(fill="both", expand=True)

        abas = ctk.CTkFrame(centro, fg_color="transparent")
        abas.pack(fill="x", padx=30, pady=(20, 0))

        btn_controle = ctk.CTkButton(abas, text="📋 Controle de Leitura", width=200, height=38,
                                      corner_radius=10, fg_color=C["vermelho"],
                                      hover_color=C["vermelho_esc"], text_color="white", font=FONTE_BOLD)
        btn_controle.pack(side="left", padx=(0, 8))

        btn_ranking = ctk.CTkButton(abas, text="🏆 Ranking Mensal", width=190, height=38,
                                     corner_radius=10, fg_color=C["inativo"],
                                     hover_color="#D1D5DB", text_color=C["inativo_txt"], font=FONTE_CORPO)
        btn_ranking.pack(side="left", padx=(0, 8))

        btn_geral = ctk.CTkButton(abas, text="🏫 Ranking Completo", width=190, height=38,
                                   corner_radius=10, fg_color=C["inativo"],
                                   hover_color="#D1D5DB", text_color=C["inativo_txt"], font=FONTE_CORPO)
        btn_geral.pack(side="left")

        area = ctk.CTkFrame(centro, fg_color="transparent")
        area.pack(fill="both", expand=True, padx=30, pady=16)

        def ativar(btn_ativo):
            for b in [btn_controle, btn_ranking, btn_geral]:
                if b == btn_ativo:
                    b.configure(fg_color=C["vermelho"], text_color="white", font=FONTE_BOLD)
                else:
                    b.configure(fg_color=C["inativo"], text_color=C["inativo_txt"], font=FONTE_CORPO)

        # ── Aba 1: Controle de leitura por turma ────────────────────
        def aba_controle():
            for w in area.winfo_children(): w.destroy()
            ativar(btn_controle)

            card = ctk.CTkFrame(area, fg_color=C["card"], corner_radius=14,
                                border_width=1, border_color=C["borda"])
            card.pack(fill="x")
            ctk.CTkLabel(card, text="Gerar planilha de controle de leitura por turma",
                         font=FONTE_SECAO, text_color=C["vermelho"]
                         ).pack(anchor="w", padx=24, pady=(20, 4))
            ctk.CTkLabel(card, text="Lista todos os alunos da turma com colunas de controle semanal e total de livros lidos.",
                         font=FONTE_PEQUENA, text_color=C["subtexto"], wraplength=650
                         ).pack(anchor="w", padx=24, pady=(0, 12))
            separador(card, pady=(0, 12))

            f = ctk.CTkFrame(card, fg_color="transparent"); f.pack(fill="x", padx=24, pady=(8, 0))
            ctk.CTkLabel(f, text="Série:", font=FONTE_CORPO, text_color=C["subtexto"]).pack(side="left")
            e_serie = ctk.CTkEntry(f, placeholder_text="Ex: 6", **estilo_entry(80))
            e_serie.pack(side="left", padx=(4, 16))

            ctk.CTkLabel(f, text="Turma:", font=FONTE_CORPO, text_color=C["subtexto"]).pack(side="left")
            seg_rel = ctk.CTkSegmentedButton(f, values=TURMAS_OPCOES, **estilo_segmented(160))
            seg_rel.set("1"); seg_rel.pack(side="left", padx=(4, 16))

            def atualizar_seg(e=None):
                s = e_serie.get().strip().upper()
                ops = turmas_da_serie(s) if s else TURMAS_OPCOES
                seg_rel.configure(values=ops); seg_rel.set(ops[0])
            e_serie.bind("<FocusOut>", atualizar_seg); e_serie.bind("<KeyRelease>", atualizar_seg)

            ctk.CTkLabel(f, text="Turno:", font=FONTE_CORPO, text_color=C["subtexto"]).pack(side="left")
            e_turno = ctk.CTkEntry(f, placeholder_text="Manhã / Tarde (opcional)", **estilo_entry(180))
            e_turno.pack(side="left", padx=(4, 16))

            ctk.CTkLabel(f, text="Professor(a):", font=FONTE_CORPO, text_color=C["subtexto"]).pack(side="left")
            e_prof = ctk.CTkEntry(f, placeholder_text="Nome do professor", **estilo_entry(180))
            e_prof.pack(side="left", padx=(4, 0))

            lbl_s = ctk.CTkLabel(card, text="", font=FONTE_PEQUENA, text_color=C["verde"], wraplength=650)
            lbl_s.pack(anchor="w", padx=24, pady=(10, 0))
            ctk.CTkButton(card, text="📥 Gerar Planilha Excel",
                          **estilo_btn_verde(220, height=42),
                          command=lambda: self._gerar_controle(e_serie, seg_rel, e_turno, e_prof, lbl_s)
                          ).pack(anchor="w", padx=24, pady=(12, 24))

        # ── Aba 2: Ranking mensal (formato da imagem) ────────────────
        def aba_ranking_mensal():
            for w in area.winfo_children(): w.destroy()
            ativar(btn_ranking)

            card = ctk.CTkFrame(area, fg_color=C["card"], corner_radius=14,
                                border_width=1, border_color=C["borda"])
            card.pack(fill="x")
            ctk.CTkLabel(card, text="Gerar planilha de ranking mensal por turma",
                         font=FONTE_SECAO, text_color=C["vermelho"]
                         ).pack(anchor="w", padx=24, pady=(20, 4))
            ctk.CTkLabel(card,
                         text="Gera a planilha no formato usado pela escola: aluno na vertical, semanas na horizontal, X marcado e total de livros lidos.",
                         font=FONTE_PEQUENA, text_color=C["subtexto"], wraplength=650
                         ).pack(anchor="w", padx=24, pady=(0, 12))
            separador(card, pady=(0, 12))

            f = ctk.CTkFrame(card, fg_color="transparent"); f.pack(fill="x", padx=24, pady=(8, 0))

            ctk.CTkLabel(f, text="Série:", font=FONTE_CORPO, text_color=C["subtexto"]).pack(side="left")
            e_serie2 = ctk.CTkEntry(f, placeholder_text="Ex: 5", **estilo_entry(80))
            e_serie2.pack(side="left", padx=(4, 16))

            ctk.CTkLabel(f, text="Turma:", font=FONTE_CORPO, text_color=C["subtexto"]).pack(side="left")
            seg_rel2 = ctk.CTkSegmentedButton(f, values=TURMAS_OPCOES, **estilo_segmented(160))
            seg_rel2.set("1"); seg_rel2.pack(side="left", padx=(4, 16))

            def atualizar_seg2(e=None):
                s = e_serie2.get().strip().upper()
                ops = turmas_da_serie(s) if s else TURMAS_OPCOES
                seg_rel2.configure(values=ops); seg_rel2.set(ops[0])
            e_serie2.bind("<FocusOut>", atualizar_seg2); e_serie2.bind("<KeyRelease>", atualizar_seg2)

            ctk.CTkLabel(f, text="Mês:", font=FONTE_CORPO, text_color=C["subtexto"]).pack(side="left")
            meses = list(MESES_PT.values())
            mes_atual = MESES_PT[datetime.now().month]
            seg_mes = ctk.CTkSegmentedButton(f, values=meses,
                                              selected_color=C["verde"],
                                              selected_hover_color=C["verde_esc"],
                                              unselected_color=C["inativo"],
                                              unselected_hover_color="#D1D5DB",
                                              text_color="white", font=("Segoe UI", 11, "bold"),
                                              corner_radius=8, width=600, height=34)
            seg_mes.set(mes_atual); seg_mes.pack(side="left", padx=(4, 16))

            ctk.CTkLabel(f, text="Professor(a):", font=FONTE_CORPO, text_color=C["subtexto"]).pack(side="left")
            e_prof2 = ctk.CTkEntry(f, placeholder_text="Nome", **estilo_entry(160))
            e_prof2.pack(side="left", padx=(4, 0))

            lbl_s2 = ctk.CTkLabel(card, text="", font=FONTE_PEQUENA, text_color=C["verde"], wraplength=650)
            lbl_s2.pack(anchor="w", padx=24, pady=(10, 0))
            ctk.CTkButton(card, text="📥 Gerar Ranking Mensal",
                          **estilo_btn_verde(230, height=42),
                          command=lambda: self._gerar_ranking_mensal(e_serie2, seg_rel2, seg_mes, e_prof2, lbl_s2)
                          ).pack(anchor="w", padx=24, pady=(12, 24))

        # ── Aba 3: Ranking completo ──────────────────────────────────
        def aba_geral():
            for w in area.winfo_children(): w.destroy()
            ativar(btn_geral)
            card = ctk.CTkFrame(area, fg_color=C["card"], corner_radius=14,
                                border_width=1, border_color=C["borda"])
            card.pack(fill="x")
            ctk.CTkLabel(card, text="Exportar ranking completo de todas as turmas",
                         font=FONTE_SECAO, text_color=C["vermelho"]
                         ).pack(anchor="w", padx=24, pady=(20, 4))
            ctk.CTkLabel(card, text="Gera uma planilha com 'Ranking Alunos' e 'Ranking Turmas'. Todas as turmas aparecem, mesmo com 0 livros.",
                         font=FONTE_PEQUENA, text_color=C["subtexto"], wraplength=650
                         ).pack(anchor="w", padx=24, pady=(0, 12))
            separador(card, pady=(0, 12))
            lbl_s3 = ctk.CTkLabel(card, text="", font=FONTE_PEQUENA, text_color=C["verde"], wraplength=650)
            lbl_s3.pack(anchor="w", padx=24, pady=(8, 0))
            ctk.CTkButton(card, text="📥 Exportar Ranking Completo",
                          **estilo_btn_verde(240, height=42),
                          command=lambda: (self.exportar_ranking(),
                                           lbl_s3.configure(text="✅  Ranking exportado para a pasta Relatorios/"))
                          ).pack(anchor="w", padx=24, pady=(12, 24))

        btn_controle.configure(command=aba_controle)
        btn_ranking.configure(command=aba_ranking_mensal)
        btn_geral.configure(command=aba_geral)
        aba_controle()

    # ── Gerador: Controle de leitura ────────────────────────────────
    def _gerar_controle(self, e_serie, seg_turma, e_turno, e_prof, lbl_status):
        serie = e_serie.get().strip().upper()
        turma = seg_turma.get().strip()
        turno = e_turno.get().strip()
        prof  = e_prof.get().strip()
        if not serie:
            messagebox.showwarning("Atenção", "Série é obrigatória!"); return

        conn = sqlite3.connect(DB_PATH); c = conn.cursor()
        if turno:
            c.execute("SELECT nome FROM alunos WHERE serie=? AND turma=? AND turno=? ORDER BY nome",
                      (serie, turma, turno.upper()))
        else:
            c.execute("SELECT nome FROM alunos WHERE serie=? AND turma=? ORDER BY nome",
                      (serie, turma))
        alunos = [r[0] for r in c.fetchall()]
        if not alunos:
            messagebox.showinfo("Info", f"Nenhum aluno encontrado na {serie}ª série, turma {turma}.")
            conn.close(); return
        if turno:
            c.execute("SELECT nome, lidos FROM ranking WHERE serie=? AND turma=? AND turno=?",
                      (serie, turma, turno.upper()))
        else:
            c.execute("SELECT nome, lidos FROM ranking WHERE serie=? AND turma=?", (serie, turma))
        dict_lidos = {r[0]: r[1] for r in c.fetchall()}
        conn.close()

        turno_label = turno.upper() if turno else "TODOS OS TURNOS"
        wb = Workbook(); ws = wb.active; ws.title = f"{serie}-T{turma}"

        ws.merge_cells("A1:F1")
        ws["A1"] = f"CONTROLE DE LEITURA — {serie}ª SÉRIE  |  TURMA {turma}  |  {turno_label}   —   {MESES_PT[datetime.now().month]}/{datetime.now().year}"
        ws["A1"].font = Font(bold=True, size=13)
        ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
        ws["A1"].fill = PatternFill("solid", fgColor="FFCC00")
        ws.row_dimensions[1].height = 28
        ws.row_dimensions[2].height = 6

        headers = ["ALUNO", "1ª Semana", "2ª Semana", "3ª Semana", "4ª Semana", "TOTAL LIDOS"]
        cinza_fill = PatternFill("solid", fgColor="D9D9D9")
        for col, h in enumerate(headers, 1):
            cell = ws.cell(3, col, h)
            cell.font = Font(bold=True, size=11)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.fill = cinza_fill
        ws.row_dimensions[3].height = 20

        verde_fill = PatternFill("solid", fgColor="E8F5E9")
        for i, aluno in enumerate(alunos, 4):
            lidos = dict_lidos.get(aluno, 0)
            ws.cell(i, 1, aluno).font = Font(size=11)
            ct = ws.cell(i, 6, lidos)
            ct.font = Font(bold=True, size=11); ct.alignment = Alignment(horizontal="center")
            for col in range(2, 6):
                c2 = ws.cell(i, col, "X" if lidos > 0 else "")
                c2.alignment = Alignment(horizontal="center")
                if lidos > 0: c2.fill = verde_fill
            ws.row_dimensions[i].height = 18

        thin = thin_border()
        for row in ws.iter_rows(min_row=3, max_row=3+len(alunos), min_col=1, max_col=6):
            for cell in row: cell.border = thin

        ws.column_dimensions["A"].width = 36
        for col in ["B","C","D","E"]: ws.column_dimensions[col].width = 14
        ws.column_dimensions["F"].width = 16

        nome_arq = f"Controle_{serie}-{turma}_{turno}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        caminho  = EXPORT_DIR / nome_arq
        wb.save(caminho)
        lbl_status.configure(text=f"✅  Arquivo gerado com sucesso!\n{caminho}")

    # ── Gerador: Ranking mensal (exatamente o formato da foto) ───────
    def _gerar_ranking_mensal(self, e_serie, seg_turma, seg_mes, e_prof, lbl_status):
        serie = e_serie.get().strip().upper()
        turma = seg_turma.get().strip()
        mes_nome = seg_mes.get().strip()         # ex: "FEVEREIRO"
        prof  = e_prof.get().strip()

        if not serie:
            messagebox.showwarning("Atenção", "Série é obrigatória!"); return

        # Busca alunos
        conn = sqlite3.connect(DB_PATH); c = conn.cursor()
        c.execute("SELECT nome FROM alunos WHERE serie=? AND turma=? ORDER BY nome",
                  (serie, turma))
        alunos = [r[0] for r in c.fetchall()]
        if not alunos:
            messagebox.showinfo("Info", f"Nenhum aluno encontrado na {serie}ª série, turma {turma}.")
            conn.close(); return
        c.execute("SELECT nome, lidos FROM ranking WHERE serie=? AND turma=?", (serie, turma))
        dict_lidos = {r[0]: r[1] for r in c.fetchall()}
        conn.close()

        wb = Workbook(); ws = wb.active
        ws.title = f"Ranking {mes_nome[:3]}"

        # ── Layout: alunos na VERTICAL (linhas), semanas na HORIZONTAL
        # Cabeçalho superior — linha 1
        escola_txt = "CENTRO DE EXC. MUL. D. JOÃO J DA MOTA E ALBUQUERQUE"
        ws.merge_cells("A1:K1")
        c1 = ws["A1"]; c1.value = escola_txt
        c1.font = Font(bold=True, size=11)
        c1.alignment = Alignment(horizontal="center", vertical="center")
        c1.fill = PatternFill("solid", fgColor="FFFFFF")
        ws.row_dimensions[1].height = 20

        # Linha 2: Mês | Turma | Professor
        ws["A2"] = f"Mês: {mes_nome}"
        ws.merge_cells("A2:B2")
        ws["A2"].font = Font(bold=True, size=11)

        ws["C2"] = f"TURMA: {serie}º ANO {turma}"
        ws.merge_cells("C2:F2")
        ws["C2"].font = Font(bold=True, size=11)
        ws["C2"].alignment = Alignment(horizontal="center")

        ws["G2"] = f"PROFESSORA: {prof}" if prof else "PROFESSORA:"
        ws.merge_cells("G2:K2")
        ws["G2"].font = Font(bold=True, size=11)
        ws.row_dimensions[2].height = 20

        # Linha 3: sub-cabeçalho "Livros lidos XXXX"
        ws.merge_cells("C3:F3")
        ws["C3"] = f"Livros lidos {datetime.now().year}"
        ws["C3"].font = Font(bold=True, size=10)
        ws["C3"].alignment = Alignment(horizontal="center")
        ws.row_dimensions[3].height = 16

        # Linha 4: cabeçalhos das colunas
        # A = ALUNA(O), B= 1ªSEM, C=2ªSEM, D=3ªSEM, E=4ªSEM, F=TOTAL
        # (transposto: aluno na col A, semanas em B-E, total F)
        cab_fill  = PatternFill("solid", fgColor="D9D9D9")
        amarelo_f = PatternFill("solid", fgColor="FFFF99")
        headers_rm = ["ALUNA(O)", "1ª\nSEMANA", "2ª\nSEMANA", "3ª\nSEMANA", "4ª\nSEMANA", "TOTAL"]
        col_widths  = [32, 9, 9, 9, 9, 8]

        for col_i, h in enumerate(headers_rm, 1):
            cell = ws.cell(4, col_i, h)
            cell.font = Font(bold=True, size=10)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.fill = cab_fill
            cell.border = thin_border()
        ws.row_dimensions[4].height = 28

        for ci, w in enumerate(col_widths, 1):
            ws.column_dimensions[get_column_letter(ci)].width = w

        # Linhas dos alunos
        rosa_fill   = PatternFill("solid", fgColor="FFE4E1")
        branco_fill = PatternFill("solid", fgColor="FFFFFF")

        for row_i, aluno in enumerate(alunos, 5):
            lidos = dict_lidos.get(aluno, 0)
            fill  = rosa_fill if (row_i % 2 == 0) else branco_fill

            # Col A: nome
            ca = ws.cell(row_i, 1, aluno)
            ca.font = Font(size=10)
            ca.alignment = Alignment(vertical="center")
            ca.fill = fill
            ca.border = thin_border()

            # Cols B-E: semanas — "X" se tem lidos
            for col_s in range(2, 6):
                cs = ws.cell(row_i, col_s, "X" if lidos > 0 else "")
                cs.font = Font(bold=True, size=11, color="0000CD")
                cs.alignment = Alignment(horizontal="center", vertical="center")
                cs.fill = fill
                cs.border = thin_border()

            # Col F: total
            ct = ws.cell(row_i, 6, lidos if lidos > 0 else "")
            ct.font = Font(bold=True, size=11)
            ct.alignment = Alignment(horizontal="center", vertical="center")
            ct.fill = amarelo_f if lidos > 0 else fill
            ct.border = thin_border()

            ws.row_dimensions[row_i].height = 16

        nome_arq = f"RankingMensal_{serie}-{turma}_{mes_nome}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        caminho  = EXPORT_DIR / nome_arq
        wb.save(caminho)
        lbl_status.configure(text=f"✅  Ranking mensal gerado!\n{caminho}")

    # ─────────────────────────────────────────
    #  EXPORTAÇÃO GENÉRICA
    # ─────────────────────────────────────────
    def exportar_dados(self, tipo):
        wb = Workbook(); ws = wb.active
        conn = sqlite3.connect(DB_PATH); c = conn.cursor()
        if tipo == "acervo":
            ws.title = "Acervo"
            ws.append(["Código","Título","Autor","Gênero","Estante","Fileira"])
            c.execute("SELECT codigo, titulo, autor, genero, estante, fileira FROM acervo ORDER BY titulo")
        else:
            ws.title = "Alunos"
            ws.append(["Nome","Série","Turma","Turno"])
            c.execute("SELECT nome, serie, turma, turno FROM alunos ORDER BY nome")
        for row in c.fetchall(): ws.append(row)
        conn.close()
        caminho = EXPORT_DIR / f"Export_{tipo}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        wb.save(caminho)
        messagebox.showinfo("Sucesso", f"{tipo.capitalize()} exportado!\n\n{caminho}")


# ─────────────────────────────────────────────
if __name__ == "__main__":
    app = BibliotecaApp()
    app.mainloop()

