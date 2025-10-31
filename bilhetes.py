import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import sqlite3
from datetime import datetime
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Alignment
import shutil
import os
import sys
import tkinter.font as tkfont

# Valores de configuração
TICKET_PRICE = 2.0  # preço por bilhete em euros
INITIAL_CASH = 100.0  # caixa inicial em euros

# --- Impressão térmica (Windows, impressora USB comum) ---
try:
    import win32print
    import win32ui
    from PIL import Image, ImageWin
    WIN32_AVAILABLE = True
except Exception:
    WIN32_AVAILABLE = False

# Imports opcionais para geração de PDF (reportlab). Se não estiverem presentes,
# funções PDF deverão falhar com uma mensagem amigável.
try:
    from reportlab.lib.pagesizes import A4
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Table
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.platypus.tables import TableStyle
    from reportlab.lib import colors
    REPORTLAB_AVAILABLE = True
except Exception:
    REPORTLAB_AVAILABLE = False

    # Configurações financeiras
    PRICE_PER_TICKET = 2.0  # Euros por bilhete
    STARTING_CASH = 100.0   # Valor inicial em caixa (euros)


def _get_default_printer_name():
    """Retorna o nome da impressora por omissão (Windows) ou None."""
    if not WIN32_AVAILABLE:
        return None
    try:
        return win32print.GetDefaultPrinter()
    except Exception:
        return None


def _send_raw_to_printer(printer_name, data_bytes):
    """Envia bytes raw para a impressora através da API win32print.

    Lança exceções em caso de falha.
    """
    if not WIN32_AVAILABLE:
        raise RuntimeError("pywin32 não está disponível no ambiente.")
    if not printer_name:
        raise ValueError("Nome da impressora não fornecido.")
    # Abrir impressora e enviar trabalho RAW
    hPrinter = None
    try:
        hPrinter = win32print.OpenPrinter(printer_name)
        # StartDocPrinter espera uma tupla (pDocument, pOutputFile, pDatatype)
        # usar RAW para enviar comandos ESC/POS
        win32print.StartDocPrinter(hPrinter, 1, ("Bilhete", None, "RAW"))
        win32print.StartPagePrinter(hPrinter)
        win32print.WritePrinter(hPrinter, data_bytes)
        win32print.EndPagePrinter(hPrinter)
        win32print.EndDocPrinter(hPrinter)
    finally:
        try:
            if hPrinter:
                win32print.ClosePrinter(hPrinter)
        except Exception:
            pass


def imprimir_bilhetes_multiplo_pdf(bilhetes, data_hora, assistente, metodo_pagamento=None, recebido=None, troco=None):
    """Gera um único PDF com uma página por bilhete (80mm largura x altura dinâmica por página).
    Cada bilhete contém: título, imagem.png (se existir) logo a seguir ao título, nº do bilhete,
    data/hora e logo.png (se existir). Depois tenta enviar o PDF para a impressora predefinida.
    """
    if not REPORTLAB_AVAILABLE:
        try:
            messagebox.showwarning(
                "Dependência em Falta",
                "A biblioteca 'reportlab' não está instalada. Instale com: pip install reportlab"
            )
        except Exception:
            print("Dependência em Falta: reportlab")
        return

    try:
        from reportlab.lib.units import mm
        from reportlab.platypus import SimpleDocTemplate, Paragraph, Image, Spacer, PageBreak
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        from reportlab.lib.enums import TA_CENTER, TA_LEFT
        from reportlab.lib.utils import ImageReader
    except Exception:
        try:
            messagebox.showwarning("Dependência em Falta", "Falha ao carregar módulos do reportlab.")
        except Exception:
            print("Falha ao carregar módulos do reportlab.")
        return

    import tempfile, time, webbrowser

    # Configurações de página compatíveis com impressora térmica 80mm
    WIDTH_MM = 80
    PAGE_HEIGHT_MM = 123  # altura por página (usar valor dentro do intervalo permitido)
    MARGIN_MM = 1

    width_pt = WIDTH_MM * mm
    height_pt = PAGE_HEIGHT_MM * mm
    margin_pt = MARGIN_MM * mm
    content_width_pt = width_pt - 2 * margin_pt

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle('TitleStyle', parent=styles['Normal'],
                                 fontName='Helvetica-Bold', fontSize=12, alignment=TA_CENTER, leading=13)
    small_style = ParagraphStyle('Small', parent=styles['Normal'],
                                 fontName='Helvetica', fontSize=10, alignment=TA_CENTER, leading=11)
    receipt_style = ParagraphStyle('Small', parent=styles['Normal'],
                                 fontName='Helvetica', fontSize=9, alignment=TA_LEFT, leading=11)

    titulo_text = "Bilhete"
    titulo_text2 = "Igreja Nossa Senhora da Oliveira"

    # preparar imagem.png (após o título) e logo.png (no final), calcular tamanhos
    base_dir = os.path.dirname(__file__)
    imagem_path = os.path.join(base_dir, "imagem.png")
    imagem_exists = False
    imagem_w_pt = imagem_h_pt = 0

    logo_path = os.path.join(base_dir, "logo.png")
    logo_exists = False
    logo_w_pt = logo_h_pt = 0

    if os.path.exists(imagem_path):
        try:
            ir = ImageReader(imagem_path)
            iw, ih = ir.getSize()
            if iw > 0:
                # reduzir imagem.png para metade da largura de conteúdo por defeito
                imagem_w_pt = content_width_pt * 0.50
                imagem_h_pt = imagem_w_pt * (ih / float(iw))
                imagem_exists = True
        except Exception:
            imagem_exists = False

    if os.path.exists(logo_path):
        try:
            ir = ImageReader(logo_path)
            iw, ih = ir.getSize()
            if iw > 0:
                # logo ocupa toda a largura de conteúdo por defeito
                logo_w_pt = content_width_pt
                logo_h_pt = logo_w_pt * (ih / float(iw))
                logo_exists = True
        except Exception:
            logo_exists = False

    # criar ficheiro PDF temporário
    tmpdir = tempfile.gettempdir()
    filename = os.path.join(tmpdir, f"bilhetes_batch_{int(time.time())}.pdf")

    story = []
    for idx, numero in enumerate(bilhetes):
        # logo
        if logo_exists and logo_h_pt > 0:
            try:
                logo_img = Image(logo_path, width=logo_w_pt, height=logo_h_pt)
                logo_img.hAlign = 'CENTER'
                story.append(logo_img)
            except Exception:
                pass

        # título
        story.append(Paragraph(titulo_text, title_style))
        story.append(Paragraph(titulo_text2, title_style))
        story.append(Spacer(1, 2 * mm))

        #preço
        story.append(Paragraph(f"Preço: 2€", title_style))
        story.append(Spacer(1, 2 * mm))

        # imagem.png logo
        if imagem_exists and imagem_h_pt > 0:
            try:
                img = Image(imagem_path, width=imagem_w_pt, height=imagem_h_pt)
                img.hAlign = 'CENTER'
                story.append(img)
                story.append(Spacer(1, 2 * mm))
            except Exception:
                # se falhar, simplesmente ignorar a imagem
                pass

        # data/hora e mensagem
        story.append(Paragraph(f"Data/Hora: {data_hora}", small_style))
        story.append(Spacer(1, 2 * mm))
        
        try:
            story.append(Paragraph(f"Donativo sem contrapartida nos termos do artigo 61 do EBF", small_style))
            story.append(Spacer(1, 2 * mm))
            story.append(Spacer(1, 2 * mm))
        except Exception:
            pass

        # método de pagamento
        try:
            if metodo_pagamento:
                story.append(Paragraph(f"Pagamento: {metodo_pagamento}", receipt_style))
            # se pagamento em numerário e valores fornecidos, mostrar recebido e troco
            if metodo_pagamento and str(metodo_pagamento).strip().lower() == 'dinheiro':
                if recebido is not None:
                    story.append(Paragraph(f"Recebido: €{float(recebido):.2f}", receipt_style))
                if troco is not None:
                    story.append(Paragraph(f"Troco: €{float(troco):.2f}", receipt_style))
                story.append(Spacer(1, 2 * mm))
        except Exception:
            pass

        # adicionar PageBreak entre bilhetes (não após o último)
        if idx != len(bilhetes) - 1:
            story.append(PageBreak())

    # gerar PDF com páginas do mesmo tamanho
    doc = SimpleDocTemplate(filename, pagesize=(width_pt, height_pt),
                            leftMargin=margin_pt, rightMargin=margin_pt,
                            topMargin=margin_pt, bottomMargin=margin_pt)
    try:
        doc.build(story)
    except Exception as e:
        try:
            messagebox.showerror("Erro PDF", f"Falha ao criar PDF dos bilhetes:\n{e}")
        except Exception:
            print("Falha ao criar PDF dos bilhetes:", e)
        return

    # tentar imprimir o PDF na impressora predefinida do sistema (Windows preferencialmente)
    try:
        if sys.platform.startswith("win") and WIN32_AVAILABLE:
            try:
                import win32api
                win32api.ShellExecute(0, "print", filename, None, ".", 0)
            except Exception:
                try:
                    os.startfile(filename, "print")
                except Exception:
                    webbrowser.open(filename)
        else:
            try:
                # em sistemas não-Windows, tenta abrir o PDF (usuário imprime manualmente)
                webbrowser.open(filename)
            except Exception:
                pass
    except Exception as e:
        try:
            messagebox.showinfo("PDF Gerado", f"PDF gerado em:\n{filename}\nImpressão automática falhou: {e}")
        except Exception:
            print("PDF gerado em:", filename, "Impressão automática falhou:", e)
# ==========================
# UTILITÁRIOS
# ==========================
def hoje_str():
    return datetime.now().strftime("%Y-%m-%d")

# Accessibility: increase font sizes for better readability
FONT_INCREASE = 2  # change this number to increase/decrease size
def AF(size, *opts):
    # returns a font tuple with increased size and any extra options (e.g., 'bold')
    try:
        sz = int(size) + FONT_INCREASE
    except Exception:
        sz = size
    return ("Segoe UI", sz) + tuple(opts)


def agora_str():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


# Increase common Tk named fonts for accessibility. Call after creating a Tk root.
def adjust_fonts(delta=FONT_INCREASE):
    names = [
        "TkDefaultFont", "TkTextFont", "TkMenuFont", "TkHeadingFont",
        "TkCaptionFont", "TkSmallCaptionFont", "TkIconFont", "TkTooltipFont"
    ]
    for n in names:
        try:
            f = tkfont.nametofont(n)
            # some fonts report negative sizes on Windows; handle gracefully
            try:
                current = int(f['size'])
            except Exception:
                current = f['size']
            try:
                f.configure(size=current + delta)
            except Exception:
                pass
        except Exception:
            pass


# ==========================
# GESTOR DE BASE DE DADOS
# ==========================
class DatabaseManager:
    def __init__(self, path="bilhetes.db"):
        self.path = path
        self.conn = sqlite3.connect(self.path, detect_types=sqlite3.PARSE_DECLTYPES | sqlite3.PARSE_COLNAMES)
        self.cursor = self.conn.cursor()
        self._criar_tabela()

    def _criar_tabela(self):
        # Cria tabela com coluna 'anotacoes' (opcional). Se a tabela já existir sem a coluna,
        # fazemos uma migração simples adicionando a coluna.
        self.cursor.execute("""
            CREATE TABLE IF NOT EXISTS registos (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                data_hora TEXT,
                assistente TEXT,
                nacionalidade TEXT,
                numero_bilhete TEXT,
                metodo_pagamento TEXT,
                fatura TEXT,
                contribuinte TEXT,
                anotacoes TEXT
            )
        """)
        self.conn.commit()

        # Verificar se a coluna 'anotacoes' existe; se não, adicioná-la (migração para versões antigas)
        try:
            self.cursor.execute("PRAGMA table_info(registos)")
            cols = [r[1] for r in self.cursor.fetchall()]
            if 'anotacoes' not in cols:
                self.cursor.execute("ALTER TABLE registos ADD COLUMN anotacoes TEXT")
                self.conn.commit()
        except Exception:
            # Se qualquer erro ocorrer aqui, não queremos quebrar a inicialização; seguir em frente
            pass

        # tabela para eventos (registos auxiliares como 'nao_entraram')
        try:
            self.cursor.execute("""
                CREATE TABLE IF NOT EXISTS eventos (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    timestamp TEXT,
                    event_type TEXT,
                    count INTEGER,
                    assistente TEXT,
                    notes TEXT
                )
            """)
            self.conn.commit()
        except Exception:
            pass

    def inserir_evento(self, event_type, count=None, assistente=None, notes=None, timestamp=None):
        try:
            ts = timestamp if timestamp is not None else datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            self.cursor.execute(
                "INSERT INTO eventos (timestamp, event_type, count, assistente, notes) VALUES (?, ?, ?, ?, ?)",
                (ts, event_type, count, assistente, notes)
            )
            self.conn.commit()
        except Exception:
            pass

    def obter_eventos_por_tipo(self, event_type, dia_str=None):
        if dia_str is None:
            dia_str = hoje_str()
        try:
            self.cursor.execute(
                "SELECT id, timestamp, event_type, count, assistente, notes FROM eventos WHERE event_type = ? AND date(timestamp) = ? ORDER BY id DESC",
                (event_type, dia_str)
            )
            return self.cursor.fetchall()
        except Exception:
            return []

    def apagar_evento_por_id(self, event_id):
        try:
            self.cursor.execute("DELETE FROM eventos WHERE id = ?", (event_id,))
            self.conn.commit()
            return True
        except Exception:
            return False

    def inserir_registo(self, data_hora, assistente, nacionalidade, numero_bilhete, metodo_pagamento, fatura, contribuinte, anotacoes=None):
        self.cursor.execute("""
            INSERT INTO registos (data_hora, assistente, nacionalidade, numero_bilhete, metodo_pagamento, fatura, contribuinte, anotacoes)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """, (data_hora, assistente, nacionalidade, numero_bilhete, metodo_pagamento, fatura, contribuinte, anotacoes))
        self.conn.commit()

    def ultimo_numero_bilhete(self):
        self.cursor.execute("SELECT numero_bilhete FROM registos ORDER BY id DESC LIMIT 1")
        row = self.cursor.fetchone()
        return row[0] if row else None

    def obter_registos_do_dia(self, dia_str=None):
        if dia_str is None:
            dia_str = hoje_str()
        # Assumimos data_hora armazenada como 'YYYY-MM-DD HH:MM:SS'
        self.cursor.execute("""
            SELECT data_hora, assistente, nacionalidade, numero_bilhete, metodo_pagamento, fatura, contribuinte, anotacoes
            FROM registos
            WHERE date(data_hora) = ?
            ORDER BY id DESC
        """, (dia_str,))
        return self.cursor.fetchall()

    def procurar_por_bilhete(self, termo):
        termo_like = f"%{termo}%"
        self.cursor.execute("""
            SELECT data_hora, assistente, nacionalidade, numero_bilhete, metodo_pagamento, fatura, contribuinte, anotacoes
            FROM registos
            WHERE numero_bilhete LIKE ?
            ORDER BY id DESC
        """, (termo_like,))
        return self.cursor.fetchall()

    def fechar(self):
        try:
            self.conn.close()
        except Exception:
            pass


# ==========================
# INTERFACE - LOGIN
# ==========================
class JanelaLogin:
    def __init__(self):
        self.login = tk.Tk()
        # aplicar ajuste de fontes para acessibilidade
        adjust_fonts()
        self.login.title("Login de Assistente")
        self.login.geometry("400x320")
        self.login.resizable(False, False)
        self.login.configure(bg="#f8fafc")
        
        # Centralizar na tela
        self.login.eval('tk::PlaceWindow . center')

        # Container principal
        main_frame = tk.Frame(self.login, bg="#f8fafc", padx=20, pady=30)
        main_frame.pack(expand=True, fill="both")

        # Título
        title_frame = tk.Frame(main_frame, bg="#f8fafc")
        title_frame.pack(pady=(0, 25))
        
        tk.Label(title_frame, text="Venda de Bilhetes", font=("Segoe UI", 20, "bold"), 
                bg="#f8fafc", fg="#2d3748").pack()
        tk.Label(title_frame, text="Sistema de Gestão", font=("Segoe UI", 12), 
                bg="#f8fafc", fg="#718096").pack(pady=(5, 0))

        # Form container
        form_frame = tk.Frame(main_frame, bg="#f8fafc")
        form_frame.pack(fill="x", pady=20)

        tk.Label(form_frame, text="Nome do Assistente:", font=("Segoe UI", 11, "bold"), 
                bg="#f8fafc", fg="#4a5568").pack(anchor="w", pady=(0, 8))

        self.entry_nome = ttk.Entry(form_frame, font=("Segoe UI", 11), width=25)
        self.entry_nome.pack(fill="x", pady=(0, 20))
        self.entry_nome.focus()

        # Bind Enter key to login
        self.entry_nome.bind('<Return>', lambda e: self.confirmar())

        # Botão de entrar
        btn_entrar = tk.Button(form_frame, text="Entrar no Sistema", 
                              font=("Segoe UI", 11, "bold"),
                              bg="#4299e1", fg="white",
                              activebackground="#3182ce",
                              activeforeground="white",
                              relief="flat",
                              padx=10, pady=10,
                              command=self.confirmar)
        btn_entrar.pack(pady=0)
        
        # Efeito hover no botão
        def on_enter(e):
            btn_entrar['background'] = '#3182ce'
        def on_leave(e):
            btn_entrar['background'] = '#4299e1'
        btn_entrar.bind("<Enter>", on_enter)
        btn_entrar.bind("<Leave>", on_leave)

        # Footer
        footer_frame = tk.Frame(main_frame, bg="#f8fafc")
        footer_frame.pack(side="bottom", pady=10)
        tk.Label(footer_frame, text="Digite seu nome para acessar o sistema", 
                font=("Segoe UI", 9), bg="#f8fafc", fg="#a0aec0").pack()

        self.login.mainloop()

    def confirmar(self):
        nome = self.entry_nome.get().strip()
        if not nome:
            messagebox.showwarning("Aviso", "Insira o nome do assistente.")
            return
        self.login.destroy()
        JanelaPrincipal(nome)


# ==========================
# INTERFACE - APLICAÇÃO PRINCIPAL
# ==========================
class JanelaPrincipal:
    def __init__(self, assistente_nome):
        self.assistente = assistente_nome
        self.dia_fechado = False

        # DB
        try:
            self.db = DatabaseManager("bilhetes.db")
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao abrir BD: {e}")
            return

        # Janela principal
        self.root = tk.Tk()
        # aplicar ajuste de fontes para acessibilidade
        adjust_fonts()
        self.root.title(f"Venda de Bilhetes - Assistente: {self.assistente}")
        self.root.geometry("1200x750")
        self.root.resizable(True, True)
        self.root.configure(bg="#f8fafc")
        
        # Centralizar na tela
        self.root.eval('tk::PlaceWindow . center')

        # Criar UI
        self._criar_interface()
        self.atualizar_tabela()
        self._atualizar_status()

        # Fechar corretamente
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.root.mainloop()

    # --------------------------
    # INTERFACE
    # --------------------------
    def _criar_interface(self):
        # Header
        header = tk.Frame(self.root, bg="#2d3748", height=80)
        header.pack(fill="x", side="top")
        header.pack_propagate(False)

        # Conteúdo do header
        header_content = tk.Frame(header, bg="#2d3748")
        header_content.pack(expand=True, fill="both", padx=30)

        tk.Label(header_content, text="Sistema de Venda de Bilhetes",
                 font=AF(16, "bold"), bg="#2d3748", fg="white").pack(side="left")

        user_frame = tk.Frame(header_content, bg="#2d3748")
        user_frame.pack(side="right")

        tk.Label(user_frame, text="Assistente:", font=AF(10),
                 bg="#2d3748", fg="#cbd5e0").pack(side="left")
        self.lbl_assistente = tk.Label(user_frame, text=self.assistente,
                                      font=AF(10, "bold"),
                                      bg="#2d3748", fg="white")
        self.lbl_assistente.pack(side="left", padx=(5, 15))

        # Botão trocar assistente com estilo moderno
        btn_trocar = tk.Button(user_frame, text="Trocar Assistente",
                               font=AF(9),
                               bg="#4a5568", fg="white",
                               activebackground="#718096",
                               activeforeground="white",
                               relief="flat",
                               padx=12, pady=4,
                               command=self._popup_trocar_assistente)
        btn_trocar.pack(side="left")

        # Container principal
        # Container principal com scroll (garante que todo o conteúdo fica acessível em ecrãs pequenos)
        container_outer = tk.Frame(self.root)
        container_outer.pack(expand=True, fill="both", padx=20, pady=20)

        canvas = tk.Canvas(container_outer, bg="#f8fafc", highlightthickness=0)
        vscroll = ttk.Scrollbar(container_outer, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vscroll.set)
        vscroll.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        # frame onde será colocado todo o conteúdo (usamos o mesmo nome main_container para compatibilidade)
        main_container = tk.Frame(canvas, bg="#f8fafc")
        self._canvas_window = canvas.create_window((0, 0), window=main_container, anchor="nw")

        # atualizar scrollregion quando o conteúdo mudar e forçar largura interna igual à do canvas
        def _on_frame_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
        def _on_canvas_configure(event):
            canvas.itemconfig(self._canvas_window, width=event.width)
        main_container.bind("<Configure>", _on_frame_configure)
        canvas.bind("<Configure>", _on_canvas_configure)

        # suporte a roda do rato (Windows/Mac/Linux)
        def _on_mousewheel(event):
            try:
                if sys.platform == "darwin":
                    delta = -1 * int(event.delta)
                else:
                    delta = -1 * int(event.delta / 120)
            except Exception:
                delta = 0
            if delta:
                canvas.yview_scroll(delta, "units")
        canvas.bind_all("<MouseWheel>", _on_mousewheel)
        canvas.bind_all("<Button-4>", lambda e: canvas.yview_scroll(-1, "units"))
        canvas.bind_all("<Button-5>", lambda e: canvas.yview_scroll(1, "units"))

        # Painel esquerdo - Formulário de venda
        left_panel = tk.Frame(main_container, bg="white", relief="flat", bd=1)
        left_panel.pack(side="left", fill="y", padx=(0, 15))

        # Título do painel
        panel_title = tk.Frame(left_panel, bg="#4299e1", height=40)
        panel_title.pack(fill="x", side="top")
        panel_title.pack_propagate(False)
        tk.Label(panel_title, text="Nova Venda", font=AF(12, "bold"), bg="#4299e1", fg="white").pack(expand=True)

        # Form container
        form_container = tk.Frame(left_panel, bg="white", padx=20, pady=20)
        form_container.pack(expand=True, fill="both")

        # Campos do formulário
        fields = [
            ("Nacionalidade:", "combo_nacionalidade"),
            ("Método de Pagamento:", "combo_pagamento"),
            ("Recibo:", "combo_fatura"),
            ("Anotações (opcional):", "entry_anotacoes"),
            ("Quantidade:", "spin_quantidade")
        ]

        nacionalidades = [
            "Português", "Brasileiro", "Espanhol", "Inglês", "Francês", "Italiano", "Asiático", "Alemão", "Outros"
        ]

        # usamos um contador de linhas para poder inserir a entrada manual logo abaixo do combo de nacionalidade
        line = 0
        for label, field_type in fields:
            tk.Label(form_container, text=label, font=AF(10, "bold"), bg="white", fg="#4a5568").grid(row=line, column=0, sticky="w", pady=12, padx=(0, 10))

            if "combo" in field_type:
                values = nacionalidades if "nacionalidade" in field_type else ["Dinheiro", "Cartão"] if "pagamento" in field_type else ["Sim", "Não"]
                default = "Português" if "nacionalidade" in field_type else "Dinheiro" if "pagamento" in field_type else "Não"
                combo = ttk.Combobox(form_container, values=values, state="readonly", font=AF(10), width=22)
                combo.set(default)
                combo.grid(row=line, column=1, sticky="ew", pady=12)
                setattr(self, field_type, combo)
                # Se este for o combo de nacionalidade, ligar o evento para mostrar entrada manual
                if "nacionalidade" in field_type:
                    # a entrada manual deve ficar imediatamente abaixo deste combo
                    self._manual_row = line + 1
                    combo.bind('<<ComboboxSelected>>', lambda e: self._on_nacionalidade_change(e))
                    # reservar linha para a entrada manual (não inserida ainda)
                    line += 1
                # Se este for o combo de fatura, reservar linha para o Nº Contribuinte
                if "fatura" in field_type:
                    self._contrib_row = line + 1
                    combo.bind('<<ComboboxSelected>>', lambda e: self._on_fatura_change(e))
                    # reservar linha para a entrada de contribuinte
                    line += 1
            elif "entry" in field_type:
                # Campo de anotações: usar um widget multi-linha (ScrolledText)
                if "anotacoes" in field_type:
                    txt = scrolledtext.ScrolledText(form_container, font=AF(10), width=36, height=4, wrap=tk.WORD)
                    txt.grid(row=line, column=1, sticky="ew", pady=12)
                    setattr(self, field_type, txt)
                else:
                    entry = ttk.Entry(form_container, font=AF(10), width=25)
                    entry.grid(row=line, column=1, sticky="ew", pady=12)
                    setattr(self, field_type, entry)
            elif "spin" in field_type:
                spin = ttk.Spinbox(form_container, from_=1, to=50, width=10, font=AF(10))
                # por padrão, 0 pessoas não entraram
                if 'nao_entraram' in field_type:
                    spin.set(0)
                else:
                    spin.set(1)
                spin.grid(row=line, column=1, sticky="w", pady=12)
                setattr(self, field_type, spin)

            # avançar para a próxima linha
            line += 1

        # criar placeholder para entrada manual (inicialmente escondido)
        self.manual_nacionalidade_var = tk.StringVar()
        self.entry_manual_nacionalidade = ttk.Entry(form_container, textvariable=self.manual_nacionalidade_var, font=AF(10), width=25)
        # criar também rótulo explícito (não gridado ainda)
        self.lbl_manual_nacionalidade = tk.Label(form_container, text="Especificar Nacionalidade:", font=AF(10, "bold"), bg="white", fg="#4a5568")
        # não grid ainda; será exibida quando necessário via _on_nacionalidade_change

        # placeholder para contribuinte (aparece apenas se Fatura == 'Sim')
        self.contribuinte_var = tk.StringVar()
        self.entry_contribuinte = ttk.Entry(form_container, textvariable=self.contribuinte_var, font=AF(10), width=25)
        self.lbl_contribuinte = tk.Label(form_container, text="Nº Contribuinte:", font=AF(10, "bold"), bg="white", fg="#4a5568")


        # Botões de ação
        btn_frame = tk.Frame(form_container, bg="white")
        # colocar os botões abaixo de todos os campos (usar 'line' para ficar abaixo da Quantidade)
        btn_frame.grid(row=line, column=0, columnspan=2, pady=(25, 10))

        # Botão Guardar
        btn_guardar = tk.Button(btn_frame, text="✓ Guardar Registo",
                               font=AF(11, "bold"),
                               bg="#48bb78", fg="white",
                               activebackground="#38a169",
                               activeforeground="white",
                               relief="flat",
                               padx=20, pady=10,
                               command=self.guardar_registo)
        btn_guardar.pack(side="left", padx=(0, 10))

        # Botão Fechar Dia
        btn_fechar = tk.Button(btn_frame, text="📊 Fechar o Dia",
                              font=AF(11, "bold"),
                              bg="#ed8936", fg="white",
                              activebackground="#dd6b20",
                              activeforeground="white",
                              relief="flat",
                              padx=20, pady=10,
                              command=self.fechar_dia)
        btn_fechar.pack(side="left")

        # Linha abaixo dos botões: registo de pessoas que não entraram
        try:
            reg_row = line + 1
            tk.Label(form_container, text="Registar Não Entraram:", font=AF(10, "bold"), bg="white", fg="#4a5568").grid(row=reg_row, column=0, sticky="w", pady=8, padx=(0,10))
            self.spin_reg_nao_entraram = ttk.Spinbox(form_container, from_=1, to=50, width=10, font=AF(10))
            self.spin_reg_nao_entraram.set(1)
            self.spin_reg_nao_entraram.grid(row=reg_row, column=1, sticky="w", pady=8)
            # botão para registar o(s) não entrado(s)
            def _on_registar_nao_entraram():
                # impedir registos após o dia estar fechado
                try:
                    if getattr(self, 'dia_fechado', False):
                        messagebox.showwarning("Aviso", "O dia já está fechado. Não é possível registar mais 'Não Entraram'.")
                        return
                except Exception:
                    pass
                try:
                    cnt = int(self.spin_reg_nao_entraram.get())
                except Exception:
                    messagebox.showwarning("Aviso", "Quantidade inválida para registo de Não Entraram.")
                    return
                if cnt <= 0:
                    messagebox.showwarning("Aviso", "Quantidade deve ser pelo menos 1.")
                    return
                # Registar um único evento com a quantidade (count) e o timestamp atual.
                ts = agora_str()
                try:
                    self.db.inserir_evento('nao_entraram', count=cnt, assistente=self.assistente, notes=None, timestamp=ts)
                    # mostrar apenas a hora ao utilizador
                    try:
                        hora = ts.split(' ')[1]
                    except Exception:
                        hora = ts
                    messagebox.showinfo("Registo Efetuado", f"Registados {cnt} não entrado(s) às {hora}")
                    self._set_status(f"{cnt} 'Não Entraram' registado(s) às {hora}.")
                    try:
                        # resetar spinbox para 1 após registo
                        self.spin_reg_nao_entraram.set(1)
                    except Exception:
                        pass
                except Exception:
                    messagebox.showwarning("Aviso", "Falha ao registar Não Entraram.")

            btn_reg_nao = tk.Button(form_container, text="Registar Não Entrou(s)", font=AF(10), bg="#f56565", fg="white", activebackground="#c53030", relief="flat", command=_on_registar_nao_entraram)
            # colocar o botão abaixo da spinbox (na próxima linha, mesma coluna da spinbox)
            btn_reg_nao.grid(row=reg_row+1, column=1, sticky="w", pady=(4, 0))
            # expor como atributo para poder ser desativado ao fechar o dia
            try:
                self.btn_reg_nao = btn_reg_nao
            except Exception:
                pass
        except Exception:
            # se algo falhar aqui, não quebrar a criação da interface
            pass

        # Painel direito - Estatísticas e dados
        right_panel = tk.Frame(main_container, bg="#f8fafc")
        right_panel.pack(side="left", expand=True, fill="both")

        # Frame de estatísticas
        stats_frame = tk.Frame(right_panel, bg="white", relief="flat", bd=1)
        stats_frame.pack(fill="x", pady=(0, 15))

        stats_title = tk.Frame(stats_frame, bg="#38a169", height=40)
        stats_title.pack(fill="x", side="top")
        stats_title.pack_propagate(False)
        tk.Label(stats_title, text="Estatísticas do Dia", font=AF(12, "bold"), bg="#38a169", fg="white").pack(expand=True)

        stats_content = tk.Frame(stats_frame, bg="white", padx=20, pady=15)
        stats_content.pack(expand=True, fill="both")

        self.lbl_total_today = tk.Label(stats_content, text="Total de bilhetes hoje: 0", font=AF(12, "bold"), bg="white", fg="#2d3748")
        self.lbl_total_today.pack(anchor="w", pady=(0, 15))

        # Caixa / Totais de pagamento
        money_frame = tk.Frame(stats_content, bg="white")
        money_frame.pack(anchor="w", pady=(0, 10))

        self.lbl_caixa_total = tk.Label(money_frame, text=f"Caixa total: €{INITIAL_CASH:.2f}", font=AF(10, "bold"), bg="white", fg="#2d3748")
        self.lbl_caixa_total.pack(anchor="w")
        self.lbl_numerario = tk.Label(money_frame, text="Numerário: €0.00", font=AF(10), bg="white", fg="#2d3748")
        self.lbl_numerario.pack(anchor="w")
        self.lbl_multibanco = tk.Label(money_frame, text="Multibanco: €0.00", font=AF(10), bg="white", fg="#2d3748")
        self.lbl_multibanco.pack(anchor="w")

        # Tabela de nacionalidades
        nat_frame = tk.Frame(stats_content, bg="white")
        nat_frame.pack(fill="both", expand=True)

        tk.Label(nat_frame, text="Distribuição por Nacionalidade:", font=AF(10, "bold"), bg="white", fg="#4a5568").pack(anchor="w")

        self.lst_nacionalidades = ttk.Treeview(nat_frame, columns=("nacionalidade", "count"), show="headings", height=8)
        self.lst_nacionalidades.heading("nacionalidade", text="Nacionalidade")
        self.lst_nacionalidades.heading("count", text="Total")
        self.lst_nacionalidades.column("nacionalidade", width=180)
        self.lst_nacionalidades.column("count", width=80, anchor="center")
        
        # Scrollbar para a treeview
        nat_scroll = ttk.Scrollbar(nat_frame, orient="vertical", command=self.lst_nacionalidades.yview)
        self.lst_nacionalidades.configure(yscrollcommand=nat_scroll.set)
        
        self.lst_nacionalidades.pack(side="left", fill="both", expand=True, pady=(8, 0))
        nat_scroll.pack(side="right", fill="y")

        # Frame de pesquisa
        search_frame = tk.Frame(right_panel, bg="white", relief="flat", bd=1)
        search_frame.pack(fill="x", pady=(0, 15))

        search_content = tk.Frame(search_frame, bg="white", padx=20, pady=15)
        search_content.pack(expand=True, fill="both")

        tk.Label(search_content, text="Pesquisar Bilhetes", font=AF(11, "bold"), bg="white", fg="#4a5568").pack(anchor="w", pady=(0, 10))

        search_controls = tk.Frame(search_content, bg="white")
        search_controls.pack(fill="x")

        tk.Label(search_controls, text="Nº do bilhete:", font=AF(10), bg="white", fg="#4a5568").pack(side="left")

        self.entry_search = ttk.Entry(search_controls, font=AF(10), width=20)
        self.entry_search.pack(side="left", padx=8)

        btn_pesquisar = tk.Button(search_controls, text="🔍 Pesquisar",
                                 font=AF(9),
                                 bg="#4299e1", fg="white",
                                 activebackground="#3182ce",
                                 activeforeground="white",
                                 relief="flat",
                                 padx=12, pady=4,
                                 command=self.pesquisar_bilhete)
        btn_pesquisar.pack(side="left", padx=(5, 10))

        btn_limpar = tk.Button(search_controls, text="Limpar Filtro",
                              font=AF(9),
                              bg="#a0aec0", fg="white",
                              activebackground="#718096",
                              activeforeground="white",
                              relief="flat",
                              padx=12, pady=4,
                              command=self.atualizar_tabela)
        btn_limpar.pack(side="left")

        # Tabela principal de registos
        table_frame = tk.Frame(right_panel, bg="white", relief="flat", bd=1)
        table_frame.pack(expand=True, fill="both")

        table_title = tk.Frame(table_frame, bg="#4a5568", height=40)
        table_title.pack(fill="x", side="top")
        table_title.pack_propagate(False)
        tk.Label(table_title, text="Registos de Hoje", font=AF(12, "bold"), bg="#4a5568", fg="white").pack(expand=True)

        table_content = tk.Frame(table_frame, bg="white")
        table_content.pack(expand=True, fill="both", padx=2, pady=2)

        # Treeview (registos do dia)
        cols = ("data_hora", "assistente", "nacionalidade", "numero_bilhete", "metodo_pagamento", "fatura", "contribuinte", "anotacoes")
        self.tree = ttk.Treeview(table_content, columns=cols, show="headings", height=15)
        
        # Configurar colunas
        # Map some internal column names to nicer display headings
        display_names = {
            'fatura': 'Recibo'
        }
        for c in cols:
            heading = display_names.get(c, c.replace("_", " ").capitalize())
            self.tree.heading(c, text=heading)
            # aumentar largura da coluna 'anotacoes'
            col_width = 220 if c == 'anotacoes' else 120
            self.tree.column(c, width=col_width, anchor="center")
        
        # Scrollbars
        v_scroll = ttk.Scrollbar(table_content, orient="vertical", command=self.tree.yview)
        h_scroll = ttk.Scrollbar(table_content, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)
        
        self.tree.pack(side="left", fill="both", expand=True)
        v_scroll.pack(side="right", fill="y")
        h_scroll.pack(side="bottom", fill="x")

        # Bind duplo-clique para mostrar detalhes (anotações completas)
        self.tree.bind('<Double-1>', self._mostrar_detalhes)

        # Status bar
        status_bar = tk.Frame(self.root, bg="#e2e8f0", height=30)
        status_bar.pack(fill="x", side="bottom")
        status_bar.pack_propagate(False)
        
        self.status_var = tk.StringVar()
        self.status_var.set("Sistema pronto - Aguardando ações")
        status_label = tk.Label(status_bar, textvariable=self.status_var, font=AF(9), bg="#e2e8f0", fg="#4a5568")
        status_label.pack(side="left", padx=15)

    # --------------------------
    # FUNÇÕES DE AÇÃO
    # --------------------------
    def _popup_trocar_assistente(self):
        if self.dia_fechado:
            messagebox.showwarning("Aviso", "O dia já está fechado! Não é possível trocar assistente.")
            return
        popup = tk.Toplevel(self.root)
        popup.title("Trocar Assistente")
        popup.geometry("320x140")
        popup.resizable(False, False)
        popup.configure(bg="#f8fafc")
        popup.transient(self.root)
        popup.grab_set()
        
        ttk.Label(popup, text="Novo nome do assistente:").pack(pady=(12, 6))
        entry = ttk.Entry(popup, width=30)
        entry.pack(pady=(0, 10))
        entry.focus()

        def confirmar():
            nome = entry.get().strip()
            if not nome:
                messagebox.showwarning("Aviso", "Insira um nome válido.")
                return
            self.assistente = nome
            self.lbl_assistente.config(text=self.assistente)
            self.root.title(f"Venda de Bilhetes - Assistente: {self.assistente}")
            popup.destroy()
            self._set_status(f"Assistente alterado para: {self.assistente}")

        ttk.Button(popup, text="Confirmar", command=confirmar).pack(pady=(0, 10))

    def _mostrar_detalhes(self, event=None):
        # mostra um popup com os detalhes da linha (especialmente as anotações completas)
        sel = self.tree.selection()
        if not sel:
            return
        item = sel[0]
        vals = self.tree.item(item, "values")
        # assumimos que anotacoes é a última coluna
        anot = vals[-1] if vals and len(vals) > 0 else ""

        popup = tk.Toplevel(self.root)
        popup.title("Anotações")
        popup.geometry("500x300")
        popup.transient(self.root)
        txt = scrolledtext.ScrolledText(popup, font=("Segoe UI", 10), wrap=tk.WORD)
        txt.pack(expand=True, fill="both", padx=10, pady=10)
        txt.insert("1.0", anot)
        txt.config(state="disabled")
        ttk.Button(popup, text="Fechar", command=popup.destroy).pack(pady=(0, 10))

    def _proximo_numero_bilhete(self):
        ultimo = self.db.ultimo_numero_bilhete()
        ano = datetime.now().year
        if ultimo and isinstance(ultimo, str) and ultimo.startswith(f"IG{ano}-"):
            try:
                n = int(ultimo.split("-")[1])
            except Exception:
                n = 0
            return n + 1
        return 1

    def _on_nacionalidade_change(self, event=None):
        try:
            val = self.combo_nacionalidade.get()
        except Exception:
            return
        # se selecionado 'Outros', mostrar entrada manual
        if val and val.lower() == 'outros':
            # inserir a entrada manual na linha definida
            try:
                # mostrar rótulo e entrada na mesma linha
                self.lbl_manual_nacionalidade.grid(row=self._manual_row, column=0, sticky='w', pady=12, padx=(0, 10))
                self.entry_manual_nacionalidade.grid(row=self._manual_row, column=1, sticky='ew', pady=12)
            except Exception:
                pass
        else:
            # esconder a entrada manual
            try:
                self.entry_manual_nacionalidade.grid_forget()
                self.lbl_manual_nacionalidade.grid_forget()
                self.manual_nacionalidade_var.set('')
            except Exception:
                pass

    def _on_fatura_change(self, event=None):
        try:
            val = self.combo_fatura.get()
        except Exception:
            return
        if val and val.lower() == 'sim':
            try:
                self.lbl_contribuinte.grid(row=self._contrib_row, column=0, sticky='w', pady=12, padx=(0, 10))
                self.entry_contribuinte.grid(row=self._contrib_row, column=1, sticky='ew', pady=12)
            except Exception:
                pass
        else:
            try:
                self.entry_contribuinte.grid_forget()
                self.lbl_contribuinte.grid_forget()
                self.contribuinte_var.set('')
            except Exception:
                pass

    def guardar_registo(self):
        if self.dia_fechado:
            messagebox.showwarning("Aviso", "O dia já está fechado! Não é possível registar bilhetes.")
            return

        nacionalidade = self.combo_nacionalidade.get()
        # Se selecionado 'Outros', exigir valor manual e usá-lo
        try:
            if nacionalidade and nacionalidade.lower() == 'outros':
                manual_val = self.manual_nacionalidade_var.get().strip()
                if not manual_val:
                    messagebox.showwarning("Aviso", "Por favor especifique a nacionalidade quando selecionar 'Outros'.")
                    return
                nacionalidade = manual_val
        except Exception:
            pass
        metodo_pagamento = self.combo_pagamento.get()
        fatura = self.combo_fatura.get()
        # contribuinte is required when fatura == 'Sim'
        contribuinte = None
        try:
            if fatura and fatura.lower() == 'sim':
                contribuinte = self.contribuinte_var.get().strip()
                if not contribuinte:
                    messagebox.showwarning("Aviso", "Por favor insira o Nº Contribuinte quando selecionar 'Sim' em Recibo.")
                    return
            else:
                # ensure empty when not required
                contribuinte = None
        except Exception:
            contribuinte = None
        try:
            quantidade = int(self.spin_quantidade.get())
            if quantidade <= 0:
                raise ValueError()
        except Exception:
            messagebox.showwarning("Aviso", "Quantidade inválida.")
            return

        # ScrolledText: obter todo o texto (multi-linha)
        try:
            anotacoes = self.entry_anotacoes.get("1.0", "end").strip() or None
        except Exception:
            # se por alguma razão não for ScrolledText (compatibilidade), tentar Entry
            anotacoes = getattr(self, 'entry_anotacoes').get().strip() or None

        # preparar números dos bilhetes (ainda não gravar — pedir pagamento primeiro)
        proximo = self._proximo_numero_bilhete()
        ano = datetime.now().year
        data_hora = agora_str()
        bilhetes = [f"IG{ano}-{proximo + i}" for i in range(quantidade)]

        # limpar campos e atualizar
        self.combo_nacionalidade.set("Português")
        # esconder/limpar entrada manual se visível
        try:
            self.entry_manual_nacionalidade.grid_forget()
            self.manual_nacionalidade_var.set("")
        except Exception:
            pass
        self.combo_pagamento.set("Dinheiro")
        self.combo_fatura.set("Não")
        self.entry_contribuinte.delete(0, tk.END)
        # limpar anotacoes
        # limpar anotacoes (ScrolleText)
        try:
            self.entry_anotacoes.delete("1.0", tk.END)
        except Exception:
            try:
                self.entry_anotacoes.delete(0, tk.END)
            except Exception:
                pass
        self.spin_quantidade.set(1)

        # Após preparar os bilhetes, dependendo do método de pagamento:
        # - se for 'Dinheiro' abrir popup para introduzir valor recebido e calcular troco
        # - se for outro método (ex. cartão/multibanco) gravar diretamente e gerar o PDF
        try:
            total_price = len(bilhetes) * TICKET_PRICE
            metodo_norm = (metodo_pagamento or "").strip().lower()
            if metodo_norm and metodo_norm != 'dinheiro':
                # pagamento por cartão: gravar registos e gerar PDF sem pedir valor recebido
                try:
                    for numero in bilhetes:
                        try:
                            self.db.inserir_registo(data_hora, self.assistente, nacionalidade, numero, metodo_pagamento, fatura, contribuinte, anotacoes)
                        except Exception:
                            pass
                    try:
                        self.atualizar_tabela()
                        self._atualizar_status()
                    except Exception:
                        pass
                    try:
                        messagebox.showinfo("Sucesso", f"Foram registados {len(bilhetes)} bilhete(s):\n{', '.join(bilhetes)}")
                        self._set_status(f"{len(bilhetes)} bilhete(s) registado(s).")
                    except Exception:
                        pass
                    try:
                        imprimir_bilhetes_multiplo_pdf(bilhetes, data_hora, self.assistente, metodo_pagamento=metodo_pagamento)
                    except Exception as e:
                        print(f"Erro ao gerar/mandar imprimir PDF dos bilhetes: {e}")
                except Exception as e:
                    messagebox.showerror("Erro", f"Erro ao gravar registos:\n{e}")
            else:
                # pagamento em numerário: pedir valor recebido via popup
                self._pedir_pagamento_e_imprimir(bilhetes, data_hora, total_price,
                                                 nacionalidade, metodo_pagamento, fatura, contribuinte, anotacoes)
        except Exception as e:
            print(f"Erro ao iniciar fluxo de pagamento: {e}")

    def _pedir_pagamento_e_imprimir(self, bilhetes, data_hora, total_price,
                                    nacionalidade=None, metodo_pagamento=None, fatura=None, contribuinte=None, anotacoes=None):
        """Abre um popup modal para introduzir o valor recebido pelo cliente, mostra o troco e confirma antes de gerar o PDF."""
        popup = tk.Toplevel(self.root)
        popup.title("Pagamento")
        popup.geometry("360x180")
        popup.transient(self.root)
        popup.grab_set()

        tk.Label(popup, text=f"Total a Pagar: €{total_price:.2f}", font=AF(11, "bold")).pack(pady=(12, 6))

        entry_frame = tk.Frame(popup)
        entry_frame.pack(pady=(6, 6))
        tk.Label(entry_frame, text="Valor Recebido: €", font=AF(10)).pack(side="left")
        recebido_var = tk.StringVar()
        recebido_entry = ttk.Entry(entry_frame, textvariable=recebido_var, width=12, font=AF(10))
        recebido_entry.pack(side="left")
        recebido_entry.focus()

        troco_label = tk.Label(popup, text="Troco: €0.00", font=AF(10))
        troco_label.pack(pady=(6, 8))

        def _update_troco(*args):
            s = recebido_var.get().strip()
            try:
                val = float(s.replace(',', '.')) if s else 0.0
                troco = val - float(total_price)
                troco_label.config(text=f"Troco: €{troco:.2f}")
            except Exception:
                troco_label.config(text="Troco: —")

        recebido_var.trace_add('write', _update_troco)

        btn_frame = tk.Frame(popup)
        btn_frame.pack(pady=(6, 8))

        def confirmar_pagamento():
            s = recebido_var.get().strip()
            try:
                received = float(s.replace(',', '.'))
            except Exception:
                messagebox.showwarning("Aviso", "Introduza um valor recebido válido (ex.: 10.00)")
                return
            if received < total_price:
                if not messagebox.askyesno("Valor Inferior", "O valor recebido é inferior ao total. Deseja continuar mesmo assim?"):
                    return
            try:
                troco = received - total_price
            except Exception:
                troco = 0.0
            # mostrar troco final antes de prosseguir
            messagebox.showinfo("Troco", f"Troco a entregar: €{troco:.2f}")
            popup.destroy()
            # após confirmação, gravar os registos no BD, atualizar UI e gerar PDF
            try:
                for numero in bilhetes:
                    try:
                        self.db.inserir_registo(data_hora, self.assistente, nacionalidade, numero, metodo_pagamento, fatura, contribuinte, anotacoes)
                    except Exception:
                        # continuar a tentar inserir outros bilhetes
                        pass
                # limpar campos e atualizar
                try:
                    self.combo_nacionalidade.set("Português")
                    self.entry_manual_nacionalidade.grid_forget()
                    self.manual_nacionalidade_var.set("")
                except Exception:
                    pass
                try:
                    self.combo_pagamento.set("Dinheiro")
                    self.combo_fatura.set("Não")
                    self.entry_contribuinte.delete(0, tk.END)
                except Exception:
                    pass
                try:
                    self.entry_anotacoes.delete("1.0", tk.END)
                except Exception:
                    try:
                        self.entry_anotacoes.delete(0, tk.END)
                    except Exception:
                        pass
                try:
                    self.spin_quantidade.set(1)
                except Exception:
                    pass

                # atualizar tabela e estatísticas
                try:
                    self.atualizar_tabela()
                    self._atualizar_status()
                except Exception:
                    pass

                # informar sucesso e gerar PDF
                try:
                    messagebox.showinfo("Sucesso", f"Foram registados {len(bilhetes)} bilhete(s):\n{', '.join(bilhetes)}")
                    self._set_status(f"{len(bilhetes)} bilhete(s) registado(s).")
                except Exception:
                    pass
                try:
                    imprimir_bilhetes_multiplo_pdf(bilhetes, data_hora, self.assistente, metodo_pagamento=metodo_pagamento, recebido=received, troco=troco)
                except Exception as e:
                    print(f"Erro ao gerar/mandar imprimir PDF dos bilhetes: {e}")
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao gravar registos após confirmação do pagamento:\n{e}")

        def cancelar():
            popup.destroy()

        ttk.Button(btn_frame, text="Confirmar", command=confirmar_pagamento).pack(side="left", padx=8)
        ttk.Button(btn_frame, text="Cancelar", command=cancelar).pack(side="left")

        popup.wait_window()

    def atualizar_tabela(self):
        # refresh table with today's data
        for ch in self.tree.get_children():
            self.tree.delete(ch)
        dados = self.db.obter_registos_do_dia()
        for idx, row in enumerate(dados):
            tag = 'evenrow' if idx % 2 == 0 else 'oddrow'
            # ensure data_hora is string
            linha = list(row)
            linha[0] = str(linha[0])
            self.tree.insert("", "end", values=linha, tags=(tag,))
        self._atualizar_status()
        self._atualizar_estatisticas()

    def pesquisar_bilhete(self):
        termo = self.entry_search.get().strip()
        if not termo:
            self.atualizar_tabela()
            return
        for ch in self.tree.get_children():
            self.tree.delete(ch)
        dados = self.db.procurar_por_bilhete(termo)
        for idx, row in enumerate(dados):
            tag = 'evenrow' if idx % 2 == 0 else 'oddrow'
            linha = list(row)
            linha[0] = str(linha[0])
            self.tree.insert("", "end", values=linha, tags=(tag,))
        self._set_status(f"Filtro: '{termo}' ({len(dados)} resultados)")

    def fechar_dia(self):
        if self.dia_fechado:
            messagebox.showinfo("Informação", "O dia já está fechado!")
            return
        if not messagebox.askyesno("Fechar o Dia", "Tem a certeza que deseja fechar o dia?"):
            return
        # pedir anotações finais (opcional)
        popup = tk.Toplevel(self.root)
        popup.title("Anotações Finais (opcional)")
        popup.geometry("500x300")
        popup.transient(self.root)
        popup.grab_set()
        tk.Label(popup, text="Anotações finais (opcional):", font=("Segoe UI", 10, "bold")).pack(anchor='w', padx=10, pady=(10, 0))
        txt_final = scrolledtext.ScrolledText(popup, font=("Segoe UI", 10), wrap=tk.WORD, height=10)
        txt_final.pack(expand=True, fill='both', padx=10, pady=10)

        def confirmar_fecho():
            try:
                self.final_notes = txt_final.get("1.0", "end").strip() or None
            except Exception:
                self.final_notes = None
            popup.destroy()
            # marcar dia fechado e seguir
            self.dia_fechado = True
            # desativar inputs
            for w in [self.combo_nacionalidade, self.combo_pagamento, self.combo_fatura,
                      self.entry_contribuinte, self.entry_anotacoes, self.entry_manual_nacionalidade, self.spin_quantidade, getattr(self, 'spin_reg_nao_entraram', None), getattr(self, 'btn_reg_nao', None)]:
                try:
                    w.config(state="disabled")
                except Exception:
                    pass
            self.gerar_excel()
            self.criar_backup()
            self._set_status("Dia fechado. Relatórios gerados.")

        def cancelar_fecho():
            popup.destroy()
            return

        btns = tk.Frame(popup)
        btns.pack(pady=(0, 10))
        ttk.Button(btns, text="Confirmar Fecho", command=confirmar_fecho).pack(side='left', padx=8)
        ttk.Button(btns, text="Cancelar", command=cancelar_fecho).pack(side='left')
        popup.wait_window()
        # defensive default
        if not hasattr(self, 'final_notes'):
            self.final_notes = None

    # --------------------------
    # BACKUP E RELATÓRIOS
    # --------------------------
    def criar_backup(self):
        hoje = hoje_str()
        pasta_backup = "backups"
        os.makedirs(pasta_backup, exist_ok=True)
        backup_nome = os.path.join(pasta_backup, f"backup_bilhetes_{hoje}.db")
        try:
            if os.path.exists(self.db.path):
                shutil.copy(self.db.path, backup_nome)
                messagebox.showinfo("Backup Criado", f"Cópia de segurança criada em:\n{backup_nome}")
            else:
                messagebox.showwarning("Aviso", "Ficheiro de BD não encontrado para backup.")
        except Exception as e:
            messagebox.showerror("Erro no Backup", f"Falha ao criar cópia de segurança:\n{e}")

    def gerar_excel(self):
        hoje = hoje_str()
        pasta = os.path.join("relatorios", hoje)
        os.makedirs(pasta, exist_ok=True)
        dados = self.db.obter_registos_do_dia()
        if not dados:
            messagebox.showinfo("Sem Dados", "Não existem registos para hoje.")
            return
        wb = Workbook()
        ws = wb.active
        ws.title = "Bilhetes do Dia"
        cabecalho = ["Data/Hora", "Assistente", "Nacionalidade", "Número Bilhete", "Método Pagamento", "Recibo", "Contribuinte", "Anotações"]
        ws.append(cabecalho)
        for col_num, _ in enumerate(cabecalho, 1):
            ws[f"{get_column_letter(col_num)}1"].font = Font(bold=True)
        for row in dados:
            # row now includes anotacoes as last element
            ws.append([str(x) if x is not None else "" for x in row])
        # aplicar wrap na coluna 'Anotações' (última coluna)
        try:
            anot_col = len(cabecalho)
            for cell in ws[get_column_letter(anot_col)]:
                cell.alignment = Alignment(wrap_text=True, vertical='top')
        except Exception:
            pass
        total = len(dados)
        ws.append([])
        ws.append(["Total de Bilhetes Vendidos:", total])
        # Calcular e adicionar totais monetários ao Excel
        try:
            cash_count = 0
            card_count = 0
            for row in dados:
                metodo = (row[4] or "").strip().lower()
                if metodo == 'dinheiro':
                    cash_count += 1
                elif metodo.startswith('cart') or 'multibanco' in metodo or 'cartão' in metodo:
                    card_count += 1
            cash_amount = cash_count * TICKET_PRICE
            card_amount = card_count * TICKET_PRICE
            numerario_total = INITIAL_CASH + cash_amount
            caixa_total = numerario_total + card_amount
            ws.append(["Numerário:", f"€{numerario_total:.2f}"])
            ws.append(["Multibanco:", f"€{card_amount:.2f}"])
            ws.append(["Caixa total:", f"€{caixa_total:.2f}"])
        except Exception:
            # não impedir criação do Excel se falhar o cálculo
            pass
        # Incluir registos 'Não Entraram' (horas) na folha, se existirem eventos para hoje
        try:
            eventos = self.db.obter_eventos_por_tipo('nao_entraram')
            if eventos:
                ws.append([])
                ws.append(["Registos 'Não Entraram' (horas):"])
                # adicionar cabeçalho simples: Hora, Assistente, Quantidade
                ws.append(["Hora", "Assistente", "Quantidade"])
                # eventos is list of (id, timestamp, event_type, count, assistente, notes)
                for ev in reversed(eventos):
                    _, ts, _, cnt, assist, notes = ev
                    try:
                        hora = ts.split(' ')[1]
                    except Exception:
                        hora = ts
                    ws.append([hora, assist or "", cnt or ""])
                # adicionar total de pessoas que não entraram
                try:
                    total_nao_entraram = sum(int(ev[3] or 0) for ev in eventos)
                    ws.append([])
                    ws.append(["Total Não Entraram:", total_nao_entraram])
                except Exception:
                    pass
        except Exception:
            pass
        # incluir anotações finais no Excel: uma linha após os totais e, opcionalmente, numa aba separada
        try:
            if hasattr(self, 'final_notes') and self.final_notes:
                # adicionar linha simples abaixo
                ws.append([])
                ws.append(["Anotações Finais:", self.final_notes])
                # criar página separada com notas (maior legibilidade)
                notas_ws = wb.create_sheet(title="Notas Finais")
                notas_ws.append(["Anotações Finais"])
                # quebrar em linhas para inserir na sheet
                for ln in (self.final_notes or '').split('\n'):
                    notas_ws.append([ln])
                # aplicar wrap na primeira coluna
                for row in notas_ws.iter_rows(min_row=1, max_col=1):
                    for cell in row:
                        try:
                            cell.alignment = Alignment(wrap_text=True, vertical='top')
                        except Exception:
                            pass
        except Exception:
            pass
        for col in ws.columns:
            max_len = max(len(str(c.value)) if c.value else 0 for c in col)
            # limitar largura máxima razoável
            width = min(max_len + 5, 100)
            ws.column_dimensions[get_column_letter(col[0].column)].width = width
        filename = os.path.join(pasta, f"Bilhetes_{hoje}.xlsx")
        try:
            wb.save(filename)
            messagebox.showinfo("Excel Gerado", f"Arquivo Excel criado: {filename}")
        except Exception as e:
            messagebox.showerror("Erro Excel", f"Falha ao criar Excel:\n{e}")

    def gerar_pdf(self):
        hoje = hoje_str()
        if not REPORTLAB_AVAILABLE:
            messagebox.showwarning("Dependência em falta", "A biblioteca 'reportlab' não está instalada. Instale com: pip install reportlab")
            return
        pasta = os.path.join("relatorios", hoje)
        os.makedirs(pasta, exist_ok=True)
        dados = self.db.obter_registos_do_dia()
        if not dados:
            messagebox.showinfo("Sem Dados", "Não existem registos para hoje.")
            return

        filename = os.path.join(pasta, f"Bilhetes_{hoje}.pdf")
        pdf = SimpleDocTemplate(filename, pagesize=A4)
        elementos = []
        styles = getSampleStyleSheet()
        elementos.append(Paragraph(f"<b>Relatório de Bilhetes - {hoje}</b>", styles["Title"]))
        cabecalho = ["Data/Hora", "Assistente", "Nacionalidade", "Nº Bilhete", "Pagamento", "Recibo", "Contribuinte", "Anotações"]
        tabela_dados = [cabecalho]
        for row in dados:
            r = list(row)
            try:
                anot_text = str(r[-1]) if r[-1] is not None else ""
            except Exception:
                anot_text = ""
            r[-1] = Paragraph(anot_text.replace('\n', '<br/>'), styles['BodyText'])
            tabela_dados.append(r)

        # Totais e summaries
        total = len(dados)
        empty_row = [""] * len(cabecalho)
        row_total = [""] * len(cabecalho)
        row_total[0] = "Total de Bilhetes Vendidos:"
        row_total[1] = total
        tabela_dados.append(empty_row)
        tabela_dados.append(row_total)

        # (campo 'Não Entraram' removido — não incluir essa linha no relatório)

        # Anotações finais (se existirem)
        try:
            if hasattr(self, 'final_notes') and self.final_notes:
                row_notes = [""] * len(cabecalho)
                row_notes[0] = Paragraph(self.final_notes.replace('\n', '<br/>'), styles['BodyText'])
                tabela_dados.append([""])  # spacer
                tabela_dados.append(row_notes)
        except Exception:
            pass

        t = Table(tabela_dados, repeatRows=1)
        t.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.lightblue),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
            ('BACKGROUND', (0, 1), (-1, -1), colors.whitesmoke)
        ]))
        elementos.append(t)
        try:
            pdf.build(elementos)
            messagebox.showinfo("PDF Gerado", f"Arquivo PDF criado: {filename}")
        except Exception as e:
            messagebox.showerror("Erro PDF", f"Falha ao criar PDF:\n{e}")

    # --------------------------
    # ESTATÍSTICAS E STATUS
    # --------------------------
    def _atualizar_estatisticas(self):
        # atualiza total e tabela por nacionalidade
        dados = self.db.obter_registos_do_dia()
        total = len(dados)
        self.lbl_total_today.config(text=f"Total de bilhetes hoje: {total}")

        # sumarizar por nacionalidade
        summary = {}
        for row in dados:
            nat = row[2] or "Outros"
            summary[nat] = summary.get(nat, 0) + 1

        # calcular totais por método de pagamento
        try:
            cash_count = 0
            card_count = 0
            for row in dados:
                metodo = (row[4] or "").strip().lower()
                if metodo == 'dinheiro':
                    cash_count += 1
                elif metodo.startswith('cart') or 'multibanco' in metodo or 'cartão' in metodo:
                    card_count += 1
        except Exception:
            cash_count = 0
            card_count = 0

        cash_amount = cash_count * TICKET_PRICE
        card_amount = card_count * TICKET_PRICE

        # Numerário deve incluir o valor inicial da caixa
        numerario_total = INITIAL_CASH + cash_amount
        # Caixa total inclui numerário (com caixa inicial) e também o multibanco
        caixa_total = numerario_total + card_amount

        # atualizar rótulos de valores monetários
        try:
            self.lbl_numerario.config(text=f"Numerário: €{numerario_total:.2f}")
            self.lbl_multibanco.config(text=f"Multibanco: €{card_amount:.2f}")
            self.lbl_caixa_total.config(text=f"Caixa total: €{caixa_total:.2f}")
        except Exception:
            pass

        # limpar lista
        for ch in self.lst_nacionalidades.get_children():
            self.lst_nacionalidades.delete(ch)
        for nat, cnt in sorted(summary.items(), key=lambda x: x[1], reverse=True):
            self.lst_nacionalidades.insert("", "end", values=(nat, cnt))

    def _atualizar_status(self):
        # atualiza tabela e estatísticas
        self._set_status("Sistema pronto")
        self._atualizar_estatisticas()

    def _set_status(self, texto, timeout_ms=5000):
        self.status_var.set(texto)
        if timeout_ms:
            self.root.after(timeout_ms, lambda: self.status_var.set("Sistema pronto - Aguardando ações"))

    # --------------------------
    # ENCERRAMENTO
    # --------------------------
    def _on_close(self):
        if messagebox.askokcancel("Sair", "Deseja sair da aplicação?"):
            try:
                self.db.fechar()
            except Exception:
                pass
            self.root.destroy()
            # exit cleanly (helps if run from a double-click)
            try:
                sys.exit(0)
            except Exception:
                pass


# ==========================
# EXECUÇÃO
# ==========================
if __name__ == "__main__":
    JanelaLogin()