import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import random
import math
from datetime import datetime

try:
    import pandas as pd
except ImportError:
    pd = None

try:
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfgen import canvas
    from reportlab.lib.units import cm
except ImportError:
    canvas = None


# ===== Config e colunas do seu XLSX =====
COL_DATA = "Data"
COL_DEZENAS = ["Dezena 1", "Dezena 2", "Dezena 3", "Dezena 4", "Dezena 5", "Dezena 6"]

MIN_N = 6
MAX_N = 20
DEFAULT_PRECO_6 = 6.00  # base para calcular valor via combinações (bate com sua tabela 6=6, 7=42, 8=168...)


# ===== Util =====
def comb(n, k):
    return math.comb(n, k)


def preco_aposta(qtd_numeros, preco_6=DEFAULT_PRECO_6):
    # Preço = preço_aposta_simples * C(n,6)
    return preco_6 * comb(qtd_numeros, 6)


def br_money(v):
    return f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def parse_date_br(s):
    # "dd/mm/aaaa"
    return datetime.strptime(str(s).strip(), "%d/%m/%Y")


def carregar_resultados_xlsx(path):
    if pd is None:
        raise RuntimeError("Pandas não está instalado. Instale com: pip install pandas openpyxl")

    df = pd.read_excel(path)

    for c in [COL_DATA] + COL_DEZENAS:
        if c not in df.columns:
            raise ValueError(
                f"Coluna '{c}' não encontrada no arquivo. Colunas encontradas: {list(df.columns)}"
            )

    df["_data"] = df[COL_DATA].apply(parse_date_br)

    for c in COL_DEZENAS:
        df[c] = df[c].astype(int)

    return df


def contar_frequencias(df):
    cont = {}
    for c in COL_DEZENAS:
        for n in df[c].tolist():
            cont[n] = cont.get(n, 0) + 1
    # ordena por freq desc e num asc
    return sorted(cont.items(), key=lambda x: (-x[1], x[0]))


def filtrar_mega_da_virada(df):
    # concursos com data 31/12
    return df[(df["_data"].dt.day == 31) & (df["_data"].dt.month == 12)].copy()


def format_jogo(nums):
    return " ".join(f"{n:02d}" for n in nums)


def amostragem_ponderada_sem_repetir(pool_nums, weights, k):
    """
    Sorteia k números únicos, ponderado por frequência.
    (Simples e eficiente para pool pequeno <= 60)
    """
    escolhidos = set()
    tentativas = 0
    while len(escolhidos) < k:
        tentativas += 1
        if tentativas > 20000:
            break
        n = random.choices(pool_nums, weights=weights, k=1)[0]
        escolhidos.add(n)
    return sorted(escolhidos)


def exportar_pdf(path_pdf, jogos, modo, arquivo_origem, n_por_jogo, pool, custo_total=None, top_preview=None):
    if canvas is None:
        raise RuntimeError("ReportLab não está instalado. Instale com: pip install reportlab")

    c = canvas.Canvas(path_pdf, pagesize=A4)
    w, h = A4

    def header():
        y = h - 2.0 * cm
        c.setFont("Helvetica-Bold", 16)
        c.drawString(2 * cm, y, "Gerador Mega-Sena - Jogos Gerados")
        y -= 0.9 * cm
        c.setFont("Helvetica", 10)
        c.drawString(2 * cm, y, f"Data/Hora: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
        y -= 0.55 * cm
        c.drawString(2 * cm, y, f"Modo: {modo}")
        y -= 0.55 * cm
        c.drawString(2 * cm, y, f"Números por jogo: {n_por_jogo} | Pool: {pool}")
        y -= 0.55 * cm
        c.drawString(2 * cm, y, f"Arquivo: {arquivo_origem}")
        y -= 0.55 * cm
        if custo_total is not None:
            c.drawString(2 * cm, y, f"Custo estimado total: {br_money(custo_total)}")
            y -= 0.55 * cm
        y -= 0.25 * cm
        c.line(2 * cm, y, w - 2 * cm, y)
        y -= 0.8 * cm
        return y

    y = header()

    c.setFont("Helvetica", 11)
    for i, jogo in enumerate(jogos, start=1):
        line = f"Jogo {i:03d}: {format_jogo(jogo)}"
        if y < 2.0 * cm:
            c.showPage()
            y = header()
            c.setFont("Helvetica", 11)
        c.drawString(2 * cm, y, line)
        y -= 0.55 * cm

    # Página extra com TOP (se tiver)
    if top_preview:
        c.showPage()
        y = h - 2.0 * cm
        c.setFont("Helvetica-Bold", 14)
        c.drawString(2 * cm, y, "Top números (preview do modo)")
        y -= 1.0 * cm
        c.setFont("Helvetica", 11)
        for line in top_preview:
            if y < 2.0 * cm:
                c.showPage()
                y = h - 2.0 * cm
                c.setFont("Helvetica", 11)
            c.drawString(2 * cm, y, line)
            y -= 0.55 * cm

    c.save()


# ===== App =====
class MegaSenaApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Gerador Mega-Sena (Tkinter) + Exportar PDF")
        self.geometry("1020x700")
        self.minsize(950, 650)

        self.df = None
        self.path_xlsx = ""

        self.jogos_gerados = []
        self.freq_cache = None
        self.top_cache = None

        self._build_ui()

    def _build_ui(self):
        # Top: arquivo
        frm_top = ttk.Frame(self, padding=12)
        frm_top.pack(fill="x")

        ttk.Label(frm_top, text="Planilha de resultados (.xlsx):").grid(row=0, column=0, sticky="w")
        self.var_path = tk.StringVar(value="")
        ttk.Entry(frm_top, textvariable=self.var_path).grid(row=0, column=1, sticky="ew", padx=8)

        ttk.Button(frm_top, text="Selecionar...", command=self.on_browse).grid(row=0, column=2, padx=6)
        ttk.Button(frm_top, text="Carregar", command=self.on_load).grid(row=0, column=3)

        frm_top.columnconfigure(1, weight=1)

        # Configs
        lf_cfg = ttk.LabelFrame(self, text="Configurações", padding=12)
        lf_cfg.pack(fill="x", padx=12, pady=(0, 10))

        ttk.Label(lf_cfg, text="Quantidade de jogos:").grid(row=0, column=0, sticky="w")
        self.var_qtd_jogos = tk.IntVar(value=1)  # padrão 1 (mas você pode mudar)
        sp_qtd = ttk.Spinbox(lf_cfg, from_=1, to=5000, textvariable=self.var_qtd_jogos, width=10, command=self.update_price)
        sp_qtd.grid(row=0, column=1, sticky="w", padx=8)

        ttk.Label(lf_cfg, text="Números por jogo (6 a 20):").grid(row=0, column=2, sticky="w")
        self.var_n_por_jogo = tk.IntVar(value=6)  # padrão 6
        sp_n = ttk.Spinbox(lf_cfg, from_=MIN_N, to=MAX_N, textvariable=self.var_n_por_jogo, width=10, command=self.update_price)
        sp_n.grid(row=0, column=3, sticky="w", padx=8)

        ttk.Label(lf_cfg, text="Modo de geração:").grid(row=1, column=0, sticky="w", pady=(10, 0))
        self.var_modo = tk.StringVar(value="Números mais sorteados (todo o período)")
        modos = [
            "Números mais sorteados (todo o período)",
            "Números mais sorteados (Mega da Virada - 31/12)",
            "Aleatórios",
        ]
        cmb = ttk.Combobox(lf_cfg, values=modos, textvariable=self.var_modo, state="readonly", width=46)
        cmb.grid(row=1, column=1, columnspan=2, sticky="w", padx=8, pady=(10, 0))
        cmb.bind("<<ComboboxSelected>>", lambda e: self.atualizar_preview_top())

        ttk.Label(lf_cfg, text="Pool (Top N números):").grid(row=1, column=3, sticky="w", pady=(10, 0))
        self.var_pool = tk.IntVar(value=30)  # padrão 30
        sp_pool = ttk.Spinbox(lf_cfg, from_=6, to=60, textvariable=self.var_pool, width=10, command=self.atualizar_preview_top)
        sp_pool.grid(row=1, column=4, sticky="w", padx=8, pady=(10, 0))

        ttk.Button(lf_cfg, text="Gerar Jogos", command=self.on_generate).grid(row=0, column=4, sticky="e")
        ttk.Button(lf_cfg, text="Exportar PDF", command=self.on_export_pdf).grid(row=0, column=5, sticky="e", padx=(8, 0))

        # Preços
        frm_price = ttk.Frame(lf_cfg)
        frm_price.grid(row=2, column=0, columnspan=6, sticky="ew", pady=(12, 0))

        self.var_preco_jogo = tk.StringVar(value="R$ 0,00")
        self.var_preco_total = tk.StringVar(value="R$ 0,00")

        ttk.Label(frm_price, text="Valor por jogo:").grid(row=0, column=0, sticky="w")
        ttk.Label(frm_price, textvariable=self.var_preco_jogo, font=("Segoe UI", 10, "bold")).grid(row=0, column=1, sticky="w", padx=6)

        ttk.Label(frm_price, text="Valor total:").grid(row=0, column=2, sticky="w", padx=(30, 0))
        ttk.Label(frm_price, textvariable=self.var_preco_total, font=("Segoe UI", 10, "bold")).grid(row=0, column=3, sticky="w", padx=6)

        self.update_price()

        # Paned output
        pan = ttk.PanedWindow(self, orient="horizontal")
        pan.pack(fill="both", expand=True, padx=12, pady=8)

        lf_left = ttk.LabelFrame(pan, text="Jogos gerados", padding=10)
        lf_right = ttk.LabelFrame(pan, text="Preview TOP / Frequências", padding=10)
        pan.add(lf_left, weight=3)
        pan.add(lf_right, weight=2)

        self.txt_out = tk.Text(lf_left, wrap="none")
        self.txt_out.pack(fill="both", expand=True)

        self.txt_top = tk.Text(lf_right, wrap="word")
        self.txt_top.pack(fill="both", expand=True)

        # Bottom actions
        frm_bottom = ttk.Frame(self, padding=(12, 0, 12, 12))
        frm_bottom.pack(fill="x")

        ttk.Button(frm_bottom, text="Copiar", command=self.on_copy).pack(side="left")
        ttk.Button(frm_bottom, text="Salvar TXT...", command=self.on_save_txt).pack(side="left", padx=8)
        ttk.Button(frm_bottom, text="Limpar", command=self.on_clear).pack(side="left")

        self.txt_out.insert("1.0", "1) Carregue a planilha .xlsx (para modos por frequência)\n2) Ajuste as opções\n3) Clique em Gerar\n")
        self.txt_top.insert("1.0", "Carregue a planilha para ver o preview de frequências.\n")

    def on_browse(self):
        path = filedialog.askopenfilename(
            title="Selecione a planilha de resultados",
            filetypes=[("Excel", "*.xlsx"), ("Todos", "*.*")]
        )
        if path:
            self.var_path.set(path)

    def on_load(self):
        path = self.var_path.get().strip()
        if not path:
            messagebox.showwarning("Atenção", "Selecione a planilha .xlsx primeiro.")
            return
        try:
            self.df = carregar_resultados_xlsx(path)
            self.path_xlsx = path
            messagebox.showinfo("OK", f"Planilha carregada! Concursos: {len(self.df)}")
            self.atualizar_preview_top()
        except Exception as e:
            messagebox.showerror("Erro ao carregar", str(e))

    def update_price(self):
        try:
            qtd = int(self.var_qtd_jogos.get())
            n = int(self.var_n_por_jogo.get())
        except Exception:
            return
        if n < MIN_N or n > MAX_N or qtd < 1:
            return

        vj = preco_aposta(n)
        vt = vj * qtd
        self.var_preco_jogo.set(br_money(vj))
        self.var_preco_total.set(br_money(vt))

    def _get_df_by_mode(self):
        modo = self.var_modo.get()
        if modo == "Números mais sorteados (Mega da Virada - 31/12)":
            if self.df is None:
                return None
            return filtrar_mega_da_virada(self.df)
        elif modo == "Números mais sorteados (todo o período)":
            return self.df
        else:
            return None  # aleatório não usa df

    def atualizar_preview_top(self):
        self.update_price()
        self.txt_top.delete("1.0", "end")

        modo = self.var_modo.get()
        pool = int(self.var_pool.get())
        n = int(self.var_n_por_jogo.get())

        if modo == "Aleatórios":
            self.freq_cache = None
            self.top_cache = None
            self.txt_top.insert("1.0", "Modo Aleatórios: sorteia números de 1 a 60.\n")
            self.txt_top.insert("end", f"Pool: 1..60 | Jogo: {n} números\n")
            return

        if self.df is None:
            self.txt_top.insert("1.0", "Carregue a planilha para ver frequências.\n")
            return

        df_use = self._get_df_by_mode()
        if df_use is None or len(df_use) == 0:
            self.txt_top.insert("1.0", "Não há dados suficientes para este modo (filtro vazio).\n")
            return

        freq = contar_frequencias(df_use)
        top = freq[:pool]
        self.freq_cache = freq
        self.top_cache = top

        self.txt_top.insert("1.0", f"{modo}\n")
        self.txt_top.insert("end", f"Concursos usados: {len(df_use)}\n")
        self.txt_top.insert("end", f"Pool = Top {pool} números | Jogo = {n} números\n\n")
        self.txt_top.insert("end", "Top 20 (número: frequência):\n")
        for num, fr in freq[:20]:
            self.txt_top.insert("end", f"{num:02d}: {fr}\n")

    def gerar_um_jogo(self):
        modo = self.var_modo.get()
        pool = int(self.var_pool.get())
        n = int(self.var_n_por_jogo.get())

        if modo == "Aleatórios":
            return sorted(random.sample(range(1, 61), n))

        if self.df is None:
            raise RuntimeError("Carregue a planilha para usar modos por frequência.")

        df_use = self._get_df_by_mode()
        if df_use is None or len(df_use) == 0:
            raise RuntimeError("Filtro do modo retornou 0 concursos. Verifique a planilha.")

        freq = contar_frequencias(df_use)
        top = freq[:pool]
        pool_nums = [x[0] for x in top]
        weights = [x[1] for x in top]

        # Ponderado pelos mais frequentes dentro do pool
        jogo = amostragem_ponderada_sem_repetir(pool_nums, weights, n)
        return jogo

    def on_generate(self):
        try:
            qtd = int(self.var_qtd_jogos.get())
            n = int(self.var_n_por_jogo.get())
            if qtd < 1:
                messagebox.showwarning("Atenção", "Quantidade de jogos deve ser >= 1.")
                return
            if n < MIN_N or n > MAX_N:
                messagebox.showwarning("Atenção", "Números por jogo devem ser entre 6 e 20.")
                return

            self.update_price()
            self.atualizar_preview_top()

            self.jogos_gerados = []
            for _ in range(qtd):
                self.jogos_gerados.append(self.gerar_um_jogo())

            self.txt_out.delete("1.0", "end")
            self.txt_out.insert("1.0", f"Modo: {self.var_modo.get()}\n")
            self.txt_out.insert("end", f"Qtd jogos: {qtd} | Números por jogo: {n} | Pool: {int(self.var_pool.get())}\n")
            self.txt_out.insert("end", f"Custo estimado total: {self.var_preco_total.get()}\n\n")

            for i, jogo in enumerate(self.jogos_gerados, start=1):
                self.txt_out.insert("end", f"Jogo {i:03d}: {format_jogo(jogo)}\n")

        except Exception as e:
            messagebox.showerror("Erro", str(e))

    def on_export_pdf(self):
        if not self.jogos_gerados:
            messagebox.showwarning("Atenção", "Gere os jogos antes de exportar.")
            return
        if canvas is None:
            messagebox.showerror("Dependência faltando", "Instale o ReportLab: pip install reportlab")
            return

        path = filedialog.asksaveasfilename(
            title="Salvar PDF",
            defaultextension=".pdf",
            filetypes=[("PDF", "*.pdf")]
        )
        if not path:
            return

        try:
            modo = self.var_modo.get()
            pool = int(self.var_pool.get())
            n = int(self.var_n_por_jogo.get())

            custo_total = preco_aposta(n) * int(self.var_qtd_jogos.get())

            top_preview_lines = None
            if modo != "Aleatórios" and self.freq_cache:
                # TOP 20 para registrar no PDF
                top_preview_lines = ["Top 20 (número: frequência):"]
                for num, fr in self.freq_cache[:20]:
                    top_preview_lines.append(f"{num:02d}: {fr}")

            exportar_pdf(
                path_pdf=path,
                jogos=self.jogos_gerados,
                modo=modo,
                arquivo_origem=self.path_xlsx if self.path_xlsx else "(não informado)",
                n_por_jogo=n,
                pool=pool,
                custo_total=custo_total,
                top_preview=top_preview_lines
            )
            messagebox.showinfo("OK", f"PDF salvo em:\n{path}")
        except Exception as e:
            messagebox.showerror("Erro ao exportar", str(e))

    def on_copy(self):
        text = self.txt_out.get("1.0", "end").strip()
        if not text:
            messagebox.showinfo("Info", "Nada para copiar.")
            return
        self.clipboard_clear()
        self.clipboard_append(text)
        messagebox.showinfo("OK", "Copiado para a área de transferência.")

    def on_save_txt(self):
        text = self.txt_out.get("1.0", "end").strip()
        if not text:
            messagebox.showinfo("Info", "Nada para salvar.")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Texto", "*.txt"), ("Todos", "*.*")]
        )
        if not path:
            return
        with open(path, "w", encoding="utf-8") as f:
            f.write(text + "\n")
        messagebox.showinfo("OK", f"Salvo em:\n{path}")

    def on_clear(self):
        self.jogos_gerados = []
        self.txt_out.delete("1.0", "end")
        self.txt_top.delete("1.0", "end")
        self.txt_out.insert("1.0", "Limpo. Gere novamente.\n")
        self.txt_top.insert("1.0", "Carregue a planilha para ver o preview de frequências.\n")


if __name__ == "__main__":
    app = MegaSenaApp()
    app.mainloop()
