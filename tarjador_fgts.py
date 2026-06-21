#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Tarjador FGTS - Execução Local
================================
Busca funcionários pelo CPF/Nome em uma pasta de arquivos Excel,
e usa os dados exatos encontrados para decidir o que tarjar no PDF.

Regras de tarjamento:
  - Só atua dentro de blocos  "Tomador: XXXX" → "Total do Tomador"
  - Linhas cujo CPF bate com o do Excel → mantém visível
  - Linhas cujo CPF não bate → tarja da coluna Nome até Total
  - Tomador não permitido → tarja o CNPJ + todos os dados do bloco

Requisitos:
  pip install pymupdf pandas openpyxl
"""

import os
import re
import glob
import queue
import threading
import unicodedata

try:
    import fitz
except ImportError:
    raise SystemExit("Instale o PyMuPDF: pip install pymupdf")

try:
    import pandas as pd
except ImportError:
    raise SystemExit("Instale o pandas/openpyxl: pip install pandas openpyxl")

import tkinter as tk
from tkinter import filedialog, messagebox, ttk, scrolledtext

# ─────────────────────────────────────────────────────────────────────────────
# CONFIGURAÇÕES
# ─────────────────────────────────────────────────────────────────────────────

# A partir de qual X (pontos PDF) a tarja começa — logo antes de "Nome Trabalhador"
X_INICIO = 72.0

# Margem direita (a tarja vai até quase a borda da página)
X_FIM_MARGEM = 4.0

# Folga vertical extra em cada linha (cobre matrícula que quebra em sub-linha)
PAD_Y = 2.5

# CPF fixo do cabeçalho do relatório (emissor) — nunca tratado como colaborador
EXCECOES_CPF = { re.sub(r"\D", "", "164.138.428-06") }

# Colunas do Excel (molde fornecido)
COL_CPF        = "CPF"
COL_NOME       = "Nome Trabalhador"
COL_MATRICULA  = "Matricula"
COL_TOMADOR    = "Tomador"

# ─────────────────────────────────────────────────────────────────────────────
# NORMALIZAÇÃO
# ─────────────────────────────────────────────────────────────────────────────

def sem_acento(s: str) -> str:
    return "".join(c for c in unicodedata.normalize("NFD", s)
                   if unicodedata.category(c) != "Mn")

def norm(s) -> str:
    if not isinstance(s, str):
        return ""
    return re.sub(r"\s+", " ", sem_acento(s).upper()).strip()

def norm_digitos(s) -> str:
    return re.sub(r"\D", "", str(s))

norm_cpf   = norm_digitos
norm_cnpj  = norm_digitos

# ─────────────────────────────────────────────────────────────────────────────
# PADRÕES
# ─────────────────────────────────────────────────────────────────────────────

CPF_RE  = re.compile(r"\b\d{3}\.\d{3}\.\d{3}-\d{2}\b")
CNPJ_RE = re.compile(r"\d{2}[\.\s]?\d{3}[\.\s]?\d{3}/\d{4}-\d{2}")

# ─────────────────────────────────────────────────────────────────────────────
# CARREGAMENTO DOS EXCELS
# ─────────────────────────────────────────────────────────────────────────────

def carrega_excels(pasta: str, log) -> pd.DataFrame:
    padroes = ["*.xlsx", "*.xls", "*.csv"]
    arquivos = []
    for p in padroes:
        arquivos += glob.glob(os.path.join(pasta, p))

    if not arquivos:
        raise FileNotFoundError(f"Nenhum Excel/CSV encontrado em: {pasta}")

    dfs = []
    for arq in arquivos:
        try:
            df = pd.read_csv(arq) if arq.lower().endswith(".csv") else pd.read_excel(arq)
            df["_arquivo"] = os.path.basename(arq)
            dfs.append(df)
            log(f"  📂 {os.path.basename(arq)} — {len(df)} linha(s)")
        except Exception as e:
            log(f"  ⚠️ Erro ao ler {os.path.basename(arq)}: {e}")

    if not dfs:
        raise ValueError("Nenhum arquivo Excel pôde ser carregado.")

    return pd.concat(dfs, ignore_index=True)


def busca_funcionarios(df: pd.DataFrame, cpfs_input: list, nomes_input: list, log) -> dict:
    """
    Busca CPFs e/ou Nomes em TODOS os Excels carregados.
    Retorna dict: { cpf_norm → set de valores permitidos para não tarjar }
    Esses valores são extraídos da linha exata do funcionário no Excel.
    """
    cpfs_busca  = { norm_cpf(c)  for c in cpfs_input  if norm_cpf(c)  }
    nomes_busca = { norm(n)      for n in nomes_input  if norm(n)      }

    # cpf_norm → dict com dados da linha
    encontrados = {}

    for _, row in df.iterrows():
        cpf_row  = norm_cpf(str(row.get(COL_CPF,  "")))
        nome_row = norm(str(row.get(COL_NOME, "")))

        bate = (cpf_row in cpfs_busca) or (nome_row and nome_row in nomes_busca)
        if not bate:
            continue

        if cpf_row not in encontrados:
            encontrados[cpf_row] = {
                "nome":       nome_row,
                "matricula":  norm(str(row.get(COL_MATRICULA, ""))),
                "tomador":    norm_cnpj(str(row.get(COL_TOMADOR, ""))),
                "arquivo":    str(row.get("_arquivo", "")),
            }
            log(f"  ✔ {nome_row}  |  CPF: ...{cpf_row[-4:]}  "
                f"  ({encontrados[cpf_row]['arquivo']})")

    return encontrados


# ─────────────────────────────────────────────────────────────────────────────
# DETECÇÃO DE BLOCOS "Tomador … Total do Tomador" NO PDF
# ─────────────────────────────────────────────────────────────────────────────

def linhas_da_pagina(page, tol: float = 3.0) -> list:
    """
    Agrupa palavras (get_text 'words') em linhas pelo centro Y,
    com tolerância de `tol` pontos. Retorna lista ordenada por Y.
    """
    grupos: dict = {}
    for w in page.get_text("words"):
        yc = round((w[1] + w[3]) / 2 / tol) * tol
        grupos.setdefault(yc, []).append(w)

    linhas = []
    for yc in sorted(grupos):
        ws = sorted(grupos[yc], key=lambda w: w[0])
        linhas.append({
            "yc":    yc,
            "y0":    min(w[1] for w in ws),
            "y1":    max(w[3] for w in ws),
            "texto": " ".join(w[4] for w in ws),
            "words": ws,
        })
    return linhas


def encontra_blocos_tomador(page) -> list:
    """
    Percorre as linhas da página e identifica pares:
      início  = linha que contém "Tomador" + CNPJ  (sem "Total do")
      fim     = linha que contém "Total do Tomador"

    Retorna lista de dicts:
      { cnpj, y_dados_inicio, y_dados_fim }
      - y_dados_inicio: logo abaixo da linha "Tomador: XXXX"
      - y_dados_fim:    topo da linha "Total do Tomador"
    """
    linhas = linhas_da_pagina(page)
    headers = []
    footers = []

    for l in linhas:
        t = l["texto"]
        tem_cnpj = bool(CNPJ_RE.search(t))

        # Cabeçalho de bloco: "Tomador" + CNPJ, SEM "Total do"
        if "Tomador" in t and tem_cnpj and "Total do Tomador" not in t:
            m = CNPJ_RE.search(t)
            cnpj = norm_cnpj(m.group(0)) if m else ""
            headers.append({ "yc": l["yc"], "y1": l["y1"], "cnpj": cnpj })

        # Rodapé de bloco
        if "Total do Tomador" in t:
            footers.append({ "y0": l["y0"] })

    blocos = []
    for h in sorted(headers, key=lambda x: x["yc"]):
        prox_footer = [f for f in footers if f["y0"] > h["y1"]]
        y_fim = min(prox_footer, key=lambda f: f["y0"])["y0"] if prox_footer else page.rect.height
        blocos.append({
            "cnpj":           h["cnpj"],
            "y_dados_inicio": h["y1"],
            "y_dados_fim":    y_fim,
        })

    return blocos


# ─────────────────────────────────────────────────────────────────────────────
# DETECÇÃO DE LINHAS DE COLABORADORES (ancoradas no CPF)
# ─────────────────────────────────────────────────────────────────────────────

def cpfs_na_zona(page, y_ini: float, y_fim: float) -> list:
    """
    Retorna lista de { cpf_norm, y_center } para cada CPF de colaborador
    encontrado dentro da faixa Y [y_ini, y_fim].
    Ignora CPFs de cabeçalho fixo.
    """
    resultado = []
    for w in page.get_text("words"):
        yc = (w[1] + w[3]) / 2
        if not (y_ini <= yc <= y_fim):
            continue
        if CPF_RE.fullmatch(w[4]):
            cpf = norm_cpf(w[4])
            if cpf not in EXCECOES_CPF:
                resultado.append({ "cpf_norm": cpf, "y_center": yc })
    resultado.sort(key=lambda x: x["y_center"])
    return resultado


def faixas_y(linhas_cpf: list, y_zona_ini: float, y_zona_fim: float) -> list:
    """
    Para cada linha (ancorada no CPF), calcula top/bottom esticando até
    a metade da distância à linha vizinha (+ PAD_Y).
    Limita ao intervalo [y_zona_ini, y_zona_fim].
    """
    n = len(linhas_cpf)
    if n == 0:
        return []

    if n > 1:
        esps = sorted(linhas_cpf[i+1]["y_center"] - linhas_cpf[i]["y_center"]
                      for i in range(n-1))
        esp = esps[len(esps)//2]
    else:
        esp = 13.0

    faixas = []
    for i, linha in enumerate(linhas_cpf):
        yc = linha["y_center"]

        top = ((linhas_cpf[i-1]["y_center"] + yc) / 2 - PAD_Y
               if i > 0
               else max(yc - esp/2 - PAD_Y, y_zona_ini))

        bottom = ((yc + linhas_cpf[i+1]["y_center"]) / 2 + PAD_Y
                  if i < n-1
                  else min(yc + esp/2 + PAD_Y, y_zona_fim))

        faixas.append({ "top": top, "bottom": bottom, "cpf_norm": linha["cpf_norm"] })

    return faixas


# ─────────────────────────────────────────────────────────────────────────────
# TARJAMENTO
# ─────────────────────────────────────────────────────────────────────────────

def redaciona_pagina(page, funcionarios: dict, tomadores_perm: set, log) -> bool:
    """
    Processa todos os blocos Tomador da página.
    - Se Tomador não está em tomadores_perm (e tomadores_perm não está vazio):
        tarja o CNPJ do Tomador + todos os dados do bloco
    - Se Tomador está permitido:
        tarja linha a linha — só linhas cujo CPF não está em `funcionarios`
    Retorna True se a página contém pelo menos um dado permitido.
    """
    blocos = encontra_blocos_tomador(page)
    if not blocos:
        return False

    contem_permitido = False
    x_fim = page.rect.width - X_FIM_MARGEM

    for bloco in blocos:
        cnpj       = bloco["cnpj"]
        y_ini      = bloco["y_dados_inicio"]
        y_fim      = bloco["y_dados_fim"]

        # Verifica se este Tomador está na lista permitida
        tomador_ok = (not tomadores_perm) or (cnpj in tomadores_perm)

        if not tomador_ok:
            # Tarja o CNPJ na linha de cabeçalho do bloco
            for inst in page.search_for(cnpj[:2] + "." + cnpj[2:5] + "." +
                                        cnpj[5:8] + "/" + cnpj[8:12] + "-" + cnpj[12:]):
                page.add_redact_annot(inst, fill=(0, 0, 0))
            # Tarja todos os dados do bloco
            page.add_redact_annot(
                fitz.Rect(X_INICIO, y_ini, x_fim, y_fim),
                fill=(0, 0, 0)
            )
            continue

        # Tomador permitido → processa linha a linha
        linhas = cpfs_na_zona(page, y_ini, y_fim)
        if not linhas:
            continue

        for faixa in faixas_y(linhas, y_ini, y_fim):
            cpf   = faixa["cpf_norm"]
            top   = faixa["top"]
            bottom = faixa["bottom"]

            if cpf in funcionarios:
                contem_permitido = True
                continue  # mantém visível

            # CPF não permitido → tarja linha inteira
            page.add_redact_annot(
                fitz.Rect(X_INICIO, top, x_fim, bottom),
                fill=(0, 0, 0)
            )

    page.apply_redactions()
    return contem_permitido


# ─────────────────────────────────────────────────────────────────────────────
# VERIFICAÇÃO RÁPIDA DE PÁGINA
# ─────────────────────────────────────────────────────────────────────────────

def pagina_relevante(texto: str, funcionarios: dict) -> bool:
    """True se a página contém pelo menos um CPF/Nome permitido."""
    for m in CPF_RE.finditer(texto):
        if norm_cpf(m.group(0)) in funcionarios:
            return True
    texto_norm = norm(texto)
    for dados in funcionarios.values():
        if dados["nome"] and dados["nome"] in texto_norm:
            return True
    return False


# ─────────────────────────────────────────────────────────────────────────────
# PROCESSAMENTO PRINCIPAL
# ─────────────────────────────────────────────────────────────────────────────

def processa_pdf(input_pdf: str, funcionarios: dict, tomadores_perm: set,
                 saida: str, log, progresso=None) -> tuple:
    log(f"\nAbrindo PDF: {os.path.basename(input_pdf)}")
    doc = fitz.open(input_pdf)
    total = doc.page_count
    doc_saida = fitz.open()
    incluidas = 0

    for num in range(total):
        if progresso:
            progresso(num + 1, total)

        page = doc[num]
        texto = page.get_text("text")

        # Pulo rápido: página sem nenhum dado relevante
        if not pagina_relevante(texto, funcionarios):
            continue

        log(f"  → Página {num+1}: dado encontrado — processando...")
        tem_permitido = redaciona_pagina(page, funcionarios, tomadores_perm, log)

        if tem_permitido:
            doc_saida.insert_pdf(doc, from_page=num, to_page=num)
            incluidas += 1

    log(f"\n{incluidas} de {total} página(s) incluídas no resultado.")

    if incluidas > 0:
        doc_saida.save(saida, garbage=4, deflate=True)
        log(f"✅ Salvo em: {saida}")
    else:
        log("⚠️ Nenhuma página correspondeu — arquivo não gerado.")

    doc_saida.close()
    doc.close()
    return incluidas, total


# ─────────────────────────────────────────────────────────────────────────────
# INTERFACE GRÁFICA
# ─────────────────────────────────────────────────────────────────────────────

class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        root.title("Tarjador FGTS")
        root.geometry("700x580")
        root.minsize(580, 460)
        root.configure(bg="#f0f0f0")

        self._pdf     = None
        self._pasta   = None
        self._fila    = queue.Queue()

        self._build_ui()
        root.after(100, self._poll)

    # ── layout ──────────────────────────────────────────────────────────────

    def _build_ui(self):
        PAD = {"padx": 14, "pady": 6}

        # ── arquivos ────────────────────────────────────────────────────────
        frm_arq = tk.LabelFrame(self.root, text=" Arquivos ", bg="#f0f0f0",
                                font=("Segoe UI", 9, "bold"))
        frm_arq.pack(fill="x", **PAD)

        self._var_pdf   = tk.StringVar(value="(nenhum PDF selecionado)")
        self._var_pasta = tk.StringVar(value="(nenhuma pasta selecionada)")

        tk.Button(frm_arq, text="Selecionar PDF…", width=20,
                  command=self._sel_pdf).grid(row=0, column=0, sticky="w",
                                              padx=8, pady=4)
        tk.Label(frm_arq, textvariable=self._var_pdf,
                 fg="#444", bg="#f0f0f0", anchor="w").grid(
                     row=0, column=1, sticky="w")

        tk.Button(frm_arq, text="Pasta dos Excels…", width=20,
                  command=self._sel_pasta).grid(row=1, column=0, sticky="w",
                                                padx=8, pady=4)
        tk.Label(frm_arq, textvariable=self._var_pasta,
                 fg="#444", bg="#f0f0f0", anchor="w").grid(
                     row=1, column=1, sticky="w")

        # ── inputs de busca ─────────────────────────────────────────────────
        frm_inp = tk.LabelFrame(self.root,
                                text=" Busca (separe múltiplos por vírgula ou Enter) ",
                                bg="#f0f0f0", font=("Segoe UI", 9, "bold"))
        frm_inp.pack(fill="x", **PAD)

        labels = ["CPF(s):", "Nome(s):", "Tomador(es):"]
        self._entry_cpfs    = self._campo(frm_inp, "CPF(s):",       0)
        self._entry_nomes   = self._campo(frm_inp, "Nome(s):",      1)
        self._entry_tomador = self._campo(frm_inp, "Tomador(es):",  2)

        tk.Label(frm_inp, text="(Tomador vazio = todos permitidos)",
                 fg="#888", bg="#f0f0f0", font=("Segoe UI", 8)).grid(
                     row=3, column=1, sticky="w", padx=8)

        # ── botão ────────────────────────────────────────────────────────────
        self._btn = tk.Button(
            self.root, text="▶  PROCESSAR",
            command=self._iniciar,
            bg="#1565c0", fg="white",
            font=("Segoe UI", 11, "bold"),
            activebackground="#0d47a1",
            relief="flat", padx=24, pady=8,
        )
        self._btn.pack(pady=(2, 6))

        # ── progresso ────────────────────────────────────────────────────────
        self._prog = ttk.Progressbar(self.root, orient="horizontal",
                                     mode="determinate")
        self._prog.pack(fill="x", padx=14, pady=(0, 6))

        # ── log ──────────────────────────────────────────────────────────────
        self._log = scrolledtext.ScrolledText(
            self.root, height=14, state="disabled",
            bg="#1e1e1e", fg="#d4d4d4",
            font=("Consolas", 9), relief="flat",
        )
        self._log.pack(fill="both", expand=True, padx=14, pady=(0, 14))

    def _campo(self, parent, label: str, row: int) -> tk.Text:
        tk.Label(parent, text=label, bg="#f0f0f0",
                 width=14, anchor="e").grid(row=row, column=0,
                                            sticky="e", padx=(8, 4), pady=3)
        entry = tk.Text(parent, height=2, width=52,
                        font=("Segoe UI", 9), relief="solid", bd=1)
        entry.grid(row=row, column=1, sticky="w", padx=(0, 8), pady=3)
        return entry

    # ── seletores ───────────────────────────────────────────────────────────

    def _sel_pdf(self):
        p = filedialog.askopenfilename(
            title="Selecione o PDF", filetypes=[("PDF", "*.pdf")])
        if p:
            self._pdf = p
            self._var_pdf.set(os.path.basename(p))

    def _sel_pasta(self):
        p = filedialog.askdirectory(title="Selecione a pasta dos arquivos Excel")
        if p:
            self._pasta = p
            self._var_pasta.set(p)

    # ── helpers de input ────────────────────────────────────────────────────

    def _lê_lista(self, widget: tk.Text) -> list:
        raw = widget.get("1.0", "end").strip()
        itens = [i.strip() for i in re.split(r"[,\n;]+", raw) if i.strip()]
        return itens

    # ── log / progresso thread-safe ─────────────────────────────────────────

    def _log_msg(self, msg: str):
        self._fila.put(("log", msg))

    def _prog_update(self, atual: int, total: int):
        self._fila.put(("prog", (atual, total)))

    def _poll(self):
        try:
            while True:
                tipo, dado = self._fila.get_nowait()
                if tipo == "log":
                    self._log.config(state="normal")
                    self._log.insert("end", dado + "\n")
                    self._log.see("end")
                    self._log.config(state="disabled")
                elif tipo == "prog":
                    a, t = dado
                    self._prog["maximum"] = t
                    self._prog["value"] = a
                elif tipo == "fim":
                    self._btn.config(state="normal")
                    if dado:
                        messagebox.showinfo("Concluído", dado)
        except queue.Empty:
            pass
        self.root.after(100, self._poll)

    # ── processamento ────────────────────────────────────────────────────────

    def _iniciar(self):
        if not self._pdf:
            messagebox.showwarning("Atenção", "Selecione um arquivo PDF.")
            return
        if not self._pasta:
            messagebox.showwarning("Atenção", "Selecione a pasta com os Excels.")
            return

        cpfs    = self._lê_lista(self._entry_cpfs)
        nomes   = self._lê_lista(self._entry_nomes)
        tomads  = self._lê_lista(self._entry_tomador)

        if not cpfs and not nomes:
            messagebox.showwarning("Atenção",
                "Informe pelo menos um CPF ou Nome para buscar.")
            return

        self._btn.config(state="disabled")
        self._log.config(state="normal")
        self._log.delete("1.0", "end")
        self._log.config(state="disabled")
        self._prog["value"] = 0

        threading.Thread(
            target=self._run,
            args=(cpfs, nomes, tomads),
            daemon=True,
        ).start()

    def _run(self, cpfs, nomes, tomads):
        try:
            log = self._log_msg

            # 1. Carrega todos os Excels
            log("Carregando arquivos Excel da pasta...")
            df = carrega_excels(self._pasta, log)
            log(f"Total de registros carregados: {len(df)}")

            # 2. Busca funcionários
            log("\nBuscando funcionários...")
            funcionarios = busca_funcionarios(df, cpfs, nomes, log)

            if not funcionarios:
                log("\n❌ Nenhum funcionário encontrado nos Excels. "
                    "Verifique os CPFs/Nomes informados.")
                self._fila.put(("fim", None))
                return

            log(f"\n{len(funcionarios)} funcionário(s) encontrado(s).")

            # 3. Tomadores permitidos (vazio = todos)
            tomadores_perm = { norm_cnpj(t) for t in tomads if norm_cnpj(t) }
            if tomadores_perm:
                log(f"Tomadores permitidos: {tomadores_perm}")
            else:
                log("Tomadores: todos permitidos (campo vazio).")

            # 4. Processa o PDF
            base, _ = os.path.splitext(self._pdf)
            saida = base + "_filtrado.pdf"

            inc, total = processa_pdf(
                self._pdf, funcionarios, tomadores_perm,
                saida, log, self._prog_update,
            )

            if inc > 0:
                self._fila.put(("fim",
                    f"{inc} de {total} página(s) salvas em:\n{saida}"))
            else:
                self._fila.put(("fim",
                    "Nenhuma página correspondeu aos critérios."))

        except Exception as e:
            self._log_msg(f"\n❌ Erro: {e}")
            import traceback
            self._log_msg(traceback.format_exc())
            self._fila.put(("fim", None))


# ─────────────────────────────────────────────────────────────────────────────
# ENTRADA
# ─────────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    root = tk.Tk()
    App(root)
    root.mainloop()
