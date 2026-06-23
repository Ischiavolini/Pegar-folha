# -*- coding: utf-8 -*-
"""
Buscador de Arquivos por CNPJ
==============================

O que esse programa faz:
1. Você escolhe uma planilha Excel com as colunas TOMADOR e CNPJ.
2. Você escolhe a pasta "raiz" onde o programa vai procurar (a pasta que tem
   uma subpasta para cada CNPJ).
3. Você escolhe a pasta de destino, onde os arquivos encontrados vão ser
   copiados.
4. O programa limpa o CNPJ (tira ponto, traço e barra), entra na pasta raiz,
   procura uma pasta com esse número (em qualquer nível, não precisa ser
   direto na raiz), e dentro dela procura um arquivo cujo nome contenha
   "CNPJ_FOLHA".
5. Quando encontra, copia esse arquivo para a pasta de destino, renomeado
   com o número do TOMADOR (mantendo a extensão original, ex: 1212.pdf).
6. Tomadores cujo CNPJ não foi encontrado em nenhuma pasta são listados em um
   arquivo "nao_encontrados.txt" na pasta de destino.
7. Se o mesmo CNPJ aparecer em mais de uma pasta, o programa usa a primeira
   que encontrar e anota o caso em "duplicados.txt" na pasta de destino,
   para você conferir manualmente depois.

Como gerar o .exe (opcional, ver instruções no final deste arquivo ou no
arquivo LEIA-ME.txt que acompanha o programa).
"""

import os
import re
import shutil
import threading
import traceback
from datetime import datetime

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

try:
    import openpyxl
except ImportError:
    openpyxl = None


# ----------------------------------------------------------------------
# Funções auxiliares (lógica pura, sem interface) - testáveis isoladamente
# ----------------------------------------------------------------------

def limpar_documento(valor):
    """Remove tudo que não for dígito de um CNPJ/CPF (pontos, traços, barras, espaços)."""
    if valor is None:
        return ""
    return re.sub(r"\D", "", str(valor))


def ler_planilha(caminho_excel):
    """
    Lê a planilha e retorna uma lista de tuplas (tomador, cnpj_limpo, cnpj_original).
    Procura as colunas pelo nome (TOMADOR e CNPJ), sem diferenciar maiúsculas/
    minúsculas e ignorando espaços nas bordas. Lança ValueError se não achar
    as colunas.
    """
    if openpyxl is None:
        raise RuntimeError(
            "A biblioteca 'openpyxl' não está instalada. "
            "Abra o Prompt de Comando e rode: pip install openpyxl"
        )

    wb = openpyxl.load_workbook(caminho_excel, data_only=True)
    sheet = wb.active

    linhas = list(sheet.iter_rows(values_only=True))
    if not linhas:
        raise ValueError("A planilha está vazia.")

    cabecalho = [str(c).strip().upper() if c is not None else "" for c in linhas[0]]

    try:
        idx_tomador = cabecalho.index("TOMADOR")
    except ValueError:
        raise ValueError("Não encontrei a coluna 'TOMADOR' na primeira linha da planilha.")

    try:
        idx_cnpj = cabecalho.index("CNPJ")
    except ValueError:
        raise ValueError("Não encontrei a coluna 'CNPJ' na primeira linha da planilha.")

    resultado = []
    for linha in linhas[1:]:
        if linha is None:
            continue
        if idx_tomador >= len(linha) or idx_cnpj >= len(linha):
            continue

        tomador_bruto = linha[idx_tomador]
        cnpj_bruto = linha[idx_cnpj]

        if tomador_bruto is None and cnpj_bruto is None:
            continue

        tomador = str(tomador_bruto).strip() if tomador_bruto is not None else ""
        cnpj_original = str(cnpj_bruto).strip() if cnpj_bruto is not None else ""
        cnpj_limpo = limpar_documento(cnpj_bruto)

        if not tomador and not cnpj_limpo:
            continue

        resultado.append((tomador, cnpj_limpo, cnpj_original))

    return resultado


def encontrar_pastas_do_cnpj(pasta_raiz, cnpj_limpo):
    """
    Procura, em qualquer nível dentro de pasta_raiz, pastas cujo nome
    (limpo de pontuação) seja igual ao cnpj_limpo. Retorna uma lista de
    caminhos "de topo" (ignora uma pasta-com-o-cnpj que esteja DENTRO de
    outra pasta-com-o-mesmo-cnpj já encontrada, já que isso é só uma pasta
    aninhada repetindo o número, não uma duplicidade real em locais
    diferentes).
    """
    todos_encontrados = []
    for atual, subpastas, _arquivos in os.walk(pasta_raiz):
        for nome_pasta in subpastas:
            if limpar_documento(nome_pasta) == cnpj_limpo:
                todos_encontrados.append(os.path.join(atual, nome_pasta))

    # Remove caminhos que estão dentro de outro caminho já encontrado
    todos_encontrados.sort(key=len)
    de_topo = []
    for caminho in todos_encontrados:
        caminho_norm = os.path.normpath(caminho)
        dentro_de_outro = any(
            caminho_norm.startswith(os.path.normpath(pai) + os.sep)
            for pai in de_topo
        )
        if not dentro_de_outro:
            de_topo.append(caminho)

    return de_topo


def encontrar_subpasta_repetida(pasta_cnpj, cnpj_limpo):
    """
    Procura, dentro de pasta_cnpj (em qualquer nível), uma subpasta cujo
    nome limpo seja igual ao cnpj_limpo - ou seja, a pasta "repetida" que
    normalmente existe dentro da pasta do CNPJ. Retorna o caminho dessa
    subpasta mais rasa encontrada, ou None se não existir nenhuma.
    """
    candidatas = []
    for atual, subpastas, _arquivos in os.walk(pasta_cnpj):
        for nome_pasta in subpastas:
            if limpar_documento(nome_pasta) == cnpj_limpo:
                candidatas.append(os.path.join(atual, nome_pasta))
    if not candidatas:
        return None
    # Pega a mais rasa (caminho mais curto), que é a primeira repetição
    candidatas.sort(key=len)
    return candidatas[0]


def encontrar_arquivo_cnpj_folha(pasta_cnpj, cnpj_limpo, log_callback=None):
    """
    Procura o arquivo de folha seguindo a regra:
    - Sempre existe, dentro da pasta do CNPJ, uma subpasta repetida com o
      mesmo nome/número. A busca pelo arquivo deve acontecer SÓ dentro
      dessa subpasta, ignorando qualquer arquivo solto fora dela (que pode
      ser um PDF de outro assunto).
    - Caso, excepcionalmente, essa subpasta não exista, cai para uma busca
      em toda a pasta do CNPJ como rede de segurança (e avisa no log).
    O nome do arquivo precisa conter "FOLHA" (sem diferenciar maiúsculas/
    minúsculas); o restante do nome (número do CNPJ, datas, etc.) pode
    variar livremente.
    Retorna o caminho completo do arquivo encontrado, ou None.
    """
    alvo = "folha"
    subpasta = encontrar_subpasta_repetida(pasta_cnpj, cnpj_limpo)

    if subpasta:
        for atual, _subpastas, arquivos in os.walk(subpasta):
            for nome_arquivo in arquivos:
                nome_sem_ext, _ext = os.path.splitext(nome_arquivo)
                if alvo in nome_sem_ext.lower():
                    return os.path.join(atual, nome_arquivo)
        # Subpasta existe mas não tem o arquivo dentro - não procura fora dela
        return None

    # Sem subpasta repetida: busca de segurança em toda a pasta do CNPJ
    if log_callback:
        log_callback(f"    Aviso: não encontrei a subpasta repetida dentro de "
                     f"'{pasta_cnpj}'. Procurando em toda a pasta como alternativa.")
    for atual, _subpastas, arquivos in os.walk(pasta_cnpj):
        for nome_arquivo in arquivos:
            nome_sem_ext, _ext = os.path.splitext(nome_arquivo)
            if alvo in nome_sem_ext.lower():
                return os.path.join(atual, nome_arquivo)
    return None


def processar(caminho_excel, pasta_busca, pasta_destino, log_callback, progresso_callback):
    """
    Executa o processo completo. log_callback(str) é chamado para cada linha
    de log. progresso_callback(atual, total) é chamado a cada item processado.
    Retorna um dicionário com estatísticas finais.
    """
    os.makedirs(pasta_destino, exist_ok=True)

    log_callback("Lendo planilha...")
    linhas = ler_planilha(caminho_excel)
    total = len(linhas)
    log_callback(f"{total} linha(s) encontrada(s) na planilha.\n")

    try:
        itens_raiz = sorted(os.listdir(pasta_busca))
        pastas_raiz = [n for n in itens_raiz if os.path.isdir(os.path.join(pasta_busca, n))]
        log_callback(f"Pasta de busca: {pasta_busca}")
        log_callback(f"Pastas encontradas direto na raiz ({len(pastas_raiz)}): "
                     f"{', '.join(pastas_raiz[:30])}"
                     f"{' ...' if len(pastas_raiz) > 30 else ''}\n")
    except Exception as e:
        log_callback(f"Aviso: não consegui listar a pasta de busca para diagnóstico ({e}).\n")

    nao_encontrados = []
    duplicados = []
    copiados = 0
    erros = []

    for i, (tomador, cnpj_limpo, cnpj_original) in enumerate(linhas, start=1):
        progresso_callback(i, total)

        if not cnpj_limpo:
            log_callback(f"[{i}/{total}] Tomador '{tomador}': CNPJ vazio/ inválido na planilha. Pulando.")
            nao_encontrados.append((tomador, cnpj_original or "(vazio)"))
            continue

        log_callback(f"[{i}/{total}] Procurando CNPJ original '{cnpj_original}' -> limpo '{cnpj_limpo}'...")

        pastas = encontrar_pastas_do_cnpj(pasta_busca, cnpj_limpo)

        if not pastas:
            log_callback(f"[{i}/{total}] Tomador '{tomador}' (CNPJ {cnpj_original}): "
                         f"pasta não encontrada.")
            nao_encontrados.append((tomador, cnpj_original))
            continue

        if len(pastas) > 1:
            log_callback(f"[{i}/{total}] Tomador '{tomador}' (CNPJ {cnpj_original}): "
                         f"ATENÇÃO - {len(pastas)} pastas com esse CNPJ encontradas. "
                         f"Usando a primeira.")
            duplicados.append((tomador, cnpj_original, pastas))

        pasta_escolhida = pastas[0]
        arquivo_origem = encontrar_arquivo_cnpj_folha(pasta_escolhida, cnpj_limpo, log_callback)

        if not arquivo_origem:
            log_callback(f"[{i}/{total}] Tomador '{tomador}' (CNPJ {cnpj_original}): "
                         f"pasta encontrada, mas nenhum arquivo 'CNPJ_FOLHA' dentro dela.")
            nao_encontrados.append((tomador, cnpj_original))
            continue

        _nome_sem_ext, ext = os.path.splitext(arquivo_origem)
        nome_tomador_seguro = re.sub(r'[\\/:*?"<>|]', "_", str(tomador)) or "SEM_TOMADOR"
        nome_destino = f"{nome_tomador_seguro}{ext}"
        caminho_destino = os.path.join(pasta_destino, nome_destino)

        # Evita sobrescrever silenciosamente se já existir um tomador com mesmo nome
        if os.path.exists(caminho_destino):
            base, ext2 = os.path.splitext(nome_destino)
            contador = 2
            while os.path.exists(caminho_destino):
                caminho_destino = os.path.join(pasta_destino, f"{base}_{contador}{ext2}")
                contador += 1

        try:
            shutil.copy2(arquivo_origem, caminho_destino)
            copiados += 1
            log_callback(f"[{i}/{total}] Tomador '{tomador}' (CNPJ {cnpj_original}): "
                         f"OK -> {os.path.basename(caminho_destino)}")
        except Exception as e:
            erros.append((tomador, cnpj_original, str(e)))
            log_callback(f"[{i}/{total}] Tomador '{tomador}' (CNPJ {cnpj_original}): "
                         f"ERRO ao copiar - {e}")

    # Grava arquivo de não encontrados
    if nao_encontrados:
        caminho_txt = os.path.join(pasta_destino, "nao_encontrados.txt")
        with open(caminho_txt, "w", encoding="utf-8") as f:
            f.write(f"Tomadores/CNPJs não encontrados - gerado em "
                    f"{datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n")
            f.write("=" * 60 + "\n\n")
            for tomador, cnpj in nao_encontrados:
                f.write(f"TOMADOR: {tomador}\tCNPJ: {cnpj}\n")
        log_callback(f"\nArquivo 'nao_encontrados.txt' gerado com {len(nao_encontrados)} item(ns).")

    # Grava arquivo de duplicados, se houver
    if duplicados:
        caminho_txt = os.path.join(pasta_destino, "duplicados.txt")
        with open(caminho_txt, "w", encoding="utf-8") as f:
            f.write(f"CNPJs com mais de uma pasta encontrada - gerado em "
                    f"{datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n")
            f.write("Foi usada sempre a PRIMEIRA pasta da lista abaixo. Confira manualmente.\n")
            f.write("=" * 60 + "\n\n")
            for tomador, cnpj, pastas in duplicados:
                f.write(f"TOMADOR: {tomador}\tCNPJ: {cnpj}\n")
                for p in pastas:
                    f.write(f"    - {p}\n")
                f.write("\n")
        log_callback(f"Arquivo 'duplicados.txt' gerado com {len(duplicados)} caso(s).")

    # Grava arquivo de erros de cópia, se houver
    if erros:
        caminho_txt = os.path.join(pasta_destino, "erros_copia.txt")
        with open(caminho_txt, "w", encoding="utf-8") as f:
            f.write(f"Erros ao copiar arquivos - gerado em "
                    f"{datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n")
            f.write("=" * 60 + "\n\n")
            for tomador, cnpj, erro in erros:
                f.write(f"TOMADOR: {tomador}\tCNPJ: {cnpj}\tERRO: {erro}\n")
        log_callback(f"Arquivo 'erros_copia.txt' gerado com {len(erros)} erro(s).")

    return {
        "total": total,
        "copiados": copiados,
        "nao_encontrados": len(nao_encontrados),
        "duplicados": len(duplicados),
        "erros": len(erros),
    }


# ----------------------------------------------------------------------
# Interface gráfica
# ----------------------------------------------------------------------

class App:
    def __init__(self, root):
        self.root = root
        root.title("Buscador de Arquivos por CNPJ")
        root.geometry("720x560")
        root.minsize(680, 520)

        self.caminho_excel = tk.StringVar()
        self.pasta_busca = tk.StringVar()
        self.pasta_destino = tk.StringVar()

        self._montar_interface()

    def _montar_interface(self):
        pad = {"padx": 10, "pady": 6}

        frame_topo = ttk.Frame(self.root)
        frame_topo.pack(fill="x", **pad)

        ttk.Label(
            frame_topo,
            text="Buscador de Arquivos por CNPJ",
            font=("Segoe UI", 14, "bold"),
        ).pack(anchor="w")
        ttk.Label(
            frame_topo,
            text="Selecione a planilha e as pastas abaixo e clique em Iniciar.",
        ).pack(anchor="w")

        # --- Planilha Excel ---
        frame1 = ttk.LabelFrame(self.root, text="1. Planilha Excel (colunas TOMADOR e CNPJ)")
        frame1.pack(fill="x", **pad)
        self._linha_selecao(frame1, self.caminho_excel, self._escolher_excel)

        # --- Pasta de busca ---
        frame2 = ttk.LabelFrame(self.root, text="2. Pasta onde estão as pastas dos CNPJs")
        frame2.pack(fill="x", **pad)
        self._linha_selecao(frame2, self.pasta_busca, self._escolher_pasta_busca)

        # --- Pasta de destino ---
        frame3 = ttk.LabelFrame(self.root, text="3. Pasta de destino (onde os arquivos serão copiados)")
        frame3.pack(fill="x", **pad)
        self._linha_selecao(frame3, self.pasta_destino, self._escolher_pasta_destino)

        # --- Botão iniciar + barra de progresso ---
        frame_acao = ttk.Frame(self.root)
        frame_acao.pack(fill="x", **pad)

        self.botao_iniciar = ttk.Button(frame_acao, text="Iniciar", command=self._iniciar)
        self.botao_iniciar.pack(side="left")

        self.barra_progresso = ttk.Progressbar(frame_acao, mode="determinate")
        self.barra_progresso.pack(side="left", fill="x", expand=True, padx=10)

        self.label_status = ttk.Label(frame_acao, text="")
        self.label_status.pack(side="left")

        # --- Log ---
        frame_log = ttk.LabelFrame(self.root, text="Andamento")
        frame_log.pack(fill="both", expand=True, **pad)

        self.texto_log = tk.Text(frame_log, wrap="word", state="disabled", height=15)
        self.texto_log.pack(side="left", fill="both", expand=True)

        scrollbar = ttk.Scrollbar(frame_log, command=self.texto_log.yview)
        scrollbar.pack(side="right", fill="y")
        self.texto_log.configure(yscrollcommand=scrollbar.set)

    def _linha_selecao(self, parent, variavel, comando):
        frame = ttk.Frame(parent)
        frame.pack(fill="x", padx=8, pady=6)
        entrada = ttk.Entry(frame, textvariable=variavel)
        entrada.pack(side="left", fill="x", expand=True)
        ttk.Button(frame, text="Escolher...", command=comando).pack(side="left", padx=6)

    def _escolher_excel(self):
        caminho = filedialog.askopenfilename(
            title="Selecione a planilha Excel",
            filetypes=[("Excel", "*.xlsx *.xlsm"), ("Todos os arquivos", "*.*")],
        )
        if caminho:
            self.caminho_excel.set(caminho)

    def _escolher_pasta_busca(self):
        pasta = filedialog.askdirectory(title="Selecione a pasta onde estão as pastas dos CNPJs")
        if pasta:
            self.pasta_busca.set(pasta)

    def _escolher_pasta_destino(self):
        pasta = filedialog.askdirectory(title="Selecione a pasta de destino")
        if pasta:
            self.pasta_destino.set(pasta)

    def _log(self, mensagem):
        self.texto_log.configure(state="normal")
        self.texto_log.insert("end", mensagem + "\n")
        self.texto_log.see("end")
        self.texto_log.configure(state="disabled")

    def _progresso(self, atual, total):
        self.barra_progresso["maximum"] = total
        self.barra_progresso["value"] = atual
        self.label_status.configure(text=f"{atual}/{total}")

    def _iniciar(self):
        excel = self.caminho_excel.get().strip()
        busca = self.pasta_busca.get().strip()
        destino = self.pasta_destino.get().strip()

        if not excel or not os.path.isfile(excel):
            messagebox.showerror("Erro", "Selecione um arquivo Excel válido.")
            return
        if not busca or not os.path.isdir(busca):
            messagebox.showerror("Erro", "Selecione uma pasta de busca válida.")
            return
        if not destino:
            messagebox.showerror("Erro", "Selecione a pasta de destino.")
            return

        self.botao_iniciar.configure(state="disabled")
        self.texto_log.configure(state="normal")
        self.texto_log.delete("1.0", "end")
        self.texto_log.configure(state="disabled")
        self.barra_progresso["value"] = 0

        thread = threading.Thread(
            target=self._executar_em_thread, args=(excel, busca, destino), daemon=True
        )
        thread.start()

    def _executar_em_thread(self, excel, busca, destino):
        try:
            resultado = processar(
                excel, busca, destino,
                log_callback=lambda msg: self.root.after(0, self._log, msg),
                progresso_callback=lambda a, t: self.root.after(0, self._progresso, a, t),
            )
            self.root.after(0, self._finalizar_sucesso, resultado, destino)
        except Exception as e:
            erro_completo = traceback.format_exc()
            self.root.after(0, self._finalizar_erro, str(e), erro_completo)

    def _finalizar_sucesso(self, resultado, destino):
        self.botao_iniciar.configure(state="normal")
        resumo = (
            f"\nConcluído!\n"
            f"Total de linhas: {resultado['total']}\n"
            f"Copiados com sucesso: {resultado['copiados']}\n"
            f"Não encontrados: {resultado['nao_encontrados']}\n"
            f"Duplicados (verificar): {resultado['duplicados']}\n"
            f"Erros de cópia: {resultado['erros']}\n"
        )
        self._log(resumo)
        messagebox.showinfo(
            "Concluído",
            f"Processo finalizado.\n\n"
            f"Copiados: {resultado['copiados']}\n"
            f"Não encontrados: {resultado['nao_encontrados']}\n"
            f"Duplicados: {resultado['duplicados']}\n"
            f"Erros: {resultado['erros']}\n\n"
            f"Veja os detalhes na pasta:\n{destino}",
        )

    def _finalizar_erro(self, mensagem, detalhe):
        self.botao_iniciar.configure(state="normal")
        self._log(f"\nERRO: {mensagem}")
        messagebox.showerror("Erro", f"Ocorreu um erro:\n\n{mensagem}")
        print(detalhe)


def main():
    root = tk.Tk()
    try:
        style = ttk.Style()
        if "vista" in style.theme_names():
            style.theme_use("vista")
    except Exception:
        pass
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
