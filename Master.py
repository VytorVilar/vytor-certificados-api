# coloque isto no topo do arquivo, antes de qualquer import que dispare o warning
import warnings

# suprime somente os UserWarning que contenham essa mensagem
warnings.filterwarnings(
    "ignore",
    message=r"pkg_resources is deprecated as an API",
    category=UserWarning
)

# opção alternativa: suprime todos os UserWarning vindos do módulo docxcompose.properties
# warnings.filterwarnings("ignore", category=UserWarning, module=r"docxcompose\.properties")

import os
import re
import time
import unicodedata
import warnings
import threading
import logging
import json
import traceback
from datetime import datetime
from pathlib import Path
from typing import Dict, Any, List, Tuple

warnings.filterwarnings("ignore", category=UserWarning, module="pkg_resources")

import pandas as pd
import chardet
from docxtpl import DocxTemplate
from docx2pdf import convert
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from colorama import Fore, init, Style

# =============================
# Inicialização e Configurações
# =============================
init(autoreset=True)

CONFIG_FILE = "config.json"
DEFAULT_CONFIG = {
    "senha_acesso": "1403",
    "diretorio_certificados": "certificados",
    "log_erros": "erros_log.csv",
    "csv_default": "CSV",
    "modelos_default": "MODELOS"
}

def load_config():
    if not Path(CONFIG_FILE).exists():
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(DEFAULT_CONFIG, f, indent=4)
        return DEFAULT_CONFIG
    with open(CONFIG_FILE, "r", encoding="utf-8") as f:
        return json.load(f)

config = load_config()

# Configuração de logs
Path("logs").mkdir(exist_ok=True)
logging.basicConfig(
    filename="logs/app.log",
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s"
)

SENHA_ACESSO = config["senha_acesso"]
DIR_CERTIFICADOS = Path(config["diretorio_certificados"])
LOG_ERROS = Path(config["log_erros"])
CPU_WORKERS = max(1, (os.cpu_count() or 2) - 1)
ATUALIZA_GUI_CADA = 5
CANCELAR_PROCESSO = False


# =============================
# Funções Auxiliares (Core)
# =============================

def limpar_nome(texto: str, max_len: int = 100) -> str:
    """Remove acentos, caracteres inválidos e limita o tamanho do nome do arquivo."""
    if not isinstance(texto, str):
        texto = str(texto) if texto is not None else ""
    texto = unicodedata.normalize('NFD', texto)
    texto = ''.join(c for c in texto if unicodedata.category(c) != 'Mn')
    texto = re.sub(r"[^\w\s-]", "", texto)
    texto = re.sub(r"\s+", "_", texto.strip())
    return texto[:max_len]

def detectar_encoding(caminho_arquivo: Path) -> str:
    with open(caminho_arquivo, 'rb') as f:
        return chardet.detect(f.read(10000)).get('encoding') or 'utf-8'

def gerar_nome_unico(path: Path) -> Path:
    if not path.exists():
        return path
    base, ext = path.stem, path.suffix
    i = 1
    while True:
        novo = path.with_name(f"{base}_{i}{ext}")
        if not novo.exists():
            return novo
        i += 1

def formatar_data(valor):
    if pd.isna(valor): return ""
    if isinstance(valor, datetime): return valor.strftime("%d/%m/%Y")
    try:
        return pd.to_datetime(valor).strftime("%d/%m/%Y")
    except:
        return str(valor)


# =============================
# Certificados - Worker
# =============================

def _processar_um_certificado(args: Tuple[Dict[str, Any], str, str, str, str]) -> Tuple[bool, Dict[str, Any]]:
    row_dict, modelo_path, tipo_saida, data_hoje, dir_certificados = args
    try:
        for c in ['nome', 'nome_empresa', 'curso', 'horas', 'data_certificado']:
            if pd.isna(row_dict.get(c)) or str(row_dict.get(c)).strip() == '':
                row_dict[c] = ""

        contexto = {k: row_dict.get(k, "") for k in
                    ['nome', 'nome_empresa', 'curso', 'horas', 'data_certificado', 'cpf', 'cnpj', 'cidade']}

        modelo_nome = Path(modelo_path).stem
        modelo_nome_limpo = limpar_nome(modelo_nome)
        nome_limpo = limpar_nome(contexto['nome'])
        empresa_limpa = limpar_nome(contexto['nome_empresa'])

        pasta_saida = Path(dir_certificados) / modelo_nome_limpo / empresa_limpa
        pasta_saida.mkdir(parents=True, exist_ok=True)

        docx_path = gerar_nome_unico(pasta_saida / f"{nome_limpo}_{data_hoje}.docx")
        pdf_path = docx_path.with_suffix(".pdf")

        doc = DocxTemplate(modelo_path)
        doc.render(contexto)
        if tipo_saida in ("DOCX", "Ambos", "PDF"):
            doc.save(docx_path)

        if tipo_saida in ("PDF", "Ambos"):
            convert(str(docx_path), str(pdf_path))

        return True, {"nome": contexto['nome'], "empresa": contexto['nome_empresa'], "status": "OK"}

    except Exception as e:
        logging.error(f"Erro ao gerar certificado: {traceback.format_exc()}")
        return False, {'erro': str(e), 'registro': row_dict, 'modelo': modelo_path}


# =============================
# Certificados - Geração
# =============================

def gerar_certificados_mp(csv_path: Path, modelos_paths: List[Path], tipo_saida: str,
                          barra_progresso: ttk.Progressbar = None, label_status: tk.Label = None):
    from multiprocessing import Pool
    global CANCELAR_PROCESSO
    CANCELAR_PROCESSO = False

    try:
        encoding_detectado = detectar_encoding(csv_path)
        df = pd.read_csv(csv_path, encoding=encoding_detectado)
        df.columns = df.columns.str.strip()
        for col in ['cpf', 'cnpj', 'cidade']:
            if col not in df.columns:
                df[col] = ''
        if "nome" not in df.columns or "nome_empresa" not in df.columns:
            messagebox.showerror("Erro", "CSV inválido: faltam colunas obrigatórias.")
            return
        data_hoje = datetime.today().strftime("%Y-%m-%d")
    except Exception as e:
        messagebox.showerror("Erro", f"Falha ao ler CSV: {e}")
        return

    registros = df.to_dict(orient='records')
    tarefas = [(row, str(modelo), tipo_saida, data_hoje, str(DIR_CERTIFICADOS))
               for modelo in modelos_paths for row in registros]

    total = len(tarefas)
    if total == 0:
        messagebox.showinfo("Aviso", "Nenhuma tarefa para processar.")
        return

    concluidos, erros, resultados = 0, [], []
    start = time.time()

    def _tick_gui(nome_atual: str = ""):
        nonlocal concluidos
        if barra_progresso is not None:
            barra_progresso["value"] = concluidos / total * 100
        if label_status is not None:
            elapsed = time.time() - start
            estimado = (elapsed / concluidos * (total - concluidos)) if concluidos > 0 else 0
            label_status.config(text=f"{concluidos}/{total} {nome_atual} | Restante: {estimado:.1f}s")

    try:
        with Pool(processes=CPU_WORKERS) as pool:
            for ok, info in pool.imap_unordered(_processar_um_certificado, tarefas, chunksize=10):
                if CANCELAR_PROCESSO:
                    pool.terminate()
                    messagebox.showinfo("Cancelado", "Processamento interrompido.")
                    return

                concluidos += 1
                if ok:
                    resultados.append(info)
                else:
                    erros.append(info)

                if concluidos % ATUALIZA_GUI_CADA == 0 or concluidos == total:
                    _tick_gui()
    except Exception as e:
        logging.error(f"Erro na geração: {traceback.format_exc()}")
        messagebox.showerror("Erro", f"Falha: {e}")
        return

    if barra_progresso: barra_progresso["value"] = 100
    if label_status: label_status.config(text=f"Finalizado: {concluidos}/{total} | Erros: {len(erros)}")

    # Relatório consolidado
    df_result = pd.DataFrame(resultados + erros)
    df_result.to_excel("relatorio_certificados.xlsx", index=False)
    df_result.to_csv("relatorio_certificados.csv", index=False, sep=";")

    messagebox.showinfo("Concluído", f"Certificados: {concluidos}/{total}\nErros: {len(erros)}")


# =============================
# PPP - Interface
# =============================

def gerar_ppp():
    caminho_planilha = filedialog.askopenfilename(title="Planilha PPP", filetypes=[("Excel", "*.xlsx")])
    caminho_modelo = filedialog.askopenfilename(title="Modelo PPP", filetypes=[("Word", "*.docx")])
    pasta_saida = filedialog.askdirectory(title="Pasta de saída")
    if not (caminho_planilha and caminho_modelo and pasta_saida): return

    df = pd.read_excel(caminho_planilha)
    df.columns = df.columns.str.strip()
    colunas_obrigatorias = ["nome_do_trabalhador", "nome_empresa", "cpf_colaborador", "data_de_emissao"]
    faltando = [c for c in colunas_obrigatorias if c not in df.columns]
    if faltando:
        messagebox.showerror("Erro", f"Colunas ausentes: {faltando}")
        return

    log_resultado = []
    for _, row in df.iterrows():
        doc = DocxTemplate(caminho_modelo)
        nome = str(row.get("nome_do_trabalhador", "")).strip()
        nome_safe = limpar_nome(nome)
        empresa = limpar_nome(str(row.get("nome_empresa", "Empresa")).strip())
        pasta_empresa = os.path.join(pasta_saida, empresa); os.makedirs(pasta_empresa, exist_ok=True)
        data_emissao = formatar_data(row.get("data_de_emissao", "")).replace("/", "-")
        caminho_salvo = os.path.join(pasta_empresa, f"{nome_safe}_{data_emissao}_PPP.docx")

        contexto = {col: row.get(col, "") for col in df.columns}
        try:
            doc.render(contexto); doc.save(caminho_salvo)
            log_resultado.append({"Trabalhador": nome, "Status": "OK"})
        except Exception as e:
            logging.error(f"Erro ao gerar PPP: {traceback.format_exc()}")
            log_resultado.append({"Trabalhador": nome, "Status": f"Erro: {e}"})

    log_df = pd.DataFrame(log_resultado)
    log_df.to_excel(os.path.join(pasta_saida, "log_ppp.xlsx"), index=False)
    log_df.to_csv(os.path.join(pasta_saida, "log_ppp.csv"), index=False, sep=";")
    messagebox.showinfo("Finalizado", "PPPs gerados com sucesso.")


# =============================
# Interfaces Tkinter
# =============================

def interface_certificados():
    def selecionar_csv():
        caminho = filedialog.askopenfilename(
            title="CSV", filetypes=[("CSV", "*.csv")],
            initialdir=config["csv_default"]
        )
        if caminho:
            entry_csv.delete(0, tk.END)
            entry_csv.insert(0, caminho)

    def selecionar_modelos():
        caminhos = filedialog.askopenfilenames(
            title="Modelos", filetypes=[("Word", "*.docx")],
            initialdir=config["modelos_default"]
        )
        if caminhos:
            entry_modelos.delete(0, tk.END)
            entry_modelos.insert(0, ";".join(caminhos))

    def iniciar_geracao_thread():
        t = threading.Thread(target=iniciar_geracao, daemon=True)
        t.start()

    def iniciar_geracao():
        csv_path = Path(entry_csv.get().strip())
        modelos_paths = [Path(p.strip()) for p in entry_modelos.get().split(';') if p.strip()]
        tipo_saida = opcao_saida.get()
        if not csv_path.exists():
            messagebox.showerror("Erro", "CSV inválido.")
            return
        if not modelos_paths:
            messagebox.showerror("Erro", "Selecione modelo.")
            return
        btn_gerar.config(state=tk.DISABLED)
        try:
            gerar_certificados_mp(csv_path, modelos_paths, tipo_saida, barra_progresso, label_status)
        finally:
            btn_gerar.config(state=tk.NORMAL)

    def cancelar():
        global CANCELAR_PROCESSO
        CANCELAR_PROCESSO = True

    janela = tk.Toplevel()
    janela.title("Certificados")
    janela.geometry("720x500")

    tk.Label(janela, text="CSV:").pack(anchor='w', padx=10)
    entry_csv = tk.Entry(janela, width=90); entry_csv.pack(padx=10)
    tk.Button(janela, text="Selecionar CSV", command=selecionar_csv).pack(padx=10, pady=5)

    tk.Label(janela, text="Modelos:").pack(anchor='w', padx=10)
    entry_modelos = tk.Entry(janela, width=90); entry_modelos.pack(padx=10)
    tk.Button(janela, text="Selecionar Modelos", command=selecionar_modelos).pack(padx=10, pady=5)

    opcao_saida = ttk.Combobox(janela, values=["DOCX", "PDF", "Ambos"], state="readonly")
    opcao_saida.set("Ambos"); opcao_saida.pack(padx=10, pady=5)

    barra_progresso = ttk.Progressbar(janela, orient="horizontal", length=600, mode="determinate"); barra_progresso.pack(pady=12)
    label_status = tk.Label(janela, text="Aguardando...", font=("Arial", 10)); label_status.pack()

    frame_botoes = tk.Frame(janela); frame_botoes.pack(pady=12)
    btn_gerar = tk.Button(frame_botoes, text="Gerar", command=iniciar_geracao_thread, bg="green", fg="white"); btn_gerar.grid(row=0, column=0, padx=5)
    btn_cancelar = tk.Button(frame_botoes, text="Cancelar", command=cancelar, bg="red", fg="white"); btn_cancelar.grid(row=0, column=1, padx=5)


def interface_ppp():
    janela = tk.Toplevel()
    janela.title("PPP")
    janela.geometry("400x200")
    tk.Button(janela, text="Gerar PPP", command=gerar_ppp, bg="blue", fg="white").pack(expand=True, pady=40)


# =============================
# Tela de senha e Menu Principal
# =============================

def menu_principal():
    menu = tk.Tk()
    menu.title("Sistema de Certificados e PPP")
    menu.geometry("500x300")
    menu.resizable(False, False)

    style = ttk.Style(menu)
    style.theme_use("clam")

    frame_top = tk.Frame(menu, bg="#2c3e50")
    frame_top.pack(fill="x")

    titulo = tk.Label(frame_top, text="Gerador de Certificados e PPP",
                      font=("Arial", 14, "bold"), fg="white", bg="#2c3e50", pady=10)
    titulo.pack()

    frame_mid = tk.Frame(menu, pady=30)
    frame_mid.pack(expand=True)

    btn1 = tk.Button(frame_mid, text="Gerar Certificados",
                     command=interface_certificados, width=25,
                     bg="#27ae60", fg="white", font=("Arial", 11, "bold"))
    btn1.grid(row=0, column=0, pady=10)

    btn2 = tk.Button(frame_mid, text="Gerar PPP",
                     command=interface_ppp, width=25,
                     bg="#2980b9", fg="white", font=("Arial", 11, "bold"))
    btn2.grid(row=1, column=0, pady=10)

    frame_bot = tk.Frame(menu, bg="#ecf0f1")
    frame_bot.pack(fill="x")
    rodape = tk.Label(frame_bot, text="© Vytola – Técnico de Segurança do Trabalho",
                      font=("Arial", 9), bg="#ecf0f1", fg="gray")
    rodape.pack(pady=5)

    menu.mainloop()

def solicitar_senha():
    def verificar():
        if entry.get() == SENHA_ACESSO:
            senha_win.destroy()
            menu_principal()
        else:
            messagebox.showerror("Erro", "Senha incorreta")

    senha_win = tk.Tk()
    senha_win.title("Acesso Restrito")
    senha_win.geometry("300x150")
    tk.Label(senha_win, text="Digite a senha:").pack(pady=10)
    entry = tk.Entry(senha_win, show="*"); entry.pack()
    tk.Button(senha_win, text="Entrar", command=verificar).pack(pady=10)
    senha_win.mainloop()

if __name__ == "__main__":
    solicitar_senha()
