import pywhatkit as kit
import pandas as pd
import time
import os
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext

clientes_df = None  # variável global para guardar os dados da planilha

def selecionar_arquivo():
    """Abre explorador para escolher o arquivo de clientes"""
    global clientes_df
    arquivo = filedialog.askopenfilename(
        title="Selecione a planilha de clientes",
        filetypes=[("Arquivos Excel", "*.xlsx")]
    )
    if arquivo:
        entry_arquivo.delete(0, tk.END)
        entry_arquivo.insert(0, arquivo)
        try:
            clientes_df = pd.read_excel(arquivo)
            messagebox.showinfo("Arquivo carregado", "Planilha carregada com sucesso!")
        except Exception as e:
            clientes_df = None
            messagebox.showerror("Erro", f"Não foi possível ler o arquivo.\n{e}")

def mostrar_campos():
    """Exibe os nomes das colunas da planilha"""
    global clientes_df
    if clientes_df is None:
        messagebox.showwarning("Atenção", "Selecione primeiro um arquivo válido.")
        return
    colunas = list(clientes_df.columns)
    campos = "\n".join([f"- {c}" for c in colunas])
    messagebox.showinfo("Campos disponíveis", f"Você pode usar na mensagem:\n\n{campos}")

def enviar_mensagens():
    global clientes_df
    arquivo = entry_arquivo.get()
    mensagem = txt_mensagem.get("1.0", tk.END).strip()

    if not arquivo or not mensagem:
        messagebox.showwarning("Atenção", "Selecione um arquivo e escreva a mensagem.")
        return

    if clientes_df is None:
        try:
            clientes_df = pd.read_excel(arquivo)
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível ler o arquivo.\n{e}")
            return

    relatorio = []
    for i, row in clientes_df.iterrows():
        dados = row.to_dict()

        numero = str(dados.get("Numero", "")).strip()
        if not numero.startswith("+"):
            numero = "+55" + numero

        try:
            texto = mensagem.format(**dados)
        except KeyError as e:
            messagebox.showerror("Erro", f"Campo {e} não encontrado na planilha.")
            return

        log_text.insert(tk.END, f"Enviando para {dados.get('Nome', '')} ({numero}): {texto}\n")
        log_text.see(tk.END)
        root.update()

        try:
            kit.sendwhatmsg_instantly(numero, texto, wait_time=10, tab_close=False)
            time.sleep(5)
            status = "Enviado"
        except Exception as e:
            status = f"Erro: {e}"

        relatorio.append({
            "Nome": dados.get("Nome", ""),
            "Numero": numero,
            "Mensagem": texto,
            "Status": status,
            "DataHora": datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        })

    # --- 🔽 Salvar relatório na pasta "output" com data/hora no nome ---
    output_dir = os.path.join(os.path.dirname(__file__), "output")
    os.makedirs(output_dir, exist_ok=True)

    data_hora = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    nome_relatorio = f"relatorio_envios_{data_hora}.xlsx"
    caminho_relatorio = os.path.join(output_dir, nome_relatorio)

    df_relatorio = pd.DataFrame(relatorio)
    df_relatorio.to_excel(caminho_relatorio, index=False)

    messagebox.showinfo("Finalizado", f"Mensagens enviadas!\nRelatório salvo em:\n{caminho_relatorio}")

# --- Interface Gráfica ---
root = tk.Tk()
root.title("Envio Automático WhatsApp - Escritório de Advocacia")
root.geometry("720x540")

# Seletor de arquivo
frame_arquivo = tk.Frame(root)
frame_arquivo.pack(pady=10, fill="x", padx=10)
tk.Label(frame_arquivo, text="Arquivo de Clientes (.xlsx):").pack(side="left")
entry_arquivo = tk.Entry(frame_arquivo, width=50)
entry_arquivo.pack(side="left", padx=5)
tk.Button(frame_arquivo, text="Selecionar", command=selecionar_arquivo).pack(side="left")
tk.Button(frame_arquivo, text="Ver Campos", command=mostrar_campos).pack(side="left", padx=5)

# Mensagem padrão
tk.Label(root, text="Mensagem Padrão (use {Nome}, {DataAudiencia}, etc.):").pack(anchor="w", padx=10)
txt_mensagem = scrolledtext.ScrolledText(root, height=5, width=80)
txt_mensagem.pack(padx=10, pady=5)
txt_mensagem.insert(tk.END, "Olá {Nome}, sua audiência será em {DataAudiencia}, no endereço {Endereco}.")

# Botão Enviar
tk.Button(root, text="Enviar Mensagens", command=enviar_mensagens, bg="green", fg="white").pack(pady=10)

# Log de envio
tk.Label(root, text="Log de Envio:").pack(anchor="w", padx=10)
log_text = scrolledtext.ScrolledText(root, height=12, width=85)
log_text.pack(padx=10, pady=5)

root.mainloop()