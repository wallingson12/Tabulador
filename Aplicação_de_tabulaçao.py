import tkinter as tk
from tkinter import filedialog, messagebox
import camelot
import pandas as pd
import os
from PIL import Image, ImageTk
from DCTF_para_excel import extrair_dctf_pdf
from Fontes_pagadoras_para_excel import *
from pgdas_para_excel import extract_pgdas_values
from Recolhimento_para_excel import extrair_recolhimento_pdf

def extrair_tabelas_para_excel(arquivo):
    try:
        tabelas = camelot.read_pdf(arquivo, pages='all', flavor='stream')
        nome_arquivo = os.path.splitext(os.path.basename(arquivo))[0]
        df_final = pd.DataFrame()

        for tabela in tabelas:
            df_final = pd.concat([df_final, tabela.df], ignore_index=True)

        output_dir = os.path.join(os.path.dirname(arquivo), nome_arquivo)
        os.makedirs(output_dir, exist_ok=True)
        output_file = os.path.join(output_dir, f'{nome_arquivo}.xlsx')
        df_final.to_excel(output_file, header=False, index=False)
        messagebox.showinfo("Sucesso", f"Tabelas extraídas e salvas em: {output_file}")

    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro ao extrair tabelas: {str(e)}")

def selecionar_arquivo():
    arquivo = filedialog.askopenfilename(filetypes=[("Arquivos PDF", "*.pdf")])
    if arquivo:
        extrair_tabelas_para_excel(arquivo)

def selecionar_diretorio_dctf():
    # Alterado para askdirectory
    diretorio = filedialog.askopenfilename()
    if diretorio:
        usar_ocr = var_ocr.get()  # Obtendo o valor do checkbox
        try:
            dctf_detalhamento, dctf_resumo, all_texts = extrair_dctf_pdf(diretorio, usar_ocr)
            messagebox.showinfo("Sucesso", "Dados DCTF extraídos e salvos no arquivo Excel com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro ao extrair DCTF: {str(e)}")

def selecionar_diretorio_recolhimento():
    diretorio = filedialog.askdirectory()
    if diretorio:
        recolhimento = extrair_recolhimento_pdf(diretorio)
        if recolhimento.empty:
            messagebox.showwarning("Atenção", "Nenhum dado foi encontrado nos PDFs selecionados.")
        else:
            arquivo_saida = os.path.join(diretorio, "recolhimento_extraido.xlsx")
            recolhimento.to_excel(arquivo_saida, index=False)
            messagebox.showinfo("Sucesso", f"Dados de recolhimento extraídos e salvos em: {arquivo_saida}")

def selecionar_diretorio_pgdas():
    diretorio = filedialog.askdirectory()
    if diretorio:
        usar_ocr = var_ocr.get()  # Obtendo o valor do checkbox
        try:
            extract_pgdas_values(diretorio, diretorio, usar_ocr=usar_ocr)  # Usar o diretório selecionado
            messagebox.showinfo("Sucesso", "Valores PGDAS extraídos e salvos com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro ao extrair os valores do PGDAS: {str(e)}")

def selecionar_fontes_pagadoras():
    arquivo = filedialog.askopenfilename(filetypes=[("Arquivos PDF", "*.pdf")])
    if arquivo:
        usar_ocr = var_ocr.get()  # Obtendo o valor do checkbox
        excel_path = os.path.splitext(arquivo)[0] + '.xlsx'  # Caminho do Excel com base no PDF

        try:
            if usar_ocr:  # Usar OCR
                extract_data_to_excel_with_ocr(arquivo, excel_path, 'Data')
                extract_values_to_excel_with_ocr(arquivo, excel_path, 'Valores')
                messagebox.showinfo("Sucesso", f"Dados extraídos com OCR e salvos em: {excel_path}")
            else:  # Sem OCR
                extract_data_to_excel(arquivo, excel_path)
                messagebox.showinfo("Sucesso", f"Dados extraídos e salvos em: {excel_path}")
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro ao extrair dados: {str(e)}")

def selecionar_diretorio_dcomp():
    diretorio = filedialog.askdirectory()
    if diretorio:
        try:
            extrair_dcomp_pdf(diretorio)
            messagebox.showinfo("Sucesso", "Dados DCOMP extraídos e salvos no arquivo Excel com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro ao extrair DCOMP: {str(e)}")

# Criação da interface gráfica
janela = tk.Tk()
janela.title("Ferramenta de Extração de PDF")

# Definindo a cor de fundo
janela.configure(bg='Dark blue')

# Carregando o logo
logo_path = 'a.jpg'  # Substitua pelo caminho do seu logo
logo_img = Image.open(logo_path)
logo_img = logo_img.resize((100, 100))  # Redimensiona a imagem se necessário
logo = ImageTk.PhotoImage(logo_img)

# Criando um label para exibir o logo
label_logo = tk.Label(janela, image=logo, bg='Dark blue')  # A cor de fundo deve combinar
label_logo.pack(pady=10)  # Adiciona um espaço vertical

frame_botoes = tk.Frame(janela)
frame_botoes.pack(padx=20, pady=20)

# Variável para controle do OCR
var_ocr = tk.BooleanVar(value=False)

# Checkbox para selecionar o uso do OCR
checkbox_ocr = tk.Checkbutton(frame_botoes, text="Usar OCR", variable=var_ocr, bg='DarkBlue', fg='white')
checkbox_ocr.grid(row=0, column=0, columnspan=2, padx=10, pady=5)

# Botões
btn_selecionar_arquivo = tk.Button(frame_botoes, text="Selecionar PDF para Extração de Tabelas", command=selecionar_arquivo)
btn_selecionar_arquivo.grid(row=1, column=0, columnspan=2, padx=10, pady=5)

btn_selecionar_diretorio_dctf = tk.Button(frame_botoes, text="Selecionar Diretório para DCTF", command=selecionar_diretorio_dctf)
btn_selecionar_diretorio_dctf.grid(row=2, column=0, columnspan=2, padx=10, pady=5)

btn_selecionar_diretorio_recolhimento = tk.Button(frame_botoes, text="Selecionar Diretório para Recolhimento", command=selecionar_diretorio_recolhimento)
btn_selecionar_diretorio_recolhimento.grid(row=3, column=0, columnspan=2, padx=10, pady=5)

btn_selecionar_diretorio_pgdas = tk.Button(frame_botoes, text="Selecionar Diretório para PGDAS", command=selecionar_diretorio_pgdas)
btn_selecionar_diretorio_pgdas.grid(row=4, column=0, columnspan=2, padx=10, pady=5)

btn_selecionar_diretorio_dcomp = tk.Button(frame_botoes, text="Selecionar Diretório para DCOMP", command=selecionar_diretorio_dcomp)
btn_selecionar_diretorio_dcomp.grid(row=5, column=0, columnspan=2, padx=10, pady=5)

btn_selecionar_fontes_pagadoras = tk.Button(frame_botoes, text="Selecionar PDF para Fontes Pagadoras", command=selecionar_fontes_pagadoras)
btn_selecionar_fontes_pagadoras.grid(row=6, column=0, columnspan=2, padx=10, pady=5)

# Iniciar o loop principal da interface
janela.mainloop()