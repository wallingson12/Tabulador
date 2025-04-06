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
from ocr_livre import processar_e_salvar_pdf_ocr

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
    diretorio = filedialog.askdirectory()

    if diretorio:
        usar_ocr = var_ocr.get()
        arquivos_pdf = [f for f in os.listdir(diretorio) if f.lower().endswith(".pdf")]

        if not arquivos_pdf:
            messagebox.showwarning("Aviso", "Nenhum arquivo PDF encontrado na pasta selecionada.")
            return

        dctf_detalhamento_total = []

        for arquivo in arquivos_pdf:
            caminho_pdf = os.path.join(diretorio, arquivo)
            try:
                dctf_detalhamento, all_texts = extrair_dctf_pdf(caminho_pdf, usar_ocr)
                dctf_detalhamento["Arquivo PDF"] = arquivo
                dctf_detalhamento_total.append(dctf_detalhamento)
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao processar {arquivo}: {str(e)}")

        if dctf_detalhamento_total:
            df_final = pd.concat(dctf_detalhamento_total, ignore_index=True)
            caminho_saida = os.path.join(diretorio, "DCTF_Consolidado.xlsx")
            df_final.to_excel(caminho_saida, index=False)
            messagebox.showinfo("Sucesso", f"Dados consolidados em: {caminho_saida}")

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
        usar_ocr = var_ocr.get()
        try:
            extract_pgdas_values(diretorio, diretorio, usar_ocr=usar_ocr)
            messagebox.showinfo("Sucesso", "Valores PGDAS extraídos e salvos com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro ao extrair os valores do PGDAS: {str(e)}")

def selecionar_fontes_pagadoras():
    arquivo = filedialog.askopenfilename(filetypes=[("Arquivos PDF", "*.pdf")])
    if arquivo:
        usar_ocr = var_ocr.get()
        excel_path = os.path.splitext(arquivo)[0] + '.xlsx'

        try:
            if usar_ocr:
                extract_data_to_excel_with_ocr(arquivo, excel_path, 'Data')
                extract_values_to_excel_with_ocr(arquivo, excel_path, 'Valores')
                messagebox.showinfo("Sucesso", f"Dados extraídos com OCR e salvos em: {excel_path}")
            else:
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

def selecionar_pdf_ocr_free():
    root = tk.Tk()
    root.withdraw()
    caminho_pdf = filedialog.askopenfilename(filetypes=[("Arquivos PDF", "*.pdf")])

    if caminho_pdf:
        try:
            caminho_tesseract = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
            output_excel = os.path.splitext(caminho_pdf)[0] + ".xlsx"

            processar_e_salvar_pdf_ocr(caminho_pdf, caminho_tesseract, output_excel, dpi=600, first_page=2, last_page=2)

            messagebox.showinfo("Sucesso", f"Dados extraídos e salvos no arquivo Excel!\n{output_excel}")
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro ao extrair os dados: {str(e)}")

def atualizar_estilo_ocr():
    if var_ocr.get():
        checkbox_ocr.config(bg="white", fg="black", text="OCR Ativado")
    else:
        checkbox_ocr.config(bg="DarkBlue", fg="white", text="Usar OCR")

janela = tk.Tk()
janela.title("Ferramenta de Extração de PDF")
janela.iconbitmap('icone.ico')
janela.configure(bg='DarkBlue')

logo_path = 'a.jpg'
logo_img = Image.open(logo_path)
logo_img = logo_img.resize((100, 100))
logo = ImageTk.PhotoImage(logo_img)
label_logo = tk.Label(janela, image=logo, bg='DarkBlue')
label_logo.pack(pady=10)

frame_botoes = tk.Frame(janela, bg='DarkBlue')
frame_botoes.pack(padx=20, pady=20)

var_ocr = tk.BooleanVar(value=False)

checkbox_ocr = tk.Checkbutton(frame_botoes, text="Usar OCR", variable=var_ocr, bg='DarkBlue', fg='white', command=atualizar_estilo_ocr)
checkbox_ocr.grid(row=0, column=0, columnspan=2, padx=10, pady=5)

botoes = [
    ("Selecionar PDF para Extração de Tabelas", selecionar_arquivo),
    ("Selecionar Diretório para DCTF", selecionar_diretorio_dctf),
    ("Selecionar Diretório para Recolhimento", selecionar_diretorio_recolhimento),
    ("Selecionar Diretório para PGDAS", selecionar_diretorio_pgdas),
    ("Selecionar Diretório para DCOMP", selecionar_diretorio_dcomp),
    ("Selecionar PDF para Fontes Pagadoras", selecionar_fontes_pagadoras)
]

for i, (texto, comando) in enumerate(botoes):
    tk.Button(frame_botoes, text=texto, command=comando).grid(row=i+1, column=0, columnspan=2, padx=10, pady=5)

janela.mainloop()