import tkinter as tk
from tkinter import filedialog, messagebox
import camelot
from collections import defaultdict
from PIL import Image, ImageTk
import re
from pdf2image import convert_from_path
import pytesseract
import pandas as pd
import os

def extrair_tabelas_para_excel(arquivo):
    #stream
    #exact
    #auto
    #lattice
    tabelas = camelot.read_pdf(arquivo, pages='all', flavor='exact')

    nome_arquivo = os.path.splitext(os.path.basename(arquivo))[0]
    df_final = pd.DataFrame()

    for tabela in tabelas:
        df_final = pd.concat([df_final, tabela.df], ignore_index=False)

    output_dir = os.path.join(os.path.dirname(arquivo), nome_arquivo)
    os.makedirs(output_dir, exist_ok=True)

    output_file = os.path.join(output_dir, f'{nome_arquivo}.xlsx')
    df_final.to_excel(output_file, header=False, index=False)

def extrair_dctf_pdf(diretorio):
    if not os.path.exists(diretorio):
        raise FileNotFoundError("O diretório especificado não foi encontrado.")

    caminho_tesseract = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
    pytesseract.pytesseract.tesseract_cmd = caminho_tesseract

    dctf_detalhamento_total = []
    dctf_resumo_total = []

    padrao_linha = r'(.+?):\s*(.+)'
    padrao_pagamento = r'PAGAMENTO\s*([\d,.]+)'
    padrao_compensacao = r'COMPENSAÇÕES\s*([\d,.]+)'
    padrao_parcelamento = r'PARCELAMENTO\s*([\d,.]+)'
    padrao_suspensao = r'SUSPENSÃO\s*([\d,.]+)'
    padrao_valor_debito = r'SOMA DOS CRÉDITOS VINCULADOS:\s*([\d,.]+)'

    for arquivo in os.listdir(diretorio):
        if arquivo.endswith(".pdf"):
            caminho_pdf = os.path.join(diretorio, arquivo)

            images = convert_from_path(caminho_pdf)

            dctf_detalhamento = defaultdict(list)
            informacoes_resumo = defaultdict(list)

            for img in images:
                img = img.convert('RGB')
                text = pytesseract.image_to_string(img, lang='por', config='--psm 4')
                linhas = text.split('\n')
                for linha in linhas:
                    match = re.match(padrao_linha, linha)
                    if match:
                        chave, valor = match.groups()
                        if chave in dctf_detalhamento:
                            dctf_detalhamento[chave].append(valor.strip())
                        elif chave in informacoes_resumo:
                            informacoes_resumo[chave].append(valor.strip())

            dctf_detalhamento_total.append(pd.DataFrame.from_dict(dctf_detalhamento))
            dctf_resumo_total.append(pd.DataFrame.from_dict(informacoes_resumo))

    dctf_detalhamento_df = pd.concat(dctf_detalhamento_total, ignore_index=True)
    dctf_resumo_df = pd.concat(dctf_resumo_total, ignore_index=True)

    dctf_detalhamento_df.to_excel('dctf_detalhamento.xlsx', index=False)
    dctf_resumo_df.to_excel('dctf_resumo.xlsx', index=False)

    return dctf_detalhamento_df, dctf_resumo_df

def extrair_recolhimento_pdf(diretorio):
    if not os.path.exists(diretorio):
        raise FileNotFoundError("O diretório especificado não foi encontrado.")

    caminho_tesseract = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
    pytesseract.pytesseract.tesseract_cmd = caminho_tesseract

    dados_totais = []

    for arquivo in os.listdir(diretorio):
        if arquivo.endswith(".pdf"):
            caminho_pdf = os.path.join(diretorio, arquivo)

            images = convert_from_path(caminho_pdf)

            padrao_linha = r'(?P<DARF>[^\s]+)\s+(?P<Data de arreacadação>\d{2}/\d{2}/\d{4})\s+(?P<Data de vencimento>\d{2}/\d{2}/\d{4})\s+(?PPeriodo de apuração\d{2}/\d{2}/\d{4})\s+(?P<Código da receita>\d{4})\s+(?P<Número de documento>\d{17})\s+(?P<Valor>[0-9.,]+)'

            for img in images:
                img = img.convert('RGB')
                text = pytesseract.image_to_string(img, lang='por', config='--psm 6')
                linhas = text.split('\n')
                for linha in linhas:
                    match = re.match(padrao_linha, linha)
                    if match:
                        dado = match.groupdict()
                        dados_totais.append(dado)

    recolhimento = pd.DataFrame(dados_totais)

    return recolhimento

def extrair_dcomp_pdf(diretorio):
    if not os.path.exists(diretorio):
        raise FileNotFoundError("O diretório especificado não foi encontrado.")

    # Configurar o caminho do Tesseract
    caminho_tesseract = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
    pytesseract.pytesseract.tesseract_cmd = caminho_tesseract

    # Expressões regulares para os padrões que queremos encontrar
    regex_patterns = {
        'CNPJ': r'CNPJ\s*(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})',
        'D ou C': r'001\. (Débito|Crédito)\s*(.*)',
        'Data de Transmissão': r'Data de Transmissão\s*(\d{2}/\d{2}/\d{4})',
        'Nome Empresarial': r'Nome Empresarial\s*(.*)',
        'Informado em Outro PER/DCOMP': r'Informado em Outro PER/DCOMP\s*(.*)',
        'PER/DCOMP Retificador': r'PER/DCOMP Retificador\s*(Sim|Não)',
        'Período de Apuração': r'Período de Apuração\s*(.*)',
        'Principal': r'Principal\s*([\d.,]+\d)',
        'Selic Acumulada': r'Selic Acumulada\s*([\d.,]+\d)',
        'Crédito Atualizado': r'Crédito Atualizado\s*([\d.,]+\d)',
        'Saldo do Crédito Original': r'Saldo do Crédito Original\s*([\d.,]+\d)',
        'Valor Original do Crédito Inicial': r'Valor Original do Crédito Inicial\s*([\d.,]+\d)',
        '0001. Período de Apuração': r'0001\. Período de Apuração\s*(.*)',
        'Código da Receita/Denominação': r'Código da Receita/Denominação\s*(.*)',
        'Débito Controlado em Processo': r'Débito Controlado em Processo\s*(.*)',
        'Multa': r'Multa\s*([\d.,]+\d)',
        'Juros': r'Juros\s*([\d.,]+\d)',
        'Total': r'Total\s*([\d.,]+\d)'
    }

    # Lista para armazenar todos os dados extraídos
    all_extracted_data = []

    # Iterar por todos os arquivos no diretório
    for filename in os.listdir(diretorio):
        if filename.endswith(".pdf"):
            filepath = os.path.join(diretorio, filename)

            # Extrair texto de cada arquivo PDF
            images = convert_from_path(filepath)
            extracted_text = ""

            for img in images:
                img = img.convert('RGB')
                text = pytesseract.image_to_string(img, lang='por', config='--psm 4')
                extracted_text += text

            # Definir o valor diretamente para 'CRÉDITO'
            extracted_data = {'Número': filename, 'CRÉDITO': 'CRÉDITO PAGAMENTO INDEVIDO OU A MAIOR'}

            # Extrair dados usando expressões regulares
            for key, pattern in regex_patterns.items():
                match = re.search(pattern, extracted_text)
                if match:
                    extracted_data[key] = match.group(1).strip()
                else:
                    extracted_data[key] = None

            # Adicionar os dados extraídos à lista
            all_extracted_data.append(extracted_data)

    # Criar DataFrame com os dados extraídos de todos os arquivos
    df = pd.DataFrame(all_extracted_data)

    # Salvar os dados extraídos de um arquivo Excel
    df.to_excel(os.path.join(diretorio, 'dcomp_extraido.xlsx'), index=False)

def extract_pgdas_values(diretorio_pdfs):
    # Configurar o caminho do Tesseract
    caminho_tesseract = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
    pytesseract.pytesseract.tesseract_cmd = caminho_tesseract

    # Lista para armazenar os valores extraídos
    all_extracted_values = []

    # Obter o diretório de destino (mesmo diretório do programa)
    diretorio_destino = os.path.dirname(os.path.abspath(__file__))

    # Iterar por todos os arquivos no diretório
    for filename in os.listdir(diretorio_pdfs):
        if filename.endswith(".pdf"):
            filepath = os.path.join(diretorio_pdfs, filename)

            # Extrair texto de todas as páginas do PDF
            extracted_text = ""
            images = convert_from_path(filepath)
            for img in images:
                img = img.convert('RGB')
                text = pytesseract.image_to_string(img, lang='por', config='--psm 4')
                extracted_text += text

            # Procurar por valores monetários em todo o texto do PDF
            money_values = re.findall(r'\b\d{1,3}(?:\.\d{3})*(?:,\d{2})\b', extracted_text)
            if money_values:
                # Armazenar índices 0 e 3
                extracted_values = {'Receita Bruta do PA (RPA) - Competência': money_values[0],
                                    'Receita bruta acumulada nos doze meses anteriores': money_values[3]}

                # Adicionar o cabeçalho para os últimos 9 valores
                last_n_values = money_values[-9:]
                headers = ['IRPJ', 'CSLL', 'COFINS', 'PIS/Pasep', 'INSS/CPP', 'ICMS', 'IPI', 'ISS', 'Total']
                for i, header in enumerate(headers):
                    extracted_values[header] = last_n_values[i] if i < len(last_n_values) else None

                # Extrair o ano do nome do arquivo
                ano = re.search(r'\b\d{4}\b', filename).group()
                if ano:
                    extracted_values['Ano'] = ano

                all_extracted_values.append(extracted_values)

    # Criar DataFrame com os valores extraídos
    df = pd.DataFrame(all_extracted_values)

    # Salvar os dados em arquivos separados para cada ano no diretório de destino
    for ano in df['Ano'].unique():
        df_ano = df[df['Ano'] == ano]
        filename = os.path.join(diretorio_destino, f'dados_{ano}.xlsx')
        df_ano.to_excel(filename, index=False)

    print("Arquivos salvos com sucesso no diretório de destino.")

def selecionar_extrair_dctf_pdf():
    diretorio = filedialog.askdirectory()
    if diretorio:
        try:
            dctf_detalhamento, dctf_resumo = extrair_dctf_pdf(diretorio)
            messagebox.showinfo("Extração Concluída", "As informações foram extraídas com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro durante a extração: {str(e)}")

def selecionar_extrair_recolhimento_pdf():
    diretorio = filedialog.askdirectory()
    if diretorio:
        try:
            recolhimento = extrair_recolhimento_pdf(diretorio)
            recolhimento.to_excel('recolhimento.xlsx', index=False)
            messagebox.showinfo("Extração Concluída", "As informações foram extraídas com sucesso e salvas no arquivo 'recolhimento.xlsx'!")
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro durante a extração: {str(e)}")

def selecionar_extrair_tabelas_pdf():
    diretorio = filedialog.askdirectory()
    if diretorio:
        try:
            extrair_tabelas_para_excel(arquivo)
            messagebox.showinfo("Extração Concluída", "As tabelas foram extraídas com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro durante a extração: {str(e)}")

def selecionar_extrair_dcomp_pdf():
    diretorio = filedialog.askdirectory()
    if diretorio:
        try:
            extrair_dcomp_pdf(diretorio)
            messagebox.showinfo("Extração Concluída", "As tabelas foram extraídas com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro durante a extração: {str(e)}")

def selecionar_extrair_pgdas_pdf():
    diretorio = filedialog.askdirectory()
    if diretorio:
        try:
            selecionar_extrair_pgdas_pdf(diretorio)
            messagebox.showinfo("Extração Concluída", "As tabelas foram extraídas com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro durante a extração: {str(e)}")

# Interface gráfica
root = tk.Tk()
root.title("Exibir Imagem")

largura_display = 300  # Largura da janela
altura_display = 350   # Altura da janela

largura_tela = root.winfo_screenwidth()
altura_tela = root.winfo_screenheight()

posicao_x = (largura_tela - largura_display) // 2
posicao_y = (altura_tela - altura_display) // 2

root.geometry(f"{largura_display}x{altura_display}+{posicao_x}+{posicao_y}")

# Carregar e exibir a imagem
icone_path = "a.jpg"
if os.path.exists(icone_path):
    imagem = Image.open(icone_path)
    imagem = imagem.resize((200, 150))  # Redimensionar a imagem para o tamanho desejado
    imagem = ImageTk.PhotoImage(imagem)

    label_imagem = tk.Label(root, image=imagem)
    label_imagem.pack()

    root.configure(bg="#00426b")

    frame_central = tk.Frame(root, bg="#00426b")
    frame_central.pack(fill=tk.BOTH, expand=True)

button = tk.Button(root, text="Extrair DCTF", command=selecionar_extrair_dctf_pdf)
button.pack(pady=5)

button = tk.Button(root, text="Extrair Recolhimento", command=selecionar_extrair_recolhimento_pdf)
button.pack(pady=5)

button = tk.Button(root, text="Extrair Tabelas", command=selecionar_extrair_tabelas_pdf)
button.pack(pady=5)

button = tk.Button(root, text="Extrair Dcomp", command=selecionar_extrair_dcomp_pdf)
button.pack(pady=5)

button = tk.Button(root, text="Extrair PGDAS", command=selecionar_extrair_pgdas_pdf)
button.pack(pady=5)

root.mainloop()
