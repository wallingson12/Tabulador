import re
import os
import PyPDF2  # Biblioteca para extrair texto diretamente de PDFs
import pandas as pd
from pdf2image import convert_from_path
import pytesseract
import numpy as np
import cv2
from PIL import Image

def pre_process(image):
    img_np = np.array(image)
    gray = cv2.cvtColor(img_np, cv2.COLOR_BGR2GRAY)
    adaptive_thresh = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 11, 2)

    kernel = np.ones((2, 2), np.uint8)
    dilated = cv2.dilate(adaptive_thresh, kernel, iterations=1)
    eroded = cv2.erode(dilated, kernel, iterations=2)

    processed_img = Image.fromarray(eroded)
    return processed_img

def extract_pgdas_values(diretorio_pdfs, diretorio_destino, usar_ocr=False):
    # Configurar o caminho do Tesseract
    caminho_tesseract = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
    pytesseract.pytesseract.tesseract_cmd = caminho_tesseract

    # Lista para armazenar os valores extraídos
    all_extracted_values = []

    # Iterar por todos os arquivos no diretório
    for filename in os.listdir(diretorio_pdfs):
        if filename.endswith(".pdf"):
            filepath = os.path.join(diretorio_pdfs, filename)
            extracted_text = ""

            if usar_ocr:
                # Extrair texto usando OCR de todas as páginas do PDF
                images = convert_from_path(filepath)
                for img in images:
                    img = img.convert('RGB')
                    processed_img = pre_process(img)  # Processar a imagem antes da extração
                    text = pytesseract.image_to_string(processed_img, lang='por', config='--psm 4')
                    extracted_text += text
            else:
                # Extrair texto diretamente do PDF usando PyPDF2 (sem OCR)
                with open(filepath, "rb") as pdf_file:
                    reader = PyPDF2.PdfReader(pdf_file)
                    for page in reader.pages:
                        extracted_text += page.extract_text() or ""

            # Procurar por valores monetários em todo o texto do PDF
            money_values = re.findall(r'\b\d{1,3}(?:\.\d{3})*(?:,\d{2})\b', extracted_text)
            if money_values:
                extracted_values = {'Receita Bruta do PA (RPA) - Competência': money_values[0],
                                    'Receita bruta acumulada nos doze meses anteriores': money_values[3] if len(money_values) > 3 else None}

                # Adicionar o cabeçalho para os últimos 9 valores
                last_n_values = money_values[-9:]
                headers = ['IRPJ', 'CSLL', 'COFINS', 'PIS/Pasep', 'INSS/CPP', 'ICMS', 'IPI', 'ISS', 'Total']
                for i, header in enumerate(headers):
                    extracted_values[header] = last_n_values[i] if i < len(last_n_values) else None

                # Extrair o ano do nome do arquivo
                ano_match = re.search(r'\b\d{4}\b', filename)
                extracted_values['Ano'] = ano_match.group() if ano_match else None

                all_extracted_values.append(extracted_values)

    # Criar DataFrame com os valores extraídos
    df = pd.DataFrame(all_extracted_values)

    # Criar diretório de destino, se não existir
    if not os.path.exists(diretorio_destino):
        os.makedirs(diretorio_destino)

    # Salvar os dados em arquivos separados para cada ano no diretório de destino
    for ano in df['Ano'].unique():
        df_ano = df[df['Ano'] == ano]
        filename = os.path.join(diretorio_destino, f'dados_{ano}.xlsx')
        df_ano.to_excel(filename, index=False)

    print("Arquivos salvos com sucesso no diretório de destino.")