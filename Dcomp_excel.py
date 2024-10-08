import re
from pdf2image import convert_from_path
import pytesseract
import pandas as pd
import os
import cv2
import numpy as np
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

def extrair_dcomp_pdf(diretorio, usar_ocr=True):
    if not os.path.exists(diretorio):
        raise FileNotFoundError("O diretório especificado não foi encontrado.")

    caminho_tesseract = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
    pytesseract.pytesseract.tesseract_cmd = caminho_tesseract

    regex_patterns = {
        'CNPJ': r'CNPJ\s*(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})',
        'D ou C': r'001\. (Débito|Crédito)\s*(.*)',
        'Data de Transmissão': r'Data de Transmissão\s*(\d{2}/\d{2}/\d{4})',
        'Nome Empresarial': r'Nome Empresarial\s*(.*)',
        'Informado em Outro PER/DCOMP': r'Informado em Outro PER/DCOMP\s*(.*)',
        'PER/DCOMP Retificador': r'PER/DCOMP Retificador\s*(Sim|Não)',
        'Período de Apuração': r'Período de Apuração\s*(.*)',
        'Principal': r'Principal\s*([\d.,]+)',
        'Selic Acumulada': r'Selic Acumulada\s*([\d.,]+)',
        'Crédito Atualizado': r'Crédito Atualizado\s*([\d.,]+)',
        'Saldo do Crédito Original': r'Saldo do Crédito Original\s*([\d.,]+)',
        'Valor Original do Crédito Inicial': r'Valor Original do Crédito Inicial\s*([\d.,]+)',
        '0001. Período de Apuração': r'0001\. Período de Apuração\s*(.*)',
        'Código da Receita/Denominação': r'Código da Receita/Denominação\s*(.*)',
        'Débito Controlado em Processo': r'Débito Controlado em Processo\s*(.*)',
        'Multa': r'Multa\s*([\d.,]+)',
        'Juros': r'Juros\s*([\d.,]+)',
        'Total': r'Total\s*([\d.,]+)'
    }

    all_extracted_data = []

    for filename in os.listdir(diretorio):
        if filename.endswith(".pdf"):
            filepath = os.path.join(diretorio, filename)
            print(f"Processando: {filepath}")
            images = convert_from_path(filepath, dpi=600)
            extracted_text = ""

            for img in images:
                img = pre_process(img)
                if usar_ocr:
                    text = pytesseract.image_to_string(img, lang='por', config='--psm 4')
                else:
                    text = ''  # Se não usar OCR, o texto será vazio

                extracted_text += text

            extracted_data = {'Número': filename, 'CRÉDITO': 'CRÉDITO PAGAMENTO INDEVIDO OU A MAIOR'}

            for key, pattern in regex_patterns.items():
                match = re.search(pattern, extracted_text)
                if match:
                    extracted_data[key] = match.group(1).strip()
                else:
                    extracted_data[key] = None

            all_extracted_data.append(extracted_data)

    df = pd.DataFrame(all_extracted_data)
    df.to_excel(os.path.join(diretorio, 'dcomp_extraido.xlsx'), index=False)
    print("Dados extraídos e salvos em 'dcomp_extraido.xlsx'")
