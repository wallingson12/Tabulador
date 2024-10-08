import pytesseract
from pdf2image import convert_from_path
import re
import pandas as pd
import numpy as np
import cv2
from PIL import Image
from tkinter import filedialog, messagebox

def preprocess_image(image):
    img_np = np.array(image)
    gray = cv2.cvtColor(img_np, cv2.COLOR_BGR2GRAY)
    adaptive_thresh = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 11, 2)
    kernel = np.ones((2, 2), np.uint8)
    dilated = cv2.dilate(adaptive_thresh, kernel, iterations=1)
    eroded = cv2.erode(dilated, kernel, iterations=2)
    processed_img = Image.fromarray(eroded)
    return processed_img

def limpar_valor(valor):
    if valor is not None:
        valor = re.sub(r'[^\d.,]', '', valor)  # Remove caracteres não numéricos
        valor = valor.replace(',', '.')  # Substitui vírgulas por pontos
        return valor
    return None

def extrair_dctf_pdf(caminho_pdf, usar_ocr=False, nome_saida_detalhamento='dctf_detalhamento.xlsx', nome_saida_resumo='dctf_resumo.xlsx'):
    caminho_tesseract = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
    pytesseract.pytesseract.tesseract_cmd = caminho_tesseract

    # Começar a partir da página 3
    images = convert_from_path(caminho_pdf, dpi=500, first_page=3)

    padrao_detalhamento = {
        "GRUPO DO TRIBUTO": r'GRUPO DO TRIBUTO\s*:?\s*(.+)',
        "CÓDIGO RECEITA": r'CÓDIGO RECEITA\s*:\s*(.+)',
        "PERIODICIDADE": r'PERIODICIDADE\s*:\s*(\S+)',
        "PA": r'PA:\s*(\d{2}/\d{2}/\d{4})',
        "Valor do Principal": r'Valor do Principal\s*:\s*([\d.,/]+)',
        "Valor da Multa": r'Valor da Multa\s*:\s*([\d.,/]+)',
        "Valor dos Juros": r'Valor dos Juros\s*:\s*([\d.,/]+)',
        "Valor Pago do Débito": r'Valor Pago do D[êéEe]{1,2}bito\s*:\s*([\d.,/]+)',
        "Valor Total do DARF": r'Valor Total do DARF\s*:\s*([\d.,/]+)'
    }

    padrao_resumo = {
        "PAGAMENTO": r'PAGAMENTO\s*\s*([\d,.\s]+)',
        "COMPENSAÇÕES": r'COMPENSAÇÕES\s*\s*([\d,.\s]+)',
        "PARCELAMENTO": r'PARCELAMENTO\s*\s*([\d,.\s]+)',
        "SUSPENSÃO": r'SUSPENSÃO\s*\s*([\d,.\s]+)',
        "SOMA DOS CRÉDITOS VINCULADOS": r'SOMA DOS CRÉDITOS VINCULADOS\s*:\s*([\d,.\s]+)'
    }

    dctf_detalhamento = {col: [] for col in padrao_detalhamento}
    informacoes_resumo = {col: [] for col in padrao_resumo}
    all_texts = []  # Lista para armazenar o texto extraído de cada página

    for img in images:
        img = img.convert('RGB')

        if usar_ocr:
            processed_img = preprocess_image(img)
            text = pytesseract.image_to_string(processed_img, lang='por', config='--psm 4')
        else:
            # Se não usar OCR, você pode adicionar uma lógica para extrair texto diretamente da imagem,
            # mas isso normalmente não é possível sem usar OCR. Aqui, apenas definimos o texto como vazio.
            text = ''

        all_texts.append(text)

        # Extração de dados do detalhamento
        for chave, padrao in padrao_detalhamento.items():
            match = re.search(padrao, text)
            if match:
                valor = match.group(1).strip()
                if chave in ["Valor do Principal", "Valor da Multa", "Valor dos Juros", "Valor Pago do Débito", "Valor Total do DARF"]:
                    valor = limpar_valor(valor)
                dctf_detalhamento[chave].append(valor)
            else:
                dctf_detalhamento[chave].append(None)

        # Extração de dados do resumo
        for chave, padrao in padrao_resumo.items():
            match = re.search(padrao, text)
            if match:
                valor = match.group(1).strip()
                informacoes_resumo[chave].append(valor)
            else:
                informacoes_resumo[chave].append(None)

    # Criar DataFrames com o número de linhas apropriado
    max_len_detalhamento = max(len(v) for v in dctf_detalhamento.values())
    max_len_resumo = max(len(v) for v in informacoes_resumo.values())

    dctf_detalhamento_df = pd.DataFrame({k: (v + [None] * (max_len_detalhamento - len(v))) for k, v in dctf_detalhamento.items()})
    dctf_resumo_df = pd.DataFrame({k: (v + [None] * (max_len_resumo - len(v))) for k, v in informacoes_resumo.items()})

    # Salvar os DataFrames em arquivos Excel
    dctf_detalhamento_df.to_excel(nome_saida_detalhamento, index=False)
    dctf_resumo_df.to_excel(nome_saida_resumo, index=False)

    return dctf_detalhamento_df, dctf_resumo_df, all_texts