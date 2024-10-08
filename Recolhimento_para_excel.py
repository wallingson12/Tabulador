import re
import pandas as pd
from PIL import Image
import numpy as np
import cv2
from pdf2image import convert_from_path
import pytesseract
import os
import camelot


def pre_process(image):
    img_np = np.array(image)
    gray = cv2.cvtColor(img_np, cv2.COLOR_BGR2GRAY)
    adaptive_thresh = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 11, 2)
    kernel = np.ones((2, 2), np.uint8)
    dilated = cv2.dilate(adaptive_thresh, kernel, iterations=1)
    eroded = cv2.erode(dilated, kernel, iterations=2)
    processed_img = Image.fromarray(eroded)
    return processed_img


def extrair_recolhimento_pdf(diretorio, usar_ocr=False):
    if not os.path.exists(diretorio):
        raise FileNotFoundError("O diretório especificado não foi encontrado.")

    caminho_tesseract = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
    pytesseract.pytesseract.tesseract_cmd = caminho_tesseract

    dados_totais = []

    for arquivo in os.listdir(diretorio):
        if arquivo.endswith(".pdf"):
            caminho_pdf = os.path.join(diretorio, arquivo)

            if usar_ocr:
                images = convert_from_path(caminho_pdf, dpi=600)
                padrao_linha = r'(?P<DARF>[^\s]+|\bDARF\s+\d+)\s+(?P<Data_de_arreacadação>\d{2}/\d{2}/\d{4})\s+(?P<Data_de_vencimento>\d{2}/\d{2}/\d{4})\s+(?P<Periodo_de_apuração>\d{2}/\d{2}/\d{4})\s+(?P<Código_da_receita>[\w]{4})\s+(?P<Número_de_documento>\d{10,17})\s+(?P<Valor>[0-9.,]+)'

                for img in images:
                    img = pre_process(img)
                    text = pytesseract.image_to_string(img, lang='por', config='--psm 6')
                    linhas = text.split('\n')
                    for linha in linhas:
                        match = re.match(padrao_linha, linha)
                        if match:
                            dado = match.groupdict()
                            dados_totais.append(dado)

            else:
                # Usando Camelot para extrair dados sem OCR
                tabelas = camelot.read_pdf(caminho_pdf, pages='all', flavor='stream')
                for tabela in tabelas:
                    for index, linha in tabela.df.iterrows():
                        # Adaptar o regex se necessário
                        match = re.match(padrao_linha, ' '.join(linha.values))
                        if match:
                            dado = match.groupdict()
                            dados_totais.append(dado)

    recolhimento = pd.DataFrame(dados_totais)
    return recolhimento