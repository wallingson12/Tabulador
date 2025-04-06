import pytesseract
from pdf2image import convert_from_path
import re
import pandas as pd
import numpy as np
import cv2
from PIL import Image

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
    if valor:
        valor = re.sub(r'[^\d.,;]', '', valor)  # Remove apenas caracteres não permitidos
        return valor
    return ''

def extrair_dctf_pdf(caminho_pdf, usar_ocr=True, nome_saida_detalhamento='dctf_detalhamento.xlsx'):
    caminho_tesseract = r"C:\\Program Files\\Tesseract-OCR\\tesseract.exe"
    pytesseract.pytesseract.tesseract_cmd = caminho_tesseract

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

    dctf_detalhamento = {col: [] for col in padrao_detalhamento}

    all_texts = []
    soma_multas = 0
    soma_juros = 0
    novo_grupo = True

    for img in images:
        img = img.convert('RGB')

        if usar_ocr:
            processed_img = preprocess_image(img)
            text = pytesseract.image_to_string(processed_img, lang='por', config='--psm 4')
        else:
            text = ''

        all_texts.append(text)

        grupo_atual = None

        for chave, padrao in padrao_detalhamento.items():
            match = re.search(padrao, text)
            if match:
                valor = match.group(1).strip()

                if chave == "GRUPO DO TRIBUTO":
                    if not novo_grupo:
                        dctf_detalhamento["Valor da Multa"].append(soma_multas)
                        dctf_detalhamento["Valor dos Juros"].append(soma_juros)
                        soma_multas, soma_juros = 0, 0

                    grupo_atual = valor
                    novo_grupo = False

                elif chave == "Valor da Multa":
                    soma_multas += float(limpar_valor(valor).replace(',', '.'))
                    continue

                elif chave == "Valor dos Juros":
                    soma_juros += float(limpar_valor(valor).replace(',', '.'))
                    continue

                dctf_detalhamento[chave].append(valor)
            else:
                if chave not in ["Valor da Multa", "Valor dos Juros"]:
                    dctf_detalhamento[chave].append(None)

    if not novo_grupo:
        dctf_detalhamento["Valor da Multa"].append(soma_multas)
        dctf_detalhamento["Valor dos Juros"].append(soma_juros)

    max_len_detalhamento = max(len(v) for v in dctf_detalhamento.values())
    dctf_detalhamento_df = pd.DataFrame(
        {k: (v + [None] * (max_len_detalhamento - len(v))) for k, v in dctf_detalhamento.items()})

    dctf_detalhamento_df.to_excel(nome_saida_detalhamento, index=False)

    return dctf_detalhamento_df, all_texts