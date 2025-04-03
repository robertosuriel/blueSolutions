import PyPDF2
import pdfplumber
import tabula
import pandas as pd
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
import io


pasta_id = "1wbPLpNj_h1i3nLCEhVx2vdYyDiIval-9"
nome_arquivo = "2025/02_007087960441.pdf"  
caminho_pdf_local = "temp_pdf.pdf"  


# Configurações globais
CONFIG = {
    "credentials_path": "credentials.json",
    "temp_pdf_path": "temp_pdf.pdf"
}

# 🔹 Função para autenticar no Google Drive 
def autenticar_drive():
    SCOPES = ["https://www.googleapis.com/auth/drive"]
    creds = service_account.Credentials.from_service_account_file(
        CONFIG["credentials_path"], scopes=SCOPES
    )
    return build("drive", "v3", credentials=creds)

# 🔹 Função para baixar PDF (com tratamento de erro melhorado)
def baixar_pdf_do_drive(pasta_id, nome_arquivo):
    try:
        drive_service = autenticar_drive()
        query = f"'{pasta_id}' in parents and name='{nome_arquivo}' and trashed=false"
        resultados = drive_service.files().list(q=query, fields="files(id,name)").execute()
        
        if not resultados.get('files'):
            raise FileNotFoundError(f"Arquivo '{nome_arquivo}' não encontrado")
            
        file_id = resultados['files'][0]['id']
        request = drive_service.files().get_media(fileId=file_id)
        pdf_bytes = io.BytesIO()
        downloader = MediaIoBaseDownload(pdf_bytes, request)
        
        while True:
            status, done = downloader.next_chunk()
            if done:
                break
                
        with open(CONFIG["temp_pdf_path"], "wb") as f:
            f.write(pdf_bytes.getvalue())
            
        return CONFIG["temp_pdf_path"]
        
    except Exception as e:
        print(f"❌ Erro ao baixar arquivo: {str(e)}")
        return None

def verificar_propriedades_pdf(caminho_pdf):
    with open(caminho_pdf, "rb") as f:
        reader = PyPDF2.PdfReader(f)
        num_paginas = len(reader.pages)
        info = reader.metadata
        
    print(f"📄 Número de Páginas: {num_paginas}")
    print(f"📋 Metadados do PDF: {info}")

def extrair_texto_pdf(caminho_pdf):
    with pdfplumber.open(caminho_pdf) as pdf:
        pagina = pdf.pages[0]  # Pegando a primeira página
        texto = pagina.extract_text()  # Extraindo texto da página
        
        print(texto)
        
def inspecionar_layout_pdf(caminho_pdf):
    with open(caminho_pdf, "rb") as f:
        reader = PyPDF2.PdfReader(f)
        pagina = reader.pages[0]
        
        # Obter as dimensões da página
        dimensoes = pagina.mediabox  # Ou use cropBox, se disponível
        print(f"📏 Dimensões da página: {dimensoes}")

def testar_colunas_individuais(caminho_pdf):
    print("🔍 Testando extração de cada coluna individualmente...")

    # Definindo os intervalos (ajuste conforme necessário)
    colunas = [
        "Itens da Fatura", "Unid.", "Quant.", "Preço Unit. (R$)", "Valor (R$)", 
        "PIS/COFINS (R$)", "Base Cálc. ICMS (R$)", "Alíquota ICMS (%)", "ICMS (R$)", 
        "Tarifa Unit. (R$)"
    ]

    areas_colunas = [
        [230, 0,   460, 90],  # Itens da Fatura
        [230, 90, 460, 120],  # Unid.
        [230, 120, 460, 160],  # Quant.
        [230, 160, 460, 205],  # Preço Unit. (R$)
        [230, 205, 460, 260],  # Valor (R$)
        [230, 260, 460, 310],  # PIS/COFINS (R$)
        [230, 310, 460, 350],  # Base Cálc. ICMS (R$)
        [230, 345, 460, 380],  # Alíquota ICMS (%)
        [230, 380, 460, 420],  # ICMS (R$)
        [230, 420, 460, 455],  # Tarifa Unit. (R$)
    ]
    
    for i, area in enumerate(areas_colunas):
        print(f"\n🔹 Testando coluna: {colunas[i]} (Área: {area})")

        tabela = tabula.read_pdf(
            caminho_pdf, pages="1", area=area, multiple_tables=True, encoding="ISO-8859-1"
        )

        if tabela:
            df = tabela[0]
            print(df.head())  # Mostra as primeiras linhas extraídas
        else:
            print("❌ Nenhum dado encontrado para esta coluna.")
            
def extrair_texto_com_visualizacao(caminho_pdf, bbox, i, ajustar_quebra):
    with pdfplumber.open(caminho_pdf) as pdf:
        pagina = pdf.pages[0]
        
        # Visualização (requer pillow)
#         img = pagina.to_image()
#         img.draw_rect(bbox, fill=None, stroke="red", stroke_width=2)
#         img.show()
        
        area = pagina.crop(bbox)
        texto = area.extract_text()

        # Se a opção estiver ativada, substituir a primeira quebra de linha
        if texto and ajustar_quebra:
            texto = texto.replace("\n", " ", 1)  

#         print(f"Texto da coluna {i}: {texto}")
        return texto  # Retorna o texto ajustado


# Lista com coordenadas e flag para ajustar a quebra de linha
area_colunas_tabela1 = [
    ([10, 230, 90, 450], False),  # Itens da Fatura
    ([90, 230, 110, 450], False),  # Unid.
    ([112, 230, 160, 450], False),  # Quant.
    ([162, 230, 205, 450], True), # Preço Unit. (R$)
    ([205, 230, 260, 450], True),  # Valor (R$)
    ([260, 230, 310, 450], True), # PIS/COFINS (R$)
    ([310, 230, 345, 450], True), # Base Cálc. ICMS (R$)
    ([350, 230, 375, 450], True), # Alíquota ICMS (%)
    ([375, 230, 420, 450], False),  # ICMS (R$)
    ([420, 230, 452, 450], True)  # Tarifa Unit. (R$)
]

area_colunas = area_colunas_tabela1


caminho_pdf = baixar_pdf_do_drive(pasta_id, nome_arquivo)

#verificar_propriedades_pdf(caminho_pdf)

#extrair_texto_pdf(caminho_pdf)

#inspecionar_layout_pdf(caminho_pdf)

#testar_colunas_individuais(caminho_pdf)

#texto = extrair_texto_com_visualizacao(caminho_pdf, area)



df = []

dados_fatura = {}  

for i, (area, ajustar_quebra) in enumerate(area_colunas, start=1):
    coluna_texto = extrair_texto_com_visualizacao(caminho_pdf, area, i, ajustar_quebra)
    
    if coluna_texto:  # Garantir que não seja None ou vazio
        linhas = coluna_texto.split("\n")  # Separar as linhas
        
        # Primeiro item como cabeçalho, restante como valores
        cabecalho = linhas[0] if linhas else f"Coluna_{i}"  
        valores = linhas[1:] if len(linhas) > 1 else [""]  # Evitar colunas vazias
        
        # Adiciona ao dicionário
        dados_fatura[cabecalho] = valores

# Criar DataFrame a partir do dicionário
df = pd.DataFrame.from_dict(dados_fatura, orient="index").transpose()

# Exibir o DataFrame
print(df)