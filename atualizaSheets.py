# from dotenv import load_dotenv
import os

os.system("pip install selenium")
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
os.system("pip install pandas")
import pandas as pd
import time
os.system("pip install requests")
import requests



from datetime import datetime, timedelta

# Carregar variáveis do arquivo .env
# load_dotenv()

# Obter as variáveis de ambiente
login_user = os.getenv('BLUE_LOGIN_USER')
login_password = os.getenv('BLUE_LOGIN_PASSWORD')


# Conectando a coelba - Login
def login():
    # Clica no botão "LOGIN"
    try:
        login_button = driver.find_element(By.XPATH, "/html/body/app-root/app-header/header/nav/div[1]/div/div/button/span[1]")
        login_button.click()
        print("Botão de login clicado com sucesso!")
        time.sleep(5)  
    except Exception as e:
        print(f"Erro ao clicar no botão: {e}")

    # Preenche os campos de login
    try:
        # Localizar e preencher o campo de e-mail
        email_field = driver.find_element(By.XPATH, "/html/body/div[2]/div[2]/div/mat-dialog-container/app-dialog-login/mat-dialog-content/section/form/mat-horizontal-stepper/div[2]/div/div/mat-form-field[1]/div/div[1]/div[3]/input")  # Substitua "email" pelo ID correto
        email_field.send_keys(login_user)

        # Localizar e preencher o campo de senha
        password_field = driver.find_element(By.XPATH, "/html/body/div[2]/div[2]/div/mat-dialog-container/app-dialog-login/mat-dialog-content/section/form/mat-horizontal-stepper/div[2]/div/div/mat-form-field[2]/div/div[1]/div[3]/input")  # Substitua "password" pelo ID correto
        password_field.send_keys(login_password)

        print("Campos de login preenchidos com sucesso!")

        # Clicar no botão de envio (ou pressionar Enter)
        submit_button = driver.find_element(By.XPATH, "/html/body/div[2]/div[2]/div/mat-dialog-container/app-dialog-login/mat-dialog-content/section/form/mat-horizontal-stepper/div[2]/div/div/div[3]/app-neo-button/button/div")  # Substitua "submit" pelo ID correto
        submit_button.click()

        print("Credenciais enviadas com sucesso!")
    except Exception as e:
        print(f"Erro ao preencher os campos de login ou enviar as credenciais: {e}")


    # Aguarda a página carregar
    time.sleep(5)
    # Recupera o Bearer Token do Local Storage
    
    bearer_token = driver.execute_script("return window.localStorage.getItem('tokenNeSe');")

    return bearer_token

# Configurar o modo headless
chrome_options = Options()
chrome_options.add_argument("--headless")  # Executar em modo headless
chrome_options.add_argument("--window-size=1920,1080")
chrome_options.add_argument("--disable-blink-features=AutomationControlled")
chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/110.0.0.0 Safari/537.36")

driver = webdriver.Chrome(options=chrome_options)

# CONECTANDO NA COELBA

# Abrir a página
driver.get("https://agenciavirtual.neoenergia.com/#/login")


time.sleep(5)
element = WebDriverWait(driver, 30).until(
    lambda driver: driver.execute_script("return document.readyState") == "complete"
)

bearer_token = login()

# Pega o que está entre "ne": e antes da próxima vírgula
tokenNeSe = bearer_token.split(":")[1].split(",")[0].strip(' "{}')


documento_cliente = login_user


url_base = f"https://apineprd.neoenergia.com/imoveis/1.1.0/clientes/{documento_cliente}/ucs"

# Parâmetros a serem adicionados
params = {
    "documento": login_user,
    "canalSolicitante": "AGC",
    "distribuidora": "COELBA",
    "usuario": "WSO2_CONEXAO",
    "indMaisUcs": "X",
    "protocolo": "123",
    "opcaoSSOS": "S",
    "tipoPerfil": "1"
}


headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/110.0.0.0 Safari/537.36",
    "Authorization": "Bearer " + tokenNeSe
}


# Fazer a requisição GET
try:
    response = requests.get(url_base, params=params, headers=headers, timeout = 30)
except requests.exceptions.Timeout:
    print("Tempo de espera na geração das ucs esgotado. Tente novamente mais tarde.")
    
    

# Verifica se a resposta foi bem-sucedida
if response.status_code == 200:
    # Converte a resposta para JSON
    dados = response.json()

    # Extrai a lista de unidades consumidoras (ucs)
    ucs = dados.get("ucs", [])

    # Extrai os códigos 'uc' em uma lista
    codigos_uc = [uc['uc'] for uc in ucs]

    # Armazenar os códigos em um DataFrame
    df_codigos = pd.DataFrame(codigos_uc, columns=["codigo_uc"])

else:
    print(f"Erro {response.status_code} ao acessar a API: {response.text}")
    

# Obtendo o primeiro código do DataFrame para gerar um numero de protocolo
primeiro_codigo = df_codigos.iloc[0]['codigo_uc']

# URL da API de protocolo
url_protocolo = "https://apineprd.neoenergia.com/protocolo/1.1.0/obterProtocolo"

headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/110.0.0.0 Safari/537.36",
    "Authorization": "Bearer " + tokenNeSe  
}

# Parâmetros da requisição, substituindo o codCliente pelo primeiro código do DataFrame
params_protocolo = {
    "distribuidora": "COEL",
    "canalSolicitante": "AGC",
    "documento": login_user,
    "codCliente": primeiro_codigo,  # Substitui o codCliente pelo primeiro código
    "recaptchaAnl": "true",
    "regiao": "NE"
}

# Fazer a requisição GET para obter o protocolo
try:
    response_protocolo = requests.get(url_protocolo, headers=headers, params=params_protocolo, timeout = 30)
except requests.exceptions.Timeout:
    print("Tempo de espera na response_protocolo esgotado. Tente novamente mais tarde.")
    
    
# Verificar se a resposta foi bem-sucedida
if response_protocolo.status_code == 200:
    try:
        # Converter resposta para JSON
        dados_protocolo = response_protocolo.json()

        
        # Pegar o valor do protocolo gerado a partir das chaves corretas
        protocolo_salesforce = dados_protocolo.get('protocoloSalesforce', None)
        protocolo_legado = dados_protocolo.get('protocoloLegado', None)
                
    except Exception as e:
        print(f"Erro ao processar a resposta: {e}")
else:
    print(f"Erro ao obter o protocolo: {response_protocolo.status_code} - {response_protocolo.text}")



import pandas as pd
import requests

# URL da API
url = "https://apineprd.neoenergia.com/multilogin/2.0.0/servicos/faturas/ucs/faturas"

# Cabeçalhos com User-Agent
headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/110.0.0.0 Safari/537.36",
    "Authorization": "Bearer " + tokenNeSe
}

# Lista de códigos a serem consultados
codigos = df_codigos['codigo_uc'].tolist()  # Convertendo a coluna para uma lista

# Lista para armazenar os dados
dados_coletados = []

# Iterar sobre os códigos e coletar os dados
for codigo in codigos:
    # Parâmetros da requisição (mantendo os outros fixos)
    params = {
        "codigo": codigo,  # Passando o código individual
        "documento": login_user,
        "canalSolicitante": "AGC",
        "usuario": "WSO2_CONEXAO",
        "protocolo": protocolo_legado,
        "tipificacao": "",
        "byPassActiv": "X",
        "documentoSolicitante": login_user,
        "documentoCliente": login_user,
        "distribuidora": "COELBA",
        "tipoPerfil": "1"
    }

    # Fazer a requisição GET
    try:
        response = requests.get(url, headers=headers, params=params, timeout = 30)
    except requests.exceptions.Timeout:
        print("Tempo de espera na geração de faturas esgotado. Tente novamente mais tarde.")
    
    
    # Verifica se a resposta foi bem-sucedida
    if response.status_code == 200:
        # Converte a resposta para JSON
        dados = response.json()

        # Extrair lista de faturas
        lista_faturas = dados.get("faturas", [])
        entrega = dados.get("entregaFaturas", {})

        if lista_faturas:
            # Adicionar cada fatura à lista de dados coletados
            for fatura in lista_faturas:
                dados_coletados.append({
                    "codigo_cliente": codigo,
                    "mesReferencia": fatura.get("mesReferencia", "N/A"),
                    "numeroFatura": fatura.get("numeroFatura", "N/A"),
                    "vencimento": fatura.get("dataVencimento", "N/A"),
                    "situação": fatura.get("statusFatura", "N/A")
                })
        else:
            # Quando não houver faturas, adicionar o código com "N/A"
            dados_coletados.append({
                "codigo_cliente": codigo,
                "mesReferencia": "N/A",
                "numeroFatura": "N/A",
                "vencimento": "N/A",
                "situação": "N/A"
            })
    else:
        # Tratamento para erro 400 (Bad Request) ou outros erros
        #print(f"Erro {response.status_code} ao buscar código {codigo}: {response.text}")
        dados_coletados.append({
            "codigo_cliente": codigo,
            "mesReferencia": "N/A",
            "numeroFatura": "N/A",
            "vencimento": "N/A",
            "situação": "N/A"
        })


import pandas as pd
from datetime import datetime, timedelta

# Criar DataFrame geral
df_geral = pd.DataFrame(dados_coletados)

# 1. Converter a coluna 'vencimento' para datetime, tratando erros
# Primeiro, vamos substituir "N/A" por NaT (Not a Time)
df_geral['vencimento'] = pd.to_datetime(df_geral['vencimento'], errors='coerce')

# 2. Calcular a data de corte (início do mês de 2 meses atrás)
data_atual = datetime.now()
data_corte = (data_atual - timedelta(days=60)).replace(day=1)  # 2 meses atrás, dia 1

# 3. Filtrar todas as faturas vencidas (independente da data)
df_vencidas = df_geral[df_geral["situação"] == "Vencida"]

# 4. Filtrar faturas com vencimento >= data_corte (ignorando NaT)
df_recentes = df_geral[df_geral['vencimento'].notna() & (df_geral['vencimento'] >= data_corte)]

# 5. Identificar a última fatura de cada cliente (caso não esteja nos filtros anteriores)
# Ordenar por cliente e vencimento (faturas sem vencimento ficarão no final)
df_ordenado = df_geral[df_geral['vencimento'] != "N/A"]
df_ordenado = df_ordenado.sort_values(by=["codigo_cliente", "vencimento"], 
                                 ascending=[True, False], 
                                 na_position='last')



# Pegar a última fatura de cada cliente (mesmo que seja NaT)
df_ultimas = df_ordenado.drop_duplicates(subset=["codigo_cliente"], keep="first")

# 6. Juntar todos os filtros e remover duplicatas
df_final = pd.concat([df_vencidas, df_recentes, df_ultimas]).drop_duplicates()

# 7. Ordenar o resultado final por cliente e vencimento
df_final = df_final.sort_values(by=["codigo_cliente", "vencimento"], 
                              ascending=[True, False], 
                              na_position='last')

df_filtrado = df_final
# Supondo que df_geral já existe
df_filtrado['vencimento'] = (
    pd.to_datetime(df_geral['vencimento'], errors='coerce')  # Converte para datetime (NaN se inválido)
    .dt.strftime('%Y-%m-%d')  # Formata para string de data
    .fillna('N/A')  # Substitui NaN por 'N/A'
)

# 2. Calcular a data de corte (início do mês de 2 meses atrás)
data_atual = datetime.now()
data_corte = (data_atual - timedelta(days=60)).replace(day=1)  # 2 meses atrás, dia 1

# 3. Filtrar todas as faturas vencidas (independente da data)
df_vencidas = df_geral[df_geral["situação"] == "Vencida"]

# 4. Filtrar faturas com vencimento >= data_corte (ignorando NaT)
df_recentes = df_geral[df_geral['vencimento'].notna() & (df_geral['vencimento'] >= data_corte)]

# 5. Identificar a última fatura de cada cliente (caso não esteja nos filtros anteriores)
# Ordenar por cliente e vencimento (faturas sem vencimento ficarão no final)
df_ordenado = df_geral.sort_values(by=["codigo_cliente", "vencimento"], 
                                 ascending=[True, False], 
                                 na_position='last')

# Pegar a última fatura de cada cliente (mesmo que seja NaT)
df_ultimas = df_ordenado.drop_duplicates(subset=["codigo_cliente"], keep="first")

# 6. Juntar todos os filtros e remover duplicatas
df_final = pd.concat([df_vencidas, df_recentes, df_ultimas]).drop_duplicates()

# 7. Ordenar o resultado final por cliente e vencimento
df_final = df_final.sort_values(by=["codigo_cliente", "vencimento"], 
                              ascending=[True, False], 
                              na_position='last')

df_filtrado = df_final
# Supondo que df_geral já existe
df_filtrado['vencimento'] = (
    pd.to_datetime(df_geral['vencimento'], errors='coerce')  # Converte para datetime (NaN se inválido)
    .dt.strftime('%Y-%m-%d')  # Formata para string de data
    .fillna('N/A')  # Substitui NaN por 'N/A'
)
os.system("pip install gspread")
os.system("pip install oauth2client")
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# 🔹 Função para autenticar e abrir a planilha do Google Sheets
def autenticar_google_sheets():
    scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
    client_sheets = gspread.authorize(creds)
    return client_sheets

# 🔹 Função para escrever no Google Sheets sem afetar as colunas com fórmulas
def escrever_no_google_sheets(df, planilha_id, nome_aba="Faturas", intervalo="A1:E"):
    client_sheets = autenticar_google_sheets()
    planilha = client_sheets.open_by_key(planilha_id)
    aba = planilha.worksheet(nome_aba)

    # Apagar apenas o conteúdo do intervalo sem remover as fórmulas nas colunas seguintes
    aba.batch_clear([intervalo])

    # Convertendo o DataFrame para uma lista de listas
    dados = [df.columns.tolist()] + df.values.tolist()  # Inclui os cabeçalhos

    # Escrevendo os dados diretamente no intervalo correto
    aba.update(intervalo, dados, value_input_option="USER_ENTERED")

    print(f"✅ Dados escritos na planilha '{nome_aba}' no intervalo '{intervalo}' com sucesso!")

# 🔹 ID da planilha do Google Sheets (substitua pelo seu ID real)
SPREADSHEET_ID = "1Ut5Y0LstIP7nhv7Jzyywc7SS7ObIPlO-3yEg-J8Pp5o"

# 🔹 Escrever os dados filtrados na planilha no intervalo correto
escrever_no_google_sheets(df_filtrado, SPREADSHEET_ID, nome_aba="Faturas", intervalo="A1:E")
