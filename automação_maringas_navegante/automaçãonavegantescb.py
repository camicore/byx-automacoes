from selenium import webdriver
import pyautogui
import os
from datetime import datetime
import time 
import win32com.client as win32


from dotenv import load_dotenv
load_dotenv()



nome_pasta = "evidencias_baixa_remessa"

URL= os.getenv("URL")
USER= os.getenv("USER")
PWD= os.getenv("PWD")
REMETENTE = (os.getenv("REMETENTE_EMAIL"), os.getenv("REMETENTE_NOME"))
email_destino = os.getenv("DESTINO_EMAIL")

# Inicia o navegador
navegador = webdriver.Chrome()
navegador.get(URL)
navegador.maximize_window()

time.sleep(3)


# Login
navegador.find_element("xpath", "/html/body/div[1]/div/div/div/div[1]/div/div/form[1]/div[1]/div/input").send_keys(USER)
navegador.find_element("xpath", "/html/body/div[1]/div/div/div/div[1]/div/div/form[1]/button[1]/span[1]/div").click()

time.sleep(2)

# Senha
navegador.find_element("xpath", "/html/body/div[1]/div/div/div/div[1]/div/div/form[1]/div[2]/div/input").send_keys(PWD)
time.sleep(2)
navegador.find_element("xpath", "/html/body/div[1]/div/div/div/div[1]/div/div/form[1]/button[1]/span[1]/div").click()

time.sleep(5)
# Aguarda o carregamento da página
# Clico no convênio correto
navegador.find_element("xpath", "/html/body/div[1]/div/div/div/div[2]/div[1]/div/div/div[4]/div").click()

time.sleep(5)


# Cria a pasta se ela não existir
if not os.path.exists(nome_pasta):
    os.makedirs(nome_pasta)

# Obtém a data atual
data_atual = datetime.now().strftime("%d-%m-%y")

# Define o nome do arquivo
nome_arquivo = f"{nome_pasta}/navegantes_cb_dia_corte_{data_atual}.png"

time.sleep(2)

# Tira o print e salva
try:
    screenshot = pyautogui.screenshot()
    screenshot.save(nome_arquivo)

    print(f"Screenshot salvo em: {nome_arquivo}")
except Exception as e:
    print(f"Erro ao tirar o print: {e}")




# Caminho da imagem com data
data_hoje = datetime.today().strftime('%d-%m-%y')
caminho_imagem = f"evidencias_baixa_remessa/navegantes_cb_dia_corte_{data_hoje}.png"

# Verifica se a imagem existe
if not os.path.exists(caminho_imagem):
    raise FileNotFoundError("Imagem não encontrada!")

# Inicializa o Outlook
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)

# Preenche os campos do e-mail
mail.To = email_destino
mail.Subject = f"Evidência do dia - Prefeitura de Navegantes {data_hoje}."
mail.Body = (
    "Olá,\n\n"
    "Segue em anexo a evidência da data de corte - Prefeitura de Navegantes, Cartão Benefício.\n\n"
    "Atenciosamente."
)
mail.Attachments.Add(os.path.abspath(caminho_imagem))

# Exibe o e-mail (ou envie com mail.Send())
mail.display()  # Use mail.Send() to send directly
