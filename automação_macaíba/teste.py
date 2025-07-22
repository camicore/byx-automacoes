from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pyautogui
import os
from datetime import datetime
import time
import win32com.client as win32
from dotenv import load_dotenv

load_dotenv()

# Variáveis de ambiente
nome_pasta = "evidencias_baixa_remessa"
URL = os.getenv("URL")
USER = os.getenv("USER")
PWD = os.getenv("PWD")
DESTINO_EMAIL = os.getenv("DESTINO_EMAIL")
COPIA_EMAIL = os.getenv("COPIA_EMAIL")

# Inicia navegador
navegador = webdriver.Chrome()
navegador.get(URL)
navegador.maximize_window()
time.sleep(3)

# Login
navegador.find_element(By.XPATH, "/html/body/div[2]/div/div[1]/div[3]/div[2]/div[1]/form/div[1]/div[1]/input").send_keys(USER)
time.sleep(2)
navegador.find_element(By.XPATH, "/html/body/div[2]/div/div[1]/div[3]/div[2]/div[1]/form/div[1]/div[2]/input").send_keys(PWD)
time.sleep(2)
pyautogui.press('enter')

# Aguarda renderização do calendário
WebDriverWait(navegador, 20).until(
    EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div[3]/div[3]/form[1]/div[2]/div/div/div[1]/div/div/div/div/div/div[2]/div/div[2]/div/table/tbody/tr/td/div/div/div[2]/div[1]/table/tbody/tr/td[6]"))
)
time.sleep(3)

# Seletores
td_xpath = "/html/body/div[1]/div[3]/div[3]/form[1]/div[2]/div/div/div[1]/div/div/div/div/div/div[2]/div/div[2]/div/table/tbody/tr/td/div/div/div[2]/div[2]/table/tbody/tr/td[5]/a/div"
evento_xpath = "/html/body/div[1]/div[3]/div[3]/form[1]/div[2]/div/div/div[1]/div/div/div/div/div/div[2]/div/div[2]/div/table/tbody/tr/td/div/div/div[2]/div[2]/table/tbody/tr/td[5]/a/div"  # IMPORTANTE: relativo ao td

# Localiza o td da data desejada
td_element = navegador.find_element(By.XPATH, td_xpath)

# Tenta encontrar o evento dentro do td
try:
    span = td_element.find_element(By.XPATH, evento_xpath)
    texto = span.text.strip()

    if texto == "Data de Corte 07/2025":
        print("✅ Evento encontrado corretamente no dia 10/07/2025.")

        # Tira print
        data_hoje = datetime.now().strftime("%d-%m-%y")
        if not os.path.exists(nome_pasta):
            os.makedirs(nome_pasta)
        nome_arquivo = f"{nome_pasta}/Pref_Macaíba_dia_corte_{data_hoje}.png"
        pyautogui.screenshot(nome_arquivo)
        print(f"🖼️ Print salvo em: {nome_arquivo}")
    else:
        raise Exception("Texto do evento diferente do esperado.")

except Exception as e:
    print(f"❌ Evento esperado não encontrado. Motivo: {e}")
    
    # Tira print
    data_hoje = datetime.now().strftime("%d-%m-%y")
    if not os.path.exists(nome_pasta):
        os.makedirs(nome_pasta)
    nome_arquivo = f"{nome_pasta}/Pref_Macaíba_dia_corte_{data_hoje}.png"
    pyautogui.screenshot(nome_arquivo)
    print(f"🖼️ Print salvo em: {nome_arquivo}")

    # Envia e-mail com print
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = DESTINO_EMAIL
    mail.CC = COPIA_EMAIL
    mail.Subject = f"⚠️ Evento não encontrado - Prefeitura Macaíba ({data_hoje})"
    mail.Body = (
        "Olá, Prezados!\n\n"
        "O evento 'Data de Corte 07/2025' não foi localizado no dia 10/07/2025.\n"
        "Segue em anexo a evidência da tela para verificação.\n\n"
        "Atenciosamente."
    )
    mail.Attachments.Add(os.path.abspath(nome_arquivo))
    mail.Display()  # Use mail.Send() para enviar automaticamente
    print("📧 E-mail gerado com sucesso.")