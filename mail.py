from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
import time
import pandas as pd
import openpyxl
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import pandas as pd

def enviar_correo(destinatario, asunto):
    # Crea el mensaje
    mensaje = MIMEMultipart()
    mensaje['From'] = 'pablo.castro_d@outlook.com'
    mensaje['To'] = destinatario
    mensaje['Subject'] = asunto



    # Configuración del servidor SMTP de Outlook
    servidor = smtplib.SMTP('smtp.office365.com', 587)
    servidor.starttls()

    # Login al servidor
    servidor.login('pablo.castro_d@outlook.com', 'telefono2708AB_')

    # Enviar el correo
    servidor.sendmail(mensaje['From'], mensaje['To'], mensaje.as_string())
    servidor.quit()

    print("Correo enviado exitosamente!")



def main():
    enviar_correo('pablo.castro@servexternos.santander.cl', 'Prueba')




    

if __name__ == "__main__":
    main()

