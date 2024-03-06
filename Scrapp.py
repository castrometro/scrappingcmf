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
from email.mime.base import MIMEBase
from email import encoders


fecha_de_hoy = time.strftime("%d/%m/%Y")
fecha_de_ayer = time.strftime("%d/%m/%Y", time.gmtime(time.time() - 86400))

def crear_excel():
    if not os.path.exists('hechos_esenciales.xlsx'):
        archivo = 'hechos_esenciales.xlsx'
        libro = openpyxl.Workbook()
        hoja = libro.active
        hoja.title = 'Hechos Esenciales'
        hoja.append(['Fecha', 'Hora', 'ID', 'Entidad', 'Materia', 'Enlace', 'ENVIADO(Y/N)'])
        libro.save(archivo)
        libro.close()
    else:
        print('El archivo "hechos_esenciales.xlsx" ya existe.')
    

def añadir_a_excel(datos):
    filas_agregadas = 0
    archivo = 'hechos_esenciales.xlsx'
    libro = openpyxl.load_workbook(archivo)
    hoja = libro.active
    for fila in datos:
        if fila[2] not in [celda.value for celda in hoja['C']]: #Si el ID no está en la columna C
            fila.append('N')
            hoja.append(fila)
            filas_agregadas += 1
        #else:
            #print('....FILA EXISTENTE.....')  
    print ('Filas agregadas: ', filas_agregadas)
    libro.save(archivo)
    libro.close()
    return 


def accederyobtenerdf():
    print('ACCEDIENDO A CMF.....')
    service = Service()

    options = webdriver.ChromeOptions()
    options.add_argument('--start-maximized')
    options.add_argument('--disable-extensions')

    driver = webdriver.Chrome(service=service, options=options)
    driver.get('https://www.cmfchile.cl/portal/principal/613/w3-channel.html')

    css_selector = 'button.btn.btn-outline-secondary.btn-sm'

    WebDriverWait(driver,20)
    time.sleep(5)

    #----VENTANA INICIAL--------
    #esperar a que el elemento sea clickeable, luego hacer click
    WebDriverWait(driver,10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, css_selector))).click()
    time.sleep(2)

    #----SCROL DOWN----------
    div_row_element = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.CSS_SELECTOR, "div.ntg-box-mb.animar.tab-pills-cmf")))
    driver.execute_script("arguments[0].scrollIntoView(true);",div_row_element)
    time.sleep(2)

    WebDriverWait(driver,10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/div[2]/div[3]/div/div[1]/div[1]/div/div/div[1]/div/table/tbody')))
    time.sleep(2)


    #-----EXTRAER DATOS--------
    print('....EXTRAYENDO DATOS.....')
    tabla = driver.find_element(By.XPATH, '/html/body/div[2]/div[2]/div[3]/div/div[1]/div[1]/div/div/div[1]/div/table')
    filas = tabla.find_elements(By.TAG_NAME, "tr")
    datos = []
    for i in range(3, len(filas)):
        elemento = filas[i].find_elements(By.TAG_NAME, "td")
        fecha_y_hora = elemento[0].text
        fecha = fecha_y_hora.split(' ')[0]
        hora = fecha_y_hora.split(' ')[1]
        id = elemento[1].text
        entidad = elemento[2].text
        materia = elemento[3].text
        enlace = elemento[1].find_element(By.TAG_NAME, "a")
        #print('fecha:',fecha)
        #print('hora:',hora)
        #print('id:',id)
        #print('entidad:',entidad)
        #print('materia:',materia)
        #print('enlace:', enlace.get_attribute('href'))
        #print('-----------------------------')
        fila = [fecha, hora, id, entidad, materia, enlace.get_attribute('href')]
        datos.append(fila)

    añadir_a_excel(datos)
    print('....FIN EXTRACCIÓN.....')
    return 0



def main():
    crear_excel()
    accederyobtenerdf()
    print('....EXCEL ACTUALIZADO.....')

        

if __name__ == "__main__":
    main()

