# Desarrollo: Enrique & David
# Proyecto:   Activaciones Caribu
# Fecha:      Agosto 2022
# Prueba de Sincrinización 2022-09-01 HOLA
                       
from multiprocessing import parent_process
import string
from selenium import webdriver                              #Se instancian las librerias necesarias
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from datetime import datetime
from tkinter import messagebox
import pyodbc
import xlrd
import time
import glob
import os
import runpy

hora=time.strftime('%H:%M', time.localtime()) #Se carga hora actual
count = 0
while str(hora)<'21:00':


      s=Service(ChromeDriverManager().install()) #abrir navegador
      driver = webdriver.Chrome(service=s)
      driver.maximize_window()

      #Conexion a BDD
      conex = pyodbc.connect('Driver={SQL Server};'
                     'Server=10.10.12.245;'
                     'Database=backoffice;'
                     'UID=sa;'
                     'PWD=C0nc3ntr42022*;'
                     'Trusted_Connection=no;')
      cursor = conex.cursor()

      # Recuperamos los registros de la tabla de usuarios
      cursor.execute("SELECT * FROM view_ventas_caribu_py")

      # Recorremos todos los registros con fetchall
      # y los volcamos en una lista de usuarios
      registros = cursor.fetchall()

      for row in registros:
            IdActivacion=(row[0])
            Telefono = (row[1])
            APELLIDO_PATERNO = (row[2])
            Apellido_Materno = (row[3])
            PRIMER_Y_SEGUNDO_NOMBRE = (row[4])
            CORREO_ELECTRONICO = (row[5])
            fecha_de_nacimiento = (row[6])
            fecha_de_nacimiento = fecha_de_nacimiento.strftime('%d/%m/%Y')
            CODIGO_POSTAL = (row[7])
            Plan = (row[10])
            GENERO = str(row[11])
            NumGenero = (row[11])
            Titulo = (row[11])       
            TipoId = (row[12])
            NumId  = (row[13])      
            RFC = (row[14])
            CALLE = (row[15])
            NUMERO_EXTERNO = (row[16])
            NUMERO_INTERNO = (row[17])
            residuo= (row[22])
            capturista="SYS"
            blanco=' '

            #count=count+1

      #cursor.close()

      driver.get('https://onix.movistar.com.mx:8443/login.action?ssoLogin=true') #Comienzan comandos selenium para interacción Web
      driver.find_element(By.XPATH, '/html/body/div[3]/div[2]/div[1]/form/table/tbody/tr[7]/td[2]/select/option[1]').click()
      driver.find_element(By.XPATH, '//*[@id="username"]').send_keys('AXM14045')
      driver.find_element(By.XPATH, '//*[@id="password"]').send_keys('Cari2022*')
      driver.find_element(By.XPATH, '/html/body/div[3]/div[2]/div[1]/form/table/tbody/tr[8]/td[2]/span/div/div').click()
      time.sleep(1)
      driver.find_element(By.CSS_SELECTOR, '#usm_continue > div:nth-child(1) > div:nth-child(1)').click()
      driver.find_element(By.CSS_SELECTOR, '#sitemap > div:nth-child(1)').click()
      driver.switch_to.default_content()
      driver.switch_to.frame(29)
      driver.find_element(By.CSS_SELECTOR, 'li.crm_sitemap_catalog_item:nth-child(8) > div:nth-child(2) > div:nth-child(1)').click()
      driver.find_element(By.CSS_SELECTOR, 'div.crm_sitemap_category:nth-child(3) > div:nth-child(2) > span:nth-child(4) > a:nth-child(1)').click()
      driver.switch_to.default_content()
      driver.switch_to.frame(30)
      driver.find_element(By.XPATH, '//*[@id="serviceNo_input_value"]').send_keys(Telefono)
      driver.find_element(By.XPATH, '/html/body/div[1]/div[3]/table/tbody/tr[4]/td/div/div/span[2]/div').click()

      time.sleep(2)
      try:
         driver.switch_to.frame(1)
         mensajeError=driver.find_element(By.XPATH, '//*[@id="zBusinessAccept_Subscriber_head"]/div[2]').text
         if len(mensajeError)>5:
            print(mensajeError)
            messagebox.showinfo(message="Error localizado", title="OSC Concentra")
            continue
      except:
         pass

      time.sleep(2)
      try:
         driver.switch_to.default_content()
         driver.switch_to.frame(30)  
         driver.find_element(By.XPATH, '/html/body/div[1]/div[4]/table/tbody/tr[2]/td/div/div[2]/div/table/tbody/tr[2]/td/div/div/div[2]/div[3]/table/tbody/tr[1]/td[10]/span/img').click()
         time.sleep(1)   
      except:
         pass
      time.sleep(2)
      # empieza
      driver.find_element(By.XPATH, '//*[@id="zBusinessAccept_Subscriber_title"]').click()
      try:
         driver.switch_to.frame(1)
         subscripcion=driver.find_element(By.XPATH, '//*[@id="datagridId_page"]/div[1]/span').text
         print(subscripcion)
         if subscripcion!='Registros Totales: 1':
            print(subscripcion)
            messagebox.showinfo(message="Error localizado", title="OSC Concentra")
            continue
      except:
         pass  
      # termina 
      time.sleep(2) 
      driver.switch_to.default_content()
      driver.switch_to.frame(30)  
      time.sleep(2) 
      driver.find_element(By.XPATH, '/html/body/div[1]/div[3]/div[1]/div[1]/ul/li[7]/div[2]/div/div[1]/label').click()
      time.sleep(2) 
      driver.switch_to.default_content()
      driver.switch_to.frame(1)  
      driver.find_element(By.XPATH, '//*[@id="AID_43920258"]/div/div').click()
      time.sleep(1)   
  
      driver.find_element(By.XPATH, f'/html/body/div[1]/div[2]/table/tbody/tr[2]/td/div/div[2]/div/table/tbody/tr[2]/td[1]/div/div[3]/div/div/select/option[{TipoId}]').click()  # INE
      driver.find_element(By.XPATH, '//*[@id="field_500012_500018_input_value"]').clear()  # NUMID
      driver.find_element(By.XPATH, '//*[@id="field_500012_500018_input_value"]').send_keys(NumId)  # NUMID
      driver.find_element(By.XPATH, '//*[@id="field_500012_500033_input_value"]').clear()  # RFC
      driver.find_element(By.XPATH, '//*[@id="field_500012_500033_input_value"]').send_keys(RFC)  # RFC
      driver.find_element(By.XPATH, '//*[@id="field_500012_500034_input_value"]').clear()  # NOMBRES
      driver.find_element(By.XPATH, '//*[@id="field_500012_500034_input_value"]').send_keys(PRIMER_Y_SEGUNDO_NOMBRE)  # NOMBRES
      driver.find_element(By.XPATH, '//*[@id="field_500012_500035_input_value"]').clear()  # APATERNO
      driver.find_element(By.XPATH, '//*[@id="field_500012_500035_input_value"]').send_keys(APELLIDO_PATERNO)  # APATERNO
      driver.find_element(By.XPATH, '//*[@id="field_500012_500036_input_value"]').clear()  # AMATERNO
      driver.find_element(By.XPATH, '//*[@id="field_500012_500036_input_value"]').send_keys(Apellido_Materno)  # AMATERNO
      time.sleep(1)
      driver.find_element(By.XPATH, '/html/body/div[1]/div[2]/table/tbody/tr[2]/td/div/div[2]/div/table/tbody/tr[4]/td[1]/div/div[3]/div[1]/div/select').click()  # GENERO
      time.sleep(1)  
      driver.find_element(By.XPATH, f'//*[@id="field_500012_500037_input_select"]/option[{GENERO}]').click()  # GENERO 
      time.sleep(2)  
      driver.find_element(By.XPATH, f'//*[@id="field_500012_500038_input_select"]/option[2]').click()  # TITULO
      driver.find_element(By.XPATH, '//*[@id="field_500012_500095_input_value"]').clear()  # FECHA NACIMIENTO
      driver.find_element(By.XPATH, '//*[@id="field_500012_500095_input_value"]').send_keys(fecha_de_nacimiento)  # FECHA NACIMIENTO
      driver.find_element(By.XPATH, '//*[@id="field_500012_500027_input_value"]').clear()  # EMAIL
      driver.find_element(By.XPATH, '//*[@id="field_500012_500027_input_value"]').send_keys(CORREO_ELECTRONICO)  # EMAIL
      
      
      time.sleep(999)   
      
   
         