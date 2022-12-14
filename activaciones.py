# Desarrollo: Enrique y David
# Proyecto:   Activaciones Caribu
# Fecha:      Agosto 2022
# Prueba de Sincrinización 2022-09-01
                       
from multiprocessing import parent_process
# import string
import psutil
import pyperclip as clipboard
from selenium import webdriver                              #Se instancian las librerias necesarias
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
# from datetime import datetime
# from tkinter import messagebox
import pyodbc
# import xlrd
import time
# import glob
# import os
# import runpy
import pyautogui

def function_activaciones():

   
   hora=time.strftime('%H:%M', time.localtime()) #Se carga hora actual
   count = 0


   #Conexion a BDD Backoffice
   conex = pyodbc.connect('Driver={SQL Server};'
               'Server=10.10.12.245;'
               'Database=backoffice;'
               'UID=sa;'
               'PWD=C0nc3ntr42022*;'
               'Trusted_Connection=no;')
   cursor = conex.cursor()

   #Conexion a BDD Telemarketing
   conex2 = pyodbc.connect('Driver={SQL Server};'
               'Server=192.168.254.18;'
               'Database=TLMKT;'
               'UID=sa;'
               'PWD=M1nistr0;'
               'Trusted_Connection=no;')
   cursor_tlmkt = conex2.cursor()

   Telefono =''
   IdActivacion=''
   # while str(hora)<'21:00':
      # Recuperamos los registros de la tabla de usuarios
   cursor.execute("SELECT * FROM view_ventas_caribu_py WHERE PRIMER_Y_SEGUNDO_NOMBRE NOT LIKE '%YADIRA%'")

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
      PlanesFiltro = (row[10])
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
         

      cursor_tlmkt.execute(f"SELECT count(*) as conteo FROM [TLMKT].[dbo].[TBL_CARIBU] where DN='{Telefono}' AND MES IN(select cast(cast(CONVERT(varchar, getdate(), 112) as varchar(6)) as int) union select case when  day(getdate())<=5 then isnull((select cast(cast(CONVERT(varchar, DATEADD(MONTH,-1,getdate()), 112) as varchar(6)) as int)),0) end)")
      registros_validos = cursor_tlmkt.fetchall()

      #cursor.close()
      conteo=0
      for row2 in registros_validos:
         conteo=(row2[0])
      if conteo!=1:
         continue
      time.sleep(1)
      s=Service(ChromeDriverManager().install()) #abrir navegador
      driver = webdriver.Chrome(service=s)
      driver.maximize_window()
      driver.get('https://onix.movistar.com.mx:8443/login.action?ssoLogin=true') #Comienzan comandos selenium para interacción Web
      # time.sleep(1)
      # driver.refresh()
      # time.sleep(2)
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
      # os.system("pause")
      driver.find_element(By.XPATH, '//*[@id="serviceNo_input_value"]').send_keys(Telefono)
      driver.find_element(By.XPATH, '/html/body/div[1]/div[3]/table/tbody/tr[4]/td/div/div/span[2]/div').click()
      


      # driver.switch_to.default_content()
      # driver.switch_to.frame(1)
      # time.sleep(1)
      # driver.find_element(By.XPATH, '/html/body/div[1]/div[3]/div[1]/div[1]/ul/li[3]/div[2]/div/div/label').click()
      # mensajeError=driver.find_element(By.XPATH, '//*[@id="zBusinessAccept_Subscriber_head"]/div[2]').text
      # print(mensajeError)
      #time.sleep(2)
      #try:
      #driver.switch_to.default_content()
      #driver.switch_to.frame(1)
      #mensajeError=driver.find_element(By.XPATH, '//*[@id="zBusinessAccept_Subscriber_head"]/div[2]').text
      #if len(mensajeError)>5:
      #   print(mensajeError)
      #   messagebox.showinfo(message="Error localizado", title="OSC Concentra")
      #   continue
      #except:
      #   pass

      time.sleep(2)
      try:
         driver.switch_to.default_content()
         driver.switch_to.frame(30)  
         driver.find_element(By.XPATH, '/html/body/div[1]/div[4]/table/tbody/tr[2]/td/div/div[2]/div/table/tbody/tr[2]/td/div/div/div[2]/div[3]/table/tbody/tr[1]/td[10]/span/img').click()
         time.sleep(1)   
      except:
         pass
      time.sleep(2)
      try:
         driver.switch_to.frame(1)
         listaNegra=driver.find_element(By.XPATH, '//*[@id="editor_value"]').text
         # if len(listaNegra)>5:
         print(listaNegra)
         cursor.execute(f"update ventas_caribu set Fecha_captura = '{PlanesFiltro}', FechaCaptura = getdate(), Observaciones = '{listaNegra}', Estatus = 'No Activada' where IdActivacion = '{IdActivacion}'") 
         conex.commit()
      # messagebox.showinfo(message=listaNegra, title="OSC Concentra")
         time.sleep(2)
         continue
      except:
         pass
      time.sleep(2)
      driver.switch_to.default_content()
      driver.switch_to.frame(30)
      driver.find_element(By.CSS_SELECTOR, '#zBusinessAccept_Subscriber_title > label').click()
      time.sleep(2)
      driver.switch_to.frame(1)
      time.sleep(2)
      TotalRegistros=driver.find_element(By.XPATH, '/html/body/div[1]/div[3]/div[2]/div/table/tbody/tr[2]/td/div/table/tbody/tr[2]/td/div/div[2]/div[4]/div[1]/span').text
      if TotalRegistros != 'Registros Totales: 1':
         cursor.execute(f"update ventas_caribu set Fecha_captura = '{PlanesFiltro}', FechaCaptura = getdate(), Observaciones = 'Registros Totales mayor a 1', Estatus = 'No Activada' where IdActivacion = '{IdActivacion}'") 
         conex.commit()
         time.sleep(2)
         continue

      driver.switch_to.default_content()
      driver.switch_to.frame(30)
      time.sleep(2)
      driver.find_element(By.XPATH, '/html/body/div[1]/div[3]/div[1]/div[1]/ul/li[7]/div[2]/div/div[1]/label').click()
      time.sleep(2)
      # driver.switch_to.default_content()
      driver.switch_to.frame(2)  
      driver.find_element(By.XPATH, '/html/body/div[1]/div[2]/table/tbody/tr[2]/td/div/div[2]/div/table/tbody/tr[4]/td/div/span[2]/div').click()   # BOTON CONTINUAR
      time.sleep(3)   
      
      driver.find_element(By.XPATH, f'/html/body/div[1]/div[2]/table/tbody/tr[2]/td/div/div[2]/div/table/tbody/tr[2]/td[1]/div/div[3]/div/div/select/option[{TipoId}]').click()  # INE
      driver.find_element(By.XPATH, '//*[@id="field_500012_500018_input_value"]').clear()  # NUMID
      driver.find_element(By.XPATH, '//*[@id="field_500012_500018_input_value"]').send_keys(NumId)  # NUMID
      driver.find_element(By.XPATH, '//*[@id="field_500012_500033_input_value"]').clear()  # RFC
      driver.find_element(By.XPATH, '//*[@id="field_500012_500033_input_value"]').send_keys(RFC)  # RFC
      driver.find_element(By.XPATH, '//*[@id="field_500012_500034_input_value"]').clear()  # NOMBRES
      time.sleep(2)  
      driver.find_element(By.XPATH, '//*[@id="field_500012_500034_input_value"]').send_keys(PRIMER_Y_SEGUNDO_NOMBRE)  # NOMBRES
      time.sleep(2)  
      driver.find_element(By.XPATH, '//*[@id="field_500012_500035_input_value"]').clear()  # APATERNO
      driver.find_element(By.XPATH, '//*[@id="field_500012_500035_input_value"]').send_keys(APELLIDO_PATERNO)  # APATERNO
      driver.find_element(By.XPATH, '//*[@id="field_500012_500036_input_value"]').clear()  # AMATERNO
      driver.find_element(By.XPATH, '//*[@id="field_500012_500036_input_value"]').send_keys(Apellido_Materno)  # AMATERNO
      time.sleep(1)
      
      if GENERO=='1':
         driver.find_element(By.XPATH, '//*[@id="field_500012_500037_input_select"]/option[2]').click()  # GENERO MASCULINO
      elif GENERO=='2':
         driver.find_element(By.XPATH, '//*[@id="field_500012_500037_input_select"]/option[3]').click()  # GENERO FEMENINO
      time.sleep(3)
      driver.find_element(By.XPATH, '//*[@id="field_500012_500038_input_select"]/option[2]').click()  # TITULO SEÑOR
      driver.find_element(By.XPATH, '//*[@id="field_500012_500095_input_value"]').clear()  # FECHA NACIMIENTO
      driver.find_element(By.XPATH, '//*[@id="field_500012_500095_input_value"]').send_keys(fecha_de_nacimiento)  # FECHA NACIMIENTO
      driver.find_element(By.XPATH, '//*[@id="field_500012_500027_input_value"]').clear()  # EMAIL
      driver.find_element(By.XPATH, '//*[@id="field_500012_500027_input_value"]').send_keys(CORREO_ELECTRONICO)  # EMAIL
      time.sleep(2)
      driver.find_element(By.XPATH, '//*[@id="field_500012_500082_input_value"]').clear()  # LIMPIA NUMERO CONTACTO 1
      driver.find_element(By.XPATH, '//*[@id="field_500012_500082_input_value"]').send_keys(Telefono)  # NUMERO CONTACTO 1
      driver.find_element(By.XPATH, '//*[@id="field_500012_500119_input"]/div').click()  # SELECCIONA PF
      driver.find_element(By.XPATH, '/html/body/div[1]/div[2]/table/tbody/tr[2]/td/div/div[2]/div/table/tbody/tr[8]/td[3]/div/div[3]/div[1]/div/select/option[2]').click()  # CLICK EN PF
      driver.find_element(By.XPATH, '/html/body/div[1]/div[2]/table/tbody/tr[2]/td/div/div[2]/div/table/tbody/tr[9]/td[3]/div/div[4]/div/div/select/option[155]').click()  # SELECCIONA PAIS
      driver.find_element(By.XPATH, '/html/body/div[1]/div[2]/table/tbody/tr[4]/td/div/div[2]/div/table/tbody/tr[2]/td[1]/div/div[3]/div[1]/div/select/option[1]').click()  # LIMPIAR ESTADO
      time.sleep(2)
      driver.find_element(By.XPATH, '//*[@id="field_500001_500008_input_value"]').clear()  # LIMPIA CODIGO POSTAL
      driver.find_element(By.XPATH, '//*[@id="field_500001_500008_input_value"]').send_keys(CODIGO_POSTAL)  # CODIGO POSTAL
      driver.find_element(By.XPATH, '//*[@id="field_500001_500005_input_value"]').clear()  # CALLE 
      driver.find_element(By.XPATH, '//*[@id="field_500001_500005_input_value"]').send_keys(CALLE)  # CALLE 
      driver.find_element(By.XPATH, '//*[@id="field_500001_500006_input_value"]').clear()  # No EXTERNO
      driver.find_element(By.XPATH, '//*[@id="field_500001_500006_input_value"]').send_keys(NUMERO_EXTERNO)  # No EXTERNO
      driver.find_element(By.XPATH, '//*[@id="field_500001_500007_input_value"]').clear()  # LIMPIA No INTERNO
      driver.find_element(By.XPATH, '//*[@id="field_500001_500007_input_value"]').send_keys(NUMERO_INTERNO)  # No INTERNO
      time.sleep(5) 
      try: 
         driver.find_element(By.CSS_SELECTOR, '#officeAddressId_titlebar > label').click()
         time.sleep(1)
         domicilioTrabajo=driver.find_element(By.CSS_SELECTOR, '#officeAddressId_titlebar > label').text
         if domicilioTrabajo == 'Dirección de Trabajo':
            time.sleep(1)   
            driver.find_element(By.XPATH, '//*[@id="officeAddress_input_container"]/div/label').click() 
            time.sleep(1) 
      except:
         pass
      
      driver.find_element(By.XPATH, '//*[@id="field_500001_500005_input_value"]').click()
      driver.find_element(By.CSS_SELECTOR, '#creditCheck > div > div').click()      # BOTON CONSULTA CREDITO
      time.sleep(3)  # TIEMPO PARA QUE ESTE DISPONIBLE EL SIGUIENTE BOTON
      driver.find_element(By.CSS_SELECTOR, '#submitCustInfo > div > div').click()   # BOTON SIGUIENTE
      
      # INICIA PANTALLA DE ELECCION DE PLAN

      time.sleep(3)
      driver.find_element(By.CSS_SELECTOR, '#queryOffer_value').send_keys(PlanesFiltro)   # BOTON CONSULTA CREDITO    
      time.sleep(2)  
      driver.find_element(By.XPATH, '//*[@id="queryOffer_search"]').click()  # BOTON BUSQUEDA
      time.sleep(3)
      driver.find_element(By.XPATH, '/html/body/div[1]/div[3]/table/tbody/tr[4]/td/div/div[2]/table/tbody/tr[2]/td/div/div[2]/div/table/tbody/tr[2]/td/div/div[2]/div[3]/table/tbody/tr/td[1]/input').click()  # CHECK BOX PLAN
      time.sleep(3)
      driver.find_element(By.XPATH, '/html/body/div[1]/div[3]/table/tbody/tr[4]/td/div/div[2]/table/tbody/tr[3]/td/div').click()  # CHECK BOX PLAN
      driver.find_element(By.XPATH, '/html/body/div[1]/div[3]/table/tbody/tr[4]/td/div/div[2]/table/tbody/tr[3]/td/div/span[2]/div/div').click()  # CHECK BOX PLAN
      time.sleep(3)
      driver.find_element(By.XPATH, '/html/body/div[1]/div[3]/table/tbody/tr[2]/td/div/div[2]/div/table/tbody/tr[2]/td/div/table/tbody/tr[3]/td[1]/div/div[3]/div/div/select/option[32]').click()  # MEJOR OFERTA COMERCIAL
      time.sleep(2)
      driver.find_element(By.CSS_SELECTOR, '#accountInfo_title > label').click()  # CLICK PESTAÑA CUENTA
      driver.find_element(By.XPATH, '/html/body/div[1]/div[4]/div[2]/div/table/tbody/tr[3]/td/div/div[2]/div[2]/div/div/table/tbody/tr[3]/td/div/table/tbody/tr[2]/td/div/div/table/tbody/tr[2]/td/div/div[1]/div[1]/div/div/div[1]/div[1]/div/cite').click()  # CLICK NUEVA CUENTA
      time.sleep(4)
      driver.find_element(By.XPATH, '/html/body/div[1]/div[4]/div[2]/div/table/tbody/tr[3]/td/div/div[2]/div[2]/div/div/table/tbody/tr[2]/td/div/table/tbody/tr[4]/td/div/table/tbody/tr[2]/td/div/table/tbody/tr[2]/td/div/table/tbody/tr[2]/td/div/div/div[1]/div[1]/div/div/div[1]/div/input').click()  # CLICK INF CLIENTE
      time.sleep(2)
      driver.find_element(By.XPATH, '/html/body/div[1]/div[4]/div[2]/div/table/tbody/tr[3]/td/div/div[2]/div[2]/div/div/table/tbody/tr[2]/td/div/table/tbody/tr[4]/td/div/table/tbody/tr[2]/td/div/table/tbody/tr[2]/td/div/table/tbody/tr[2]/td/div/div/div[2]/div/table/tbody/tr[7]/td[1]/div/div[3]/div[1]/div/input').send_keys(Telefono)  # TELEFONO FACTURA
      driver.find_element(By.XPATH, '/html/body/div[1]/div[4]/div[2]/div/table/tbody/tr[3]/td/div/div[2]/div[2]/div/div/table/tbody/tr[2]/td/div/table/tbody/tr[4]/td/div/table/tbody/tr[2]/td/div/table/tbody/tr[2]/td/div/table/tbody/tr[2]/td/div/div/div[2]/div/table/tbody/tr[8]/td[2]/div/div[3]/div/div/select/option[23]').click()  # CFDI
      time.sleep(4)
      pyautogui.press('tab')
      time.sleep(1)
      pyautogui.press('enter')
      time.sleep(1)
      driver.find_element(By.XPATH, '/html/body/div[1]/div[4]/div[2]/div/table/tbody/tr[3]/td/div/div[2]/div[2]/div/div/table/tbody/tr[2]/td/div/table/tbody/tr[4]/td/div/table/tbody/tr[2]/td/div/table/tbody/tr[2]/td/div/table/tbody/tr[9]/td/div/div[1]/div[1]/div/div/div[1]/div/input').click()  # DOMICILIO FACTURA
      time.sleep(1)
      driver.find_element(By.XPATH, '/html/body/div[1]/div[4]/div[2]/div/table/tbody/tr[3]/td/div/div[2]/div[2]/div/div/table/tbody/tr[2]/td/div/table/tbody/tr[6]/td/div').click()  # FRAME BOTON CONFIRMAR
      driver.find_element(By.XPATH, '/html/body/div[1]/div[4]/div[2]/div/table/tbody/tr[3]/td/div/div[2]/div[2]/div/div/table/tbody/tr[2]/td/div/table/tbody/tr[6]/td/div/span[1]/div/div').click()  # BOTON CONFIRMAR
      time.sleep(4)
      pyautogui.press('tab')
      time.sleep(1)
      pyautogui.press('enter')
      time.sleep(1)

      
      driver.find_element(By.XPATH, '/html/body/div[1]/div[4]/div[2]/div/table/tbody/tr[3]/td/div/div[1]/div[1]/ul').click()  # CLIC FRAME                                 
      
      try:
         driver.find_element(By.XPATH, '/html/body/div[1]/div[4]/div[2]/div/table/tbody/tr[3]/td/div/div[2]/div[2]/div/div/table/tbody/tr[3]/td/div/table/tbody/tr[2]/td/div/div/table/tbody/tr[2]/td/div/div[2]/div/table/tbody/tr[2]/td/div/div/div[2]/div[3]/table/tbody/tr[1]/td[5]/span/label').click()
         checkActivo=driver.find_element(By.XPATH, '/html/body/div[1]/div[4]/div[2]/div/table/tbody/tr[3]/td/div/div[2]/div[2]/div/div/table/tbody/tr[3]/td/div/table/tbody/tr[2]/td/div/div/table/tbody/tr[2]/td/div/div[2]/div/table/tbody/tr[2]/td/div/div/div[2]/div[3]/table/tbody/tr[1]/td[5]/span/label').text
         print(f"AAAAAAAAEEEE{checkActivo}")
         if checkActivo == 'Activo/a':
            driver.find_element(By.XPATH, '/html/body/div[1]/div[4]/div[2]/div/table/tbody/tr[3]/td/div/div[2]/div[2]/div/div/table/tbody/tr[3]/td/div/table/tbody/tr[2]/td/div/div/table/tbody/tr[2]/td/div/div[2]/div/table/tbody/tr[2]/td/div/div/div[2]/div[3]/table/tbody/tr[1]/td[1]/input').click()  # CLICC CHECK 1 
            time.sleep(1)
            driver.find_element(By.XPATH, '/html/body/div[1]/div[4]/div[2]/div/table/tbody/tr[3]/td/div/div[2]/div[2]/div/div/table/tbody/tr[3]/td/div/table/tbody/tr[2]/td/div/div/table/tbody/tr[2]/td/div/div[2]/div/table/tbody/tr[2]/td/div/div/div[2]/div[3]/table/tbody/tr[2]/td[1]/input').click()  # CLIC CHECK 2    
            time.sleep(1)                                              
            driver.find_element(By.XPATH, '/html/body/div[1]/div[4]/div[2]/div/table/tbody/tr[3]/td/div/div[2]/div[2]/div/div/table/tbody/tr[3]/td/div/table/tbody/tr[2]/td/div/div/table/tbody/tr[2]/td/div/div[2]/div/table/tbody/tr[2]/td/div/div/div[2]/div[3]/table/tbody/tr[3]/td[1]/input').click()  # CLIC CHECK 3
         else:
            driver.find_element(By.XPATH, '/html/body/div[1]/div[4]/div[2]/div/table/tbody/tr[3]/td/div/div[2]/div[2]/div/div/table/tbody/tr[3]/td/div/table/tbody/tr[2]/td/div/div/table/tbody/tr[2]/td/div/div[2]/div/table/tbody/tr[2]/td/div/div/div[2]/div[3]/table/tbody/tr[1]/td[1]/input').click()  # CLICC CHECK 1 
            time.sleep(1)
            driver.find_element(By.XPATH, '/html/body/div[1]/div[4]/div[2]/div/table/tbody/tr[3]/td/div/div[2]/div[2]/div/div/table/tbody/tr[3]/td/div/table/tbody/tr[2]/td/div/div/table/tbody/tr[2]/td/div/div[2]/div/table/tbody/tr[2]/td/div/div/div[2]/div[3]/table/tbody/tr[2]/td[1]/input').click()  # CLIC CHECK 2     
         # # driver.find_element(By.XPATH, '/html/body/div[1]/div[4]/div[2]/div/table/tbody/tr[3]/td/div/div[1]/div[1]/ul').click()  # CLIC FRAME                            
         # driver.find_element(By.XPATH, '/html/body/div[1]/div[4]/div[2]/div/table/tbody/tr[3]/td/div/div[2]/div[2]/div/div/table/tbody/tr[3]/td/div/table/tbody/tr[2]/td/div/div/table/tbody/tr[2]/td/div/div[2]/div/table/tbody/tr[2]/td/div/div/div[2]/div[3]/table/tbody/tr[1]/td[1]/input').click()  # CLICC CHECK 1          
         # driver.find_element(By.XPATH, '/html/body/div[1]/div[4]/div[2]/div/table/tbody/tr[3]/td/div/div[2]/div[2]/div/div/table/tbody/tr[3]/td/div/table/tbody/tr[2]/td/div/div/table/tbody/tr[2]/td/div/div[2]/div/table/tbody/tr[2]/td/div/div/div[2]/div[3]/table/tbody/tr[2]/td[1]/input').click()  # CLIC CHECK 2                                                  
         # driver.find_element(By.XPATH, '/html/body/div[1]/div[4]/div[2]/div/table/tbody/tr[3]/td/div/div[2]/div[2]/div/div/table/tbody/tr[3]/td/div/table/tbody/tr[2]/td/div/div/table/tbody/tr[2]/td/div/div[2]/div/table/tbody/tr[2]/td/div/div/div[2]/div[3]/table/tbody/tr[3]/td[1]/input').click()  # CLIC CHECK 3        
      except:
         pass
      
      # try:
      #    driver.find_element(By.XPATH, '/html/body/div[1]/div[4]/div[2]/div/table/tbody/tr[3]/td/div/div[1]/div[1]/ul').click()  # CLIC FRAME                                 
      #    driver.find_element(By.XPATH, '/html/body/div[1]/div[4]/div[2]/div/table/tbody/tr[3]/td/div/div[2]/div[2]/div/div/table/tbody/tr[3]/td/div/table/tbody/tr[2]/td/div/div/table/tbody/tr[2]/td/div/div[2]/div/table/tbody/tr[2]/td/div/div/div[2]/div[3]/table/tbody/tr[1]/td[1]/input').click()  # CLICC CHECK 1                                       
      #    driver.find_element(By.XPATH, '/html/body/div[1]/div[4]/div[2]/div/table/tbody/tr[3]/td/div/div[2]/div[2]/div/div/table/tbody/tr[3]/td/div/table/tbody/tr[2]/td/div/div/table/tbody/tr[2]/td/div/div[2]/div/table/tbody/tr[2]/td/div/div/div[2]/div[3]/table/tbody/tr[2]/td[1]/input').click()  # CLIC CHECK 2   
      #    driver.find_element(By.XPATH, '/html/body/div[1]/div[4]/div[2]/div/table/tbody/tr[3]/td/div/div[2]/div[2]/div/div/table/tbody/tr[3]/td/div/table/tbody/tr[2]/td/div/div/table/tbody/tr[2]/td/div/div[2]/div/table/tbody/tr[2]/td/div/div/div[2]/div[3]/table/tbody/tr[3]/td[1]/input').click()  # CLIC CHECK 3
                                                         
      # except:
      #    pass   
      # os.system("pause")
            
      driver.find_element(By.XPATH, '/html/body/div[1]/div[4]/div[2]/div/table/tbody/tr[3]/td/div/div[1]/div[1]/ul').click()  # CLIC FRAME                                 
      time.sleep(4)
      driver.find_element(By.XPATH, '/html/body/div[1]/div[4]/div[2]/div/table/tbody/tr[3]/td/div/div[2]/div[2]/div/div/table/tbody/tr[3]/td/div/table/tbody/tr[2]/td/div/div/table/tbody/tr[3]/td/div/div[2]/div/table/tbody/tr[2]/td/div/div/div[2]/div[3]/table/tbody/tr[2]/td[5]/input').click()  # CLICK POSPAGO 
      time.sleep(2)
      driver.find_element(By.XPATH, '/html/body/div[1]/div[5]').click()  # CLICK FRAME 
      driver.find_element(By.XPATH, '/html/body/div[1]/div[5]/span[2]/div/div').click()  # CLICK BOTON SIGUIENTE
      time.sleep(5)
      driver.find_element(By.XPATH, '/html/body').click()  # CLICK FRAME 
      driver.find_element(By.XPATH, '/html/body/div[1]/div[3]/table/tbody/tr[4]/td/div/table/tbody/tr[4]/td/div/table/tbody/tr[2]/td[1]/div/span[2]/div/div').click()  # CLICK ENVIAR FOLIO
      time.sleep(5)
         
      #pyautogui.press('tab')
      #time.sleep(1)
      #pyautogui.press('enter')

      folio = driver.current_window_handle
      print(folio)
      driver.switch_to.window (folio)
      driver.find_element(By.XPATH, '//*[@id="winmsg0"]/div/div[2]/div[3]/div[3]/span[1]/div/div').click()  #CLICK ACEPTAR
      time.sleep(9)
      # pyautogui.press('tab')



      # pos = pyautogui.position()
      # print(f"EQUIS{pos[0]}")
      # print(f"y{pos[1]}")
      pyautogui.dragTo(383, 292, 0.5)
      # pos2=pos[0]-3
      time.sleep(1)
      pyautogui.dragRel(380, None, 0.5, button="left")
      time.sleep(1)
      pyautogui.hotkey('ctrl','c')
      # time.sleep(1)
      numFolio=clipboard.paste()
      
      numFolio=numFolio.replace(".","")
      esNumero = numFolio.isnumeric()
      print(esNumero)
      # time.sleep(1)
      if esNumero:

         cursor.execute(f"update ventas_caribu set Fecha_captura='{PlanesFiltro}', FechaCaptura=getdate(), Observaciones='Sin Observaciones', Estatus='Activada', folio='{numFolio}', Capturista='SYSPy' where IdActivacion = '{IdActivacion}'") 
         conex.commit()
      else:

         cursor.execute(f"update ventas_caribu set Fecha_captura='{PlanesFiltro}', FechaCaptura=getdate(), Observaciones='{numFolio}', Estatus='No Activada', folio='{numFolio}', Capturista='SYSPy' where IdActivacion = '{IdActivacion}'") 
         conex.commit()

      time.sleep(2)
      #os.system('pause')
      
      continue

# function_activaciones()


