#import module
import psutil
import os
from datetime import datetime
from activaciones import function_activaciones
import time
hora=time.strftime('%H:%M', time.localtime()) #Se carga hora actual
count = 0
while str(hora)<'21:00':
  # check if chrome is open
  ejecutando="chrome.exe" in (i.name() for i in psutil.process_iter())

  if ejecutando:
    print("Se está ejecutando")
  else:
    
    # os.chdir("C:/Users/Brandon Cruz Romero/Documents/WFM/PY-CARIBU-AUTO/automatizacion-caribu-py")
    # os.startfile("Activaciones.bat")
    print("No se está ejecutando")
    try:

      function_activaciones()
    except:
      continue
  time.sleep(15)




