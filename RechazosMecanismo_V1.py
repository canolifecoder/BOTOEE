import snap7
import snap7.client as c
import struct
from snap7.util import *
from snap7.types import *
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import time
import keyboard
import numpy as np
import os
import datetime

#Configuración de Spreadsheet
scope = ['https://www.googleapis.com/auth/drive','https://spreadsheets.google.com/feeds']
creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
client = gspread.authorize(creds)
sheet = client.open('Cadencia').worksheet('Cadencia_Linea')
sheet1 = client.open('Cadencia').worksheet('Rechazos_Mecanismo')

#Condiciones Iniciales
Flag=0
Flag1=0
MinutosActual=0
Cond=1

#Configuración del PLC
IP = '172.17.15.10'
RACK = 0
SLOT = 2

#Función de lectura del PLC
def ReadMemory(plc,Area,DB,byte,bit,datatype):
    result = plc.read_area(areas[Area],DB,byte,datatype)
    if datatype==S7WLBit:
        return get_bool(result,0,bit)   
    elif datatype==S7WLByte or datatype==S7WLWord:
        return get_int(result,0)
    elif datatype==S7WLReal:
        return get_real(result,0)
    elif datatype==S7WLDWord:
        return get_dword(result,0)
    else:
        return None

#Función de escritura al PLC
def WriteMemory(plc,Area,DB,byte,bit,datatype,value):
    result = plc.read_area(areas[Area],DB,byte,datatype)
    if datatype==S7WLBit:
        set_bool(result,0,bit,value)
    elif datatype==S7WLByte or datatype==S7WLWord:
        set_int(result,0,value)
    elif datatype==S7WLReal:
        set_real(result,0,value)
    elif datatype==S7WLDWord:
        set_dword(result,0,value)
    plc.write_area(areas[Area],DB,byte,result)

#Ciclo Principal
if __name__=='__main__':

    while Cond==1:
        try:

            plc = c.Client()
            plc.connect(IP, RACK, SLOT)

            respuesta = c.Client().get_connected()

            if respuesta == 0:
                print('PLC OK')

                Archivo = open('Datos1.txt','r') # Abrir el archivo en modo lectura
                a = Archivo.readlines() # Se lee las variables del archivo
                Archivo.close() # Se cierra el archivo
                DiaActual = int(a[2])
                
                Dia = datetime.datetime.today().day # Actualizo el dìa actual desde el sistema

                if DiaActual != Dia: # Ingresa si el día es diferente al ultimo día almacenado
                    Archivo = open('Datos1.txt','w')
                    a[1] = str(int(a[1])+24)+'\n' # Dirección de apuntador en la hoja de Google Sheets
                    a[2] = str(Dia)+'\n' # Actualizar día en el archivo
                    Archivo.writelines(a)
                    Archivo.close
                    DiaActual = Dia # Actualización del día actual

                Hora = time.strftime("%H:%M:%S") # Obtener la hora del sistema
                Fecha = time.strftime("%d/%m/%y") # Obtener la fecha del sistema
                Minutos = time.strftime("%M") # Obtener la fecha del sistema

                j=0
                Rechazos = np.zeros((24), dtype=int)

                if (int(Minutos) % 5 == 0) or (Flag==1):
                    
                    for i in range(0,24):

                        Flag=1
                        b = int(a[1])+i

                        Rechazos[i] = ReadMemory(plc,'DB',21,j,0,S7WLWord)
                        #time.sleep(1)
                        sheet1.update_cell(b,1, Fecha)  # Imprime la fecha en la hoja de google sheets
                        #time.sleep(1)
                        sheet1.update_cell(b,2, str(Rechazos[i])) # Imprime la cantidad de rechazos de mecanismos en google sheets
                        #time.sleep(1)
                        sheet1.update_cell(b,3, str(i)) # Imprime el rango de hora en google sheets
                        time.sleep(4)
                        j = j+2

                        print('Posicion',i)
                        print(j)
                        print(Rechazos[i])

                Flag=0
                #Archivo = open('Datos1.txt','w')
                #a[2] = str(Dia)+'\n' # Actualizar día actual
                #Archivo.writelines(a)
                #Archivo.close
                
                time.sleep(1)    

                if keyboard.is_pressed('p'): # Si presiono la tecla 'p' se detiene el proceso
                    print('Se termino la ejecución del programa')
                    Cond=0        

            else:
                print('PLC sin conexión')  

                if keyboard.is_pressed('p'): # Si presiono la tecla 'p' se detiene el proceso
                    print('Se termino la ejecución del programa')
                    Cond=0
                                      

        except:
            Hora = time.strftime("%H:%M:%S") # Obtener la hora del sistema
            Fecha = time.strftime("%d/%m/%y") # Obtener la fecha del sistema
            print(Fecha, Hora, 'Error de comunicación')

            if keyboard.is_pressed('p'): # Si presiono la tecla 'p' se detiene el proceso
                print('Se termino la ejecución del programa')
                Cond=0