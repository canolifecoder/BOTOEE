#!/usr/bin/python

import gspread
import pywhatkit
import datetime
import time

espera_Minuto = True
espera_Hora = True

#CONEXION CON EL JSON:
gc = gspread.service_account(
    filename='enviocadencia-8ffd859b0128.json')
gc2 = gspread.service_account(
    filename='enviocadencia-f5961987ea87.json')
gc3 = gspread.service_account(
    filename='named-haven-340115-ac55768dd57c.json')
gc4 = gspread.service_account(
    filename='named-haven-340115-8092bf8dd87f.json'
)


tiempo = datetime.datetime.now()
Dia = datetime.datetime.today().weekday()
Hora = int(tiempo.hour)
Minuto = int(tiempo.minute)+1
Minuto2 = int(tiempo.minute)
Segundo= int(tiempo.second)

sh = gc3.open("Gestión Célula Hora a Hora Células")
#HORA EMSABLE MECANISMOS::
worksheet = sh.get_worksheet(0)
HoraEM=worksheet.col_values(2)

#HORA CONJUNTO SUSPENCIÓN::
worksheet = sh.get_worksheet(1)
HoraCS=worksheet.col_values(2)

#HORA ENSAMBLE GABINETE:::
worksheet = sh.get_worksheet(2)
HoraEG=worksheet.col_values(2)

#HORA TESTEO FINAL::
worksheet = sh.get_worksheet(3)
HoraTF=worksheet.col_values(2)

#HORA TAPA MOVIL::
worksheet = sh.get_worksheet(4)
HoraTM=worksheet.col_values(2)

#HORA TAPA FIJA::
worksheet = sh.get_worksheet(5)
HoraTF2=worksheet.col_values(2)


Archivo2 = open('ControlHora.txt','r') # Abrir el archivo en modo lectura
b = Archivo2.readlines() # Se lee las variables del archivo
Archivo2.close() # Se cierra el archivo
HoraActual = int(b[7])

Hora = int(tiempo.hour)
Dia = datetime.datetime.today().day # Actualizo el dìa actual desde el sistema
if HoraActual != Hora: # Ingresa si el día es diferente al ultimo día almacenado
    Archivo2 = open('ControlHora.txt','w')
    b[7] = str(Hora)+'\n' # Dirección de apuntador en la hoja de Google Sheets
    Archivo2.writelines(b)
    Archivo2.close
    HoraActual = Hora # Actualización del día actual 
    print(Hora)

HoraEM2=str(HoraEM[-1])
print(HoraEM2)
Archivo2 = open('ControlHora.txt','w')
b[1] = str(HoraEM[-1]) + "\n"
b[2] = str(HoraCS[-1]) + "\n"
b[3] = str(HoraEG[-1]) + "\n"
b[4] = str(HoraTF[-1]) + "\n"
b[5] = str(HoraTM[-1]) + "\n"
b[6] = str(HoraTF2[-1]) + "\n"
Archivo2.writelines(b)
Archivo2.close
print("Escrito exitosamente")


#EMSAMBLE MECANISMOS:::::::::::::::
print("EMSAMBLE MECANISMOS:--------")
EmsambleMecanismos = sh.get_worksheet(0)
UnidadesFabricadasEM =  EmsambleMecanismos.col_values(5)
print("Unidades producidas: "+UnidadesFabricadasEM[-1])
MensajeUnidadesFabricadasEM ="*Unidades producidas*: "+ UnidadesFabricadasEM[-1]

EmsambleMecanismos = sh.get_worksheet(0)
HoraEM=EmsambleMecanismos.col_values(2)

Archivo2 = open('ControlHora.txt','r') # Abrir el archivo en modo lectura
b = Archivo2.readlines() # Se lee las variables del archivo
HoraGuardadaEM = str(b[1])
Archivo2.close() # Se cierra el archivo

print(HoraGuardadaEM)
print(HoraEM[-1])
if str(HoraEM[-1])+"\n" == HoraGuardadaEM:
    print ("No hay nuevos reportes en la Hora actual")
else:
    print("Se continua con el registro con normalidad")
