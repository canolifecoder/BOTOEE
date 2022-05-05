#!/usr/bin/python

import gspread
import pywhatkit
import datetime
import time


#CONEXION CON EL JSON:
gc = gspread.service_account(
    filename='enviocadencia-8ffd859b0128.json')
gc2 = gspread.service_account(
    filename='enviocadencia-f5961987ea87.json')

tiempo = datetime.datetime.now()
Dia = datetime.datetime.today().weekday()
Hora = int(tiempo.hour)
Minuto = int(tiempo.minute)+1
Minuto2 = int(tiempo.minute)
Segundo= int(tiempo.second)

sh2= gc2.open("DATOS UNIDADES PRODUCIDAS Z3")
sh = gc.open("Cadencia")

# Seleccionar primera hoja
worksheet = sh.get_worksheet(0)
worksheet2 = sh2.get_worksheet(0)

unidadesProducidas_list=worksheet2.col_values(2)

Unidades = unidadesProducidas_list[-1]
Unidades3 = unidadesProducidas_list[-2]
Unidades2 = int(Unidades)
Unidades4 = int(Unidades3)
unidadesFinal=0

if Unidades2>=Unidades4:
    UnidadesFinal=Unidades2
    print(UnidadesFinal)
if Unidades2<Unidades4:
    UnidadesFinal=Unidades4
    print(UnidadesFinal)

Archivo = open('UnidadesLinea.txt','r') # Abrir el archivo en modo lectura
a = Archivo.readlines() # Se lee las variables del archivo
Archivo.close() # Se cierra el archivo
DiaActual = int(a[1])

Dia = datetime.datetime.today().day # Actualizo el dìa actual desde el sistema
if DiaActual != Dia: # Ingresa si el día es diferente al ultimo día almacenado
    Archivo = open('UnidadesLinea.txt','w')
    a[0] = str(0)+'\n' # Dirección de apuntador en la hoja de Google Sheets
    a[1] = str(Dia)+'\n' # Actualizar día en el archivo
    Archivo.writelines(a)
    Archivo.close
    DiaActual = Dia # Actualización del día actual
if Minuto2==35:
    try:
        acumulado= int(a[0])
        Archivo.close
        acumulado2= acumulado+UnidadesFinal
        acumulado3 = str(acumulado2)
        Convertido=str(UnidadesFinal)
        Archivo = open('UnidadesLinea.txt','w')
        a[0] = str(acumulado3)+'\n' # Dirección de apuntador en la hoja de Google Sheets
        Archivo.writelines(a)
        Archivo.close
        print(acumulado3)

    except:
            print("Error, no se capturo el acumulado")