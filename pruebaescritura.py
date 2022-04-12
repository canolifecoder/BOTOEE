import datetime
import time

tiempo = datetime.datetime.now() 
Dia = datetime.datetime.today().weekday()
Hora = int(tiempo.hour)
Minuto = int(tiempo.minute)+1
Minuto2 = int(tiempo.minute)
Segundo= int(tiempo.second)


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

flag=0  

if  (Minuto2 == 9):
        acumulado= int(a[0])
        Archivo.close
        acumulado2= acumulado+5
        acumulado3 = str(acumulado2)
        Archivo = open('UnidadesLinea.txt','w')
        a[0] = str(acumulado3)+'\n' # Dirección de apuntador en la hoja de Google Sheets

        Archivo.writelines(a)
        Archivo.close
        print(acumulado3)


""" if Minuto2 != 25:
            flag=0 """