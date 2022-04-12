
from ast import If
from tkinter.tix import Tree
from xmlrpc import client
import gspread
import time


# Introducir datos
""" worksheet.update('A1', 'segundo')
worksheet.update('A2', 'ENLACE')
worksheet.update('B1', 'prueba')
worksheet.update('B2', 'SOBREESCRIBIR')
worksheet.update('C1', 'ENLACESSSPS')
worksheet.update('C2', 'PROPUESTAS')
 """

import pywhatkit
import datetime
#import emoji

i = 0
espera_Minuto = True
espera_Hora = True

while i == 0:
    tiempo = datetime.datetime.now()
    Hora = int(tiempo.hour)
    Minuto = int(tiempo.minute)+1
    Minuto2 = int(tiempo.minute)
    Segundo= int(tiempo.second)
        #CONDICIONAL PARA SELECCIONAR LA HORA:
    if Hora >=6 and Hora <= 22:
        if espera_Minuto:
            print("Esperando minuto de envio...")
            espera_Minuto= False
    else:
        if espera_Hora:
            print ("Esperando la hora de envío...")
            espera_Hora = False
        i=0
        #CONDICIONAL PARA SELECCIONAR EL MIN Y EL RANGO DE SEGUNDOS
    if Minuto2==0 and Segundo>=45 and Segundo<=49:
        #CONEXION CON EL JSON:
        gc = gspread.service_account(
            filename='enviocadencia-8ffd859b0128.json')
        gc2 = gspread.service_account(
            filename='enviocadencia-f5961987ea87.json')

        # Abrir por titulo LA HOJA DE CALCULO
        sh = gc.open("Cadencia")
        sh2= gc2.open("DATOS UNIDADES PRODUCIDAS Z3")

        # Seleccionar primera hoja
        worksheet = sh.get_worksheet(0)
        worksheet2 = sh2.get_worksheet(0)

        # Se selecciona la columna deseada
        unidadesProducidas_list=worksheet2.col_values(2)
        cadencia_List = worksheet.col_values(3)
        hora_list = worksheet.col_values(4)

        #se imprime el mensaje antes de enviarlo:
        print(cadencia_List[-1], "--", hora_list[-1],
              "-----",unidadesProducidas_list[-1])

        #Se introduce en variables -- operaciones
        cadencia1 = cadencia_List[-1]
        hora1 = hora_list[-1]
        Unidades = unidadesProducidas_list[-1]
        resultadocadencia= int(cadencia1)
        resultadounidades= int(Unidades)
        hora2 = str(int(hora1)-1)
        eficiencia = resultadounidades/resultadocadencia*100
        redondeado = round(eficiencia,2)
        resultado = str(redondeado)
        
    if Minuto==3:
        #Se envia el mensaje por WPP
        try:
            mensaje = "*Informe linea de emsamble* "+hora2+ "-"+hora1+ "\n*Cadencia:* " +cadencia1+ "\n*Unidades Producidas:* "+ Unidades +"\n*Eficiencia de la linea:* "+ resultado +"%"
            pywhatkit.sendwhatmsg_to_group(
                "LEsZN7aH2TVHRAzaZUp3qI", mensaje, Hora, Minuto, 45, True, 20)
            print("Mensaje enviado")
            print(mensaje)
            espera_Minuto = True
            espera_Hora = True
        except:
            print("Error!! El mensaje no pudo ser enviado")
            i=0
    if Minuto==50:
        #Se envia el mensaje por WPP
        try:
            mensaje2 = "*Recuerde enviar el informe:*\nhttps://mail.google.com/mail/u/0/?ogbl#inbox"
            pywhatkit.sendwhatmsg_to_group(
                "LEsZN7aH2TVHRAzaZUp3qI", mensaje2, Hora, Minuto, 45, True, 20)
            print("Mensaje enviado")
            print(mensaje2)
            espera_Minuto = True
            espera_Hora = True
        except:
            print("Error!! El mensaje no pudo ser enviado")
    time.sleep(1)
    
else:
    print("No hubo accion", Minuto, "minutos")
i = 0
