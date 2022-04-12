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

while True:
    tiempo = datetime.datetime.now()
    Dia = datetime.datetime.today().weekday()
    Hora = int(tiempo.hour)
    Minuto = int(tiempo.minute)+1
    Minuto2 = int(tiempo.minute)
    Segundo= int(tiempo.second)

        #CONDICIONAL PARA SELECCIONAR LA HORA:
    if (Hora >=6 and Hora <= 22) and (Dia != 6): # Rango de horas de operación y el día de trabajo (Lun - Sab)
        if espera_Minuto:
            print("Esperando minuto de envio...")
            espera_Minuto= False

            #CONDICIONAL PARA SELECCIONAR EL MIN Y EL RANGO DE SEGUNDOS
        if (Minuto2==0 and Segundo>=45) and Segundo<=49:
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
            print("Cadencia:",cadencia_List[-1], "--HORA:", hora_list[-1],
                "-Unidades 1:",unidadesProducidas_list[-1],"-Unidades 2:", unidadesProducidas_list[-2])

            #Se introduce en variables -- operaciones
            cadencia1 = cadencia_List[-1]
            hora1 = hora_list[-1]
            Unidades = unidadesProducidas_list[-1]
            Unidades2 =unidadesProducidas_list[-2]
            ResultadoUnidad=0
            if Unidades>=Unidades2:
                ResultadoUnidad = Unidades
            #elif Unidades==Unidades2:
             #   ResultadoUnidad = Unidades
            elif Unidades<Unidades2:
                 ResultadoUnidad = Unidades2  
                 
            resultadocadencia= int(cadencia1)
            resultadounidades= int(ResultadoUnidad)
            hora2 = str(int(hora1)-1)
            eficiencia = resultadounidades/resultadocadencia*100
            redondeado = round(eficiencia,2)
            resultado = str(redondeado)
            print(ResultadoUnidad)
        if Minuto==3:

            #Se envia el mensaje por WPP

            try:
                mensaje = "*Informe linea de ensamble "+hora2+ "-"+hora1+ "*\nCadencia: " +cadencia1+ "\nUnidades Producidas: "+ ResultadoUnidad +"\nEficiencia de la linea: "+ resultado +"%"
                pywhatkit.sendwhatmsg_to_group(
                "Jsol5eaS9yX80IGakvF7bF", mensaje, Hora, Minuto, 50, True, 20)
                print("Mensaje enviado")
                print(mensaje)

                espera_Minuto = True
                espera_Hora = True
            except:
                print("Error!! El mensaje no pudo ser enviado")
                
        # if Minuto==50:
        #     #Se envia el mensaje por WPP
        #     try:
        #         mensaje2 = "*Recuerde diligenciar el formulario:*\nhttps://mail.google.com/mail/u/0/?ogbl#inbox"
        #         pywhatkit.sendwhatmsg_to_group(
        #             "LEsZN7aH2TVHRAzaZUp3qI", mensaje2, Hora, Minuto, 56, True, 20)
        #         print("Mensaje enviado")
        #         print(mensaje2)
        #         espera_Minuto = True
        #         espera_Hora = True
        #     except:
        #         print("Error!! El mensaje no pudo ser enviado")
        time.sleep(1)

    else:
        if espera_Hora:
            print ("Esperando la hora de envío...")
            espera_Hora = False