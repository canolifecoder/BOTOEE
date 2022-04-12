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
f=1
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
        if (Minuto2==47 and Segundo>=45) and Segundo<=49:
            # Abrir por titulo LA HOJA DE CALCULO
            sh = gc.open("Cadencia")
            sh2= gc2.open("DATOS UNIDADES PRODUCIDAS Z3")

            if Minuto2 !=48:
                f=1
                print("SE CAMBIO LA FLAG==",f)

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
            Unidades2 = int(Unidades)
            resultadocadencia= int(cadencia1)
            resultadounidades= int(Unidades)
            hora2 = str(int(hora1)-1)
            eficiencia = resultadounidades/resultadocadencia*100
            redondeado = round(eficiencia,2)
            resultado = str(redondeado)
            
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
            espera_Minuto = True
            espera_Hora = True

        if  Minuto2 == 48 and f==1:
            try:
                acumulado= int(a[0])
                Archivo.close
                acumulado2= acumulado+Unidades2
                acumulado3 = str(acumulado2)
                Archivo = open('UnidadesLinea.txt','w')
                a[0] = str(acumulado3)+'\n' # Dirección de apuntador en la hoja de Google Sheets
                Archivo.writelines(a)
                Archivo.close
                print(acumulado3)
                espera_Minuto = True
                espera_Hora = True
                f=0
            except:
                 print("Error, no se capturo el acumulado")
            
            
        if Minuto==3:

            #Se envia el mensaje por WPP

            try:
                mensaje = "*Informe linea de ensamble "+hora2+ "-"+hora1+ "*\nCadencia: " +cadencia1+ "\nUnidades Producidas: "+ Unidades +"\nEficiencia de la linea: "+ resultado +"%"
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