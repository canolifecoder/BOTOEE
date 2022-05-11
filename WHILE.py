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

SinNovedad=0
while True:
    tiempo = datetime.datetime.now()
    Dia = datetime.datetime.today().weekday()
    Hora = int(tiempo.hour)
    Minuto = int(tiempo.minute)+1
    Minuto2 = int(tiempo.minute)
    Segundo= int(tiempo.second)

    #CONDICIONAL PARA SELECCIONAR LA HORA:
    if (Hora>=6 and Hora<=22) and (Dia != 6): # Rango de horas de operación y el día de trabajo (Lun - Sab)
        if espera_Minuto:
            print("1 Esperando minuto de envio...")
            espera_Minuto= False

        if (Minuto2==0 and Segundo>=35) and Segundo<=50:

            sh = gc3.open("Gestión Célula Hora a Hora Células")
            #HORA ENSAMBLE MECANISMOS::
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
            
            # Abrir por titulo LA HOJA DE CALCULO
            sh2 = gc.open("Cadencia")
            sh3= gc2.open("DATOS UNIDADES PRODUCIDAS Z3")
            worksheet2 = sh3.get_worksheet(0)
            unidadesProducidas_list=worksheet2.col_values(2)
            Unidades = unidadesProducidas_list[-1]
            Unidades3 = unidadesProducidas_list[-2]
            Unidades2 = int(Unidades)
            Unidades4 = int(Unidades3)
            unidadesFinal=0
            f=1
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

        if Minuto2 == 1 and f==1:
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
                acumulado3="*Unidades acumuladas:* "+acumulado3
                f=0
                if espera_Minuto:
                    print("1 Esperando minuto de envio...")
                    espera_Minuto= False
            except:
                 print("Error, no se capturo el acumulado")    
        
            #CONDICIONAL PARA SELECCIONAR EL MIN Y EL RANGO DE SEGUNDOS
        if (Minuto2==3 and Segundo>=40) and Segundo<=45:
            ##CADENCIA:::::
            sh = gc3.open("Gestión Célula Hora a Hora Células")
            cadencia = gc.open("Cadencia")
            cadencia = cadencia.get_worksheet(0)
            cadenciaList = cadencia.col_values(3)
            print("Cadencia de la linea: "+cadenciaList[-1])
            MensajeCadencia = "*Cadencia de la linea:* "+ cadenciaList[-1]

            if str(HoraEM[-1])+"\n" == HoraGuardadaEM:
                print ("Ensamble mecanismos, no reporto")
                mensaje="*ENSAMBLE MECANISMOS*/n*La celula no reporto.*"
            else:
                #EMSAMBLE MECANISMOS:::::::::::::::
                print("EMSAMBLE MECANISMOS:--------")
                EmsambleMecanismos = sh.get_worksheet(0)
                UnidadesFabricadasEM =  EmsambleMecanismos.col_values(5)
                print("Unidades producidas: "+UnidadesFabricadasEM[-1])
                MensajeUnidadesFabricadasEM ="*Unidades producidas*: "+ UnidadesFabricadasEM[-1]

                EmsambleMecanismos = sh.get_worksheet(0)
                ParoProgramadoEM =  EmsambleMecanismos.col_values(8)
                if ParoProgramadoEM[-1]=="Si":
                    RazonParoProgramadoEM =  EmsambleMecanismos.col_values(9)
                    TiempoParoProgramadoEM =  EmsambleMecanismos.col_values(10)
                    mensajeParoProgramadoEM="*Paro programado - Tiempo:* "+TiempoParoProgramadoEM[-1]+" min, *Razon:* "+RazonParoProgramadoEM[-1]
                    print (mensajeParoProgramadoEM)
                else:
                    mensajeParoProgramadoEM=""
                    print(mensajeParoProgramadoEM)

                IncidentesEM =  EmsambleMecanismos.col_values(11)
                if IncidentesEM[-1]=="Si":
                    DescrIncidenteEM =  EmsambleMecanismos.col_values(13)
                    ValidarParoIncidentesEM =  EmsambleMecanismos.col_values(14)
                    if ValidarParoIncidentesEM[-1]=="Si":    
                        TiempoIncidenteEM =  EmsambleMecanismos.col_values(15)
                        mensajeIncidenteEM="*Incidente y/o accidente ambiental y/o SST - Tiempo:* "+TiempoIncidenteEM[-1]+" min, *Razon:* "+DescrIncidenteEM[-1]
                    else:
                        mensajeIncidenteEM="*Incidente y/o accidente ambiental y/o SST - Razon:* "+DescrIncidenteEM[-1]
                    print (mensajeIncidenteEM)
                else:
                    mensajeIncidenteEM=""
                    print(mensajeIncidenteEM)

                ServiciosPublicosEM =  EmsambleMecanismos.col_values(16)
                if ServiciosPublicosEM[-1]=="Si":
                    DescrServiciosPublicosEM =  EmsambleMecanismos.col_values(18)
                    TiempoServiciosPublicosEM =  EmsambleMecanismos.col_values(17)
                    mensajeServiciosPublicosEM="*afectacion por falta de servicios públicos - Tiempo:* "+TiempoServiciosPublicosEM[-1]+" min, *Razon:* "+DescrServiciosPublicosEM[-1]+""
                    print (mensajeServiciosPublicosEM)
                else:
                    mensajeServiciosPublicosEM=""
                    print(mensajeServiciosPublicosEM)

                MaquinaEM =  EmsambleMecanismos.col_values(19)
                if MaquinaEM[-1]=="Si":
                    DescrMaquinaEM=  EmsambleMecanismos.col_values(22)
                    TiempoMaquinaEM =  EmsambleMecanismos.col_values(20)
                    mensajeMaquinaEM="*afectacion por maquina - Tiempo:* "+TiempoMaquinaEM[-1]+" min, *Razon:* "+DescrMaquinaEM[-1]+""
                    print (mensajeMaquinaEM)
                else:
                    mensajeMaquinaEM=""
                    print(mensajeMaquinaEM)

                ManoObraEM =  EmsambleMecanismos.col_values(23)
                if ManoObraEM[-1]=="Si":
                    DescrManoObraEM=  EmsambleMecanismos.col_values(27)
                    TiempoManoObraEM =  EmsambleMecanismos.col_values(24)
                    mensajeManoObraEM="*Hubo afectacion en las unidades por Mano de Obra - Tiempo:* "+TiempoManoObraEM[-1]+" min, *Razon:* "+DescrManoObraEM[-1]+""
                    print (mensajeManoObraEM)
                else:
                    mensajeManoObraEM=""
                    print(mensajeManoObraEM)


                MateriaPrimaEM =  EmsambleMecanismos.col_values(28)
                if MateriaPrimaEM[-1]=="Si":
                    DescrMateriaPrimaEM=  EmsambleMecanismos.col_values(32)
                    TiempoMateriaPrimaEM =  EmsambleMecanismos.col_values(29)
                    mensajeMateriaPrimaEM="*Hubo afectacion en las unidades por Materia Prima - Tiempo:* "+TiempoMateriaPrimaEM[-1]+" min, *Razon:* "+DescrMateriaPrimaEM[-1]+""
                    print (mensajeMateriaPrimaEM)
                else:
                    mensajeMateriaPrimaEM=""
                    print(mensajeMateriaPrimaEM)

                UnidadesPorMetodoEM =  EmsambleMecanismos.col_values(33)
                if UnidadesPorMetodoEM[-1]=="Si":
                    DescrUnidadesPorMetodoEM=  EmsambleMecanismos.col_values(36)
                    TiempoUnidadesPorMetodoEM =  EmsambleMecanismos.col_values(34)
                    mensajeUnidadesPorMetodoEM="*Hubo afectacion en las unidades por Metodo - Tiempo:* "+TiempoUnidadesPorMetodoEM[-1]+" min, *Razon:* "+DescrUnidadesPorMetodoEM[-1]+""
                    print (mensajeUnidadesPorMetodoEM)
                else:
                    mensajeUnidadesPorMetodoEM=""
                    print(mensajeUnidadesPorMetodoEM)

                ScrapEM =  EmsambleMecanismos.col_values(37)
                if ScrapEM[-1]=="Si":
                    DescrScrapEM=  EmsambleMecanismos.col_values(38)
                    CantidadScrapEM =  EmsambleMecanismos.col_values(39)
                    mensajeScrapEM="*Se genero SCRAP - Cantidad:* "+CantidadScrapEM[-1]+", *Razon:* "+DescrScrapEM[-1]+""
                    print (mensajeScrapEM)
                else:
                    mensajeScrapEM=""
                    print(mensajeScrapEM)


                UnidadesReprocesadasEM =  EmsambleMecanismos.col_values(41)
                if UnidadesReprocesadasEM[-1]=="Si":
                    CantidadReprocesadasEM=  EmsambleMecanismos.col_values(42)

                    mensajeUnidadesReprocesadasEM="*Se reprocesaron unidades - Cantidad:* "+CantidadReprocesadasEM[-1]+""
                    print (mensajeUnidadesReprocesadasEM)
                else:
                    mensajeUnidadesReprocesadasEM=""
                    print(mensajeUnidadesReprocesadasEM)
                    
                OeeEM= EmsambleMecanismos.col_values(48)
                OeeEmsambleMecanismos = OeeEM[-1]
                mensajeOeeEM = "OEE: "+ OeeEM[-1]
                print("*OEE*: "+OeeEmsambleMecanismos)
                
                mensaje="*GESTION CELULA HORA A HORA:*" 

                if MensajeCadencia!="":
                    mensaje = mensaje + "\n\n" +MensajeCadencia
                mensaje=mensaje+"\n*EMSAMBLE MECANISMOS*"
                if MensajeUnidadesFabricadasEM!="":
                    mensaje=mensaje+"\n\n"+MensajeUnidadesFabricadasEM
                if mensajeParoProgramadoEM!="":
                    mensaje=mensaje+"\n"+mensajeParoProgramadoEM
                    SinNovedad=2
                if mensajeIncidenteEM!="":
                    mensaje=mensaje+"\n"+mensajeIncidenteEM
                    SinNovedad=2
                if mensajeServiciosPublicosEM!="":
                    mensaje=mensaje+"\n"+mensajeServiciosPublicosEM
                    SinNovedad=2
                if mensajeMaquinaEM!="":
                    mensaje=mensaje+"\n"+mensajeMaquinaEM
                    SinNovedad=2
                if mensajeManoObraEM!="":
                    mensaje=mensaje+"\n"+mensajeManoObraEM
                    SinNovedad=2
                if mensajeMateriaPrimaEM!="":
                    mensaje=mensaje+"\n"+mensajeMateriaPrimaEM
                    SinNovedad=2
                if mensajeUnidadesPorMetodoEM!="":
                    mensaje=mensaje+"\n"+mensajeUnidadesPorMetodoEM
                    SinNovedad=2
                if mensajeUnidadesPorMetodoEM!="":
                    mensaje=mensaje+"\n"+mensajeUnidadesPorMetodoEM
                    SinNovedad=2
                if mensajeScrapEM!="":
                    mensaje=mensaje+"\n"+mensajeScrapEM
                    SinNovedad=2
                if mensajeUnidadesReprocesadasEM!="":
                    mensaje=mensaje+"\n"+mensajeUnidadesReprocesadasEM
                    SinNovedad=2

                if SinNovedad!=2:
                    mensaje=mensaje+"\n*No se reportaron novedades*"
                
                SinNovedad=0

                if mensajeOeeEM!="" and OeeEmsambleMecanismos!="#DIV/0!":
                    mensaje=mensaje+"\n"+mensajeOeeEM
            
    
#-----------------------------------------------------------------------------------------------------------
        if espera_Minuto:
            print("2 Esperando minuto de envio...")
            espera_Minuto= False

        #CONDICIONAL PARA SELECCIONAR EL MIN Y EL RANGO DE SEGUNDOS
        if (Minuto2==4 and Segundo>=40) and Segundo<=45:
            #VALIDAR QUE SI HAYA UN REPORTE
            if str(HoraCS[-1])+"\n" == HoraGuardadaCS:
                print ("Conjunto suspencion, no reporto")
                mensaje2="*CONJUNTO SUSPENCION*/n*La celula no reporto.*"
            else:
                #CONJUNTO SUSPENCIÓN:::::::::::::::
                print("CONJUNTO SUSPENCIÓN:---------")

                ConjuntoSuspencion = sh.get_worksheet(1)

                UnidadesFabricadasCJ=  ConjuntoSuspencion.col_values(5)
                print("Unidades Fabricadas: "+UnidadesFabricadasCJ[-1])
                MensajeUnidadesFabricadasCJ="*Unidades Producidas:* "+ UnidadesFabricadasCJ[-1]
                #PAROS PROGRAMADOS::
                ParoProgramadoCS =  ConjuntoSuspencion.col_values(8)
                if ParoProgramadoCS[-1]=="Si":
                    RazonParoProgramadoCS =  ConjuntoSuspencion.col_values(9)
                    TiempoParoProgramadoCS =  ConjuntoSuspencion.col_values(10)
                    mensajeParoProgramadoCS="*Paro programado - Tiempo:* "+TiempoParoProgramadoCS[-1]+" min, *Razon:* "+RazonParoProgramadoCS[-1]
                    print (mensajeParoProgramadoCS)
                else:
                    mensajeParoProgramadoCS=""
                    print(mensajeParoProgramadoCS)

                #INCIDENTES SST:::
                IncidentesCS =  ConjuntoSuspencion.col_values(11)
                if IncidentesCS[-1]=="Si":
                    DescrIncidenteCS =  ConjuntoSuspencion.col_values(13)
                    ValidarParoIncidentesCS =  ConjuntoSuspencion.col_values(14)
                    if ValidarParoIncidentesCS[-1]=="Si":    
                        TiempoIncidenteCS =  ConjuntoSuspencion.col_values(16)
                        mensajeIncidenteCS="*Incidente y/o accidente ambiental y/o SST - Tiempo:* "+TiempoIncidenteCS[-1]+" min, *Razon:* "+DescrIncidenteCS[-1]
                    else:
                        mensajeIncidenteCS="*Incidente y/o accidente ambiental y/o SST - Razon:* "+DescrIncidenteCS[-1] + ", No se generó paro."
                    print (mensajeIncidenteCS)
                else:
                    mensajeIncidenteCS=""
                    print(mensajeIncidenteCS)

                #SERVICIOS PUBLICOS::::
                ServiciosPublicosCS =  ConjuntoSuspencion.col_values(16)
                if ServiciosPublicosCS[-1]=="Si":
                    DescrServiciosPublicosCS =  ConjuntoSuspencion.col_values(18)
                    TiempoServiciosPublicosCS =  ConjuntoSuspencion.col_values(17)
                    mensajeServiciosPublicosCS="*afectacion por falta de servicios públicos - Tiempo:* "+TiempoServiciosPublicosCS[-1]+" min, *Razon:* "+DescrServiciosPublicosCS[-1]+""
                    print (mensajeServiciosPublicosCS)
                else:
                    mensajeServiciosPublicosCS=""
                    print(mensajeServiciosPublicosCS)

                #MAQUINA:::::::::::
                MaquinaCS =  ConjuntoSuspencion.col_values(19)
                if MaquinaCS[-1]=="Si":
                    DescrMaquinaCS=  ConjuntoSuspencion.col_values(22)
                    TiempoMaquinaCS =  ConjuntoSuspencion.col_values(20)
                    mensajeMaquinaCS="*afectacion por maquina - Tiempo:* "+TiempoMaquinaCS[-1]+" min, *Razon:* "+DescrMaquinaCS[-1]+""
                    print (mensajeMaquinaCS)
                else:
                    mensajeMaquinaCS=""
                    print(mensajeMaquinaCS)

                #MANO DE OBRA:::::::
                ManoObraCS =  ConjuntoSuspencion.col_values(23)
                if ManoObraCS[-1]=="Si":
                    DescrManoObraCS=  ConjuntoSuspencion.col_values(26)
                    TiempoManoObraCS =  ConjuntoSuspencion.col_values(24)
                    mensajeManoObraCS="*Hubo afectacion en las unidades por Mano de Obra - Tiempo:* "+TiempoManoObraCS[-1]+" min, *Razon:* "+DescrManoObraCS[-1]+""
                    print (mensajeManoObraCS)
                else:
                    mensajeManoObraCS=""
                    print(mensajeManoObraCS)

                #MATERIA PRIMA:::::::::::
                MateriaPrimaCS =  ConjuntoSuspencion.col_values(27)
                if MateriaPrimaCS[-1]=="Si":
                    DescrMateriaPrimaCS=  ConjuntoSuspencion.col_values(31)
                    TiempoMateriaPrimaCS =  ConjuntoSuspencion.col_values(28)
                    mensajeMateriaPrimaCS="*Hubo afectacion en las unidades por Materia Prima - Tiempo:* "+TiempoMateriaPrimaCS[-1]+" min, *Razon:* "+DescrMateriaPrimaCS[-1]+""
                    print (mensajeMateriaPrimaCS)
                else:
                    mensajeMateriaPrimaCS=""
                    print(mensajeMateriaPrimaCS)


                #POR METODO:::::::::::::
                UnidadesPorMetodoCS =  ConjuntoSuspencion.col_values(32)
                if UnidadesPorMetodoCS[-1]=="Si":
                    DescrUnidadesPorMetodoCS=  ConjuntoSuspencion.col_values(35)
                    TiempoUnidadesPorMetodoCS =  ConjuntoSuspencion.col_values(33)
                    mensajeUnidadesPorMetodoCS="*Hubo afectacion en las unidades por Metodo - Tiempo:* "+TiempoUnidadesPorMetodoCS[-1]+" min, *Razon:* "+DescrUnidadesPorMetodoCS[-1]+""
                    print (mensajeUnidadesPorMetodoCS)
                else:
                    mensajeUnidadesPorMetodoCS=""
                    print(mensajeUnidadesPorMetodoCS)


                ##SCRAPPPP::::::::::::::::::::::::::::::::
                ScrapCS =  ConjuntoSuspencion.col_values(36)
                if ScrapCS[-1]=="Si":
                    DescrScrapCS=  ConjuntoSuspencion.col_values(38)
                    CantidadScrapCS =  ConjuntoSuspencion.col_values(39)
                    mensajeScrapCS="*Se genero SCRAP - Cantidad:* "+CantidadScrapCS[-1]+", *Razon:* "+DescrScrapCS[-1]+""
                    print (mensajeScrapCS)
                else:
                    mensajeScrapCS=""
                    print(mensajeScrapCS)

                ##UNIDADES REPROCESADAS:::::::::::::
                UnidadesReprocesadasCS =  ConjuntoSuspencion.col_values(40)
                if UnidadesReprocesadasCS[-1]=="Si":
                    CantidadReprocesadasCS=  ConjuntoSuspencion.col_values(41)

                    mensajeUnidadesReprocesadasCS="*Se reprocesaron unidades - Cantidad:* "+CantidadReprocesadasCS[-1]+""
                    print (mensajeUnidadesReprocesadasCS)
                else:
                    mensajeUnidadesReprocesadasCS=""
                    print(mensajeUnidadesReprocesadasCS)

                ##OEE::::::::::::::::::::::
                OeeCS= ConjuntoSuspencion.col_values(47)
                OeeConjuntoSuspencion = OeeCS[-1]
                print("OEE: "+OeeConjuntoSuspencion)
                MensajeOeeCS="*OEE:* "+OeeCS[-1]

                print("Entró")
                mensaje2="\n*CONJUNTO SUSPENCION*" 
                if MensajeUnidadesFabricadasCJ!="":
                    mensaje2=mensaje2+"\n\n"+MensajeUnidadesFabricadasCJ
                if mensajeParoProgramadoCS!="":
                    mensaje2=mensaje2+"\n"+mensajeParoProgramadoCS
                    SinNovedad=2
                if mensajeIncidenteCS!="":
                    mensaje2=mensaje2+"\n"+mensajeIncidenteCS
                    SinNovedad=2
                if mensajeServiciosPublicosCS!="":
                    mensaje2=mensaje2+"\n"+mensajeServiciosPublicosCS
                    SinNovedad=2
                if mensajeMaquinaCS!="":
                    mensaje2=mensaje2+"\n"+mensajeMaquinaCS
                    SinNovedad=2
                if mensajeManoObraCS!="":
                    mensaje2=mensaje2+"\n"+mensajeManoObraCS
                    SinNovedad=2
                if mensajeMateriaPrimaCS!="":
                    mensaje2=mensaje2+"\n"+mensajeMateriaPrimaCS
                    SinNovedad=2
                if mensajeUnidadesPorMetodoCS!="":
                    mensaje2=mensaje2+"\n"+mensajeUnidadesPorMetodoCS
                    SinNovedad=2
                if mensajeUnidadesPorMetodoCS!="":
                    mensaje2=mensaje2+"\n"+mensajeUnidadesPorMetodoCS
                    SinNovedad=2
                if mensajeScrapCS!="":
                    mensaje2=mensaje2+"\n"+mensajeScrapCS
                    SinNovedad=2
                if mensajeUnidadesReprocesadasCS!="":
                    mensaje2=mensaje2+"\n"+mensajeUnidadesReprocesadasCS
                    SinNovedad=2

                if SinNovedad!=2:
                    mensaje2=mensaje2+"\n*No se reportaron novedades*"

                SinNovedad=0
        
                if MensajeOeeCS!="" and OeeConjuntoSuspencion!="#DIV/0!":
                    mensaje2=mensaje2+"\n"+MensajeOeeCS


        if espera_Minuto:
            print("2 Esperando minuto de envio...")
            espera_Minuto= False

        #CONDICIONAL PARA SELECCIONAR EL MIN Y EL RANGO DE SEGUNDOS
        if (Minuto2==5 and Segundo>=40) and Segundo<=45:

            #VALIDAR QUE SI HAYA UN REPORTE
            if str(HoraEG[-1])+"\n" == HoraGuardadaEG:
                print ("Ensamble gabinete, no reporto")
                mensaje3="*EMSAMBLE GABINETE*/n*La celula no reporto.*"
            else:

                #--------------------------------------------------------------------------------------------------
                #EMSAMBLE GABINETE:::::::::::::::
                print("EMSAMBLE GABINETE:---------")

                #SELECCION DE LA HOJA::
                EmsambleGabinete = sh.get_worksheet(2)
                #SELECCIONAR LA REFERENCIA::::: COPA 1 -- COPA 2 HACEB -- COPA 2 WHIRLPOOL
                UnidadesFabricadasEG=  EmsambleGabinete.col_values(5)
                print("Unidades Fabricadas: "+ UnidadesFabricadasEG[-1])
                MensajeUnidadesFabricadasEG="*Unidades Producidas:* "+UnidadesFabricadasEG[-1]
                #SELECCION DE COPA 1:::::
                SelectReferenciaEG = EmsambleGabinete.col_values(7)
                if SelectReferenciaEG[-1]=="Copa 1.0":
                    print("COPA 1::::")
                    #PAROS PROGRAMADOS::
                    ParoProgramadoEGCP1 = EmsambleGabinete.col_values(8)
                    if ParoProgramadoEGCP1[-1]=="Si":
                        RazonParoProgramadoEGCP1 =  EmsambleGabinete.col_values(9)
                        TiempoParoProgramadoEGCP1 =  EmsambleGabinete.col_values(10)
                        mensajeParoProgramadoEGCP1="*Paro programado - Tiempo:* "+TiempoParoProgramadoEGCP1[-1]+" min, *Razon:* "+RazonParoProgramadoEGCP1[-1]
                        print(mensajeParoProgramadoEGCP1)
                    else:
                        mensajeParoProgramadoEGCP1=""
                        print(mensajeParoProgramadoEGCP1)
                    
                    #INCIDENTES::
                    IncidenteEGCP1=EmsambleGabinete.col_values(11)
                    if IncidenteEGCP1[-1]=="Si":
                        DescrIncidenteEGCP1=EmsambleGabinete.col_values(13)
                        ValidarParoIncidenteEGCP1=EmsambleGabinete.col_values(14)
                        mensajeIncidenteEGCP1="*Incidente y/o accidente ambiental y/o SST: Razon:* "+DescrIncidenteEGCP1[-1] + " No se generó paro"
                        print(mensajeIncidenteEGCP1)
                        if ValidarParoIncidenteEGCP1[-1]=="Si":   
                            TiempoIncidenteEGCP1=EmsambleGabinete.col_values(15)
                            mensajeIncidenteEGCP1="*Incidente y/o accidente ambiental y/o SST - Tiempo:* "+TiempoIncidenteEGCP1[-1]+" min, *Razon:* "+DescrIncidenteEGCP1[-1]
                            print (mensajeIncidenteEGCP1)
                        else:
                            #DescrIncidenteEGCP1=EmsambleGabinete.col_values(12)
                            mensajeIncidenteEGCP1="*Incidente y/o accidente ambiental y/o SST: Razon:* "+DescrIncidenteEGCP1[-1] + " No se generó paro"
                            print (mensajeIncidenteEGCP1)
                    else:
                        mensajeIncidenteEGCP1=""
                        print(mensajeIncidenteEGCP1)

                ##SERVICIOS PUBLICOS COPA1::
                    ServiciosPublicosEGCP1=EmsambleGabinete.col_values(16)
                    if ServiciosPublicosEGCP1[-1]=="Si":
                        DescrServiciosPublicosEGCP1=EmsambleGabinete.col_values(18)
                        TiempoServiciosPublucosEGCP1=EmsambleGabinete.col_values(17)
                        mensajeServiciosPublicosEGCP1="*Hubo afectacion en las unidades del hora a hora por falta de servicios públicos: Razon:* "+DescrServiciosPublicosEGCP1[-1]+ " -*Tiempo:* :"+TiempoServiciosPublucosEGCP1[-1]+"min"
                        print(mensajeServiciosPublicosEGCP1)
                    else:
                        mensajeServiciosPublicosEGCP1=""
                        print(mensajeServiciosPublicosEGCP1)
                #POR MAQUINA COPA1:::
                    MaquinaEGCP1=EmsambleGabinete.col_values(19)
                    if MaquinaEGCP1[-1]=="Si":
                        DescrMaquinaEGCP1=EmsambleGabinete.col_values(22)
                        TiempoMaquinaEGCP1=EmsambleGabinete.col_values(20)
                        mensajeMaquinaEGCP1="*Hubo afectacion en las unidades por Maquina/ Equipo: Razon:* "+DescrMaquinaEGCP1[-1]+ " - *Tiempo:* "+TiempoMaquinaEGCP1[-1]+"min" 
                        print(mensajeMaquinaEGCP1)
                    else:
                        mensajeMaquinaEGCP1=""
                        print(mensajeMaquinaEGCP1)

                #POR MANO DE OBRA COPA1::::::::
                    ManoDeObraEGCP1=EmsambleGabinete.col_values(23)
                    if ManoDeObraEGCP1[-1]=="Si":
                        DescrManoDeObraEGCP1=EmsambleGabinete.col_values(27)
                        TiempoManoDeObraEGCP1=EmsambleGabinete.col_values(24)
                        mensajeManoDeObraEGCP1="*Hubo afectacion en las unidades por Mano De Obra: Razon:* "+DescrManoDeObraEGCP1[-1]+ " - *Tiempo:* "+TiempoManoDeObraEGCP1[-1]+"min" 
                        print(mensajeManoDeObraEGCP1)
                    else:
                        mensajeManoDeObraEGCP1=""
                        print(mensajeManoDeObraEGCP1)

                #MATERIA PRIMA COPA1::::

                    MateriaPrimaEGCP1=EmsambleGabinete.col_values(28)
                    if MateriaPrimaEGCP1[-1]=="Si":
                        DescrMateriaPrimaEGCP1=EmsambleGabinete.col_values(32)
                        TiempoMateriaPrimaEGCP1=EmsambleGabinete.col_values(29)
                        mensajeMateriaPrimaEGCP1="*Hubo afectacion en las unidades por Materia Prima: Razon:* "+DescrMateriaPrimaEGCP1[-1]+ " - *Tiempo:* "+TiempoMateriaPrimaEGCP1[-1]+"min" 
                        print(mensajeMateriaPrimaEGCP1)
                    else:
                        mensajeMateriaPrimaEGCP1=""
                        print(mensajeMateriaPrimaEGCP1)

                #POR METODO COPA1:::
                    MetodoEGCP1=EmsambleGabinete.col_values(33)
                    if MetodoEGCP1[-1]=="Si":
                        DescrMetodoEGCP1=EmsambleGabinete.col_values(36)
                        TiempoMetodoEGCP1=EmsambleGabinete.col_values(34)
                        mensajeMetodoEGCP1="*Hubo afectacion en las unidades por Metodo: Razon:* "+DescrMetodoEGCP1[-1]+ "- *Tiempo:* "+TiempoMetodoEGCP1[-1]+"min" 
                        print(mensajeMetodoEGCP1)
                    else:
                        mensajeMetodoEGCP1=""
                        print(mensajeMetodoEGCP1)

                #SCRAP COPA1::::::::::
                    ScrapEGCP1=EmsambleGabinete.col_values(37)
                    if ScrapEGCP1[-1]=="Si":
                        DescrScrapEGCP1=EmsambleGabinete.col_values(39)
                        CantidadScrapEGCP1=EmsambleGabinete.col_values(40)
                        mensajeScrapEGCP1="*Se generó SCRAP: Cantidad:* "+CantidadScrapEGCP1[-1]+" - *Razon:* "+DescrScrapEGCP1[-1]
                        print(mensajeScrapEGCP1)
                    else:
                        mensajeScrapEGCP1=""
                        print(mensajeScrapEGCP1)

                #REPROCESADAS COPA1::::::::
                    UnidadesReprocesadasEGP1 =  EmsambleGabinete.col_values(41)
                    if UnidadesReprocesadasEGP1[-1]=="Si":
                        CantidadReprocesadasEGP1=  EmsambleGabinete.col_values(42)

                        mensajeUnidadesReprocesadasEGP1="*Se reprocesaron unidades - Cantidad:* "+CantidadReprocesadasEGP1[-1]+""
                        print (mensajeUnidadesReprocesadasEGP1)
                    else:
                        mensajeUnidadesReprocesadasEGP1=""
                        print(mensajeUnidadesReprocesadasEGP1)

                ##SELECCIONA COPA 2 HACEB::::::::::::::::::::::

                if SelectReferenciaEG[-1]=="Copa 2.0 Haceb":
                    print("COPA 2 HACEB::::")
                    #PAROS PROGRAMADOS::
                    ParoProgramadoEGCP2H = EmsambleGabinete.col_values(43)
                    if ParoProgramadoEGCP2H[-1]=="Si":
                        RazonParoProgramadoEGCP2H =  EmsambleGabinete.col_values(44)
                        TiempoParoProgramadoEGCP2H =  EmsambleGabinete.col_values(45)
                        mensajeParoProgramadoEGCP2H="*Paro programado - Tiempo:* "+TiempoParoProgramadoEGCP2H[-1]+" min, *Razon:* "+RazonParoProgramadoEGCP2H[-1]
                        print(mensajeParoProgramadoEGCP2H)
                    else:
                        mensajeParoProgramadoEGCP2H=""
                        print(mensajeParoProgramadoEGCP2H)
                    
                    #INCIDENTES::
                    IncidenteEGCP2H=EmsambleGabinete.col_values(46)
                    if IncidenteEGCP2H[-1]=="Si":
                        DescrIncidenteEGCP2H=EmsambleGabinete.col_values(48)
                        ValidarParoIncidenteEMCP2H=EmsambleGabinete.col_values(49)
                        mensajeIncidenteEGCP2H="*Incidente y/o accidente ambiental y/o SST: Razon:* "+DescrIncidenteEGCP2H[-1]+ " no se generó paro."
                        print(mensajeIncidenteEGCP2H)
                        if ValidarParoIncidenteEMCP2H[-1]=="Si":   
                            TiempoIncidenteEGCP2H=EmsambleGabinete.col_values(50)
                            mensajeIncidenteEGCP2H="*Incidente y/o accidente ambiental y/o SST - Tiempo:* "+TiempoIncidenteEGCP2H[-1]+" min, *Razon:* "+DescrIncidenteEGCP2H[-1]
                            print (mensajeIncidenteEGCP2H)
                        else:
                            #DescrIncidenteEGCP1=EmsambleGabinete.col_values(12)
                            mensajeIncidenteEGCP2H="*Incidente y/o accidente ambiental y/o SST: Razon:* "+DescrIncidenteEGCP2H[-1] + " no se generó paro."
                            print (mensajeIncidenteEGCP2H)
                    else:
                        mensajeIncidenteEGCP2H=""
                        print(mensajeIncidenteEGCP2H)

                ##SERVICIOS PUBLICOS COPA2::
                    ServiciosPublicosEGCP2H=EmsambleGabinete.col_values(51)
                    if ServiciosPublicosEGCP2H[-1]=="Si":
                        DescrServiciosPublicosEGCP2H=EmsambleGabinete.col_values(53)
                        TiempoServiciosPublicosEGCP2H=EmsambleGabinete.col_values(52)
                        mensajeServiciosPublicosEGCP2H="*Hubo afectacion en las unidades del hora a hora por falta de servicios públicos: Razon:* "+DescrServiciosPublicosEGCP2H[-1]+ " -*Tiempo:* :"+TiempoServiciosPublicosEGCP2H[-1]+"min"
                        print(mensajeServiciosPublicosEGCP2H)
                    else:
                        mensajeServiciosPublicosEGCP2H=""
                        print(mensajeServiciosPublicosEGCP2H)
                #POR MAQUINA COPA2:::
                    MaquinaEGCP2H=EmsambleGabinete.col_values(54)
                    if MaquinaEGCP2H[-1]=="Si":
                        DescrMaquinaEGCP2H=EmsambleGabinete.col_values(57)
                        TiempoMaquinaEGCP2H=EmsambleGabinete.col_values(55)
                        mensajeMaquinaEGCP2H="*Hubo afectacion en las unidades por Maquina/ Equipo: Razon:* "+DescrMaquinaEGCP2H[-1]+ " - *Tiempo:* "+TiempoMaquinaEGCP2H[-1]+"min" 
                        print(mensajeMaquinaEGCP2H)
                    else:
                        mensajeMaquinaEGCP2H=""
                        print(mensajeMaquinaEGCP2H)

                #POR MANO DE OBRA COPA2::::::::
                    ManoDeObraEGCP2H=EmsambleGabinete.col_values(58)
                    if ManoDeObraEGCP2H[-1]=="Si":
                        DescrManoDeObraEGCP2H=EmsambleGabinete.col_values(62)
                        TiempoManoDeObraEGCP2H=EmsambleGabinete.col_values(59)
                        mensajeManoDeObraEGCP2H="*Hubo afectacion en las unidades por Mano De Obra: Razon:* "+DescrManoDeObraEGCP2H[-1]+ " - *Tiempo:* "+TiempoManoDeObraEGCP2H[-1]+"min" 
                        print(mensajeManoDeObraEGCP2H)
                    else:
                        mensajeManoDeObraEGCP2H=""
                        print(mensajeManoDeObraEGCP2H)

                #MATERIA PRIMA COPA2::::

                    MateriaPrimaEGCP2H=EmsambleGabinete.col_values(63)
                    if MateriaPrimaEGCP2H[-1]=="Si":
                        DescrMateriaPrimaEGCP2H=EmsambleGabinete.col_values(67)
                        TiempoMateriaPrimaEGCP2H=EmsambleGabinete.col_values(64)
                        mensajeMateriaPrimaEGCP2H="*Hubo afectacion en las unidades por Materia Prima: Razon:* "+DescrMateriaPrimaEGCP2H[-1]+ " - *Tiempo:* "+TiempoMateriaPrimaEGCP2H[-1]+"min" 
                        print(mensajeMateriaPrimaEGCP2H)
                    else:
                        mensajeMateriaPrimaEGCP2H=""
                        print(mensajeMateriaPrimaEGCP2H)

                #POR METODO COPA2:::
                    MetodoEGCP2H=EmsambleGabinete.col_values(68)
                    if MetodoEGCP2H[-1]=="Si":
                        DescrMetodoEGCP2H=EmsambleGabinete.col_values(71)
                        TiempoMetodoEGCP2H=EmsambleGabinete.col_values(69)
                        mensajeMetodoEGCP2H="*Hubo afectacion en las unidades por Metodo: Razon:* "+DescrMetodoEGCP2H[-1]+ "- *Tiempo:* "+TiempoMetodoEGCP2H[-1]+"min" 
                        print(mensajeMetodoEGCP2H)
                    else:
                        mensajeMetodoEGCP2H=""
                        print(mensajeMetodoEGCP2H)

                #SCRAP COPA2::::::::::
                    ScrapEGCP2H=EmsambleGabinete.col_values(72)
                    if ScrapEGCP2H[-1]=="Si":
                        DescrScrapEGCP2H=EmsambleGabinete.col_values(74)
                        CantidadScrapEGCP2H=EmsambleGabinete.col_values(75)
                        mensajeScrapEGCP2H="*Se generó SCRAP: Cantidad:* "+CantidadScrapEGCP2H[-1]+" - *Razon:* "+DescrScrapEGCP2H[-1]
                        print(mensajeScrapEGCP2H)
                    else:
                        mensajeScrapEGCP2H=""
                        print(mensajeScrapEGCP2H)

                #REPROCESADAS COPA2::::::::
                    UnidadesReprocesadasEGCP2H =  EmsambleGabinete.col_values(76)
                    if UnidadesReprocesadasEGCP2H[-1]=="Si":
                        CantidadReprocesadasEGCP2H=  EmsambleGabinete.col_values(77)

                        mensajeUnidadesReprocesadasEGCP2H="*Se reprocesaron unidades - Cantidad:* "+CantidadReprocesadasEGCP2H[-1]+""
                        print (mensajeUnidadesReprocesadasEGCP2H)
                    else:
                        mensajeUnidadesReprocesadasEGCP2H=""
                        print(mensajeUnidadesReprocesadasEGCP2H)

                ##COPA 2 WHIRLPOOL::::::::::::::
                # EMSAMBLE MECANISMOS COPA 2 WHIRLPOOL__:::::
                if SelectReferenciaEG[-1]=="Copa 2.0 Whirlpool":
                    print("COPA 2 Whirlpool::::::")
                    #PAROS PROGRAMADOS::::
                    ParoProgramadoEGCP2W = EmsambleGabinete.col_values(78)
                    if ParoProgramadoEGCP2W[-1]=="Si":
                        RazonParoProgramadoEGCP2W =  EmsambleGabinete.col_values(79)
                        TiempoParoProgramadoEGCP2W =  EmsambleGabinete.col_values(80)
                        mensajeParoProgramadoEGCP2W="*Paro programado - Tiempo:* "+TiempoParoProgramadoEGCP2W[-1]+" min, *Razon:* "+RazonParoProgramadoEGCP2W[-1]
                        print(mensajeParoProgramadoEGCP2W)
                    else:
                        mensajeParoProgramadoEGCP2W=""
                        print(mensajeParoProgramadoEGCP2W)
                    
                    #INCIDENTES WHIRLPOOL COPA 2:::::
                    IncidenteEGCP2W=EmsambleGabinete.col_values(81)
                    if IncidenteEGCP2W[-1]=="Si":
                        DescrIncidenteEGCP2W=EmsambleGabinete.col_values(83)
                        ValidarParoIncidenteEMCP2W=EmsambleGabinete.col_values(84)
                        mensajeIncidenteEGCP2W="*Incidente y/o accidente ambiental y/o SST: Razon:* "+DescrIncidenteEGCP2W[-1]+ " no se generó paro."
                        print(mensajeIncidenteEGCP2W)
                        if ValidarParoIncidenteEMCP2W[-1]=="Si":   
                            TiempoIncidenteEGCP2W=EmsambleGabinete.col_values(85)
                            mensajeIncidenteEGCP2W="*Incidente y/o accidente ambiental y/o SST - Tiempo:* "+TiempoIncidenteEGCP2W[-1]+" min, *Razon:* "+DescrIncidenteEGCP2W[-1]
                            print (mensajeIncidenteEGCP2W)
                        else:
                            #DescrIncidenteEGCP1=EmsambleGabinete.col_values(12)
                            mensajeIncidenteEGCP2W="*Incidente y/o accidente ambiental y/o SST: Razon:* "+DescrIncidenteEGCP2W[-1] + " no se generó paro."
                            print (mensajeIncidenteEGCP2W)
                    else:
                        mensajeIncidenteEGCP2H=""
                        print(mensajeIncidenteEGCP2H)

                ##SERVICIOS PUBLICOS COPA2 WHIRPOOL:::
                    ServiciosPublicosEGCP2W=EmsambleGabinete.col_values(86)
                    if ServiciosPublicosEGCP2W[-1]=="Si":
                        DescrServiciosPublicosEGCP2W=EmsambleGabinete.col_values(88)
                        TiempoServiciosPublicosEGCP2W=EmsambleGabinete.col_values(87)
                        mensajeServiciosPublicosEGCP2W="*Hubo afectacion en las unidades del hora a hora por falta de servicios públicos: Razon:* "+DescrServiciosPublicosEGCP2W[-1]+ " -*Tiempo:* :"+TiempoServiciosPublicosEGCP2W[-1]+"min"
                        print(mensajeServiciosPublicosEGCP2W)
                    else:
                        mensajeServiciosPublicosEGCP2W=""
                        print(mensajeServiciosPublicosEGCP2W)

                #POR MAQUINA COPA2 WHIRLPOOL::::::::
                    MaquinaEGCP2W=EmsambleGabinete.col_values(89)
                    if MaquinaEGCP2W[-1]=="Si":
                        DescrMaquinaEGCP2W=EmsambleGabinete.col_values(92)
                        TiempoMaquinaEGCP2W=EmsambleGabinete.col_values(90)
                        mensajeMaquinaEGCP2W="*Hubo afectacion en las unidades por Maquina/ Equipo: Razon:* "+DescrMaquinaEGCP2W[-1]+ " - *Tiempo:* "+TiempoMaquinaEGCP2W[-1]+"min" 
                        print(mensajeMaquinaEGCP2W)
                    else:
                        mensajeMaquinaEGCP2W=""
                        print(mensajeMaquinaEGCP2W)

                #POR MANO DE OBRA COPA2 WHIRLPOOL::::::::
                    ManoDeObraEGCP2W=EmsambleGabinete.col_values(93)
                    if ManoDeObraEGCP2W[-1]=="Si":
                        DescrManoDeObraEGCP2W=EmsambleGabinete.col_values(97)
                        TiempoManoDeObraEGCP2W=EmsambleGabinete.col_values(94)
                        mensajeManoDeObraEGCP2W="*Hubo afectacion en las unidades por Mano De Obra: Razon:* "+DescrManoDeObraEGCP2W[-1]+ " - *Tiempo:* "+TiempoManoDeObraEGCP2W[-1]+"min" 
                        print(mensajeManoDeObraEGCP2W)
                    else:
                        mensajeManoDeObraEGCP2W=""
                        print(mensajeManoDeObraEGCP2W)

                #MATERIA PRIMA COPA2 WHIRPOOL::::

                    MateriaPrimaEGCP2W=EmsambleGabinete.col_values(98)
                    if MateriaPrimaEGCP2W[-1]=="Si":
                        DescrMateriaPrimaEGCP2W=EmsambleGabinete.col_values(102)
                        TiempoMateriaPrimaEGCP2W=EmsambleGabinete.col_values(99)
                        mensajeMateriaPrimaEGCP2W="*Hubo afectacion en las unidades por Materia Prima: Razon:* "+DescrMateriaPrimaEGCP2W[-1]+ " - *Tiempo:* "+TiempoMateriaPrimaEGCP2W[-1]+"min" 
                        print(mensajeMateriaPrimaEGCP2W)
                    else:
                        mensajeMateriaPrimaEGCP2W=""
                        print(mensajeMateriaPrimaEGCP2W)

                #POR METODO COPA2 WHIRLPOOL:::
                    MetodoEGCP2W=EmsambleGabinete.col_values(103)
                    if MetodoEGCP2W[-1]=="Si":
                        DescrMetodoEGCP2W=EmsambleGabinete.col_values(106)
                        TiempoMetodoEGCP2W=EmsambleGabinete.col_values(104)
                        mensajeMetodoEGCP2W="*Hubo afectacion en las unidades por Metodo: Razon:* "+DescrMetodoEGCP2W[-1]+ "- *Tiempo:* "+TiempoMetodoEGCP2W[-1]+"min" 
                        print(mensajeMetodoEGCP2W)
                    else:
                        mensajeMetodoEGCP2W=""
                        print(mensajeMetodoEGCP2W)

                #SCRAP COPA2 WHIRLPOOL::::::::::
                    ScrapEGCP2W=EmsambleGabinete.col_values(107)
                    if ScrapEGCP2W[-1]=="Si":
                        DescrScrapEGCP2W=EmsambleGabinete.col_values(109)
                        CantidadScrapEGCP2W=EmsambleGabinete.col_values(110)
                        mensajeScrapEGCP2W="*Se generó SCRAP: Cantidad:* "+CantidadScrapEGCP2W[-1]+" - *Razon:* "+DescrScrapEGCP2W[-1]
                        print(mensajeScrapEGCP2W)
                    else:
                        mensajeScrapEGCP2W=""
                        print(mensajeScrapEGCP2W)

                #REPROCESADAS COPA2::::::::
                    UnidadesReprocesadasEGCP2W =  EmsambleGabinete.col_values(111)
                    if UnidadesReprocesadasEGCP2W[-1]=="Si":
                        CantidadReprocesadasEGCP2W=  EmsambleGabinete.col_values(112)

                        mensajeUnidadesReprocesadasEGCP2W="*Se reprocesaron unidades - Cantidad:* "+CantidadReprocesadasEGCP2W[-1]+""
                        print (mensajeUnidadesReprocesadasEGCP2W)
                    else:
                        mensajeUnidadesReprocesadasEGCP2W=""
                        print(mensajeUnidadesReprocesadasEGCP2W)

                OeeEG= EmsambleGabinete.col_values(118)
                OeeEmsableGabinete = OeeEG[-1]
                print("OEE: " +OeeEmsableGabinete)
                MensajeOeeEG="*OEE:* "+ OeeEmsableGabinete

                print("Entró")

                mensaje3="\n*ENSAMBLE GABINETE*" 
                if MensajeUnidadesFabricadasEG!="":
                    mensaje3=mensaje3+"\n\n"+MensajeUnidadesFabricadasEG

                if SelectReferenciaEG[-1]=="Copa 1.0":
                    mensaje3=mensaje3+"\n\n*COPA 1.0*"
                
                    if mensajeParoProgramadoEGCP1!="":
                        mensaje3=mensaje3+"\n"+mensajeParoProgramadoEGCP1
                        SinNovedad=2
                    if mensajeIncidenteEGCP1!="":
                        mensaje3=mensaje3+"\n"+mensajeIncidenteEGCP1
                        SinNovedad=2
                    if mensajeServiciosPublicosEGCP1!="":
                        mensaje3=mensaje3+"\n"+mensajeServiciosPublicosEGCP1
                        SinNovedad=2
                    if mensajeMaquinaEGCP1!="":
                        mensaje3=mensaje3+"\n"+mensajeMaquinaEGCP1
                        SinNovedad=2
                    if mensajeManoDeObraEGCP1!="":
                        mensaje3=mensaje3+"\n"+mensajeManoDeObraEGCP1
                        SinNovedad=2
                    if mensajeMateriaPrimaEGCP1!="":
                        mensaje3=mensaje3+"\n"+mensajeMateriaPrimaEGCP1
                        SinNovedad=2
                    if mensajeMetodoEGCP1!="":
                        mensaje3=mensaje3+"\n"+mensajeMetodoEGCP1
                        SinNovedad=2
                    if mensajeScrapEGCP1!="":
                        mensaje3=mensaje3+"\n"+mensajeScrapEGCP1
                        SinNovedad=2
                    if mensajeUnidadesReprocesadasEGP1!="":
                        mensaje3=mensaje3+"\n"+mensajeUnidadesReprocesadasEGP1
                        SinNovedad=2
                
                if SelectReferenciaEG[-1]=="Copa 2.0 Haceb":
                    mensaje3=mensaje3+"\n\n*COPA 2.0 HACEB*"
                
                    if mensajeParoProgramadoEGCP2H!="":
                        mensaje3=mensaje3+"\n"+mensajeParoProgramadoEGCP2H
                        SinNovedad=2
                    if mensajeIncidenteEGCP2H!="":
                        mensaje3=mensaje3+"\n"+mensajeIncidenteEGCP2H
                        SinNovedad=2
                    if mensajeServiciosPublicosEGCP2H!="":
                        mensaje3=mensaje3+"\n"+mensajeServiciosPublicosEGCP2H
                        SinNovedad=2
                    if mensajeMaquinaEGCP2H!="":
                        mensaje3=mensaje3+"\n"+mensajeMaquinaEGCP2H
                        SinNovedad=2
                    if mensajeManoDeObraEGCP2H!="":
                        mensaje3=mensaje3+"\n"+mensajeManoDeObraEGCP2H
                        SinNovedad=2
                    if mensajeMateriaPrimaEGCP2H!="":
                        mensaje3=mensaje3+"\n"+mensajeMateriaPrimaEGCP2H
                        SinNovedad=2
                    if mensajeMetodoEGCP2H!="":
                        mensaje3=mensaje3+"\n"+mensajeMetodoEGCP2H
                        SinNovedad=2
                    if mensajeScrapEGCP2H!="":
                        mensaje3=mensaje3+"\n"+mensajeScrapEGCP2H
                        SinNovedad=2
                    if mensajeUnidadesReprocesadasEGCP2H!="":
                        mensaje3=mensaje3+"\n"+mensajeUnidadesReprocesadasEGCP2H
                        SinNovedad=2

                if SelectReferenciaEG[-1]=="Copa 2.0 Whirlpool":
                    mensaje3=mensaje3+"\n\n*COPA 2.0 WHIRLPOOL*"
                
                    if mensajeParoProgramadoEGCP2W!="":
                        mensaje3=mensaje3+"\n"+mensajeParoProgramadoEGCP2W
                        SinNovedad=2
                    if mensajeIncidenteEGCP2W!="":
                        mensaje3=mensaje3+"\n"+mensajeIncidenteEGCP2W
                        SinNovedad=2
                    if mensajeServiciosPublicosEGCP2W!="":
                        mensaje3=mensaje3+"\n"+mensajeServiciosPublicosEGCP2W
                        SinNovedad=2
                    if mensajeMaquinaEGCP2W!="":
                        mensaje3=mensaje3+"\n"+mensajeMaquinaEGCP2W
                        SinNovedad=2
                    if mensajeManoDeObraEGCP2W!="":
                        mensaje3=mensaje3+"\n"+mensajeManoDeObraEGCP2W
                        SinNovedad=2
                    if mensajeMateriaPrimaEGCP2W!="":
                        mensaje3=mensaje3+"\n"+mensajeMateriaPrimaEGCP2W
                        SinNovedad=2
                    if mensajeMetodoEGCP2W!="":
                        mensaje3=mensaje3+"\n"+mensajeMetodoEGCP2W
                        SinNovedad=2
                    if mensajeScrapEGCP2W!="":
                        mensaje3=mensaje3+"\n"+mensajeScrapEGCP2W
                        SinNovedad=2
                    if mensajeUnidadesReprocesadasEGCP2W!="":
                        mensaje3=mensaje3+"\n"+mensajeUnidadesReprocesadasEGCP2W
                        SinNovedad=2

                if SinNovedad!=2:
                    mensaje3=mensaje3+"\n*No se reportaron novedades*"
                
                SinNovedad=0

                if MensajeOeeEG!="" and OeeEmsableGabinete!="#DIV/0!":
                    mensaje3=mensaje3+"\n"+MensajeOeeEG

        if espera_Minuto:
            print("2 Esperando minuto de envio...")
            espera_Minuto= False

        #CONDICIONAL PARA SELECCIONAR EL MIN Y EL RANGO DE SEGUNDOS
        if (Minuto2==6 and Segundo>=40) and Segundo<=45:

            #VALIDAR QUE SI HAYA UN REPORTE
            if str(HoraTF[-1])+"\n" == HoraGuardadaTF:
                print ("Testeo final, no reporto")
                mensaje6="*TESTEO FINAL*/n/n*La celula no reporto.*"
            else:

                #TESTEO FINAL:::::::::::::::
                print("TESTEO FINAL::---------")

                #SELECCION DE LA HOJA::
                TesteoFinal = sh.get_worksheet(3)
                #SELECCIONAR LA REFERENCIA::::: COPA 1 -- COPA 2 HACEB -- COPA 2 WHIRLPOOL
                UnidadesFabricadasTF=  TesteoFinal.col_values(5)
                print("Unidades Fabricadas: "+UnidadesFabricadasTF[-1])
                MensajeUnidadesFabricadasTF="*Unidades Producidas:* "+UnidadesFabricadasTF[-1]
                #SELECCION DE COPA 1:::::
                SelectReferenciaTF = TesteoFinal.col_values(7)
                if SelectReferenciaTF[-1]=="Copa 1.0":
                    print("COPA 1::::")
                    #PAROS PROGRAMADOS::
                    ParoProgramadoTFCP1 = TesteoFinal.col_values(8)
                    if ParoProgramadoTFCP1[-1]=="Si":
                        RazonParoProgramadoTFCP1 =  TesteoFinal.col_values(9)
                        TiempoParoProgramadoTFCP1 =  TesteoFinal.col_values(10)
                        mensajeParoProgramadoTFCP1="*Paro programado - Tiempo:* "+TiempoParoProgramadoTFCP1[-1]+" min, *Razon:* "+RazonParoProgramadoTFCP1[-1]
                        print(mensajeParoProgramadoTFCP1)
                    else:
                        mensajeParoProgramadoTFCP1=""
                        print(mensajeParoProgramadoTFCP1)
                    
                    #INCIDENTES::
                    IncidenteTFCP1=TesteoFinal.col_values(11)
                    if IncidenteTFCP1[-1]=="Si":
                        DescrIncidenteTFCP1=TesteoFinal.col_values(13)
                        ValidarParoIncidenteTFCP1=TesteoFinal.col_values(14)
                        mensajeIncidenteTFCP1="*Incidente y/o accidente ambiental y/o SST: Razon:* "+DescrIncidenteTFCP1[-1] + " No se generó paro"
                        print(mensajeIncidenteTFCP1)
                        if ValidarParoIncidenteTFCP1[-1]=="Si":   
                            TiempoIncidenteTFCP1=TesteoFinal.col_values(15)
                            mensajeIncidenteTFCP1="*Incidente y/o accidente ambiental y/o SST - Tiempo:* "+TiempoIncidenteTFCP1[-1]+" min, *Razon:* "+DescrIncidenteTFCP1[-1]
                            print (mensajeIncidenteTFCP1)
                        else:
                            DescrIncidenteTFCP1=TesteoFinal.col_values(13)
                            mensajeIncidenteTFCP1="*Incidente y/o accidente ambiental y/o SST: Razon:* "+DescrIncidenteTFCP1[-1] + " No se generó paro"
                            print (mensajeIncidenteTFCP1)
                    else:
                        mensajeIncidenteTFCP1=""
                        print(mensajeIncidenteTFCP1)

                ##SERVICIOS PUBLICOS COPA1::
                    ServiciosPublicosTFCP1=TesteoFinal.col_values(16)
                    if ServiciosPublicosTFCP1[-1]=="Si":
                        DescrServiciosPublicosTFCP1=TesteoFinal.col_values(18)
                        TiempoServiciosPublucosTFCP1=TesteoFinal.col_values(17)
                        mensajeServiciosPublicosTFCP1="*Hubo afectacion en las unidades del hora a hora por falta de servicios públicos: Razon:* "+DescrServiciosPublicosTFCP1[-1]+ " -*Tiempo:* :"+TiempoServiciosPublucosTFCP1[-1]+"min"
                        print(mensajeServiciosPublicosTFCP1)
                    else:
                        mensajeServiciosPublicosTFCP1=""
                        print(mensajeServiciosPublicosTFCP1)
                #POR MAQUINA COPA1:::
                    MaquinaTFCP1=TesteoFinal.col_values(19)
                    if MaquinaTFCP1[-1]=="Si":
                        DescrMaquinaTFCP1=TesteoFinal.col_values(22)
                        TiempoMaquinaTFCP1=TesteoFinal.col_values(20)
                        mensajeMaquinaTFCP1="*Hubo afectacion en las unidades por Maquina/ Equipo: Razon:* "+DescrMaquinaTFCP1[-1]+ " - *Tiempo:* "+TiempoMaquinaTFCP1[-1]+"min" 
                        print(mensajeMaquinaTFCP1)
                    else:
                        mensajeMaquinaTFCP1=""
                        print(mensajeMaquinaTFCP1)

                #POR MANO DE OBRA COPA1::::::::
                    ManoDeObraTFCP1=TesteoFinal.col_values(23)
                    if ManoDeObraTFCP1[-1]=="Si":
                        DescrManoDeObraTFCP1=TesteoFinal.col_values(27)
                        TiempoManoDeObraTFCP1=TesteoFinal.col_values(24)
                        mensajeManoDeObraTFCP1="*Hubo afectacion en las unidades por Mano De Obra: Razon:* "+DescrManoDeObraTFCP1[-1]+ " - *Tiempo:* "+TiempoManoDeObraTFCP1[-1]+"min" 
                        print(mensajeManoDeObraTFCP1)
                    else:
                        mensajeManoDeObraTFCP1=""
                        print(mensajeManoDeObraTFCP1)

                #MATERIA PRIMA COPA1::::

                    MateriaPrimaTFCP1=TesteoFinal.col_values(28)
                    if MateriaPrimaTFCP1[-1]=="Si":
                        DescrMateriaPrimaTFCP1=TesteoFinal.col_values(32)
                        TiempoMateriaPrimaTFCP1=TesteoFinal.col_values(29)
                        mensajeMateriaPrimaTFCP1="*Hubo afectacion en las unidades por Materia Prima: Razon:* "+DescrMateriaPrimaTFCP1[-1]+ " - *Tiempo:* "+TiempoMateriaPrimaTFCP1[-1]+"min" 
                        print(mensajeMateriaPrimaTFCP1)
                    else:
                        mensajeMateriaPrimaTFCP1=""
                        print(mensajeMateriaPrimaTFCP1)

                #POR METODO COPA1:::
                    MetodoTFCP1=TesteoFinal.col_values(33)
                    if MetodoTFCP1[-1]=="Si":
                        DescrMetodoTFCP1=TesteoFinal.col_values(36)
                        TiempoMetodoTFCP1=TesteoFinal.col_values(34)
                        mensajeMetodoTFCP1="*Hubo afectacion en las unidades por Método: Razon:* "+DescrMetodoTFCP1[-1]+ "- *Tiempo:* "+TiempoMetodoTFCP1[-1]+"min" 
                        print(mensajeMetodoTFCP1)
                    else:
                        mensajeMetodoTFCP1=""
                        print(mensajeMetodoTFCP1)

                #SCRAP COPA1::::::::::
                    ScrapTFCP1=TesteoFinal.col_values(37)
                    if ScrapTFCP1[-1]=="Si":
                        DescrScrapTFCP1=TesteoFinal.col_values(39)
                        CantidadScrapTFCP1=TesteoFinal.col_values(40)
                        mensajeScrapTFCP1="*Se generó SCRAP: Cantidad:* "+CantidadScrapTFCP1[-1]+" - *Razon:* "+DescrScrapTFCP1[-1]
                        print(mensajeScrapTFCP1)
                    else:
                        mensajeScrapTFCP1=""
                        print(mensajeScrapTFCP1)

                #REPROCESADAS COPA1::::::::
                    UnidadesReprocesadasTFCP1 =  TesteoFinal.col_values(41)
                    if UnidadesReprocesadasTFCP1[-1]=="Si":
                        CantidadReprocesadasTFCP1=  TesteoFinal.col_values(42)

                        mensajeUnidadesReprocesadasTFCP1="*Se reprocesaron unidades - Cantidad:* "+CantidadReprocesadasTFCP1[-1]+""
                        print (mensajeUnidadesReprocesadasTFCP1)
                    else:
                        mensajeUnidadesReprocesadasTFCP1=""
                        print(mensajeUnidadesReprocesadasTFCP1)

                ##SELECCIONA COPA 2 HACEB::::::::::::::::::::::

                if SelectReferenciaTF[-1]=="Copa 2.0 Haceb":
                    print("COPA 2 HACEB::::")
                    #PAROS PROGRAMADOS::
                    ParoProgramadoTFCP2H = TesteoFinal.col_values(43)
                    if ParoProgramadoTFCP2H[-1]=="Si":
                        RazonParoProgramadoTFCP2H =  TesteoFinal.col_values(44)
                        TiempoParoProgramadoTFCP2H =  TesteoFinal.col_values(45)
                        mensajeParoProgramadoTFCP2H="*Paro programado - Tiempo:* "+TiempoParoProgramadoTFCP2H[-1]+" min, *Razon:* "+RazonParoProgramadoTFCP2H[-1]
                        print(mensajeParoProgramadoTFCP2H)
                    else:
                        mensajeParoProgramadoTFCP2H=""
                        print(mensajeParoProgramadoTFCP2H)
                    
                    #INCIDENTES::
                    IncidenteTFCP2H=TesteoFinal.col_values(46)
                    if IncidenteTFCP2H[-1]=="Si":
                        DescrIncidenteTFCP2H=TesteoFinal.col_values(48)
                        ValidarParoIncidenteTFCP2H=TesteoFinal.col_values(49)
                        mensajeIncidenteTFCP2H="*Incidente y/o accidente ambiental y/o SST: Razon:* "+DescrIncidenteTFCP2H[-1]+ " no se generó paro."
                        print(mensajeIncidenteTFCP2H)
                        if ValidarParoIncidenteTFCP2H[-1]=="Si":   
                            TiempoIncidenteTFCP2H=TesteoFinal.col_values(50)
                            mensajeIncidenteTFCP2H="*Incidente y/o accidente ambiental y/o SST - Tiempo:* "+TiempoIncidenteTFCP2H[-1]+" min, *Razon:* "+DescrIncidenteTFCP2H[-1]
                            print (mensajeIncidenteTFCP2H)
                        else:
                            #DescrIncidenteEGCP1=TesteoFinal.col_values(12)
                            mensajeIncidenteTFCP2H="*Incidente y/o accidente ambiental y/o SST: Razon:* "+DescrIncidenteTFCP2H[-1] + " no se generó paro."
                            print (mensajeIncidenteTFCP2H)
                    else:
                        mensajeIncidenteTFCP2H=""
                        print(mensajeIncidenteTFCP2H)

                ##SERVICIOS PUBLICOS COPA2::
                    ServiciosPublicosTFCP2H=TesteoFinal.col_values(51)
                    if ServiciosPublicosTFCP2H[-1]=="Si":
                        DescrServiciosPublicosTFCP2H=TesteoFinal.col_values(53)
                        TiempoServiciosPublicosTFCP2H=TesteoFinal.col_values(52)
                        mensajeServiciosPublicosTFCP2H="*Hubo afectacion en las unidades del hora a hora por falta de servicios públicos: Razon:* "+DescrServiciosPublicosTFCP2H[-1]+ " -*Tiempo:* :"+TiempoServiciosPublicosTFCP2H[-1]+"min"
                        print(mensajeServiciosPublicosTFCP2H)
                    else:
                        mensajeServiciosPublicosTFCP2H=""
                        print(mensajeServiciosPublicosTFCP2H)
                #POR MAQUINA COPA2:::
                    MaquinaTFCP2H=TesteoFinal.col_values(54)
                    if MaquinaTFCP2H[-1]=="Si":
                        DescrMaquinaTFCP2H=TesteoFinal.col_values(57)
                        TiempoMaquinaTFCP2H=TesteoFinal.col_values(55)
                        mensajeMaquinaTFCP2H="*Hubo afectacion en las unidades por Maquina/ Equipo: Razon:* "+DescrMaquinaTFCP2H[-1]+ " - *Tiempo:* "+TiempoMaquinaTFCP2H[-1]+"min" 
                        print(mensajeMaquinaTFCP2H)
                    else:
                        mensajeMaquinaTFCP2H=""
                        print(mensajeMaquinaTFCP2H)

                #POR MANO DE OBRA COPA2::::::::
                    ManoDeObraTFCP2H=TesteoFinal.col_values(58)
                    if ManoDeObraTFCP2H[-1]=="Si":
                        DescrManoDeObraTFCP2H=TesteoFinal.col_values(62)
                        TiempoManoDeObraTFCP2H=TesteoFinal.col_values(59)
                        mensajeManoDeObraTFCP2H="*Hubo afectacion en las unidades por Mano De Obra: Razon:* "+DescrManoDeObraTFCP2H[-1]+ " - *Tiempo:* "+TiempoManoDeObraTFCP2H[-1]+"min" 
                        print(mensajeManoDeObraTFCP2H)
                    else:
                        mensajeManoDeObraTFCP2H=""
                        print(mensajeManoDeObraTFCP2H)

                #MATERIA PRIMA COPA2::::

                    MateriaPrimaTFCP2H=TesteoFinal.col_values(63)
                    if MateriaPrimaTFCP2H[-1]=="Si":
                        DescrMateriaPrimaTFCP2H=TesteoFinal.col_values(67)
                        TiempoMateriaPrimaTFCP2H=TesteoFinal.col_values(64)
                        mensajeMateriaPrimaTFCP2H="*Hubo afectacion en las unidades por Materia Prima: Razon:* "+DescrMateriaPrimaTFCP2H[-1]+ " - *Tiempo:* "+TiempoMateriaPrimaTFCP2H[-1]+"min" 
                        print(mensajeMateriaPrimaTFCP2H)
                    else:
                        mensajeMateriaPrimaTFCP2H=""
                        print(mensajeMateriaPrimaTFCP2H)

                #POR METODO COPA2:::
                    MetodoTFCP2H=TesteoFinal.col_values(68)
                    if MetodoTFCP2H[-1]=="Si":
                        DescrMetodoTFCP2H=TesteoFinal.col_values(71)
                        TiempoMetodoTFCP2H=TesteoFinal.col_values(69)
                        mensajeMetodoTFCP2H="*Hubo afectacion en las unidades por Metodo: Razon:* "+DescrMetodoTFCP2H[-1]+ "- *Tiempo:* "+TiempoMetodoTFCP2H[-1]+"min" 
                        print(mensajeMetodoTFCP2H)
                    else:
                        mensajeMetodoTFCP2H=""
                        print(mensajeMetodoTFCP2H)

                #SCRAP COPA2::::::::::
                    ScrapTFCP2H=TesteoFinal.col_values(72)
                    if ScrapTFCP2H[-1]=="Si":
                        DescrScrapTFCP2H=TesteoFinal.col_values(74)
                        CantidadScrapTFCP2H=TesteoFinal.col_values(75)
                        mensajeScrapTFCP2H="*Se generó SCRAP: Cantidad:* "+CantidadScrapTFCP2H[-1]+" - *Razon:* "+DescrScrapTFCP2H[-1]
                        print(mensajeScrapTFCP2H)
                    else:
                        mensajeScrapTFCP2H=""
                        print(mensajeScrapTFCP2H)

                #REPROCESADAS COPA2::::::::
                    UnidadesReprocesadasTFCP2H =  TesteoFinal.col_values(76)
                    if UnidadesReprocesadasTFCP2H[-1]=="Si":
                        CantidadReprocesadasTFCP2H=  TesteoFinal.col_values(77)

                        mensajeUnidadesReprocesadasTFCP2H="*Se reprocesaron unidades - Cantidad:* "+CantidadReprocesadasTFCP2H[-1]+""
                        print (mensajeUnidadesReprocesadasTFCP2H)
                    else:
                        mensajeUnidadesReprocesadasTFCP2H=""
                        print(mensajeUnidadesReprocesadasTFCP2H)

                ##COPA 2 WHIRLPOOL::::::::::::::
                # TESTEO FINAL COPA 2 WHIRLPOOL__:::::
                if SelectReferenciaTF[-1]=="Copa 2.0 Whirlpool":
                    print("COPA 2 Whirlpool::::::")
                    #PAROS PROGRAMADOS::::
                    ParoProgramadoTFCP2W = TesteoFinal.col_values(78)
                    if ParoProgramadoTFCP2W[-1]=="Si":
                        RazonParoProgramadoTFCP2W =  TesteoFinal.col_values(79)
                        TiempoParoProgramadoTFCP2W =  TesteoFinal.col_values(80)
                        mensajeParoProgramadoTFCP2W="*Paro programado - Tiempo:* "+TiempoParoProgramadoTFCP2W[-1]+" min, *Razon:* "+RazonParoProgramadoTFCP2W[-1]
                        print(mensajeParoProgramadoTFCP2W)
                    else:
                        mensajeParoProgramadoTFCP2W=""
                        print(mensajeParoProgramadoTFCP2W)
                    
                    #INCIDENTES WHIRLPOOL COPA 2:::::
                    IncidenteTFCP2W=TesteoFinal.col_values(81)
                    if IncidenteTFCP2W[-1]=="Si":
                        DescrIncidenteTFCP2W=TesteoFinal.col_values(83)
                        ValidarParoIncidenteTFCP2W=TesteoFinal.col_values(84)
                        mensajeIncidenteTFCP2W="*Incidente y/o accidente ambiental y/o SST: Razon:* "+DescrIncidenteTFCP2W[-1]+ " no se generó paro."
                        print(mensajeIncidenteTFCP2W)
                        if ValidarParoIncidenteTFCP2W[-1]=="Si":   
                            TiempoIncidenteTFCP2W=TesteoFinal.col_values(85)
                            mensajeIncidenteTFCP2W="*Incidente y/o accidente ambiental y/o SST - Tiempo:* "+TiempoIncidenteTFCP2W[-1]+" min, *Razon:* "+DescrIncidenteTFCP2W[-1]
                            print (mensajeIncidenteTFCP2W)
                        else:
                            #DescrIncidenteEGCP1=TesteoFinal.col_values(12)
                            mensajeIncidenteTFCP2W="*Incidente y/o accidente ambiental y/o SST: Razon:* "+DescrIncidenteTFCP2W[-1] + " no se generó paro."
                            print (mensajeIncidenteTFCP2W)
                    else:
                        mensajeIncidenteTFCP2W=""
                        print(mensajeIncidenteTFCP2W)

                ##SERVICIOS PUBLICOS COPA2 WHIRPOOL:::
                    ServiciosPublicosTFCP2W=TesteoFinal.col_values(86)
                    if ServiciosPublicosTFCP2W[-1]=="Si":
                        DescrServiciosPublicosTFCP2W=TesteoFinal.col_values(88)
                        TiempoServiciosPublicosTFCP2W=TesteoFinal.col_values(87)
                        mensajeServiciosPublicosTFCP2W="*Hubo afectacion en las unidades del hora a hora por falta de servicios públicos: Razon:* "+DescrServiciosPublicosTFCP2W[-1]+ " -*Tiempo:* :"+TiempoServiciosPublicosTFCP2W[-1]+"min"
                        print(mensajeServiciosPublicosTFCP2W)
                    else:
                        mensajeServiciosPublicosTFCP2W=""
                        print(mensajeServiciosPublicosTFCP2W)

                #POR MAQUINA COPA2 WHIRLPOOL::::::::
                    MaquinaTFCP2W=TesteoFinal.col_values(89)
                    if MaquinaTFCP2W[-1]=="Si":
                        DescrMaquinaTFCP2W=TesteoFinal.col_values(92)
                        TiempoMaquinaTFCP2W=TesteoFinal.col_values(90)
                        mensajeMaquinaTFCP2W="*Hubo afectacion en las unidades por Maquina/ Equipo: Razon:* "+DescrMaquinaTFCP2W[-1]+ " - *Tiempo:* "+TiempoMaquinaTFCP2W[-1]+"min" 
                        print(mensajeMaquinaTFCP2W)
                    else:
                        mensajeMaquinaTFCP2W=""
                        print(mensajeMaquinaTFCP2W)

                #POR MANO DE OBRA COPA2 WHIRLPOOL::::::::
                    ManoDeObraTFCP2W=TesteoFinal.col_values(92)
                    if ManoDeObraTFCP2W[-1]=="Si":
                        DescrManoDeObraTFCP2W=TesteoFinal.col_values(96)
                        TiempoManoDeObraTFCP2W=TesteoFinal.col_values(93)
                        mensajeManoDeObraTFCP2W="*Hubo afectacion en las unidades por Mano De Obra: Razon:* "+DescrManoDeObraTFCP2W[-1]+ " - *Tiempo:* "+TiempoManoDeObraTFCP2W[-1]+"min" 
                        print(mensajeManoDeObraTFCP2W)
                    else:
                        mensajeManoDeObraTFCP2W=""
                        print(mensajeManoDeObraTFCP2W)

                #MATERIA PRIMA COPA2 WHIRPOOL::::

                    MateriaPrimaTFCP2W=TesteoFinal.col_values(97)
                    if MateriaPrimaTFCP2W[-1]=="Si":
                        DescrMateriaPrimaTFCP2W=TesteoFinal.col_values(101)
                        TiempoMateriaPrimaTFCP2W=TesteoFinal.col_values(98)
                        mensajeMateriaPrimaTFCP2W="*Hubo afectacion en las unidades por Materia Prima: Razon:* "+DescrMateriaPrimaTFCP2W[-1]+ " - *Tiempo:* "+TiempoMateriaPrimaTFCP2W[-1]+"min" 
                        print(mensajeMateriaPrimaTFCP2W)
                    else:
                        mensajeMateriaPrimaTFCP2W=""
                        print(mensajeMateriaPrimaTFCP2W)

                #POR METODO COPA2 WHIRLPOOL:::
                    MetodoTFCP2W=TesteoFinal.col_values(102)
                    if MetodoTFCP2W[-1]=="Si":
                        DescrMetodoTFCP2W=TesteoFinal.col_values(105)
                        TiempoMetodoTFCP2W=TesteoFinal.col_values(103)
                        mensajeMetodoTFCP2W="*Hubo afectacion en las unidades por Metodo: Razon:* "+DescrMetodoTFCP2W[-1]+ "- *Tiempo:* "+TiempoMetodoTFCP2W[-1]+"min" 
                        print(mensajeMetodoTFCP2W)
                    else:
                        mensajeMetodoTFCP2W=""
                        print(mensajeMetodoTFCP2W)

                #SCRAP COPA2 WHIRLPOOL::::::::::
                    ScrapTFCP2W=TesteoFinal.col_values(106)
                    if ScrapTFCP2W[-1]=="Si":
                        DescrScrapTFCP2W=TesteoFinal.col_values(108)
                        CantidadScrapTFCP2W=TesteoFinal.col_values(109)
                        mensajeScrapTFCP2W="*Se generó SCRAP: Cantidad:* "+CantidadScrapTFCP2W[-1]+" - *Razon:* "+DescrScrapTFCP2W[-1]
                        print(mensajeScrapTFCP2W)
                    else:
                        mensajeScrapTFCP2W=""
                        print(mensajeScrapTFCP2W)

                #REPROCESADAS COPA2::::::::
                    UnidadesReprocesadasTFCP2W =  TesteoFinal.col_values(110)
                    if UnidadesReprocesadasTFCP2W[-1]=="Si":
                        CantidadReprocesadasTFCP2W=  TesteoFinal.col_values(111)

                        mensajeUnidadesReprocesadasTFCP2W="*Se reprocesaron unidades - Cantidad:* "+CantidadReprocesadasTFCP2W[-1]+""
                        print (mensajeUnidadesReprocesadasTFCP2W)
                    else:
                        mensajeUnidadesReprocesadasTFCP2W=""
                        print(mensajeUnidadesReprocesadasTFCP2W)


                OeeTF= TesteoFinal.col_values(117)
                OeeTesteoFinal = OeeTF[-1]
                print("OEE: " + OeeTesteoFinal)
                MensajeOeeTF="*OEE:* " + OeeTesteoFinal

                mensaje6="\n*TESTEO FINAL*" 
                if MensajeUnidadesFabricadasTF!="":
                    mensaje6=mensaje6+"\n\n"+MensajeUnidadesFabricadasTF

                if SelectReferenciaTF[-1]=="Copa 1.0":
                    mensaje6=mensaje6+"\n\n*COPA 1.0*"
                
                    if mensajeParoProgramadoTFCP1!="":
                        mensaje6=mensaje6+"\n"+mensajeParoProgramadoTFCP1
                        SinNovedad=2
                    if mensajeIncidenteTFCP1!="":
                        mensaje6=mensaje6+"\n"+mensajeIncidenteTFCP1
                        SinNovedad=2
                    if mensajeServiciosPublicosTFCP1!="":
                        mensaje6=mensaje6+"\n"+mensajeServiciosPublicosTFCP1
                        SinNovedad=2
                    if mensajeMaquinaTFCP1!="":
                        mensaje6=mensaje6+"\n"+mensajeMaquinaTFCP1
                        SinNovedad=2
                    if mensajeManoDeObraTFCP1!="":
                        mensaje6=mensaje6+"\n"+mensajeManoDeObraTFCP1
                        SinNovedad=2
                    if mensajeMateriaPrimaTFCP1!="":
                        mensaje6=mensaje6+"\n"+mensajeMateriaPrimaTFCP1
                        SinNovedad=2
                    if mensajeMetodoTFCP1!="":
                        mensaje6=mensaje6+"\n"+mensajeMetodoTFCP1
                        SinNovedad=2
                    if mensajeScrapTFCP1!="":
                        mensaje6=mensaje6+"\n"+mensajeScrapTFCP1
                        SinNovedad=2
                    if mensajeUnidadesReprocesadasTFCP1!="":
                        mensaje6=mensaje6+"\n"+mensajeUnidadesReprocesadasTFCP1
                        SinNovedad=2
                
                if SelectReferenciaTF[-1]=="Copa 2.0 Haceb":
                    mensaje6=mensaje6+"\n\n*COPA 2.0 HACEB*"
                
                    if mensajeParoProgramadoTFCP2H!="":
                        mensaje6=mensaje6+"\n"+mensajeParoProgramadoTFCP2H
                        SinNovedad=2
                    if mensajeIncidenteTFCP2H!="":
                        mensaje6=mensaje6+"\n"+mensajeIncidenteTFCP2H
                        SinNovedad=2
                    if mensajeServiciosPublicosTFCP2H!="":
                        mensaje6=mensaje6+"\n"+mensajeServiciosPublicosTFCP2H
                        SinNovedad=2
                    if mensajeMaquinaTFCP2H!="":
                        mensaje6=mensaje6+"\n"+mensajeMaquinaTFCP2H
                        SinNovedad=2
                    if mensajeManoDeObraTFCP2H!="":
                        mensaje6=mensaje6+"\n"+mensajeManoDeObraTFCP2H
                        SinNovedad=2
                    if mensajeMateriaPrimaTFCP2H!="":
                        mensaje6=mensaje6+"\n"+mensajeMateriaPrimaTFCP2H
                        SinNovedad=2
                    if mensajeMetodoTFCP2H!="":
                        mensaje6=mensaje6+"\n"+mensajeMetodoTFCP2H
                        SinNovedad=2
                    if mensajeScrapTFCP2H!="":
                        mensaje6=mensaje6+"\n"+mensajeScrapTFCP2H
                        SinNovedad=2
                    if mensajeUnidadesReprocesadasTFCP2H!="":
                        mensaje6=mensaje6+"\n"+mensajeUnidadesReprocesadasTFCP2H
                        SinNovedad=2

                if SelectReferenciaTF[-1]=="Copa 2.0 Whirlpool":
                    mensaje6=mensaje6+"\n\n*COPA 2.0 WHIRLPOOL*"
                
                    if mensajeParoProgramadoTFCP2W!="":
                        mensaje6=mensaje6+"\n"+mensajeParoProgramadoTFCP2W
                        SinNovedad=2
                    if mensajeIncidenteTFCP2W!="":
                        mensaje6=mensaje6+"\n"+mensajeIncidenteTFCP2W
                        SinNovedad=2
                    if mensajeServiciosPublicosTFCP2W!="":
                        mensaje6=mensaje6+"\n"+mensajeServiciosPublicosTFCP2W
                        SinNovedad=2
                    if mensajeMaquinaTFCP2W!="":
                        mensaje6=mensaje6+"\n"+mensajeMaquinaTFCP2W
                        SinNovedad=2
                    if mensajeManoDeObraTFCP2W!="":
                        mensaje6=mensaje6+"\n"+mensajeManoDeObraTFCP2W
                        SinNovedad=2
                    if mensajeMateriaPrimaTFCP2W!="":
                        mensaje6=mensaje6+"\n"+mensajeMateriaPrimaTFCP2W
                        SinNovedad=2
                    if mensajeMetodoTFCP2W!="":
                        mensaje6=mensaje6+"\n"+mensajeMetodoTFCP2W
                        SinNovedad=2
                    if mensajeScrapTFCP2W!="":
                        mensaje6=mensaje6+"\n"+mensajeScrapTFCP2W
                        SinNovedad=2
                    if mensajeUnidadesReprocesadasTFCP2W!="":
                        mensaje6=mensaje6+"\n"+mensajeUnidadesReprocesadasTFCP2W
                        SinNovedad=2

                if SinNovedad!=2:
                    mensaje6=mensaje6+"\n*No se reportaron novedades*"

                SinNovedad=0

                if MensajeOeeTF!="" and OeeTesteoFinal!="#DIV/0!":
                    mensaje6=mensaje6+"\n"+MensajeOeeTF

        if espera_Minuto:
            print("2 Esperando minuto de envio...")
            espera_Minuto= False

        #CONDICIONAL PARA SELECCIONAR EL MIN Y EL RANGO DE SEGUNDOS
        if (Minuto2==7 and Segundo>=40) and Segundo<=45:

            #VALIDAR QUE SI HAYA UN REPORTE
            if str(HoraTM[-1])+"\n" == HoraGuardadaTM:
                print ("Tapa movil, no reporto")
                mensaje4="*TAPA MOVIL*/n/n*La celula no reporto.*"
            else:
                #TAPA MOVIL:::::::::::::::
                print("TAPA MOVIL::---------")
                #SELECCION DE LA HOJA::
                TapaMovil = sh.get_worksheet(4)
                #SELECCIONAR LA REFERENCIA::::: COPA 1 -- COPA 2 HACEB -- COPA 2 WHIRLPOOL
                UnidadesFabricadasTM=  TapaMovil.col_values(5)
                print("Unidades Fabricadas: "+UnidadesFabricadasTM[-1])
                MensajeUnidadesFabricadasTM="*Unidades Producidas:* "+ UnidadesFabricadasTM[-1]
                #SELECCION DE COPA 1:::::
                SelectReferenciaTM = TapaMovil.col_values(7)
                if SelectReferenciaTM[-1]=="Copa 1.0":
                    print("COPA 1::::")
                    #PAROS PROGRAMADOS TAPA MOVIL COPA 1::::
                    ParoProgramadoTMCP1 = TapaMovil.col_values(8)
                    if ParoProgramadoTMCP1[-1]=="Si":
                        RazonParoProgramadoTMCP1 =  TapaMovil.col_values(9)
                        TiempoParoProgramadoTMCP1 =  TapaMovil.col_values(10)
                        mensajeParoProgramadoTMCP1="*Paro programado - Tiempo:* "+TiempoParoProgramadoTMCP1[-1]+" min, *Razon:* "+RazonParoProgramadoTMCP1[-1]
                        print(mensajeParoProgramadoTMCP1)
                    else:
                        mensajeParoProgramadoTMCP1=""
                        print(mensajeParoProgramadoTMCP1)
                    
                    #INCIDENTES TAPA MOVIL COPA 1::
                    IncidenteTMCP1=TapaMovil.col_values(11)
                    if IncidenteTMCP1[-1]=="Si":
                        DescrIncidenteTMCP1=TapaMovil.col_values(13)
                        ValidarParoIncidenteTMCP1=TapaMovil.col_values(14)
                        mensajeIncidenteTMCP1="*Incidente y/o accidente ambiental y/o SST: Razon:* "+DescrIncidenteTMCP1[-1] + " No se generó paro"
                        print(mensajeIncidenteTMCP1)
                        if ValidarParoIncidenteTMCP1[-1]=="Si":   
                            TiempoIncidenteTMCP1=TapaMovil.col_values(15)
                            mensajeIncidenteTMCP1="*Incidente y/o accidente ambiental y/o SST - Tiempo:* "+TiempoIncidenteTMCP1[-1]+" min, *Razon:* "+DescrIncidenteTMCP1[-1]
                            print (mensajeIncidenteTMCP1)
                        else:
                            #DescrIncidenteTMCP1=TapaMovil.col_values(12)
                            mensajeIncidenteTMCP1="*Incidente y/o accidente ambiental y/o SST: Razon:* "+DescrIncidenteTMCP1[-1] + " No se generó paro"
                            print (mensajeIncidenteTMCP1)
                    else:
                        mensajeIncidenteTMCP1=""
                        print(mensajeIncidenteTMCP1)

                ##SERVICIOS PUBLICOS  TAPA MOVIL COPA1:::.
                    ServiciosPublicosTMCP1=TapaMovil.col_values(16)
                    if ServiciosPublicosTMCP1[-1]=="Si":
                        DescrServiciosPublicosTMCP1=TapaMovil.col_values(18)
                        TiempoServiciosPublucosTMCP1=TapaMovil.col_values(17)
                        mensajeServiciosPublicosTMCP1="*Hubo afectación en las unidades del hora a hora por falta de servicios públicos: Razon:* "+DescrServiciosPublicosTMCP1[-1]+ " -*Tiempo:* :"+TiempoServiciosPublucosTMCP1[-1]+"min"
                        print(mensajeServiciosPublicosTMCP1)
                    else:
                        mensajeServiciosPublicosTMCP1=""
                        print(mensajeServiciosPublicosTMCP1)
                #POR MAQUINA COPA1 TAPA MOVIL:::
                    MaquinaTMCP1=TapaMovil.col_values(19)
                    if MaquinaTMCP1[-1]=="Si":
                        DescrMaquinaTMCP1=TapaMovil.col_values(22)
                        TiempoMaquinaTMCP1=TapaMovil.col_values(20)
                        mensajeMaquinaTMCP1="*Hubo afectación en las unidades por Maquina/ Equipo: Razon:* "+DescrMaquinaTMCP1[-1]+ " - *Tiempo:* "+TiempoMaquinaTMCP1[-1]+"min" 
                        print(mensajeMaquinaTMCP1)
                    else:
                        mensajeMaquinaTMCP1=""
                        print(mensajeMaquinaTMCP1)

                #POR MANO DE OBRA COPA1 TAPA MOVIL::::::::
                    ManoDeObraTMCP1=TapaMovil.col_values(23)
                    if ManoDeObraTMCP1[-1]=="Si":
                        DescrManoDeObraTMCP1=TapaMovil.col_values(27)
                        TiempoManoDeObraTMCP1=TapaMovil.col_values(24)
                        mensajeManoDeObraTMCP1="*Hubo afectación en las unidades por Mano De Obra: Razon:* "+DescrManoDeObraTMCP1[-1]+ " - *Tiempo:* "+TiempoManoDeObraTMCP1[-1]+"min" 
                        print(mensajeManoDeObraTMCP1)
                    else:
                        mensajeManoDeObraTMCP1=""
                        print(mensajeManoDeObraTMCP1)

                #MATERIA PRIMA COPA1 TAPA MOVIL::::

                    MateriaPrimaTMCP1=TapaMovil.col_values(28)
                    if MateriaPrimaTMCP1[-1]=="Si":
                        DescrMateriaPrimaTMCP1=TapaMovil.col_values(32)
                        TiempoMateriaPrimaTMCP1=TapaMovil.col_values(29)
                        mensajeMateriaPrimaTMCP1="*Hubo afectación en las unidades por Materia Prima: Razon:* "+DescrMateriaPrimaTMCP1[-1]+ " - *Tiempo:* "+TiempoMateriaPrimaTMCP1[-1]+"min" 
                        print(mensajeMateriaPrimaTMCP1)
                    else:
                        mensajeMateriaPrimaTMCP1=""
                        print(mensajeMateriaPrimaTMCP1)

                #POR METODO COPA1 TAPA MOVIL:::
                    MetodoTMCP1=TapaMovil.col_values(33)
                    if MetodoTMCP1[-1]=="Si":
                        DescrMetodoTMCP1=TapaMovil.col_values(36)
                        TiempoMetodoTMCP1=TapaMovil.col_values(34)
                        mensajeMetodoTMCP1="*Hubo afectación en las unidades por Método: Razon:* "+DescrMetodoTMCP1[-1]+ "- *Tiempo:* "+TiempoMetodoTMCP1[-1]+"min" 
                        print(mensajeMetodoTMCP1)
                    else:
                        mensajeMetodoTMCP1=""
                        print(mensajeMetodoTMCP1)

                #SCRAP COPA1 TAPA MOVIL::::::::::
                    ScrapTMCP1=TapaMovil.col_values(37)
                    if ScrapTMCP1[-1]=="Si":
                        DescrScrapTMCP1=TapaMovil.col_values(39)
                        CantidadScrapTMCP1=TapaMovil.col_values(40)
                        mensajeScrapTMCP1="*Se generó SCRAP: Cantidad:* "+CantidadScrapTMCP1[-1]+" - *Razon:* "+DescrScrapTMCP1[-1]
                        print(mensajeScrapTMCP1)
                    else:
                        mensajeScrapTMCP1=""
                        print(mensajeScrapTMCP1)

                #REPROCESADAS COPA1 TAPA MOVIL::::::::
                    UnidadesReprocesadasTMCP1 =  TapaMovil.col_values(41)
                    if UnidadesReprocesadasTMCP1[-1]=="Si":
                        CantidadReprocesadasTMCP1=  TapaMovil.col_values(42)

                        mensajeUnidadesReprocesadasTMCP1="*Se reprocesaron unidades - Cantidad:* "+CantidadReprocesadasTMCP1[-1]+""
                        print (mensajeUnidadesReprocesadasTMCP1)
                    else:
                        mensajeUnidadesReprocesadasTMCP1=""
                        print(mensajeUnidadesReprocesadasTMCP1)

                ##SELECCIONA COPA 2 HACEB TAPA MOVIL::::::::::::::::::::::

                if SelectReferenciaTM[-1]=="Copa 2.0 Haceb":
                    print("COPA 2 HACEB::::")
                    #PAROS PROGRAMADOS TAPA MOVIL COPA 2 HACEB::
                    ParoProgramadoTMCP2H = TapaMovil.col_values(43)
                    if ParoProgramadoTMCP2H[-1]=="Si":
                        RazonParoProgramadoTMCP2H =  TapaMovil.col_values(44)
                        TiempoParoProgramadoTMCP2H =  TapaMovil.col_values(45)
                        mensajeParoProgramadoTMCP2H="*Paro programado - Tiempo:* "+TiempoParoProgramadoTMCP2H[-1]+" min, *Razon:* "+RazonParoProgramadoTMCP2H[-1]
                        print(mensajeParoProgramadoTMCP2H)
                    else:
                        mensajeParoProgramadoTMCP2H=""
                        print(mensajeParoProgramadoTMCP2H)
                    
                    #INCIDENTES TAPA MOVIL COPA 2 HACEB::
                    IncidenteTMCP2H=TapaMovil.col_values(46)
                    if IncidenteTMCP2H[-1]=="Si":
                        DescrIncidenteTMCP2H=TapaMovil.col_values(48)
                        ValidarParoIncidenteEMCP2H=TapaMovil.col_values(49)
                        mensajeIncidenteTMCP2H="*Incidente y/o accidente ambiental y/o SST: Razon:* "+DescrIncidenteTMCP2H[-1]+ " no se generó paro."
                        print(mensajeIncidenteTMCP2H)
                        if ValidarParoIncidenteEMCP2H[-1]=="Si":   
                            TiempoIncidenteTMCP2H=TapaMovil.col_values(50)
                            mensajeIncidenteTMCP2H="*Incidente y/o accidente ambiental y/o SST - Tiempo:* "+TiempoIncidenteTMCP2H[-1]+" min, *Razon:* "+DescrIncidenteTMCP2H[-1]
                            print (mensajeIncidenteTMCP2H)
                        else:
                            #DescrIncidenteTMCP1=TapaMovil.col_values(12)
                            mensajeIncidenteTMCP2H="*Incidente y/o accidente ambiental y/o SST: Razon:* "+DescrIncidenteTMCP2H[-1] + " no se generó paro."
                            print (mensajeIncidenteTMCP2H)
                    else:
                        mensajeIncidenteTMCP2H=""
                        print(mensajeIncidenteTMCP2H)
                
                ##SERVICIOS PUBLICOS COPA2 TAPA MOVIL::
                    ServiciosPublicosTMCP2H=TapaMovil.col_values(51)
                    if ServiciosPublicosTMCP2H[-1]=="Si":
                        DescrServiciosPublicosTMCP2H=TapaMovil.col_values(53)
                        TiempoServiciosPublicosTMCP2H=TapaMovil.col_values(52)
                        mensajeServiciosPublicosTMCP2H="*Hubo afectación en las unidades del hora a hora por falta de servicios públicos: Razon:* "+DescrServiciosPublicosTMCP2H[-1]+ " -*Tiempo:* :"+TiempoServiciosPublicosTMCP2H[-1]+"min"
                        print(mensajeServiciosPublicosTMCP2H)
                    else:
                        mensajeServiciosPublicosTMCP2H=""
                        print(mensajeServiciosPublicosTMCP2H)
                #POR MAQUINA COPA2 TAPA MOVIL:::
                    MaquinaTMCP2H=TapaMovil.col_values(54)
                    if MaquinaTMCP2H[-1]=="Si":
                        DescrMaquinaTMCP2H=TapaMovil.col_values(57)
                        TiempoMaquinaTMCP2H=TapaMovil.col_values(55)
                        mensajeMaquinaTMCP2H="*Hubo afectación en las unidades por Maquina/ Equipo: Razon:* "+DescrMaquinaTMCP2H[-1]+ " - *Tiempo:* "+TiempoMaquinaTMCP2H[-1]+"min" 
                        print(mensajeMaquinaTMCP2H)
                    else:
                        mensajeMaquinaTMCP2H=""
                        print(mensajeMaquinaTMCP2H)

                #POR MANO DE OBRA COPA2 TAPA MOVIL::::::::
                    ManoDeObraTMCP2H=TapaMovil.col_values(58)
                    if ManoDeObraTMCP2H[-1]=="Si":
                        DescrManoDeObraTMCP2H=TapaMovil.col_values(62)
                        TiempoManoDeObraTMCP2H=TapaMovil.col_values(59)
                        mensajeManoDeObraTMCP2H="*Hubo afectación en las unidades por Mano De Obra: Razon:* "+DescrManoDeObraTMCP2H[-1]+ " - *Tiempo:* "+TiempoManoDeObraTMCP2H[-1]+"min" 
                        print(mensajeManoDeObraTMCP2H)
                    else:
                        mensajeManoDeObraTMCP2H=""
                        print(mensajeManoDeObraTMCP2H)

                #MATERIA PRIMA COPA2 TAPA MOVIL::::

                    MateriaPrimaTMCP2H=TapaMovil.col_values(63)
                    if MateriaPrimaTMCP2H[-1]=="Si":
                        DescrMateriaPrimaTMCP2H=TapaMovil.col_values(67)
                        TiempoMateriaPrimaTMCP2H=TapaMovil.col_values(64)
                        mensajeMateriaPrimaTMCP2H="*Hubo afectación en las unidades por Materia Prima: Razon:* "+DescrMateriaPrimaTMCP2H[-1]+ " - *Tiempo:* "+TiempoMateriaPrimaTMCP2H[-1]+"min" 
                        print(mensajeMateriaPrimaTMCP2H)
                    else:
                        mensajeMateriaPrimaTMCP2H=""
                        print(mensajeMateriaPrimaTMCP2H)

                #POR METODO COPA2 TAPA MOVIL:::
                    MetodoTMCP2H=TapaMovil.col_values(68)
                    if MetodoTMCP2H[-1]=="Si":
                        DescrMetodoTMCP2H=TapaMovil.col_values(71)
                        TiempoMetodoTMCP2H=TapaMovil.col_values(69)
                        mensajeMetodoTMCP2H="*Hubo afectación en las unidades por Método: Razon:* "+DescrMetodoTMCP2H[-1]+ "- *Tiempo:* "+TiempoMetodoTMCP2H[-1]+"min" 
                        print(mensajeMetodoTMCP2H)
                    else:
                        mensajeMetodoTMCP2H=""
                        print(mensajeMetodoTMCP2H)

                #SCRAP COPA2 TAPA MOVIL::::::::::
                    ScrapTMCP2H=TapaMovil.col_values(72)
                    if ScrapTMCP2H[-1]=="Si":
                        DescrScrapTMCP2H=TapaMovil.col_values(74)
                        CantidadScrapTMCP2H=TapaMovil.col_values(75)
                        mensajeScrapTMCP2H="*Se generó SCRAP: Cantidad:* "+CantidadScrapTMCP2H[-1]+" - *Razon:* "+DescrScrapTMCP2H[-1]
                        print(mensajeScrapTMCP2H)
                    else:
                        mensajeScrapTMCP2H=""
                        print(mensajeScrapTMCP2H)

                #REPROCESADAS COPA2 TAPA MOVIL::::::::
                    UnidadesReprocesadasTMCP2H =  TapaMovil.col_values(76)
                    if UnidadesReprocesadasTMCP2H[-1]=="Si":
                        CantidadReprocesadasTMCP2H=  TapaMovil.col_values(77)

                        mensajeUnidadesReprocesadasTMCP2H="*Se reprocesaron unidades - Cantidad:* "+CantidadReprocesadasTMCP2H[-1]+""
                        print (mensajeUnidadesReprocesadasTMCP2H)
                    else:
                        mensajeUnidadesReprocesadasTMCP2H=""
                        print(mensajeUnidadesReprocesadasTMCP2H)

                ##COPA 2 WHIRLPOOL TAPA MOVIL::::::::::::::
                # TAPA MOVIL COPA 2 WHIRLPOOL__:::::
                if SelectReferenciaTM[-1]=="Copa 2.0 Whirlpool":
                    print("COPA 2 Whirlpool::::::")
                    #PAROS PROGRAMADOS::::
                    ParoProgramadoTMCP2W = TapaMovil.col_values(78)
                    if ParoProgramadoTMCP2W[-1]=="Si":
                        RazonParoProgramadoTMCP2W =  TapaMovil.col_values(79)
                        TiempoParoProgramadoTMCP2W =  TapaMovil.col_values(80)
                        mensajeParoProgramadoTMCP2W="*Paro programado - Tiempo:* "+TiempoParoProgramadoTMCP2W[-1]+" min, *Razon:* "+RazonParoProgramadoTMCP2W[-1]
                        print(mensajeParoProgramadoTMCP2W)
                    else:
                        mensajeParoProgramadoTMCP2W=""
                        print(mensajeParoProgramadoTMCP2W)
                    
                    #INCIDENTES WHIRLPOOL COPA 2 TAPA MOVIL:::::
                    IncidenteTMCP2W=TapaMovil.col_values(81)
                    if IncidenteTMCP2W[-1]=="Si":
                        DescrIncidenteTMCP2W=TapaMovil.col_values(83)
                        ValidarParoIncidenteEMCP2W=TapaMovil.col_values(84)
                        mensajeIncidenteTMCP2W="*Incidente y/o accidente ambiental y/o SST: Razon:* "+DescrIncidenteTMCP2W[-1]+ " no se generó paro."
                        print(mensajeIncidenteTMCP2W)
                        if ValidarParoIncidenteEMCP2W[-1]=="Si":   
                            TiempoIncidenteTMCP2W=TapaMovil.col_values(85)
                            mensajeIncidenteTMCP2W="*Incidente y/o accidente ambiental y/o SST - Tiempo:* "+TiempoIncidenteTMCP2W[-1]+" min, *Razon:* "+DescrIncidenteTMCP2W[-1]
                            print (mensajeIncidenteTMCP2W)
                        else:
                            #DescrIncidenteTMCP1=TapaMovil.col_values(12)
                            mensajeIncidenteTMCP2W="*Incidente y/o accidente ambiental y/o SST: Razon:* "+DescrIncidenteTMCP2W[-1] + " no se generó paro."
                            print (mensajeIncidenteTMCP2W)
                    else:
                        mensajeIncidenteTMCP2W=""
                        print(mensajeIncidenteTMCP2W)

                ##SERVICIOS PUBLICOS COPA2 WHIRPOOL TAPA MOVIL:::
                    ServiciosPublicosTMCP2W=TapaMovil.col_values(86)
                    if ServiciosPublicosTMCP2W[-1]=="Si":
                        DescrServiciosPublicosTMCP2W=TapaMovil.col_values(88)
                        TiempoServiciosPublicosTMCP2W=TapaMovil.col_values(87)
                        mensajeServiciosPublicosTMCP2W="*Hubo afectación en las unidades del hora a hora por falta de servicios públicos: Razon:* "+DescrServiciosPublicosTMCP2W[-1]+ " -*Tiempo:* :"+TiempoServiciosPublicosTMCP2W[-1]+"min"
                        print(mensajeServiciosPublicosTMCP2W)
                    else:
                        mensajeServiciosPublicosTMCP2W=""
                        print(mensajeServiciosPublicosTMCP2W)

                #POR MAQUINA COPA2 WHIRLPOOL TAPA MOVIL::::::::
                    MaquinaTMCP2W=TapaMovil.col_values(89)
                    if MaquinaTMCP2W[-1]=="Si":
                        DescrMaquinaTMCP2W=TapaMovil.col_values(92)
                        TiempoMaquinaTMCP2W=TapaMovil.col_values(90)
                        mensajeMaquinaTMCP2W="*Hubo afectación en las unidades por Maquina/ Equipo: Razon:* "+DescrMaquinaTMCP2W[-1]+ " - *Tiempo:* "+TiempoMaquinaTMCP2W[-1]+"min" 
                        print(mensajeMaquinaTMCP2W)
                    else:
                        mensajeMaquinaTMCP2W=""
                        print(mensajeMaquinaTMCP2W)

                #POR MANO DE OBRA COPA2 WHIRLPOOL TAPA MOVIL::::::::
                    ManoDeObraTMCP2W=TapaMovil.col_values(93)
                    if ManoDeObraTMCP2W[-1]=="Si":
                        DescrManoDeObraTMCP2W=TapaMovil.col_values(97)
                        TiempoManoDeObraTMCP2W=TapaMovil.col_values(94)
                        mensajeManoDeObraTMCP2W="*Hubo afectación en las unidades por Mano De Obra: Razon:* "+DescrManoDeObraTMCP2W[-1]+ " - *Tiempo:* "+TiempoManoDeObraTMCP2W[-1]+"min" 
                        print(mensajeManoDeObraTMCP2W)
                    else:
                        mensajeManoDeObraTMCP2W=""
                        print(mensajeManoDeObraTMCP2W)

                #MATERIA PRIMA COPA2 WHIRPOOL TAPA MOVIL::::

                    MateriaPrimaTMCP2W=TapaMovil.col_values(98)
                    if MateriaPrimaTMCP2W[-1]=="Si":
                        DescrMateriaPrimaTMCP2W=TapaMovil.col_values(102)
                        TiempoMateriaPrimaTMCP2W=TapaMovil.col_values(99)
                        mensajeMateriaPrimaTMCP2W="*Hubo afectación en las unidades por Materia Prima: Razon:* "+DescrMateriaPrimaTMCP2W[-1]+ " - *Tiempo:* "+TiempoMateriaPrimaTMCP2W[-1]+"min" 
                        print(mensajeMateriaPrimaTMCP2W)
                    else:
                        mensajeMateriaPrimaTMCP2W=""
                        print(mensajeMateriaPrimaTMCP2W)

                #POR METODO COPA2 WHIRLPOOL TAPA MOVIL:::
                    MetodoTMCP2W=TapaMovil.col_values(103)
                    if MetodoTMCP2W[-1]=="Si":
                        DescrMetodoTMCP2W=TapaMovil.col_values(106)
                        TiempoMetodoTMCP2W=TapaMovil.col_values(104)
                        mensajeMetodoTMCP2W="*Hubo afectación en las unidades por Método: Razon:* "+DescrMetodoTMCP2W[-1]+ "- *Tiempo:* "+TiempoMetodoTMCP2W[-1]+"min" 
                        print(mensajeMetodoTMCP2W)
                    else:
                        mensajeMetodoTMCP2W=""
                        print(mensajeMetodoTMCP2W)

                #SCRAP COPA2 WHIRLPOOL TAPA MOVIL::::::::::
                    ScrapTMCP2W=TapaMovil.col_values(107)
                    if ScrapTMCP2W[-1]=="Si":
                        DescrScrapTMCP2W=TapaMovil.col_values(109)
                        CantidadScrapTMCP2W=TapaMovil.col_values(110)
                        mensajeScrapTMCP2W="*Se generó SCRAP: Cantidad:* "+CantidadScrapTMCP2W[-1]+" - *Razon:* "+DescrScrapTMCP2W[-1]
                        print(mensajeScrapTMCP2W)
                    else:
                        mensajeScrapTMCP2W=""
                        print(mensajeScrapTMCP2W)

                #REPROCESADAS COPA2::::::::
                    UnidadesReprocesadasTMCP2W =  TapaMovil.col_values(111)
                    if UnidadesReprocesadasTMCP2W[-1]=="Si":
                        CantidadReprocesadasTMCP2W=  TapaMovil.col_values(112)

                        mensajeUnidadesReprocesadasTMCP2W="*Se reprocesaron unidades - Cantidad:* "+CantidadReprocesadasTMCP2W[-1]+""
                        print (mensajeUnidadesReprocesadasTMCP2W)
                    else:
                        mensajeUnidadesReprocesadasTMCP2W=""
                        print(mensajeUnidadesReprocesadasTMCP2W)


                OeeTM=TapaMovil.col_values(118)
                OeeTapaMovil = OeeTM[-1]
                print("Porcentaje OEE: "+ OeeTapaMovil)
                MensajeOeeTM="OEE: "+OeeTapaMovil

                mensaje4="\n*TAPA MOVIL*" 
                if MensajeUnidadesFabricadasTM!="":
                    mensaje4=mensaje4+"\n\n"+MensajeUnidadesFabricadasTM

                if SelectReferenciaTM[-1]=="Copa 1.0":
                    mensaje4=mensaje4+"\n\n*COPA 1.0*"
                
                    if mensajeParoProgramadoTMCP1!="":
                        mensaje4=mensaje4+"\n"+mensajeParoProgramadoTMCP1
                        SinNovedad=2
                    if mensajeIncidenteTMCP1!="":
                        mensaje4=mensaje4+"\n"+mensajeIncidenteTMCP1
                        SinNovedad=2
                    if mensajeServiciosPublicosTMCP1!="":
                        mensaje4=mensaje4+"\n"+mensajeServiciosPublicosTMCP1
                        SinNovedad=2
                    if mensajeMaquinaTMCP1!="":
                        mensaje4=mensaje4+"\n"+mensajeMaquinaTMCP1
                        SinNovedad=2
                    if mensajeManoDeObraTMCP1!="":
                        mensaje4=mensaje4+"\n"+mensajeManoDeObraTMCP1
                        SinNovedad=2
                    if mensajeMateriaPrimaTMCP1!="":
                        mensaje4=mensaje4+"\n"+mensajeMateriaPrimaTMCP1
                        SinNovedad=2
                    if mensajeMetodoTMCP1!="":
                        mensaje4=mensaje4+"\n"+mensajeMetodoTMCP1
                        SinNovedad=2
                    if mensajeScrapTMCP1!="":
                        mensaje4=mensaje4+"\n"+mensajeScrapTMCP1
                        SinNovedad=2
                    if mensajeUnidadesReprocesadasTMCP1!="":
                        mensaje4=mensaje4+"\n"+mensajeUnidadesReprocesadasTMCP1
                        SinNovedad=2
                
                if SelectReferenciaTM[-1]=="Copa 2.0 Haceb":
                    mensaje4=mensaje4+"\n\n*COPA 2.0 HACEB*"
                
                    if mensajeParoProgramadoTMCP2H!="":
                        mensaje4=mensaje4+"\n"+mensajeParoProgramadoTMCP2H
                        SinNovedad=2
                    if mensajeIncidenteTMCP2H!="":
                        mensaje4=mensaje4+"\n"+mensajeIncidenteTMCP2H
                        SinNovedad=2
                    if mensajeServiciosPublicosTMCP2H!="":
                        mensaje4=mensaje4+"\n"+mensajeServiciosPublicosTMCP2H
                        SinNovedad=2
                    if mensajeMaquinaTMCP2H!="":
                        mensaje4=mensaje4+"\n"+mensajeMaquinaTMCP2H
                        SinNovedad=2
                    if mensajeManoDeObraTMCP2H!="":
                        mensaje4=mensaje4+"\n"+mensajeManoDeObraTMCP2H
                        SinNovedad=2
                    if mensajeMateriaPrimaTMCP2H!="":
                        mensaje4=mensaje4+"\n"+mensajeMateriaPrimaTMCP2H
                        SinNovedad=2
                    if mensajeMetodoTMCP2H!="":
                        mensaje4=mensaje4+"\n"+mensajeMetodoTMCP2H
                        SinNovedad=2
                    if mensajeScrapTMCP2H!="":
                        mensaje4=mensaje4+"\n"+mensajeScrapTMCP2H
                        SinNovedad=2
                    if mensajeUnidadesReprocesadasTMCP2H!="":
                        mensaje4=mensaje4+"\n"+mensajeUnidadesReprocesadasTMCP2H
                        SinNovedad=2

                if SelectReferenciaTM[-1]=="Copa 2.0 Whirlpool":
                    mensaje4=mensaje4+"\n\n*COPA 2.0 WHIRLPOOL*"
                
                    if mensajeParoProgramadoTMCP2W!="":
                        mensaje4=mensaje4+"\n"+mensajeParoProgramadoTMCP2W
                        SinNovedad=2
                    if mensajeIncidenteTMCP2W!="":
                        mensaje4=mensaje4+"\n"+mensajeIncidenteTMCP2W
                        SinNovedad=2
                    if mensajeServiciosPublicosTMCP2W!="":
                        mensaje4=mensaje4+"\n"+mensajeServiciosPublicosTMCP2W
                        SinNovedad=2
                    if mensajeMaquinaTMCP2W!="":
                        mensaje4=mensaje4+"\n"+mensajeMaquinaTMCP2W
                        SinNovedad=2
                    if mensajeManoDeObraTMCP2W!="":
                        mensaje4=mensaje4+"\n"+mensajeManoDeObraTMCP2W
                        SinNovedad=2
                    if mensajeMateriaPrimaTMCP2W!="":
                        mensaje4=mensaje4+"\n"+mensajeMateriaPrimaTMCP2W
                        SinNovedad=2
                    if mensajeMetodoTMCP2W!="":
                        mensaje4=mensaje4+"\n"+mensajeMetodoTMCP2W
                        SinNovedad=2
                    if mensajeScrapTMCP2W!="":
                        mensaje4=mensaje4+"\n"+mensajeScrapTMCP2W
                        SinNovedad=2
                    if mensajeUnidadesReprocesadasTMCP2W!="":
                        mensaje4=mensaje4+"\n"+mensajeUnidadesReprocesadasTMCP2W
                        SinNovedad=2         

                if SinNovedad!=2:
                    mensaje4=mensaje4+"\n*No se reportaron novedades*"   
                
                SinNovedad=0

                if MensajeOeeTM!="" and OeeTapaMovil!="#DIV/0!":
                    mensaje4=mensaje4+"\n"+MensajeOeeTM

        if espera_Minuto:
            print("2 Esperando minuto de envio...")
            espera_Minuto= False

        #CONDICIONAL PARA SELECCIONAR EL MIN Y EL RANGO DE SEGUNDOS
        if (Minuto2==8 and Segundo>=40) and Segundo<=45:

            #VALIDAR QUE SI HAYA UN REPORTE
            if str(HoraTF2[-1])+"\n" == HoraGuardadaTF2:
                print ("Tapa fija, no reporto")
                mensaje5="*TAPA FIJA*/n/n*La celula no reporto.*"
            else:
                #TAPA FIJA:::::::::::::::::
                print("TAPA FIJA:::---------")

                #SELECCION DE LA HOJA::
                TapaFija = sh.get_worksheet(5)
                #SELECCIONAR LA REFERENCIA:::::Back Panel -- COPA 2 WHIRLPOOL -- AGIPELER --- BACK PANEL --- IMPELER --- QUASAR
                UnidadesFabricadasTF2=  TapaFija.col_values(5)
                print(UnidadesFabricadasTF2[-1])
                mensajeUnidadesFabricadasTF2="*Unidades producidas:* "+UnidadesFabricadasTF2[-1]
                #SELECCION DE Agipeller:::::
                SelectReferenciaTF2 = TapaFija.col_values(7)
                if SelectReferenciaTF2[-1]=="Agipeller":
                    print("Agipeller::::")
                    #PAROS PROGRAMADOS TAPA FIJA Agipeller::::
                    ParoProgramadoTF2AGI = TapaFija.col_values(8)
                    if ParoProgramadoTF2AGI[-1]=="Si":
                        RazonParoProgramadoTF2AGI =  TapaFija.col_values(9)
                        TiempoParoProgramadoTF2AGI =  TapaFija.col_values(10)
                        mensajeParoProgramadoTF2AGI="*Paro programado - Tiempo:* "+TiempoParoProgramadoTF2AGI[-1]+" min, *Razon:* "+RazonParoProgramadoTF2AGI[-1]
                        print(mensajeParoProgramadoTF2AGI)
                    else:
                        mensajeParoProgramadoTF2AGI=""
                        print(mensajeParoProgramadoTF2AGI)
                    
                    #INCIDENTES TAPA FIJA Agipeller::
                    IncidenteTF2AGI=TapaFija.col_values(11)
                    if IncidenteTF2AGI[-1]=="Si":
                        DescrIncidenteTF2AGI=TapaFija.col_values(13)
                        ValidarParoIncidenteTF2AGI=TapaFija.col_values(14)
                        mensajeIncidenteTF2AGI="*Incidente y/o accidente ambiental y/o SST: Razon:* "+DescrIncidenteTF2AGI[-1] + " No se generó paro"
                        print(mensajeIncidenteTF2AGI)
                        if ValidarParoIncidenteTF2AGI[-1]=="Si":   
                            TiempoIncidenteTF2AGI=TapaFija.col_values(15)
                            mensajeIncidenteTF2AGI="*Incidente y/o accidente ambiental y/o SST - Tiempo:* "+TiempoIncidenteTF2AGI[-1]+" min, *Razon:* "+DescrIncidenteTF2AGI[-1]
                            print (mensajeIncidenteTF2AGI)
                        else:
                            #DescrIncidenteTF2AGI=TapaFija.col_values(12)
                            mensajeIncidenteTF2AGI="*Incidente y/o accidente ambiental y/o SST: Razon:* "+DescrIncidenteTF2AGI[-1] + " No se generó paro"
                            print (mensajeIncidenteTF2AGI)
                    else:
                        mensajeIncidenteTF2AGI=""
                        print(mensajeIncidenteTF2AGI)

                ##SERVICIOS PUBLICOS  TAPA FIJA Agipeller:::.
                    ServiciosPublicosTF2AGI=TapaFija.col_values(16)
                    if ServiciosPublicosTF2AGI[-1]=="Si":
                        DescrServiciosPublicosTF2AGI=TapaFija.col_values(18)
                        TiempoServiciosPublucosTF2AGI=TapaFija.col_values(17)
                        mensajeServiciosPublicosTF2AGI="*Hubo afectación en las unidades del hora a hora por falta de servicios públicos: Razon:* "+DescrServiciosPublicosTF2AGI[-1]+ " -*Tiempo:* :"+TiempoServiciosPublucosTF2AGI[-1]+"min"
                        print(mensajeServiciosPublicosTF2AGI)
                    else:
                        mensajeServiciosPublicosTF2AGI=""
                        print(mensajeServiciosPublicosTF2AGI)
                #POR MAQUINA Agipeller TAPA FIJA:::
                    MaquinaTF2AGI=TapaFija.col_values(19)
                    if MaquinaTF2AGI[-1]=="Si":
                        DescrMaquinaTF2AGI=TapaFija.col_values(22)
                        TiempoMaquinaTF2AGI=TapaFija.col_values(20)
                        mensajeMaquinaTF2AGI="*Hubo afectación en las unidades por Maquina/ Equipo: Razon:* "+DescrMaquinaTF2AGI[-1]+ " - *Tiempo:* "+TiempoMaquinaTF2AGI[-1]+"min" 
                        print(mensajeMaquinaTF2AGI)
                    else:
                        mensajeMaquinaTF2AGI=""
                        print(mensajeMaquinaTF2AGI)

                #POR MANO DE OBRA Agipeller TAPA FIJA::::::::
                    ManoDeObraTF2AGI=TapaFija.col_values(23)
                    if ManoDeObraTF2AGI[-1]=="Si":
                        DescrManoDeObraTF2AGI=TapaFija.col_values(27)
                        TiempoManoDeObraTF2AGI=TapaFija.col_values(24)
                        mensajeManoDeObraTF2AGI="*Hubo afectación en las unidades por Mano De Obra: Razon:* "+DescrManoDeObraTF2AGI[-1]+ " - *Tiempo:* "+TiempoManoDeObraTF2AGI[-1]+"min" 
                        print(mensajeManoDeObraTF2AGI)
                    else:
                        mensajeManoDeObraTF2AGI=""
                        print(mensajeManoDeObraTF2AGI)

                #MATERIA PRIMA Agipeller TAPA FIJA::::

                    MateriaPrimaTF2AGI=TapaFija.col_values(28)
                    if MateriaPrimaTF2AGI[-1]=="Si":
                        DescrMateriaPrimaTF2AGI=TapaFija.col_values(32)
                        TiempoMateriaPrimaTF2AGI=TapaFija.col_values(29)
                        mensajeMateriaPrimaTF2AGI="*Hubo afectación en las unidades por Materia Prima: Razon:* "+DescrMateriaPrimaTF2AGI[-1]+ " - *Tiempo:* "+TiempoMateriaPrimaTF2AGI[-1]+"min" 
                        print(mensajeMateriaPrimaTF2AGI)
                    else:
                        mensajeMateriaPrimaTF2AGI=""
                        print(mensajeMateriaPrimaTF2AGI)

                #POR METODO Agipeller TAPA FIJA:::
                    MetodoTF2AGI=TapaFija.col_values(33)
                    if MetodoTF2AGI[-1]=="Si":
                        DescrMetodoTF2AGI=TapaFija.col_values(36)
                        TiempoMetodoTF2AGI=TapaFija.col_values(34)
                        mensajeMetodoTF2AGI="*Hubo afectación en las unidades por Método: Razon:* "+DescrMetodoTF2AGI[-1]+ "- *Tiempo:* "+TiempoMetodoTF2AGI[-1]+"min" 
                        print(mensajeMetodoTF2AGI)
                    else:
                        mensajeMetodoTF2AGI=""
                        print(mensajeMetodoTF2AGI)

                #SCRAP Agipeller TAPA FIJA::::::::::
                    ScrapTF2AGI=TapaFija.col_values(37)
                    if ScrapTF2AGI[-1]=="Si":
                        DescrScrapTF2AGI=TapaFija.col_values(39)
                        CantidadScrapTF2AGI=TapaFija.col_values(40)
                        mensajeScrapTF2AGI="*Se generó SCRAP: Cantidad:* "+CantidadScrapTF2AGI[-1]+" - *Razon:* "+DescrScrapTF2AGI[-1]
                        print(mensajeScrapTF2AGI)
                    else:
                        mensajeScrapTF2AGI=""
                        print(mensajeScrapTF2AGI)

                #REPROCESADAS Agipeller TAPA FIJA::::::::
                    UnidadesReprocesadasTF2AGI =  TapaFija.col_values(41)
                    if UnidadesReprocesadasTF2AGI[-1]=="Si":
                        CantidadReprocesadasTF2AGI=  TapaFija.col_values(42)

                        mensajeUnidadesReprocesadasTF2AGI="*Se reprocesaron unidades - Cantidad:* "+CantidadReprocesadasTF2AGI[-1]+""
                        print (mensajeUnidadesReprocesadasTF2AGI)
                    else:
                        mensajeUnidadesReprocesadasTF2AGI=""
                        print(mensajeUnidadesReprocesadasTF2AGI)

                ##SELECCIONA Back Panel TAPA FIJA::::::::::::::::::::::________________________________________
                #_____________________________________________________

                if SelectReferenciaTF2[-1]=="Back Panel":
                    print("Back Panel:::")
                    #PAROS PROGRAMADOS TAPA FIJA Back Panel::
                    ParoProgramadoTF2BP = TapaFija.col_values(43)
                    if ParoProgramadoTF2BP[-1]=="Si":
                        RazonParoProgramadoTF2BP =  TapaFija.col_values(44)
                        TiempoParoProgramadoTF2BP =  TapaFija.col_values(45)
                        mensajeParoProgramadoTF2BP="*Paro programado - Tiempo:* "+TiempoParoProgramadoTF2BP[-1]+" min, *Razon:* "+RazonParoProgramadoTF2BP[-1]
                        print(mensajeParoProgramadoTF2BP)
                    else:
                        mensajeParoProgramadoTF2BP=""
                        print(mensajeParoProgramadoTF2BP)
                    
                    #INCIDENTES TAPA FIJA Back Panel::
                    IncidenteTF2BP=TapaFija.col_values(46)
                    if IncidenteTF2BP[-1]=="Si":
                        DescrIncidenteTF2BP=TapaFija.col_values(48)
                        ValidarParoIncidenteEMCP2H=TapaFija.col_values(49)
                        mensajeIncidenteTF2BP="*Incidente y/o accidente ambiental y/o SST: Razon:* "+DescrIncidenteTF2BP[-1]+ " no se generó paro."
                        print(mensajeIncidenteTF2BP)
                        if ValidarParoIncidenteEMCP2H[-1]=="Si":   
                            TiempoIncidenteTF2BP=TapaFija.col_values(50)
                            mensajeIncidenteTF2BP="*Incidente y/o accidente ambiental y/o SST - Tiempo:* "+TiempoIncidenteTF2BP[-1]+" min, *Razon:* "+DescrIncidenteTF2BP[-1]
                            print (mensajeIncidenteTF2BP)
                        else:
                            #DescrIncidenteTF2AGI=TapaFija.col_values(12)
                            mensajeIncidenteTF2BP="*Incidente y/o accidente ambiental y/o SST: Razon:* "+DescrIncidenteTF2BP[-1] + " no se generó paro."
                            print (mensajeIncidenteTF2BP)
                    else:
                        mensajeIncidenteTF2BP=""
                        print(mensajeIncidenteTF2BP)
                
                ##SERVICIOS PUBLICOS backpanel TAPA FIJA::
                    ServiciosPublicosTF2BP=TapaFija.col_values(51)
                    if ServiciosPublicosTF2BP[-1]=="Si":
                        DescrServiciosPublicosTF2BP=TapaFija.col_values(53)
                        TiempoServiciosPublicosTF2BP=TapaFija.col_values(52)
                        mensajeServiciosPublicosTF2BP="*Hubo afectación en las unidades del hora a hora por falta de servicios públicos: Razon:* "+DescrServiciosPublicosTF2BP[-1]+ " -*Tiempo:* :"+TiempoServiciosPublicosTF2BP[-1]+"min"
                        print(mensajeServiciosPublicosTF2BP)
                    else:
                        mensajeServiciosPublicosTF2BP=""
                        print(mensajeServiciosPublicosTF2BP)
                #POR MAQUINA backpanel TAPA FIJA:::
                    MaquinaTF2BP=TapaFija.col_values(54)
                    if MaquinaTF2BP[-1]=="Si":
                        DescrMaquinaTF2BP=TapaFija.col_values(57)
                        TiempoMaquinaTF2BP=TapaFija.col_values(55)
                        mensajeMaquinaTF2BP="*Hubo afectación en las unidades por Maquina/ Equipo: Razon:* "+DescrMaquinaTF2BP[-1]+ " - *Tiempo:* "+TiempoMaquinaTF2BP[-1]+"min" 
                        print(mensajeMaquinaTF2BP)
                    else:
                        mensajeMaquinaTF2BP=""
                        print(mensajeMaquinaTF2BP)

                #POR MANO DE OBRA backpanel TAPA FIJA::::::::
                    ManoDeObraTF2BP=TapaFija.col_values(58)
                    if ManoDeObraTF2BP[-1]=="Si":
                        DescrManoDeObraTF2BP=TapaFija.col_values(62)
                        TiempoManoDeObraTF2BP=TapaFija.col_values(59)
                        mensajeManoDeObraTF2BP="*Hubo afectación en las unidades por Mano De Obra: Razon:* "+DescrManoDeObraTF2BP[-1]+ " - *Tiempo:* "+TiempoManoDeObraTF2BP[-1]+"min" 
                        print(mensajeManoDeObraTF2BP)
                    else:
                        mensajeManoDeObraTF2BP=""
                        print(mensajeManoDeObraTF2BP)

                #MATERIA PRIMA COPA2 TAPA FIJA::::

                    MateriaPrimaTF2BP=TapaFija.col_values(63)
                    if MateriaPrimaTF2BP[-1]=="Si":
                        DescrMateriaPrimaTF2BP=TapaFija.col_values(67)
                        TiempoMateriaPrimaTF2BP=TapaFija.col_values(64)
                        mensajeMateriaPrimaTF2BP="*Hubo afectación en las unidades por Materia Prima: Razon:* "+DescrMateriaPrimaTF2BP[-1]+ " - *Tiempo:* "+TiempoMateriaPrimaTF2BP[-1]+"min" 
                        print(mensajeMateriaPrimaTF2BP)
                    else:
                        mensajeMateriaPrimaTF2BP=""
                        print(mensajeMateriaPrimaTF2BP)

                #POR METODO COPA2 TAPA FIJA:::
                    MetodoTF2BP=TapaFija.col_values(68)
                    if MetodoTF2BP[-1]=="Si":
                        DescrMetodoTF2BP=TapaFija.col_values(71)
                        TiempoMetodoTF2BP=TapaFija.col_values(69)
                        mensajeMetodoTF2BP="*Hubo afectación en las unidades por Método: Razon:* "+DescrMetodoTF2BP[-1]+ "- *Tiempo:* "+TiempoMetodoTF2BP[-1]+"min" 
                        print(mensajeMetodoTF2BP)
                    else:
                        mensajeMetodoTF2BP=""
                        print(mensajeMetodoTF2BP)

                #SCRAP COPA2 TAPA FIJA::::::::::
                    ScrapTF2BP=TapaFija.col_values(72)
                    if ScrapTF2BP[-1]=="Si":
                        DescrScrapTF2BP=TapaFija.col_values(74)
                        CantidadScrapTF2BP=TapaFija.col_values(75)
                        mensajeScrapTF2BP="*Se generó SCRAP: Cantidad:* "+CantidadScrapTF2BP[-1]+" - *Razon:* "+DescrScrapTF2BP[-1]
                        print(mensajeScrapTF2BP)
                    else:
                        mensajeScrapTF2BP=""
                        print(mensajeScrapTF2BP)

                #REPROCESADAS COPA2 TAPA FIJA::::::::
                    UnidadesReprocesadasTF2BP =  TapaFija.col_values(76)
                    if UnidadesReprocesadasTF2BP[-1]=="Si":
                        CantidadReprocesadasTF2BP=  TapaFija.col_values(77)

                        mensajeUnidadesReprocesadasTF2BP="*Se reprocesaron unidades - Cantidad:* "+CantidadReprocesadasTF2BP[-1]+""
                        print (mensajeUnidadesReprocesadasTF2BP)
                    else:
                        mensajeUnidadesReprocesadasTF2BP=""
                        print(mensajeUnidadesReprocesadasTF2BP)

                ##Copa 2.0 Haceb TAPA FIJA::::::::::::::
                # TAPA FIJA Copa 2.0 Haceb__:::::
                if SelectReferenciaTF2[-1]=="Copa 2.0 Haceb":
                    print("Copa 2.0 Haceb::::::")
                    #PAROS PROGRAMADOS::::
                    ParoProgramadoTF2CP2H = TapaFija.col_values(78)
                    if ParoProgramadoTF2CP2H[-1]=="Si":
                        RazonParoProgramadoTF2CP2H =  TapaFija.col_values(79)
                        TiempoParoProgramadoTF2CP2H =  TapaFija.col_values(80)
                        mensajeParoProgramadoTF2CP2H="*Paro programado - Tiempo:* "+TiempoParoProgramadoTF2CP2H[-1]+" min, *Razon:* "+RazonParoProgramadoTF2CP2H[-1]
                        print(mensajeParoProgramadoTF2CP2H)
                    else:
                        mensajeParoProgramadoTF2CP2H=""
                        print(mensajeParoProgramadoTF2CP2H)
                    
                    #INCIDENTES Copa 2.0 Haceb TAPA FIJA:::::
                    IncidenteTF2CP2H=TapaFija.col_values(81)
                    if IncidenteTF2CP2H[-1]=="Si":
                        DescrIncidenteTF2CP2H=TapaFija.col_values(83)
                        ValidarParoIncidenteEMCP2W=TapaFija.col_values(84)
                        mensajeIncidenteTF2CP2H="*Incidente y/o accidente ambiental y/o SST: Razon:* "+DescrIncidenteTF2CP2H[-1]+ " no se generó paro."
                        print(mensajeIncidenteTF2CP2H)
                        if ValidarParoIncidenteEMCP2W[-1]=="Si":   
                            TiempoIncidenteTF2CP2H=TapaFija.col_values(85)
                            mensajeIncidenteTF2CP2H="*Incidente y/o accidente ambiental y/o SST - Tiempo:* "+TiempoIncidenteTF2CP2H[-1]+" min, *Razon:* "+DescrIncidenteTF2CP2H[-1]
                            print (mensajeIncidenteTF2CP2H)
                        else:
                            #DescrIncidenteTF2AGI=TapaFija.col_values(12)
                            mensajeIncidenteTF2CP2H="*Incidente y/o accidente ambiental y/o SST: Razon:* "+DescrIncidenteTF2CP2H[-1] + " no se generó paro."
                            print (mensajeIncidenteTF2CP2H)
                    else:
                        mensajeIncidenteTF2CP2H=""
                        print(mensajeIncidenteTF2CP2H)

                ##SERVICIOS PUBLICOS Copa 2.0 Haceb TAPA FIJA:::
                    ServiciosPublicosTF2CP2H=TapaFija.col_values(86)
                    if ServiciosPublicosTF2CP2H[-1]=="Si":
                        DescrServiciosPublicosTF2CP2H=TapaFija.col_values(88)
                        TiempoServiciosPublicosTF2CP2H=TapaFija.col_values(87)
                        mensajeServiciosPublicosTF2CP2H="*Hubo afectación en las unidades del hora a hora por falta de servicios públicos: Razon:* "+DescrServiciosPublicosTF2CP2H[-1]+ " -*Tiempo:* :"+TiempoServiciosPublicosTF2CP2H[-1]+"min"
                        print(mensajeServiciosPublicosTF2CP2H)
                    else:
                        mensajeServiciosPublicosTF2CP2H=""
                        print(mensajeServiciosPublicosTF2CP2H)

                #POR MAQUINA Copa 2.0 Haceb TAPA FIJA::::::::
                    MaquinaTF2CP2H=TapaFija.col_values(89)
                    if MaquinaTF2CP2H[-1]=="Si":
                        DescrMaquinaTF2CP2H=TapaFija.col_values(92)
                        TiempoMaquinaTF2CP2H=TapaFija.col_values(90)
                        mensajeMaquinaTF2CP2H="*Hubo afectación en las unidades por Maquina/ Equipo: Razon:* "+DescrMaquinaTF2CP2H[-1]+ " - *Tiempo:* "+TiempoMaquinaTF2CP2H[-1]+"min" 
                        print(mensajeMaquinaTF2CP2H)
                    else:
                        mensajeMaquinaTF2CP2H=""
                        print(mensajeMaquinaTF2CP2H)

                #POR MANO DE OBRA Copa 2.0 Haceb TAPA FIJA::::::::
                    ManoDeObraTF2CP2H=TapaFija.col_values(93)
                    if ManoDeObraTF2CP2H[-1]=="Si":
                        DescrManoDeObraTF2CP2H=TapaFija.col_values(97)
                        TiempoManoDeObraTF2CP2H=TapaFija.col_values(94)
                        mensajeManoDeObraTF2CP2H="*Hubo afectación en las unidades por Mano De Obra: Razon:* "+DescrManoDeObraTF2CP2H[-1]+ " - *Tiempo:* "+TiempoManoDeObraTF2CP2H[-1]+"min" 
                        print(mensajeManoDeObraTF2CP2H)
                    else:
                        mensajeManoDeObraTF2CP2H=""
                        print(mensajeManoDeObraTF2CP2H)

                #MATERIA PRIMA Copa 2.0 Haceb TAPA FIJA::::

                    MateriaPrimaTF2CP2H=TapaFija.col_values(98)
                    if MateriaPrimaTF2CP2H[-1]=="Si":
                        DescrMateriaPrimaTF2CP2H=TapaFija.col_values(102)
                        TiempoMateriaPrimaTF2CP2H=TapaFija.col_values(99)
                        mensajeMateriaPrimaTF2CP2H="*Hubo afectación en las unidades por Materia Prima: Razon:* "+DescrMateriaPrimaTF2CP2H[-1]+ " - *Tiempo:* "+TiempoMateriaPrimaTF2CP2H[-1]+"min" 
                        print(mensajeMateriaPrimaTF2CP2H)
                    else:
                        mensajeMateriaPrimaTF2CP2H=""
                        print(mensajeMateriaPrimaTF2CP2H)

                #POR METODO Copa 2.0 Haceb TAPA FIJA:::
                    MetodoTF2CP2H=TapaFija.col_values(103)
                    if MetodoTF2CP2H[-1]=="Si":
                        DescrMetodoTF2CP2H=TapaFija.col_values(106)
                        TiempoMetodoTF2CP2H=TapaFija.col_values(104)
                        mensajeMetodoTF2CP2H="*Hubo afectación en las unidades por Método: Razon:* "+DescrMetodoTF2CP2H[-1]+ "- *Tiempo:* "+TiempoMetodoTF2CP2H[-1]+"min" 
                        print(mensajeMetodoTF2CP2H)
                    else:
                        mensajeMetodoTF2CP2H=""
                        print(mensajeMetodoTF2CP2H)

                #SCRAP Copa 2.0 Haceb TAPA FIJA::::::::::
                    ScrapTF2CP2H=TapaFija.col_values(107)
                    if ScrapTF2CP2H[-1]=="Si":
                        DescrScrapTF2CP2H=TapaFija.col_values(109)
                        CantidadScrapTF2CP2H=TapaFija.col_values(110)
                        mensajeScrapTF2CP2H="*Se generó SCRAP: Cantidad:* "+CantidadScrapTF2CP2H[-1]+" - *Razon:* "+DescrScrapTF2CP2H[-1]
                        print(mensajeScrapTF2CP2H)
                    else:
                        mensajeScrapTF2CP2H=""
                        print(mensajeScrapTF2CP2H)

                #REPROCESADAS Copa 2.0 Haceb::::::::
                    UnidadesReprocesadasTF2CP2H =  TapaFija.col_values(111)
                    if UnidadesReprocesadasTF2CP2H[-1]=="Si":
                        CantidadReprocesadasTF2CP2H=  TapaFija.col_values(112)

                        mensajeUnidadesReprocesadasTF2CP2H="*Se reprocesaron unidades - Cantidad:* "+CantidadReprocesadasTF2CP2H[-1]+""
                        print (mensajeUnidadesReprocesadasTF2CP2H)
                    else:
                        mensajeUnidadesReprocesadasTF2CP2H=""
                        print(mensajeUnidadesReprocesadasTF2CP2H)


                ##COPA 2WHIRLPOOL::::::::::::__________________________________________________________________

                ##COPA 2WHIRLPOOL Haceb TAPA FIJA::::::::::::::
                # TAPA FIJA COPA 2WHIRLPOOL Haceb__:::::
                if SelectReferenciaTF2[-1]=="Copa 2.0 Whirlpool":
                    print("COPA 2 WHIRLPOOL::::::")
                    #PAROS PROGRAMADOS::::
                    ParoProgramadoTF2CP2W = TapaFija.col_values(113)
                    if ParoProgramadoTF2CP2W[-1]=="Si":
                        RazonParoProgramadoTF2CP2W =  TapaFija.col_values(114)
                        TiempoParoProgramadoTF2CP2W =  TapaFija.col_values(115)
                        mensajeParoProgramadoTF22CP2W="*Paro programado - Tiempo:* "+TiempoParoProgramadoTF2CP2W[-1]+" min, *Razon:* "+RazonParoProgramadoTF2CP2W[-1]
                        print(mensajeParoProgramadoTF22CP2W)
                    else:
                        mensajeParoProgramadoTF22CP2W=""
                        print(mensajeParoProgramadoTF22CP2W)
                    
                    #INCIDENTES Copa 2.0 Whirlpool TAPA FIJA:::::
                    IncidenteTF2CP2W=TapaFija.col_values(116)
                    if IncidenteTF2CP2W[-1]=="Si":
                        DescrIncidenteTF2CP2W=TapaFija.col_values(118)
                        ValidarParoIncidenteEMCP2W=TapaFija.col_values(119)
                        mensajeIncidenteTF2CP2W="*Incidente y/o accidente ambiental y/o SST: Razon:* "+DescrIncidenteTF2CP2W[-1]+ " no se generó paro."
                        print(mensajeIncidenteTF2CP2W)
                        if ValidarParoIncidenteEMCP2W[-1]=="Si":   
                            TiempoIncidenteTF2CP2W=TapaFija.col_values(120)
                            mensajeIncidenteTF2CP2W="*Incidente y/o accidente ambiental y/o SST - Tiempo:* "+TiempoIncidenteTF2CP2W[-1]+" min, *Razon:* "+DescrIncidenteTF2CP2W[-1]
                            print (mensajeIncidenteTF2CP2W)
                        else:
                            #DescrIncidenteTF2AGI=TapaFija.col_values(12)
                            mensajeIncidenteTF2CP2W="*Incidente y/o accidente ambiental y/o SST: Razon:* "+DescrIncidenteTF2CP2W[-1] + " no se generó paro."
                            print (mensajeIncidenteTF2CP2W)
                    else:
                        mensajeIncidenteTF2CP2W=""
                        print(mensajeIncidenteTF2CP2W)

                ##SERVICIOS PUBLICOS Copa 2.0 Whirlpool TAPA FIJA:::
                    ServiciosPublicosTF2CP2W=TapaFija.col_values(121)
                    if ServiciosPublicosTF2CP2W[-1]=="Si":
                        DescrServiciosPublicosTF2CP2W=TapaFija.col_values(123)
                        TiempoServiciosPublicosTF2CP2W=TapaFija.col_values(122)
                        mensajeServiciosPublicosTF2CP2W="*Hubo afectación en las unidades del hora a hora por falta de servicios públicos: Razon:* "+DescrServiciosPublicosTF2CP2W[-1]+ " -*Tiempo:* :"+TiempoServiciosPublicosTF2CP2W[-1]+"min"
                        print(mensajeServiciosPublicosTF2CP2W)
                    else:
                        mensajeServiciosPublicosTF2CP2W=""
                        print(mensajeServiciosPublicosTF2CP2W)

                #POR MAQUINA Copa 2.0 Haceb TAPA FIJA::::::::
                    MaquinaTF2CP2W=TapaFija.col_values(124)
                    if MaquinaTF2CP2W[-1]=="Si":
                        DescrMaquinaTF2CP2W=TapaFija.col_values(127)
                        TiempoMaquinaTF2CP2W=TapaFija.col_values(125)
                        mensajeMaquinaTF2CP2W="*Hubo afectación en las unidades por Maquina/ Equipo: Razon:* "+DescrMaquinaTF2CP2W[-1]+ " - *Tiempo:* "+TiempoMaquinaTF2CP2W[-1]+"min" 
                        print(mensajeMaquinaTF2CP2W)
                    else:
                        mensajeMaquinaTF2CP2W=""
                        print(mensajeMaquinaTF2CP2W)

                #POR MANO DE OBRA Copa 2.0 Whirlpool TAPA FIJA::::::::
                    ManoDeObraTF2CP2W=TapaFija.col_values(128)
                    if ManoDeObraTF2CP2W[-1]=="Si":
                        DescrManoDeObraTF2CP2W=TapaFija.col_values(132)
                        TiempoManoDeObraTF2CP2W=TapaFija.col_values(129)
                        mensajeManoDeObraTF2CP2W="*Hubo afectación en las unidades por Mano De Obra: Razon:* "+DescrManoDeObraTF2CP2W[-1]+ " - *Tiempo:* "+TiempoManoDeObraTF2CP2W[-1]+"min" 
                        print(mensajeManoDeObraTF2CP2W)
                    else:
                        mensajeManoDeObraTF2CP2W=""
                        print(mensajeManoDeObraTF2CP2W)

                #MATERIA PRIMA Copa 2.0 Whirlpool TAPA FIJA::::

                    MateriaPrimaTF2CP2W=TapaFija.col_values(133)
                    if MateriaPrimaTF2CP2W[-1]=="Si":
                        DescrMateriaPrimaTF2CP2W=TapaFija.col_values(137)
                        TiempoMateriaPrimaTF2CP2W=TapaFija.col_values(174)
                        mensajeMateriaPrimaTF2CP2W="*Hubo afectación en las unidades por Materia Prima: Razon:* "+DescrMateriaPrimaTF2CP2W[-1]+ " - *Tiempo:* "+TiempoMateriaPrimaTF2CP2W[-1]+"min" 
                        print(mensajeMateriaPrimaTF2CP2W)
                    else:
                        mensajeMateriaPrimaTF2CP2W=""
                        print(mensajeMateriaPrimaTF2CP2W)

                #POR METODO Copa 2.0 Whirlpool TAPA FIJA:::
                    MetodoTF2CP2W=TapaFija.col_values(138)
                    if MetodoTF2CP2W[-1]=="Si":
                        DescrMetodoTF2CP2W=TapaFija.col_values(141)
                        TiempoMetodoTF2CP2W=TapaFija.col_values(139)
                        mensajeMetodoTF2CP2W="*Hubo afectación en las unidades por Método: Razon:* "+DescrMetodoTF2CP2W[-1]+ "- *Tiempo:* "+TiempoMetodoTF2CP2W[-1]+"min" 
                        print(mensajeMetodoTF2CP2W)
                    else:
                        mensajeMetodoTF2CP2W=""
                        print(mensajeMetodoTF2CP2W)

                #SCRAP Copa 2.0 Haceb TAPA FIJA::::::::::
                    ScrapTF2CP2W=TapaFija.col_values(142)
                    if ScrapTF2CP2W[-1]=="Si":
                        DescrScrapTF2CP2W=TapaFija.col_values(144)
                        CantidadScrapTF2CP2W=TapaFija.col_values(145)
                        mensajeScrapTF2CP2W="*Se generó SCRAP: Cantidad:* "+CantidadScrapTF2CP2W[-1]+" - *Razon:* "+DescrScrapTF2CP2W[-1]
                        print(mensajeScrapTF2CP2W)
                    else:
                        mensajeScrapTF2CP2W=""
                        print(mensajeScrapTF2CP2W)

                #REPROCESADAS Copa 2.0 Haceb::::::::
                    UnidadesReprocesadasTF2CP2W =  TapaFija.col_values(146)
                    if UnidadesReprocesadasTF2CP2W[-1]=="Si":
                        CantidadReprocesadasTF2CP2W = TapaFija.col_values(147)

                        mensajeUnidadesReprocesadasTF2CP2W="*Se reprocesaron unidades - Cantidad:* "+CantidadReprocesadasTF2CP2W[-1]+""
                        print (mensajeUnidadesReprocesadasTF2CP2W)
                    else:
                        mensajeUnidadesReprocesadasTF2CP2W=""
                        print(mensajeUnidadesReprocesadasTF2CP2W)


                ##IMPELLER TAPA FIJA:::::::::::::::::::::::::::::::::::::::::::::::::
                # ___________________________________________________


                if SelectReferenciaTF2[-1]=="Impeller":
                    print("Impeller::::::")
                    #PAROS PROGRAMADOS::::
                    ParoProgramadoTF2IMPELLER = TapaFija.col_values(148)
                    if ParoProgramadoTF2IMPELLER[-1]=="Si":
                        RazonParoProgramadoTF2IMPELLER =  TapaFija.col_values(149)
                        TiempoParoProgramadoTF2IMPELLER =  TapaFija.col_values(150)
                        mensajeParoProgramadoTF2IMPELLER="*Paro programado - Tiempo:* "+TiempoParoProgramadoTF2IMPELLER[-1]+" min, *Razon:* "+RazonParoProgramadoTF2IMPELLER[-1]
                        print(mensajeParoProgramadoTF2IMPELLER)
                    else:
                        mensajeParoProgramadoTF2IMPELLER=""
                        print(mensajeParoProgramadoTF2IMPELLER)
                    
                    #INCIDENTES Impeller TAPA FIJA:::::
                    IncidenteTF2IMPELLER=TapaFija.col_values(151)
                    if IncidenteTF2IMPELLER[-1]=="Si":
                        DescrIncidenteTF2IMPELLER=TapaFija.col_values(153)
                        ValidarParoIncidenteEMCP2W=TapaFija.col_values(154)
                        mensajeIncidenteTF2IMPELLER="*Incidente y/o accidente ambiental y/o SST: Razon:* "+DescrIncidenteTF2IMPELLER[-1]+ " no se generó paro."
                        print(mensajeIncidenteTF2IMPELLER)
                        if ValidarParoIncidenteEMCP2W[-1]=="Si":   
                            TiempoIncidenteTF2IMPELLER=TapaFija.col_values(155)
                            mensajeIncidenteTF2IMPELLER="*Incidente y/o accidente ambiental y/o SST - Tiempo:* "+TiempoIncidenteTF2IMPELLER[-1]+" min, *Razon:* "+DescrIncidenteTF2IMPELLER[-1]
                            print (mensajeIncidenteTF2IMPELLER)
                        else:
                            #DescrIncidenteTF2AGI=TapaFija.col_values(12)
                            mensajeIncidenteTF2IMPELLER="*Incidente y/o accidente ambiental y/o SST: Razon:* "+DescrIncidenteTF2IMPELLER[-1] + " no se generó paro."
                            print (mensajeIncidenteTF2IMPELLER)
                    else:
                        mensajeIncidenteTF2IMPELLER=""
                        print(mensajeIncidenteTF2IMPELLER)

                ##SERVICIOS PUBLICOS Copa 2.0 Whirlpool TAPA FIJA:::
                    ServiciosPublicosTF2IMPELLER=TapaFija.col_values(156)
                    if ServiciosPublicosTF2IMPELLER[-1]=="Si":
                        DescrServiciosPublicosTF2IMPELLER=TapaFija.col_values(158)
                        TiempoServiciosPublicosTF2IMPELLER=TapaFija.col_values(157)
                        mensajeServiciosPublicosTF2IMPELLER="*Hubo afectación en las unidades del hora a hora por falta de servicios públicos: Razon:* "+DescrServiciosPublicosTF2IMPELLER[-1]+ " -*Tiempo:* :"+TiempoServiciosPublicosTF2IMPELLER[-1]+"min"
                        print(mensajeServiciosPublicosTF2IMPELLER)
                    else:
                        mensajeServiciosPublicosTF2IMPELLER=""
                        print(mensajeServiciosPublicosTF2IMPELLER)

                #POR MAQUINA Impeller TAPA FIJA::::::::
                    MaquinaTF2IMPELLER=TapaFija.col_values(159)
                    if MaquinaTF2IMPELLER[-1]=="Si":
                        DescrMaquinaTF2IMPELLER=TapaFija.col_values(162)
                        TiempoMaquinaTF2IMPELLER=TapaFija.col_values(160)
                        mensajeMaquinaTF2IMPELLER="*Hubo afectación en las unidades por Maquina/ Equipo: Razon:* "+DescrMaquinaTF2IMPELLER[-1]+ " - *Tiempo:* "+TiempoMaquinaTF2IMPELLER[-1]+"min" 
                        print(mensajeMaquinaTF2IMPELLER)
                    else:
                        mensajeMaquinaTF2IMPELLER=""
                        print(mensajeMaquinaTF2IMPELLER)

                #POR MANO DE OBRA Impeller TAPA FIJA::::::::
                    ManoDeObraTF2IMPELLER=TapaFija.col_values(163)
                    if ManoDeObraTF2IMPELLER[-1]=="Si":
                        DescrManoDeObraTF2IMPELLER=TapaFija.col_values(167)
                        TiempoManoDeObraTF2IMPELLER=TapaFija.col_values(164)
                        mensajeManoDeObraTF2IMPELLER="*Hubo afectación en las unidades por Mano De Obra: Razon:* "+DescrManoDeObraTF2IMPELLER[-1]+ " - *Tiempo:* "+TiempoManoDeObraTF2IMPELLER[-1]+"min" 
                        print(mensajeManoDeObraTF2IMPELLER)
                    else:
                        mensajeManoDeObraTF2IMPELLER=""
                        print(mensajeManoDeObraTF2IMPELLER)

                #MATERIA PRIMA Impeller TAPA FIJA::::

                    MateriaPrimaTF2IMPELLER=TapaFija.col_values(168)
                    if MateriaPrimaTF2IMPELLER[-1]=="Si":
                        DescrMateriaPrimaTF2IMPELLER=TapaFija.col_values(172)
                        TiempoMateriaPrimaTF2IMPELLER=TapaFija.col_values(169)
                        mensajeMateriaPrimaTF2IMPELLER="*Hubo afectación en las unidades por Materia Prima: Razon:* "+DescrMateriaPrimaTF2IMPELLER[-1]+ " - *Tiempo:* "+TiempoMateriaPrimaTF2IMPELLER[-1]+"min" 
                        print(mensajeMateriaPrimaTF2IMPELLER)
                    else:
                        mensajeMateriaPrimaTF2IMPELLER=""
                        print(mensajeMateriaPrimaTF2IMPELLER)

                #POR METODO Impeller TAPA FIJA:::
                    MetodoTF2IMPELLER=TapaFija.col_values(173)
                    if MetodoTF2IMPELLER[-1]=="Si":
                        DescrMetodoTF2IMPELLER=TapaFija.col_values(176)
                        TiempoMetodoTF2IMPELLER=TapaFija.col_values(174)
                        mensajeMetodoTF2IMPELLER="*Hubo afectación en las unidades por Método: Razon:* "+DescrMetodoTF2IMPELLER[-1]+ "- *Tiempo:* "+TiempoMetodoTF2IMPELLER[-1]+"min" 
                        print(mensajeMetodoTF2IMPELLER)
                    else:
                        mensajeMetodoTF2IMPELLER=""
                        print(mensajeMetodoTF2IMPELLER)

                #SCRAP Impeller TAPA FIJA::::::::::
                    ScrapTF2IMPELLER=TapaFija.col_values(177)
                    if ScrapTF2IMPELLER[-1]=="Si":
                        DescrScrapTF2IMPELLER=TapaFija.col_values(179)
                        CantidadScrapTF2IMPELLER=TapaFija.col_values(180)
                        mensajeScrapTF2IMPELLER="*Se generó SCRAP: Cantidad:* "+CantidadScrapTF2IMPELLER[-1]+" - *Razon:* "+DescrScrapTF2IMPELLER[-1]
                        print(mensajeScrapTF2IMPELLER)
                    else:
                        mensajeScrapTF2IMPELLER=""
                        print(mensajeScrapTF2IMPELLER)

                #REPROCESADAS Impeller Haceb::::::::
                    UnidadesReprocesadasTF2IMPELLER =  TapaFija.col_values(181)
                    if UnidadesReprocesadasTF2IMPELLER[-1]=="Si":
                        CantidadReprocesadasTF2IMPELLER = TapaFija.col_values(182)

                        mensajeUnidadesReprocesadasTF2IMPELLER="*Se reprocesaron unidades - Cantidad:* "+CantidadReprocesadasTF2IMPELLER[-1]+""
                        print (mensajeUnidadesReprocesadasTF2IMPELLER)
                    else:
                        mensajeUnidadesReprocesadasTF2IMPELLER=""
                        print(mensajeUnidadesReprocesadasTF2IMPELLER)


                ##QUASAR TAPA FIJA:::::::::::::::::::::::::::::::::::::::::::::::::
                # ___________________________________________________

                if SelectReferenciaTF2[-1]=="Quasar":
                    print("Quasar::::::::::")
                    #PAROS PROGRAMADOS:::::::::::
                    #________
                    ParoProgramadoTF2QUASAR = TapaFija.col_values(183)
                    if ParoProgramadoTF2QUASAR[-1]=="Si":
                        RazonParoProgramadoTF2QUASAR =  TapaFija.col_values(184)
                        TiempoParoProgramadoTF2QUASAR =  TapaFija.col_values(185)
                        mensajeParoProgramadoTF2QUASAR="*Paro programado - Tiempo:* "+TiempoParoProgramadoTF2QUASAR[-1]+" min, *Razon:* "+RazonParoProgramadoTF2QUASAR[-1]
                        print(mensajeParoProgramadoTF2QUASAR)
                    else:
                        mensajeParoProgramadoTF2QUASAR=""
                        print(mensajeParoProgramadoTF2QUASAR)
                    
                    #INCIDENTES Quasar TAPA FIJA:::::
                    IncidenteTF2QUASAR=TapaFija.col_values(186)
                    if IncidenteTF2QUASAR[-1]=="Si":
                        DescrIncidenteTF2QUASAR=TapaFija.col_values(188)
                        ValidarParoIncidenteEMCP2W=TapaFija.col_values(189)
                        mensajeIncidenteTF2QUASAR="*Incidente y/o accidente ambiental y/o SST: Razon:* "+DescrIncidenteTF2QUASAR[-1]+ " no se generó paro."
                        print(mensajeIncidenteTF2QUASAR)
                        if ValidarParoIncidenteEMCP2W[-1]=="Si":   
                            TiempoIncidenteTF2QUASAR=TapaFija.col_values(190)
                            mensajeIncidenteTF2QUASAR="*Incidente y/o accidente ambiental y/o SST - Tiempo:* "+TiempoIncidenteTF2QUASAR[-1]+" min, *Razon:* "+DescrIncidenteTF2QUASAR[-1]
                            print (mensajeIncidenteTF2QUASAR)
                        else:
                            #DescrIncidenteTF2AGI=TapaFija.col_values(12)
                            mensajeIncidenteTF2QUASAR="*Incidente y/o accidente ambiental y/o SST: Razon:* "+DescrIncidenteTF2QUASAR[-1] + " no se generó paro."
                            print (mensajeIncidenteTF2QUASAR)
                    else:
                        mensajeIncidenteTF2QUASAR=""
                        print(mensajeIncidenteTF2QUASAR)

                ##SERVICIOS PUBLICOS Quasar TAPA FIJA:::
                    ServiciosPublicosTF2QUASAR=TapaFija.col_values(191)
                    if ServiciosPublicosTF2QUASAR[-1]=="Si":
                        DescrServiciosPublicosTF2QUASAR=TapaFija.col_values(193)
                        TiempoServiciosPublicosTF2QUASAR=TapaFija.col_values(192)
                        mensajeServiciosPublicosTF2QUASAR="*Hubo afectación en las unidades del hora a hora por falta de servicios públicos: Razon:* "+DescrServiciosPublicosTF2QUASAR[-1]+ " -*Tiempo:* :"+TiempoServiciosPublicosTF2QUASAR[-1]+"min"
                        print(mensajeServiciosPublicosTF2QUASAR)
                    else:
                        mensajeServiciosPublicosTF2QUASAR=""
                        print(mensajeServiciosPublicosTF2QUASAR)

                #POR MAQUINA Quasar TAPA FIJA::::::::
                    MaquinaTF2QUASAR=TapaFija.col_values(194)
                    if MaquinaTF2QUASAR[-1]=="Si":
                        DescrMaquinaTF2QUASAR=TapaFija.col_values(197)
                        TiempoMaquinaTF2QUASAR=TapaFija.col_values(195)
                        mensajeMaquinaTF2QUASAR="*Hubo afectación en las unidades por Maquina/ Equipo: Razon:* "+DescrMaquinaTF2QUASAR[-1]+ " - *Tiempo:* "+TiempoMaquinaTF2QUASAR[-1]+"min" 
                        print(mensajeMaquinaTF2QUASAR)
                    else:
                        mensajeMaquinaTF2QUASAR=""
                        print(mensajeMaquinaTF2QUASAR)

                #POR MANO DE OBRA Quasar TAPA FIJA::::::::
                    ManoDeObraTF2QUASAR=TapaFija.col_values(198)
                    if ManoDeObraTF2QUASAR[-1]=="Si":
                        DescrManoDeObraTF2QUASAR=TapaFija.col_values(202)
                        TiempoManoDeObraTF2QUASAR=TapaFija.col_values(199)
                        mensajeManoDeObraTF2QUASAR="*Hubo afectación en las unidades por Mano De Obra: Razon:* "+DescrManoDeObraTF2QUASAR[-1]+ " - *Tiempo:* "+TiempoManoDeObraTF2QUASAR[-1]+"min" 
                        print(mensajeManoDeObraTF2QUASAR)
                    else:
                        mensajeManoDeObraTF2QUASAR=""
                        print(mensajeManoDeObraTF2QUASAR)

                #MATERIA PRIMA Quasar TAPA FIJA::::

                    MateriaPrimaTF2QUASAR=TapaFija.col_values(203)
                    if MateriaPrimaTF2QUASAR[-1]=="Si":
                        DescrMateriaPrimaTF2QUASAR=TapaFija.col_values(207)
                        TiempoMateriaPrimaTF2QUASAR=TapaFija.col_values(204)
                        mensajeMateriaPrimaTF2QUASAR="*Hubo afectación en las unidades por Materia Prima: Razon:* "+DescrMateriaPrimaTF2QUASAR[-1]+ " - *Tiempo:* "+TiempoMateriaPrimaTF2QUASAR[-1]+"min" 
                        print(mensajeMateriaPrimaTF2QUASAR)
                    else:
                        mensajeMateriaPrimaTF2QUASAR=""
                        print(mensajeMateriaPrimaTF2QUASAR)

                #POR METODO Quasar TAPA FIJA:::
                    MetodoTF2QUASAR=TapaFija.col_values(208)
                    if MetodoTF2QUASAR[-1]=="Si":
                        DescrMetodoTF2QUASAR=TapaFija.col_values(211)
                        TiempoMetodoTF2QUASAR=TapaFija.col_values(209)
                        mensajeMetodoTF2QUASAR="*Hubo afectación en las unidades por Método: Razon:* "+DescrMetodoTF2QUASAR[-1]+ "- *Tiempo:* "+TiempoMetodoTF2QUASAR[-1]+"min" 
                        print(mensajeMetodoTF2QUASAR)
                    else:
                        mensajeMetodoTF2QUASAR=""
                        print(mensajeMetodoTF2QUASAR)

                #SCRAP Quasar TAPA FIJA::::::::::
                    ScrapTF2QUASAR=TapaFija.col_values(212)
                    if ScrapTF2QUASAR[-1]=="Si":
                        DescrScrapTF2QUASAR=TapaFija.col_values(214)
                        CantidadScrapTF2QUASAR=TapaFija.col_values(215)
                        mensajeScrapTF2QUASAR="*Se generó SCRAP: Cantidad:* "+CantidadScrapTF2QUASAR[-1]+" - *Razon:* "+DescrScrapTF2QUASAR[-1]
                        print(mensajeScrapTF2QUASAR)
                    else:
                        mensajeScrapTF2QUASAR=""
                        print(mensajeScrapTF2QUASAR)

                #REPROCESADAS Quasar Haceb::::::::
                    UnidadesReprocesadasTF2QUASAR =  TapaFija.col_values(216)
                    if UnidadesReprocesadasTF2QUASAR[-1]=="Si":
                        CantidadReprocesadasTF2QUASAR = TapaFija.col_values(217)

                        mensajeUnidadesReprocesadasTF2QUASAR="*Se reprocesaron unidades - Cantidad:* "+CantidadReprocesadasTF2QUASAR[-1]+""
                        print (mensajeUnidadesReprocesadasTF2QUASAR)
                    else:
                        mensajeUnidadesReprocesadasTF2QUASAR=""
                        print(mensajeUnidadesReprocesadasTF2QUASAR)

                OeeTF2= TapaFija.col_values(223)
                OeeTapaFija = OeeTF2[-1]
                print("Porcentaje OEE: "+ OeeTapaFija)
                MensajeOeeTF2="OEE: "+OeeTapaFija

            if espera_Minuto:
                print("Esperando minuto para envio de wpp...")
                espera_Minuto= False

            if Minuto2==9:
            #Se envia el mensaje por WPP
                print("Entró")

                mensaje5="\n*TAPA FIJA*" 

                if mensajeUnidadesFabricadasTF2!="":
                    mensaje5=mensaje5+"\n\n"+mensajeUnidadesFabricadasTF2

                if SelectReferenciaTF2[-1]=="Agipeller":
                    mensaje5=mensaje5+"\n\n*AGIPELLER*"
                
                    if mensajeParoProgramadoTF2AGI!="":
                        mensaje5=mensaje5+"\n"+mensajeParoProgramadoTF2AGI
                        SinNovedad=2
                    if mensajeIncidenteTF2AGI!="":
                        mensaje5=mensaje5+"\n"+mensajeIncidenteTF2AGI
                        SinNovedad=2
                    if mensajeServiciosPublicosTF2AGI!="":
                        mensaje5=mensaje5+"\n"+mensajeServiciosPublicosTF2AGI
                        SinNovedad=2
                    if mensajeMaquinaTF2AGI!="":
                        mensaje5=mensaje5+"\n"+mensajeMaquinaTF2AGI
                        SinNovedad=2
                    if mensajeManoDeObraTF2AGI!="":
                        mensaje5=mensaje5+"\n"+mensajeManoDeObraTF2AGI
                        SinNovedad=2
                    if mensajeMateriaPrimaTF2AGI!="":
                        mensaje5=mensaje5+"\n"+mensajeMateriaPrimaTF2AGI
                        SinNovedad=2
                    if mensajeMetodoTF2AGI!="":
                        mensaje5=mensaje5+"\n"+mensajeMetodoTF2AGI
                        SinNovedad=2
                    if mensajeScrapTF2AGI!="":
                        mensaje5=mensaje5+"\n"+mensajeScrapTF2AGI
                        SinNovedad=2
                    if mensajeUnidadesReprocesadasTF2AGI!="":
                        mensaje5=mensaje5+"\n"+mensajeUnidadesReprocesadasTF2AGI
                        SinNovedad=2
                
                if SelectReferenciaTF2[-1]=="Back Panel":
                    mensaje5=mensaje5+"\n\n*BACK PANEL*"
                
                    if mensajeParoProgramadoTF2BP!="":
                        mensaje5=mensaje5+"\n"+mensajeParoProgramadoTF2BP
                        SinNovedad=2
                    if mensajeIncidenteTF2BP!="":
                        mensaje5=mensaje5+"\n"+mensajeIncidenteTF2BP
                        SinNovedad=2
                    if mensajeServiciosPublicosTF2BP!="":
                        mensaje5=mensaje5+"\n"+mensajeServiciosPublicosTF2BP
                        SinNovedad=2
                    if mensajeMaquinaTF2BP!="":
                        mensaje5=mensaje5+"\n"+mensajeMaquinaTF2BP
                        SinNovedad=2
                    if mensajeManoDeObraTF2BP!="":
                        mensaje5=mensaje5+"\n"+mensajeManoDeObraTF2BP
                        SinNovedad=2
                    if mensajeMateriaPrimaTF2BP!="":
                        mensaje5=mensaje5+"\n"+mensajeMateriaPrimaTF2BP
                        SinNovedad=2
                    if mensajeMetodoTF2BP!="":
                        mensaje5=mensaje5+"\n"+mensajeMetodoTF2BP
                        SinNovedad=2
                    if mensajeScrapTF2BP!="":
                        mensaje5=mensaje5+"\n"+mensajeScrapTF2BP
                        SinNovedad=2
                    if mensajeUnidadesReprocesadasTF2BP!="":
                        mensaje5=mensaje5+"\n"+mensajeUnidadesReprocesadasTF2BP
                        SinNovedad=2
                if SelectReferenciaTF2[-1]=="Copa 2.0 Haceb":
                    mensaje5=mensaje5+"\n\n*COPA 2.0 HACEB*"
                
                    if mensajeParoProgramadoTF2CP2H!="":
                        mensaje5=mensaje5+"\n"+mensajeParoProgramadoTF2CP2H
                        SinNovedad=2
                    if mensajeIncidenteTF2CP2H!="":
                        mensaje5=mensaje5+"\n"+mensajeIncidenteTF2CP2H
                        SinNovedad=2
                    if mensajeServiciosPublicosTF2CP2H!="":
                        mensaje5=mensaje5+"\n"+mensajeServiciosPublicosTF2CP2H
                        SinNovedad=2
                    if mensajeMaquinaTF2CP2H!="":
                        mensaje5=mensaje5+"\n"+mensajeMaquinaTF2CP2H
                        SinNovedad=2
                    if mensajeManoDeObraTF2CP2H!="":
                        mensaje5=mensaje5+"\n"+mensajeManoDeObraTF2CP2H
                        SinNovedad=2
                    if mensajeMateriaPrimaTF2CP2H!="":
                        mensaje5=mensaje5+"\n"+mensajeMateriaPrimaTF2CP2H
                        SinNovedad=2
                    if mensajeMetodoTF2CP2H!="":
                        mensaje5=mensaje5+"\n"+mensajeMetodoTF2CP2H
                        SinNovedad=2
                    if mensajeScrapTF2CP2H!="":
                        mensaje5=mensaje5+"\n"+mensajeScrapTF2CP2H
                        SinNovedad=2
                    if mensajeUnidadesReprocesadasTF2CP2H!="":
                        mensaje5=mensaje5+"\n"+mensajeUnidadesReprocesadasTF2CP2H
                        SinNovedad=2

                if SelectReferenciaTF2[-1]=="Copa 2.0 Whirlpool":
                    mensaje5=mensaje5+"\n\n*COPA 2.0 WHIRLPOOL*"

                    if mensajeParoProgramadoTF22CP2W!="":
                        mensaje5=mensaje5+"\n"+mensajeParoProgramadoTF22CP2W
                        SinNovedad=2
                    if mensajeIncidenteTF2CP2W!="":
                        mensaje5=mensaje5+"\n"+mensajeIncidenteTF2CP2W
                        SinNovedad=2
                    if mensajeServiciosPublicosTF2CP2W!="":
                        mensaje5=mensaje5+"\n"+mensajeServiciosPublicosTF2CP2W
                        SinNovedad=2
                    if mensajeMaquinaTF2CP2W!="":
                        mensaje5=mensaje5+"\n"+mensajeMaquinaTF2CP2W
                        SinNovedad=2
                    if mensajeManoDeObraTF2CP2W!="":
                        mensaje5=mensaje5+"\n"+mensajeManoDeObraTF2CP2W
                        SinNovedad=2
                    if mensajeMateriaPrimaTF2CP2W!="":
                        mensaje5=mensaje5+"\n"+mensajeMateriaPrimaTF2CP2W
                        SinNovedad=2
                    if mensajeMetodoTF2CP2W!="":
                        mensaje5=mensaje5+"\n"+mensajeMetodoTF2CP2W
                        SinNovedad=2
                    if mensajeScrapTF2CP2W!="":
                        mensaje5=mensaje5+"\n"+mensajeScrapTF2CP2W
                        SinNovedad=2
                    if mensajeUnidadesReprocesadasTF2CP2W!="":
                        mensaje5=mensaje5+"\n"+mensajeUnidadesReprocesadasTF2CP2W    
                        SinNovedad=2 

                if SelectReferenciaTF2[-1]=="Impeller":
                    mensaje5=mensaje5+"\n\n*IMPELLER*"
                
                    if mensajeParoProgramadoTF2IMPELLER!="":
                        mensaje5=mensaje5+"\n"+mensajeParoProgramadoTF2IMPELLER
                        SinNovedad=2
                    if mensajeIncidenteTF2IMPELLER!="":
                        mensaje5=mensaje5+"\n"+mensajeIncidenteTF2IMPELLER
                        SinNovedad=2
                    if mensajeServiciosPublicosTF2IMPELLER!="":
                        mensaje5=mensaje5+"\n"+mensajeServiciosPublicosTF2IMPELLER
                        SinNovedad=2
                    if mensajeMaquinaTF2IMPELLER!="":
                        mensaje5=mensaje5+"\n"+mensajeMaquinaTF2IMPELLER
                        SinNovedad=2
                    if mensajeManoDeObraTF2IMPELLER!="":
                        mensaje5=mensaje5+"\n"+mensajeManoDeObraTF2IMPELLER
                        SinNovedad=2
                    if mensajeMateriaPrimaTF2IMPELLER!="":
                        mensaje5=mensaje5+"\n"+mensajeMateriaPrimaTF2IMPELLER
                        SinNovedad=2
                    if mensajeMetodoTF2IMPELLER!="":
                        mensaje5=mensaje5+"\n"+mensajeMetodoTF2IMPELLER
                        SinNovedad=2
                    if mensajeScrapTF2IMPELLER!="":
                        mensaje5=mensaje5+"\n"+mensajeScrapTF2IMPELLER
                        SinNovedad=2
                    if mensajeUnidadesReprocesadasTF2IMPELLER!="":
                        mensaje5=mensaje5+"\n"+mensajeUnidadesReprocesadasTF2IMPELLER
                        SinNovedad=2

                if SelectReferenciaTF2[-1]=="Quasar":
                    mensaje5=mensaje5+"\n\n*QUASAR*"
                
                    if mensajeParoProgramadoTF2QUASAR!="":
                        mensaje5=mensaje5+"\n"+mensajeParoProgramadoTF2QUASAR
                        SinNovedad=2
                    if mensajeIncidenteTF2QUASAR!="":
                        mensaje5=mensaje5+"\n"+mensajeIncidenteTF2QUASAR
                        SinNovedad=2
                    if mensajeServiciosPublicosTF2QUASAR!="":
                        mensaje5=mensaje5+"\n"+mensajeServiciosPublicosTF2QUASAR
                        SinNovedad=2
                    if mensajeMaquinaTF2QUASAR!="":
                        mensaje5=mensaje5+"\n"+mensajeMaquinaTF2QUASAR
                        SinNovedad=2
                    if mensajeManoDeObraTF2QUASAR!="":
                        mensaje5=mensaje5+"\n"+mensajeManoDeObraTF2QUASAR
                        SinNovedad=2
                    if mensajeMateriaPrimaTF2QUASAR!="":
                        mensaje5=mensaje5+"\n"+mensajeMateriaPrimaTF2QUASAR
                        SinNovedad=2
                    if mensajeMetodoTF2QUASAR!="":
                        mensaje5=mensaje5+"\n"+mensajeMetodoTF2QUASAR
                        SinNovedad=2
                    if mensajeScrapTF2QUASAR!="":
                        mensaje5=mensaje5+"\n"+mensajeScrapTF2QUASAR
                        SinNovedad=2
                    if mensajeUnidadesReprocesadasTF2QUASAR!="":
                        mensaje5=mensaje5+"\n"+mensajeUnidadesReprocesadasTF2QUASAR     
                        SinNovedad=2
                
                if SinNovedad!=2:
                    mensaje5=mensaje5+"\n*No se reportaron novedades*"
                
                SinNovedad=0

                if MensajeOeeTF2!="" and OeeTapaFija!="#DIV/0!":
                    mensaje5=mensaje5+"\n"+MensajeOeeTF2
            
            mensajefinal=mensaje+"\n"+mensaje2+"\n"+mensaje3+"\n"+mensaje4+"\n"+mensaje5+"\n"+mensaje6+"\n\n"+acumulado3
            
            try:
                pywhatkit.sendwhatmsg_to_group(
                    "HPuk6hQ3c4p1SSAgyQCTbE", mensajefinal, Hora, Minuto, 30, True, 20)
                print("Mensaje enviado")
                print(mensaje)

                espera_Minuto = True
                espera_Hora = True
                print(mensajefinal)
            except:
                print("Error!! El mensaje no pudo ser enviado")

        time.sleep(1)

    else:
        if espera_Hora:
            print ("Esperando la hora de envío...")
            espera_Hora = False