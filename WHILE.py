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
            print("1 Esperando minuto de envio...")
            espera_Minuto= False

            #CONDICIONAL PARA SELECCIONAR EL MIN Y EL RANGO DE SEGUNDOS
        if (Minuto2==56 and Segundo>=40) and Segundo<=45:
            ##CADENCIA:::::
            sh = gc3.open("Gestión Célula Hora a Hora Células")
            cadencia = gc.open("Cadencia")
            cadencia = cadencia.get_worksheet(0)
            cadenciaList = cadencia.col_values(3)
            print("Cadencia de las celulas: "+cadenciaList[-1])
            MensajeCadencia = "Cadencia de las celulas: "+ cadenciaList[-1]

            #EMSAMBLE MECANISMOS:::::::::::::::
            print("EMSAMBLE MECANISMOS:--------")
            EmsambleMecanismos = sh.get_worksheet(0)
            UnidadesFabricadasEM =  EmsambleMecanismos.col_values(5)
            print("Unidades producidas: "+UnidadesFabricadasEM[-1])
            MensajeUnidadesFabricadasEM ="Unidades producidas: "+ UnidadesFabricadasEM[-1]

            EmsambleMecanismos = sh.get_worksheet(0)
            ParoProgramadoEM =  EmsambleMecanismos.col_values(9)
            if ParoProgramadoEM[-1]=="Si":
                RazonParoProgramadoEM =  EmsambleMecanismos.col_values(10)
                TiempoParoProgramadoEM =  EmsambleMecanismos.col_values(11)
                mensajeParoProgramadoEM="*Paro programado - Tiempo:* "+TiempoParoProgramadoEM[-1]+" min, *Razón:* "+RazonParoProgramadoEM[-1]
                print (mensajeParoProgramadoEM)
            else:
                mensajeParoProgramadoEM=""
                print(mensajeParoProgramadoEM)

            IncidentesEM =  EmsambleMecanismos.col_values(12)
            if IncidentesEM[-1]=="Si":
                DescrIncidenteEM =  EmsambleMecanismos.col_values(14)
                ValidarParoIncidentesEM =  EmsambleMecanismos.col_values(15)
                if ValidarParoIncidentesEM[-1]=="Si":    
                    TiempoIncidenteEM =  EmsambleMecanismos.col_values(16)
                    mensajeIncidenteEM="*Incidente y/o accidente ambiental y/o SST - Tiempo:* "+TiempoIncidenteEM[-1]+" min, *Razón:* "+DescrIncidenteEM[-1]
                else:
                    mensajeIncidenteEM="*Incidente y/o accidente ambiental y/o SST - Razón:* "+DescrIncidenteEM[-1]
                print (mensajeIncidenteEM)
            else:
                mensajeIncidenteEM=""
                print(mensajeIncidenteEM)

            ServiciosPublicosEM =  EmsambleMecanismos.col_values(17)
            if ServiciosPublicosEM[-1]=="Si":
                DescrServiciosPublicosEM =  EmsambleMecanismos.col_values(19)
                TiempoServiciosPublicosEM =  EmsambleMecanismos.col_values(18)
                mensajeServiciosPublicosEM="*Afectación por falta de servicios públicos - Tiempo:* "+TiempoServiciosPublicosEM[-1]+" min, *Razón:* "+DescrServiciosPublicosEM[-1]+""
                print (mensajeServiciosPublicosEM)
            else:
                mensajeServiciosPublicosEM=""
                print(mensajeServiciosPublicosEM)

            MaquinaEM =  EmsambleMecanismos.col_values(20)
            if MaquinaEM[-1]=="Si":
                DescrMaquinaEM=  EmsambleMecanismos.col_values(23)
                TiempoMaquinaEM =  EmsambleMecanismos.col_values(21)
                mensajeMaquinaEM="*Afectación por maquina - Tiempo:* "+TiempoMaquinaEM[-1]+" min, *Razón:* "+DescrMaquinaEM[-1]+""
                print (mensajeMaquinaEM)
            else:
                mensajeMaquinaEM=""
                print(mensajeMaquinaEM)

            ManoObraEM =  EmsambleMecanismos.col_values(24)
            if ManoObraEM[-1]=="Si":
                DescrManoObraEM=  EmsambleMecanismos.col_values(28)
                TiempoManoObraEM =  EmsambleMecanismos.col_values(25)
                mensajeManoObraEM="*Hubo afectación en las unidades por Mano de Obra - Tiempo:* "+TiempoManoObraEM[-1]+" min, *Razón:* "+DescrManoObraEM[-1]+""
                print (mensajeManoObraEM)
            else:
                mensajeManoObraEM=""
                print(mensajeManoObraEM)


            MateriaPrimaEM =  EmsambleMecanismos.col_values(29)
            if MateriaPrimaEM[-1]=="Si":
                DescrMateriaPrimaEM=  EmsambleMecanismos.col_values(33)
                TiempoMateriaPrimaEM =  EmsambleMecanismos.col_values(30)
                mensajeMateriaPrimaEM="*Hubo afectación en las unidades por Materia Prima - Tiempo:* "+TiempoMateriaPrimaEM[-1]+" min, *Razón:* "+DescrMateriaPrimaEM[-1]+""
                print (mensajeMateriaPrimaEM)
            else:
                mensajeMateriaPrimaEM=""
                print(mensajeMateriaPrimaEM)

            UnidadesPorMetodoEM =  EmsambleMecanismos.col_values(34)
            if UnidadesPorMetodoEM[-1]=="Si":
                DescrUnidadesPorMetodoEM=  EmsambleMecanismos.col_values(37)
                TiempoUnidadesPorMetodoEM =  EmsambleMecanismos.col_values(35)
                mensajeUnidadesPorMetodoEM="*Hubo afectación en las unidades por Método - Tiempo:* "+TiempoUnidadesPorMetodoEM[-1]+" min, *Razón:* "+DescrUnidadesPorMetodoEM[-1]+""
                print (mensajeUnidadesPorMetodoEM)
            else:
                mensajeUnidadesPorMetodoEM=""
                print(mensajeUnidadesPorMetodoEM)

            ScrapEM =  EmsambleMecanismos.col_values(38)
            if ScrapEM[-1]=="Si":
                DescrScrapEM=  EmsambleMecanismos.col_values(39)
                CantidadScrapEM =  EmsambleMecanismos.col_values(41)
                mensajeScrapEM="*Se genero SCRAP - Cantidad:* "+CantidadScrapEM[-1]+", *Razón:* "+DescrScrapEM[-1]+""
                print (mensajeScrapEM)
            else:
                mensajeScrapEM=""
                print(mensajeScrapEM)


            UnidadesReprocesadasEM =  EmsambleMecanismos.col_values(42)
            if UnidadesReprocesadasEM[-1]=="Si":
                CantidadReprocesadasEM=  EmsambleMecanismos.col_values(43)

                mensajeUnidadesReprocesadasEM="*Se reprocesaron unidades - Cantidad:* "+CantidadReprocesadasEM[-1]+""
                print (mensajeUnidadesReprocesadasEM)
            else:
                mensajeUnidadesReprocesadasEM=""
                print(mensajeUnidadesReprocesadasEM)
                
            OeeEM= EmsambleMecanismos.col_values(48)
            OeeEmsambleMecanismos = OeeEM[-1]
            mensajeOeeEM = "OEE: "+ OeeEM[-1]
            print("OEE: "+OeeEmsambleMecanismos)
            
            mensaje="*GESTION CELULA HORA A HORA:*" 
            
            if MensajeCadencia!="":
                mensaje = mensaje + "\n\n" +MensajeCadencia
            mensaje=mensaje+"\n\n*EMSAMBLE MECANISMOS*"
            if MensajeUnidadesFabricadasEM!="":
                mensaje=mensaje+"\n\n"+MensajeUnidadesFabricadasEM
            if mensajeParoProgramadoEM!="":
                mensaje=mensaje+"\n"+mensajeParoProgramadoEM
            if mensajeIncidenteEM!="":
                mensaje=mensaje+"\n"+mensajeIncidenteEM
            if mensajeServiciosPublicosEM!="":
                mensaje=mensaje+"\n"+mensajeServiciosPublicosEM
            if mensajeMaquinaEM!="":
                mensaje=mensaje+"\n"+mensajeMaquinaEM
            if mensajeManoObraEM!="":
                mensaje=mensaje+"\n"+mensajeManoObraEM
            if mensajeMateriaPrimaEM!="":
                mensaje=mensaje+"\n"+mensajeMateriaPrimaEM
            if mensajeUnidadesPorMetodoEM!="":
                mensaje=mensaje+"\n"+mensajeUnidadesPorMetodoEM
            if mensajeUnidadesPorMetodoEM!="":
                mensaje=mensaje+"\n"+mensajeUnidadesPorMetodoEM
            if mensajeScrapEM!="":
                mensaje=mensaje+"\n"+mensajeScrapEM
            if mensajeUnidadesReprocesadasEM!="":
                mensaje=mensaje+"\n"+mensajeUnidadesReprocesadasEM
            if mensajeOeeEM!="":
                mensaje=mensaje+"\n"+mensajeOeeEM
            print (mensaje)
    
#-----------------------------------------------------------------------------------------------------------
        if espera_Minuto:
            print("2 Esperando minuto de envio...")
            espera_Minuto= False

            #CONDICIONAL PARA SELECCIONAR EL MIN Y EL RANGO DE SEGUNDOS
        if (Minuto2==57 and Segundo>=40) and Segundo<=45:
            #CONJUNTO SUSPENCIÓN:::::::::::::::
            print("CONJUNTO SUSPENCIÓN:---------")

            ConjuntoSuspencion = sh.get_worksheet(1)

            UnidadesFabricadasCJ=  ConjuntoSuspencion.col_values(5)
            print("Unidades Fabricadas: "+UnidadesFabricadasCJ[-1])
            MensajeUnidadesFabricadasCJ="Unidades Producidas: "+ UnidadesFabricadasCJ[-1]
            #PAROS PROGRAMADOS::
            ParoProgramadoCS =  ConjuntoSuspencion.col_values(9)
            if ParoProgramadoCS[-1]=="Si":
                RazonParoProgramadoCS =  ConjuntoSuspencion.col_values(10)
                TiempoParoProgramadoCS =  ConjuntoSuspencion.col_values(11)
                mensajeParoProgramadoCS="*Paro programado - Tiempo:* "+TiempoParoProgramadoCS[-1]+" min, *Razón:* "+RazonParoProgramadoCS[-1]
                print (mensajeParoProgramadoCS)
            else:
                mensajeParoProgramadoCS=""
                print(mensajeParoProgramadoCS)

            #INCIDENTES SST:::
            IncidentesCS =  ConjuntoSuspencion.col_values(12)
            if IncidentesCS[-1]=="Si":
                DescrIncidenteCS =  ConjuntoSuspencion.col_values(14)
                ValidarParoIncidentesCS =  ConjuntoSuspencion.col_values(15)
                if ValidarParoIncidentesCS[-1]=="Si":    
                    TiempoIncidenteCS =  ConjuntoSuspencion.col_values(16)
                    mensajeIncidenteCS="*Incidente y/o accidente ambiental y/o SST - Tiempo:* "+TiempoIncidenteCS[-1]+" min, *Razón:* "+DescrIncidenteCS[-1]
                else:
                    mensajeIncidenteCS="*Incidente y/o accidente ambiental y/o SST - Razón:* "+DescrIncidenteCS[-1] + ", No se generó paro."
                print (mensajeIncidenteCS)
            else:
                mensajeIncidenteCS=""
                print(mensajeIncidenteCS)

            #SERVICIOS PUBLICOS::::
            ServiciosPublicosCS =  ConjuntoSuspencion.col_values(17)
            if ServiciosPublicosCS[-1]=="Si":
                DescrServiciosPublicosCS =  ConjuntoSuspencion.col_values(19)
                TiempoServiciosPublicosCS =  ConjuntoSuspencion.col_values(18)
                mensajeServiciosPublicosCS="*Afectación por falta de servicios públicos - Tiempo:* "+TiempoServiciosPublicosCS[-1]+" min, *Razón:* "+DescrServiciosPublicosCS[-1]+""
                print (mensajeServiciosPublicosCS)
            else:
                mensajeServiciosPublicosCS=""
                print(mensajeServiciosPublicosCS)

            #MAQUINA:::::::::::
            MaquinaCS =  ConjuntoSuspencion.col_values(20)
            if MaquinaCS[-1]=="Si":
                DescrMaquinaCS=  ConjuntoSuspencion.col_values(23)
                TiempoMaquinaCS =  ConjuntoSuspencion.col_values(21)
                mensajeMaquinaCS="*Afectación por maquina - Tiempo:* "+TiempoMaquinaCS[-1]+" min, *Razón:* "+DescrMaquinaCS[-1]+""
                print (mensajeMaquinaCS)
            else:
                mensajeMaquinaCS=""
                print(mensajeMaquinaCS)

            #MANO DE OBRA:::::::
            ManoObraCS =  ConjuntoSuspencion.col_values(24)
            if ManoObraCS[-1]=="Si":
                DescrManoObraCS=  ConjuntoSuspencion.col_values(27)
                TiempoManoObraCS =  ConjuntoSuspencion.col_values(25)
                mensajeManoObraCS="*Hubo afectación en las unidades por Mano de Obra - Tiempo:* "+TiempoManoObraCS[-1]+" min, *Razón:* "+DescrManoObraCS[-1]+""
                print (mensajeManoObraCS)
            else:
                mensajeManoObraCS=""
                print(mensajeManoObraCS)

            #MATERIA PRIMA:::::::::::
            MateriaPrimaCS =  ConjuntoSuspencion.col_values(28)
            if MateriaPrimaCS[-1]=="Si":
                DescrMateriaPrimaCS=  ConjuntoSuspencion.col_values(32)
                TiempoMateriaPrimaCS =  ConjuntoSuspencion.col_values(29)
                mensajeMateriaPrimaCS="*Hubo afectación en las unidades por Materia Prima - Tiempo:* "+TiempoMateriaPrimaCS[-1]+" min, *Razón:* "+DescrMateriaPrimaCS[-1]+""
                print (mensajeMateriaPrimaCS)
            else:
                mensajeMateriaPrimaCS=""
                print(mensajeMateriaPrimaCS)


            #POR METODO:::::::::::::
            UnidadesPorMetodoCS =  ConjuntoSuspencion.col_values(33)
            if UnidadesPorMetodoCS[-1]=="Si":
                DescrUnidadesPorMetodoCS=  ConjuntoSuspencion.col_values(36)
                TiempoUnidadesPorMetodoCS =  ConjuntoSuspencion.col_values(34)
                mensajeUnidadesPorMetodoCS="*Hubo afectación en las unidades por Método - Tiempo:* "+TiempoUnidadesPorMetodoCS[-1]+" min, *Razón:* "+DescrUnidadesPorMetodoCS[-1]+""
                print (mensajeUnidadesPorMetodoCS)
            else:
                mensajeUnidadesPorMetodoCS=""
                print(mensajeUnidadesPorMetodoCS)


            ##SCRAPPPP::::::::::::::::::::::::::::::::
            ScrapCS =  ConjuntoSuspencion.col_values(37)
            if ScrapCS[-1]=="Si":
                DescrScrapCS=  ConjuntoSuspencion.col_values(39)
                CantidadScrapCS =  ConjuntoSuspencion.col_values(40)
                mensajeScrapCS="*Se genero SCRAP - Cantidad:* "+CantidadScrapCS[-1]+", *Razón:* "+DescrScrapCS[-1]+""
                print (mensajeScrapCS)
            else:
                mensajeScrapCS=""
                print(mensajeScrapCS)

            ##UNIDADES REPROCESADAS:::::::::::::
            UnidadesReprocesadasCS =  ConjuntoSuspencion.col_values(41)
            if UnidadesReprocesadasCS[-1]=="Si":
                CantidadReprocesadasCS=  ConjuntoSuspencion.col_values(42)

                mensajeUnidadesReprocesadasCS="*Se reprocesaron unidades - Cantidad:* "+CantidadReprocesadasCS[-1]+""
                print (mensajeUnidadesReprocesadasCS)
            else:
                mensajeUnidadesReprocesadasCS=""
                print(mensajeUnidadesReprocesadasCS)

            ##OEE::::::::::::::::::::::
            OeeCS= ConjuntoSuspencion.col_values(47)
            OeeConjuntoSuspencion = OeeCS[-1]
            print("OEE: "+OeeConjuntoSuspencion)
            MensajeOeeCS="OEE: "+OeeCS[-1]

        if espera_Minuto:
            print("Esperando minuto para envio de wpp...")
            espera_Minuto= False

        if Minuto2==58:
        #Se envia el mensaje por WPP
            print("Entró")
            mensaje2="\n\n*CONJUNTO SUSPENCION*" 
            if MensajeUnidadesFabricadasCJ!="":
                mensaje2=mensaje2+"\n\n"+MensajeUnidadesFabricadasCJ
            if mensajeParoProgramadoCS!="":
                mensaje2=mensaje2+"\n"+mensajeParoProgramadoCS
            if mensajeIncidenteCS!="":
                mensaje2=mensaje2+"\n"+mensajeIncidenteCS
            if mensajeServiciosPublicosCS!="":
                mensaje2=mensaje2+"\n"+mensajeServiciosPublicosCS
            if mensajeMaquinaCS!="":
                mensaje2=mensaje2+"\n"+mensajeMaquinaCS
            if mensajeManoObraCS!="":
                mensaje2=mensaje2+"\n"+mensajeManoObraCS
            if mensajeMateriaPrimaCS!="":
                mensaje2=mensaje2+"\n"+mensajeMateriaPrimaCS
            if mensajeUnidadesPorMetodoCS!="":
                mensaje2=mensaje2+"\n"+mensajeUnidadesPorMetodoCS
            if mensajeUnidadesPorMetodoCS!="":
                mensaje2=mensaje2+"\n"+mensajeUnidadesPorMetodoCS
            if mensajeScrapCS!="":
                mensaje2=mensaje2+"\n"+mensajeScrapCS
            if mensajeUnidadesReprocesadasCS!="":
                mensaje2=mensaje2+"\n"+mensajeUnidadesReprocesadasCS
            if MensajeOeeCS!="":
                mensaje2=mensaje2+"\n"+MensajeOeeCS

            mensajefinal=mensaje+"\n"+mensaje2

            try:
                pywhatkit.sendwhatmsg_to_group(
                    "HPuk6hQ3c4p1SSAgyQCTbE", mensajefinal, Hora, Minuto, 40, True, 20)
                print("Mensaje enviado")
                print(mensaje)

                espera_Minuto = True
                espera_Hora = True
            except:
                print("Error!! El mensaje no pudo ser enviado")

        time.sleep(1)

    else:
        if espera_Hora:
            print ("Esperando la hora de envío...")
            espera_Hora = False     