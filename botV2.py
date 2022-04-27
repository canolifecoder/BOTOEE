#!/usr/bin/python

from tkinter.tix import Select
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
    filename='named-haven-340115-ac55768dd57c.json'
)

sh = gc3.open("Gestión Célula Hora a Hora Células")

#___________________________________________________________________
#___________________________________________________________________



#EMSAMBLE GABINETE:::::::::::::::
print("TESTEO FINAL::---------")

#SELECCION DE LA HOJA::
TesteoFinal = sh.get_worksheet(3)
#SELECCIONAR LA REFERENCIA::::: COPA 1 -- COPA 2 HACEB -- COPA 2 WHIRLPOOL
UnidadesFabricadasEG=  TesteoFinal.col_values(5)
print(UnidadesFabricadasEG[-1])
#SELECCION DE COPA 1:::::
SelectReferenciaEG = TesteoFinal.col_values(7)
if SelectReferenciaEG[-1]=="Copa 1.0":
    print("COPA 1::::")
    #PAROS PROGRAMADOS::
    ParoProgramadoEGCP1 = TesteoFinal.col_values(8)
    if ParoProgramadoEGCP1[-1]=="Si":
        RazonParoProgramadoEGCP1 =  TesteoFinal.col_values(9)
        TiempoParoProgramadoEGCP1 =  TesteoFinal.col_values(10)
        mensajeParoProgramadoEGCP1="*Paro programado - Tiempo:* "+TiempoParoProgramadoEGCP1[-1]+" min, *Razón:* "+RazonParoProgramadoEGCP1[-1]
        print(mensajeParoProgramadoEGCP1)
    else:
        mensajeParoProgramadoEGCP1=""
        print(mensajeParoProgramadoEGCP1)
    
    #INCIDENTES::
    IncidenteEGCP1=TesteoFinal.col_values(11)
    if IncidenteEGCP1[-1]=="Si":
        DescrIncidenteEGCP1=TesteoFinal.col_values(13)
        ValidarParoIncidenteEGCP1=TesteoFinal.col_values(14)
        mensajeIncidenteEGCP1="*Incidente y/o accidente ambiental y/o SST: Razón:* "+DescrIncidenteEGCP1[-1] + " No se generó paro"
        print(mensajeIncidenteEGCP1)
        if ValidarParoIncidenteEGCP1[-1]=="Si":   
            TiempoIncidenteEGCP1=TesteoFinal.col_values(15)
            mensajeIncidenteEGCP1="*Incidente y/o accidente ambiental y/o SST - Tiempo:* "+TiempoIncidenteEGCP1[-1]+" min, *Razón:* "+DescrIncidenteEGCP1[-1]
            print (mensajeIncidenteEGCP1)
        else:
            #DescrIncidenteEGCP1=TesteoFinal.col_values(12)
            mensajeIncidenteEGCP1="*Incidente y/o accidente ambiental y/o SST: Razón:* "+DescrIncidenteEGCP1[-1] + " No se generó paro"
            print (mensajeIncidenteEGCP1)
    else:
        mensajeIncidenteEGCP1=""
        print(mensajeIncidenteEGCP1)

##SERVICIOS PUBLICOS COPA1::
    ServiciosPublicosEGCP1=TesteoFinal.col_values(16)
    if ServiciosPublicosEGCP1[-1]=="Si":
        DescrServiciosPublicosEGCP1=TesteoFinal.col_values(18)
        TiempoServiciosPublucosEGCP1=TesteoFinal.col_values(17)
        mensajeServiciosPublicosEGCP1="*Hubo afectación en las unidades del hora a hora por falta de servicios públicos: Razón:* "+DescrServiciosPublicosEGCP1[-1]+ " -*Tiempo:* :"+TiempoServiciosPublucosEGCP1[-1]+"min"
        print(mensajeServiciosPublicosEGCP1)
    else:
        mensajeServiciosPublicosEGCP1=""
        print(mensajeServiciosPublicosEGCP1)
#POR MAQUINA COPA1:::
    MaquinaEGCP1=TesteoFinal.col_values(19)
    if MaquinaEGCP1[-1]=="Si":
        DescrMaquinaEGCP1=TesteoFinal.col_values(22)
        TiempoMaquinaEGCP1=TesteoFinal.col_values(20)
        mensajeMaquinaEGCP1="*Hubo afectación en las unidades por Maquina/ Equipo: Razón:* "+DescrMaquinaEGCP1[-1]+ " - *Tiempo:* "+TiempoMaquinaEGCP1[-1]+"min" 
        print(mensajeMaquinaEGCP1)
    else:
        mensajeMaquinaEGCP1=""
        print(mensajeMaquinaEGCP1)

#POR MANO DE OBRA COPA1::::::::
    ManoDeObraEGCP1=TesteoFinal.col_values(23)
    if ManoDeObraEGCP1[-1]=="Si":
        DescrManoDeObraEGCP1=TesteoFinal.col_values(27)
        TiempoManoDeObraEGCP1=TesteoFinal.col_values(24)
        mensajeManoDeObraEGCP1="*Hubo afectación en las unidades por Mano De Obra: Razón:* "+DescrManoDeObraEGCP1[-1]+ " - *Tiempo:* "+TiempoManoDeObraEGCP1[-1]+"min" 
        print(mensajeManoDeObraEGCP1)
    else:
        mensajeManoDeObraEGCP1=""
        print(mensajeManoDeObraEGCP1)

#MATERIA PRIMA COPA1::::

    MateriaPrimaEGCP1=TesteoFinal.col_values(28)
    if MateriaPrimaEGCP1[-1]=="Si":
        DescrMateriaPrimaEGCP1=TesteoFinal.col_values(32)
        TiempoMateriaPrimaEGCP1=TesteoFinal.col_values(29)
        mensajeMateriaPrimaEGCP1="*Hubo afectación en las unidades por Materia Prima: Razón:* "+DescrMateriaPrimaEGCP1[-1]+ " - *Tiempo:* "+TiempoMateriaPrimaEGCP1[-1]+"min" 
        print(mensajeMateriaPrimaEGCP1)
    else:
        mensajeMateriaPrimaEGCP1=""
        print(mensajeMateriaPrimaEGCP1)

#POR METODO COPA1:::
    MetodoEGCP1=TesteoFinal.col_values(33)
    if MetodoEGCP1[-1]=="Si":
        DescrMetodoEGCP1=TesteoFinal.col_values(36)
        TiempoMetodoEGCP1=TesteoFinal.col_values(34)
        mensajeMetodoEGCP1="*Hubo afectación en las unidades por Método: Razón:* "+DescrMetodoEGCP1[-1]+ "- *Tiempo:* "+TiempoMetodoEGCP1[-1]+"min" 
        print(mensajeMetodoEGCP1)
    else:
        mensajeMetodoEGCP1=""
        print(mensajeMetodoEGCP1)

#SCRAP COPA1::::::::::
    ScrapEGCP1=TesteoFinal.col_values(37)
    if ScrapEGCP1[-1]=="Si":
        DescrScrapEGCP1=TesteoFinal.col_values(39)
        CantidadScrapEGCP1=TesteoFinal.col_values(40)
        mensajeScrapEGCP1="*Se generó SCRAP: Cantidad:* "+CantidadScrapEGCP1[-1]+" - *Razón:* "+DescrScrapEGCP1[-1]
        print(mensajeScrapEGCP1)
    else:
        mensajeScrapEGCP1=""
        print(mensajeScrapEGCP1)

#REPROCESADAS COPA1::::::::
    UnidadesReprocesadasEGP1 =  TesteoFinal.col_values(41)
    if UnidadesReprocesadasEGP1[-1]=="Si":
        CantidadReprocesadasEGP1=  TesteoFinal.col_values(42)

        mensajeUnidadesReprocesadasEGP1="*Se reprocesaron unidades - Cantidad:* "+CantidadReprocesadasEGP1[-1]+""
        print (mensajeUnidadesReprocesadasEGP1)
    else:
        mensajeUnidadesReprocesadasEGP1=""
        print(mensajeUnidadesReprocesadasEGP1)

##SELECCIONA COPA 2 HACEB::::::::::::::::::::::

if SelectReferenciaEG[-1]=="Copa 2.0 Haceb":
    print("COPA 2 HACEB::::")
    #PAROS PROGRAMADOS::
    ParoProgramadoEGCP2H = TesteoFinal.col_values(43)
    if ParoProgramadoEGCP2H[-1]=="Si":
        RazonParoProgramadoEGCP2H =  TesteoFinal.col_values(44)
        TiempoParoProgramadoEGCP2H =  TesteoFinal.col_values(45)
        mensajeParoProgramadoEGCP2H="*Paro programado - Tiempo:* "+TiempoParoProgramadoEGCP2H[-1]+" min, *Razón:* "+RazonParoProgramadoEGCP2H[-1]
        print(mensajeParoProgramadoEGCP2H)
    else:
        mensajeParoProgramadoEGCP2H=""
        print(mensajeParoProgramadoEGCP2H)
    
    #INCIDENTES::
    IncidenteEGCP2H=TesteoFinal.col_values(46)
    if IncidenteEGCP2H[-1]=="Si":
        DescrIncidenteEGCP2H=TesteoFinal.col_values(48)
        ValidarParoIncidenteEMCP2H=TesteoFinal.col_values(49)
        mensajeIncidenteEGCP2H="*Incidente y/o accidente ambiental y/o SST: Razón:* "+DescrIncidenteEGCP2H[-1]+ " no se generó paro."
        print(mensajeIncidenteEGCP2H)
        if ValidarParoIncidenteEMCP2H[-1]=="Si":   
            TiempoIncidenteEGCP2H=TesteoFinal.col_values(50)
            mensajeIncidenteEGCP2H="*Incidente y/o accidente ambiental y/o SST - Tiempo:* "+TiempoIncidenteEGCP2H[-1]+" min, *Razón:* "+DescrIncidenteEGCP2H[-1]
            print (mensajeIncidenteEGCP2H)
        else:
            #DescrIncidenteEGCP1=TesteoFinal.col_values(12)
            mensajeIncidenteEGCP2H="*Incidente y/o accidente ambiental y/o SST: Razón:* "+DescrIncidenteEGCP2H[-1] + " no se generó paro."
            print (mensajeIncidenteEGCP2H)
    else:
        mensajeIncidenteEGCP2H=""
        print(mensajeIncidenteEGCP2H)

##SERVICIOS PUBLICOS COPA2::
    ServiciosPublicosEGCP2H=TesteoFinal.col_values(51)
    if ServiciosPublicosEGCP2H[-1]=="Si":
        DescrServiciosPublicosEGCP2H=TesteoFinal.col_values(53)
        TiempoServiciosPublicosEGCP2H=TesteoFinal.col_values(52)
        mensajeServiciosPublicosEGCP2H="*Hubo afectación en las unidades del hora a hora por falta de servicios públicos: Razón:* "+DescrServiciosPublicosEGCP2H[-1]+ " -*Tiempo:* :"+TiempoServiciosPublicosEGCP2H[-1]+"min"
        print(mensajeServiciosPublicosEGCP2H)
    else:
        mensajeServiciosPublicosEGCP2H=""
        print(mensajeServiciosPublicosEGCP2H)
#POR MAQUINA COPA2:::
    MaquinaEGCP2H=TesteoFinal.col_values(54)
    if MaquinaEGCP2H[-1]=="Si":
        DescrMaquinaEGCP2H=TesteoFinal.col_values(57)
        TiempoMaquinaEGCP2H=TesteoFinal.col_values(55)
        mensajeMaquinaEGCP2H="*Hubo afectación en las unidades por Maquina/ Equipo: Razón:* "+DescrMaquinaEGCP2H[-1]+ " - *Tiempo:* "+TiempoMaquinaEGCP2H[-1]+"min" 
        print(mensajeMaquinaEGCP2H)
    else:
        mensajeMaquinaEGCP2H=""
        print(mensajeMaquinaEGCP2H)

#POR MANO DE OBRA COPA2::::::::
    ManoDeObraEGCP2H=TesteoFinal.col_values(58)
    if ManoDeObraEGCP2H[-1]=="Si":
        DescrManoDeObraEGCP2H=TesteoFinal.col_values(62)
        TiempoManoDeObraEGCP2H=TesteoFinal.col_values(59)
        mensajeManoDeObraEGCP2H="*Hubo afectación en las unidades por Mano De Obra: Razón:* "+DescrManoDeObraEGCP2H[-1]+ " - *Tiempo:* "+TiempoManoDeObraEGCP2H[-1]+"min" 
        print(mensajeManoDeObraEGCP2H)
    else:
        mensajeManoDeObraEGCP2H=""
        print(mensajeManoDeObraEGCP2H)

#MATERIA PRIMA COPA2::::

    MateriaPrimaEGCP2H=TesteoFinal.col_values(63)
    if MateriaPrimaEGCP2H[-1]=="Si":
        DescrMateriaPrimaEGCP2H=TesteoFinal.col_values(67)
        TiempoMateriaPrimaEGCP2H=TesteoFinal.col_values(64)
        mensajeMateriaPrimaEGCP2H="*Hubo afectación en las unidades por Materia Prima: Razón:* "+DescrMateriaPrimaEGCP2H[-1]+ " - *Tiempo:* "+TiempoMateriaPrimaEGCP2H[-1]+"min" 
        print(mensajeMateriaPrimaEGCP2H)
    else:
        mensajeMateriaPrimaEGCP2H=""
        print(mensajeMateriaPrimaEGCP2H)

#POR METODO COPA2:::
    MetodoEGCP2H=TesteoFinal.col_values(68)
    if MetodoEGCP2H[-1]=="Si":
        DescrMetodoEGCP2H=TesteoFinal.col_values(71)
        TiempoMetodoEGCP2H=TesteoFinal.col_values(69)
        mensajeMetodoEGCP2H="*Hubo afectación en las unidades por Método: Razón:* "+DescrMetodoEGCP2H[-1]+ "- *Tiempo:* "+TiempoMetodoEGCP2H[-1]+"min" 
        print(mensajeMetodoEGCP2H)
    else:
        mensajeMetodoEGCP2H=""
        print(mensajeMetodoEGCP2H)

#SCRAP COPA2::::::::::
    ScrapEGCP2H=TesteoFinal.col_values(72)
    if ScrapEGCP2H[-1]=="Si":
        DescrScrapEGCP2H=TesteoFinal.col_values(74)
        CantidadScrapEGCP2H=TesteoFinal.col_values(75)
        mensajeScrapEGCP2H="*Se generó SCRAP: Cantidad:* "+CantidadScrapEGCP2H[-1]+" - *Razón:* "+DescrScrapEGCP2H[-1]
        print(mensajeScrapEGCP2H)
    else:
        mensajeScrapEGCP2H=""
        print(mensajeScrapEGCP2H)

#REPROCESADAS COPA2::::::::
    UnidadesReprocesadasEGCP2H =  TesteoFinal.col_values(76)
    if UnidadesReprocesadasEGCP2H[-1]=="Si":
        CantidadReprocesadasEGCP2H=  TesteoFinal.col_values(77)

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
    ParoProgramadoEGCP2W = TesteoFinal.col_values(78)
    if ParoProgramadoEGCP2W[-1]=="Si":
        RazonParoProgramadoEGCP2W =  TesteoFinal.col_values(79)
        TiempoParoProgramadoEGCP2W =  TesteoFinal.col_values(80)
        mensajeParoProgramadoEGCP2W="*Paro programado - Tiempo:* "+TiempoParoProgramadoEGCP2W[-1]+" min, *Razón:* "+RazonParoProgramadoEGCP2W[-1]
        print(mensajeParoProgramadoEGCP2W)
    else:
        mensajeParoProgramadoEGCP2W=""
        print(mensajeParoProgramadoEGCP2W)
    
    #INCIDENTES WHIRLPOOL COPA 2:::::
    IncidenteEGCP2W=TesteoFinal.col_values(81)
    if IncidenteEGCP2W[-1]=="Si":
        DescrIncidenteEGCP2W=TesteoFinal.col_values(83)
        ValidarParoIncidenteEMCP2W=TesteoFinal.col_values(84)
        mensajeIncidenteEGCP2W="*Incidente y/o accidente ambiental y/o SST: Razón:* "+DescrIncidenteEGCP2W[-1]+ " no se generó paro."
        print(mensajeIncidenteEGCP2W)
        if ValidarParoIncidenteEMCP2W[-1]=="Si":   
            TiempoIncidenteEGCP2W=TesteoFinal.col_values(85)
            mensajeIncidenteEGCP2W="*Incidente y/o accidente ambiental y/o SST - Tiempo:* "+TiempoIncidenteEGCP2W[-1]+" min, *Razón:* "+DescrIncidenteEGCP2W[-1]
            print (mensajeIncidenteEGCP2W)
        else:
            #DescrIncidenteEGCP1=TesteoFinal.col_values(12)
            mensajeIncidenteEGCP2W="*Incidente y/o accidente ambiental y/o SST: Razón:* "+DescrIncidenteEGCP2W[-1] + " no se generó paro."
            print (mensajeIncidenteEGCP2W)
    else:
        mensajeIncidenteEGCP2H=""
        print(mensajeIncidenteEGCP2H)

##SERVICIOS PUBLICOS COPA2 WHIRPOOL:::
    ServiciosPublicosEGCP2W=TesteoFinal.col_values(86)
    if ServiciosPublicosEGCP2W[-1]=="Si":
        DescrServiciosPublicosEGCP2W=TesteoFinal.col_values(87)
        TiempoServiciosPublicosEGCP2W=TesteoFinal.col_values(88)
        mensajeServiciosPublicosEGCP2W="*Hubo afectación en las unidades del hora a hora por falta de servicios públicos: Razón:* "+DescrServiciosPublicosEGCP2W[-1]+ " -*Tiempo:* :"+TiempoServiciosPublicosEGCP2W[-1]+"min"
        print(mensajeServiciosPublicosEGCP2W)
    else:
        mensajeServiciosPublicosEGCP2W=""
        print(mensajeServiciosPublicosEGCP2W)

#POR MAQUINA COPA2 WHIRLPOOL::::::::
    MaquinaEGCP2W=TesteoFinal.col_values(89)
    if MaquinaEGCP2W[-1]=="Si":
        DescrMaquinaEGCP2W=TesteoFinal.col_values(92)
        TiempoMaquinaEGCP2W=TesteoFinal.col_values(90)
        mensajeMaquinaEGCP2W="*Hubo afectación en las unidades por Maquina/ Equipo: Razón:* "+DescrMaquinaEGCP2W[-1]+ " - *Tiempo:* "+TiempoMaquinaEGCP2W[-1]+"min" 
        print(mensajeMaquinaEGCP2W)
    else:
        mensajeMaquinaEGCP2W=""
        print(mensajeMaquinaEGCP2W)

#POR MANO DE OBRA COPA2 WHIRLPOOL::::::::
    ManoDeObraEGCP2W=TesteoFinal.col_values(93)
    if ManoDeObraEGCP2W[-1]=="Si":
        DescrManoDeObraEGCP2W=TesteoFinal.col_values(97)
        TiempoManoDeObraEGCP2W=TesteoFinal.col_values(94)
        mensajeManoDeObraEGCP2W="*Hubo afectación en las unidades por Mano De Obra: Razón:* "+DescrManoDeObraEGCP2W[-1]+ " - *Tiempo:* "+TiempoManoDeObraEGCP2W[-1]+"min" 
        print(mensajeManoDeObraEGCP2W)
    else:
        mensajeManoDeObraEGCP2W=""
        print(mensajeManoDeObraEGCP2W)

#MATERIA PRIMA COPA2 WHIRPOOL::::

    MateriaPrimaEGCP2W=TesteoFinal.col_values(98)
    if MateriaPrimaEGCP2W[-1]=="Si":
        DescrMateriaPrimaEGCP2W=TesteoFinal.col_values(102)
        TiempoMateriaPrimaEGCP2W=TesteoFinal.col_values(99)
        mensajeMateriaPrimaEGCP2W="*Hubo afectación en las unidades por Materia Prima: Razón:* "+DescrMateriaPrimaEGCP2W[-1]+ " - *Tiempo:* "+TiempoMateriaPrimaEGCP2W[-1]+"min" 
        print(mensajeMateriaPrimaEGCP2W)
    else:
        mensajeMateriaPrimaEGCP2W=""
        print(mensajeMateriaPrimaEGCP2W)

#POR METODO COPA2 WHIRLPOOL:::
    MetodoEGCP2W=TesteoFinal.col_values(103)
    if MetodoEGCP2W[-1]=="Si":
        DescrMetodoEGCP2W=TesteoFinal.col_values(106)
        TiempoMetodoEGCP2W=TesteoFinal.col_values(104)
        mensajeMetodoEGCP2W="*Hubo afectación en las unidades por Método: Razón:* "+DescrMetodoEGCP2W[-1]+ "- *Tiempo:* "+TiempoMetodoEGCP2W[-1]+"min" 
        print(mensajeMetodoEGCP2W)
    else:
        mensajeMetodoEGCP2W=""
        print(mensajeMetodoEGCP2W)

#SCRAP COPA2 WHIRLPOOL::::::::::
    ScrapEGCP2W=TesteoFinal.col_values(107)
    if ScrapEGCP2W[-1]=="Si":
        DescrScrapEGCP2W=TesteoFinal.col_values(109)
        CantidadScrapEGCP2W=TesteoFinal.col_values(110)
        mensajeScrapEGCP2W="*Se generó SCRAP: Cantidad:* "+CantidadScrapEGCP2W[-1]+" - *Razón:* "+DescrScrapEGCP2W[-1]
        print(mensajeScrapEGCP2W)
    else:
        mensajeScrapEGCP2W=""
        print(mensajeScrapEGCP2W)

#REPROCESADAS COPA2::::::::
    UnidadesReprocesadasEGCP2W =  TesteoFinal.col_values(111)
    if UnidadesReprocesadasEGCP2W[-1]=="Si":
        CantidadReprocesadasEGCP2W=  TesteoFinal.col_values(112)

        mensajeUnidadesReprocesadasEGCP2W="*Se reprocesaron unidades - Cantidad:* "+CantidadReprocesadasEGCP2W[-1]+""
        print (mensajeUnidadesReprocesadasEGCP2W)
    else:
        mensajeUnidadesReprocesadasEGCP2W=""
        print(mensajeUnidadesReprocesadasEGCP2W)


OeeEG= TesteoFinal.col_values(118)
OeeEmsableGabinete = OeeEG[-1]
print(OeeEmsableGabinete)