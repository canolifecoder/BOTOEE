#!/usr/bin/python

from tkinter.tix import Select
import gspread
import pywhatkit
import datetime
import time


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


    
sh = gc3.open("Gestión Célula Hora a Hora Células")
""" sh2 = gc4.open("Gestión Célula Hora a Hora Células") """




#TESTEO FINAL:::::::::::::::
print("TESTEO FINAL::---------")

#SELECCION DE LA HOJA::
TesteoFinal = sh.get_worksheet(3)
#SELECCIONAR LA REFERENCIA::::: COPA 1 -- COPA 2 HACEB -- COPA 2 WHIRLPOOL
UnidadesFabricadasTF=  TesteoFinal.col_values(5)
print("Unidades Fabricadas: "+UnidadesFabricadasTF[-1])
MensajeUnidadesFabricadasTF="Unidades Producidas: "+UnidadesFabricadasTF[-1]
#SELECCION DE COPA 1:::::
SelectReferenciaTF = TesteoFinal.col_values(7)
if SelectReferenciaTF[-1]=="Copa 1.0":
    print("COPA 1::::")
    #PAROS PROGRAMADOS::
    ParoProgramadoTFCP1 = TesteoFinal.col_values(8)
    if ParoProgramadoTFCP1[-1]=="Si":
        RazonParoProgramadoTFCP1 =  TesteoFinal.col_values(9)
        TiempoParoProgramadoTFCP1 =  TesteoFinal.col_values(10)
        mensajeParoProgramadoTFCP1="*Paro programado - Tiempo:* "+TiempoParoProgramadoTFCP1[-1]+" min, *Razón:* "+RazonParoProgramadoTFCP1[-1]
        print(mensajeParoProgramadoTFCP1)
    else:
        mensajeParoProgramadoTFCP1=""
        print(mensajeParoProgramadoTFCP1)
    
    #INCIDENTES::
    IncidenteTFCP1=TesteoFinal.col_values(11)
    if IncidenteTFCP1[-1]=="Si":
        DescrIncidenteTFCP1=TesteoFinal.col_values(13)
        ValidarParoIncidenteTFCP1=TesteoFinal.col_values(14)
        mensajeIncidenteTFCP1="*Incidente y/o accidente ambiental y/o SST: Razón:* "+DescrIncidenteTFCP1[-1] + " No se generó paro"
        print(mensajeIncidenteTFCP1)
        if ValidarParoIncidenteTFCP1[-1]=="Si":   
            TiempoIncidenteTFCP1=TesteoFinal.col_values(15)
            mensajeIncidenteTFCP1="*Incidente y/o accidente ambiental y/o SST - Tiempo:* "+TiempoIncidenteTFCP1[-1]+" min, *Razón:* "+DescrIncidenteTFCP1[-1]
            print (mensajeIncidenteTFCP1)
        else:
            DescrIncidenteTFCP1=TesteoFinal.col_values(12)
            mensajeIncidenteTFCP1="*Incidente y/o accidente ambiental y/o SST: Razón:* "+DescrIncidenteTFCP1[-1] + " No se generó paro"
            print (mensajeIncidenteTFCP1)
    else:
        mensajeIncidenteTFCP1=""
        print(mensajeIncidenteTFCP1)

##SERVICIOS PUBLICOS COPA1::
    ServiciosPublicosTFCP1=TesteoFinal.col_values(16)
    if ServiciosPublicosTFCP1[-1]=="Si":
        DescrServiciosPublicosTFCP1=TesteoFinal.col_values(18)
        TiempoServiciosPublucosTFCP1=TesteoFinal.col_values(17)
        mensajeServiciosPublicosTFCP1="*Hubo afectación en las unidades del hora a hora por falta de servicios públicos: Razón:* "+DescrServiciosPublicosTFCP1[-1]+ " -*Tiempo:* :"+TiempoServiciosPublucosTFCP1[-1]+"min"
        print(mensajeServiciosPublicosTFCP1)
    else:
        mensajeServiciosPublicosTFCP1=""
        print(mensajeServiciosPublicosTFCP1)
#POR MAQUINA COPA1:::
    MaquinaTFCP1=TesteoFinal.col_values(19)
    if MaquinaTFCP1[-1]=="Si":
        DescrMaquinaTFCP1=TesteoFinal.col_values(22)
        TiempoMaquinaTFCP1=TesteoFinal.col_values(20)
        mensajeMaquinaTFCP1="*Hubo afectación en las unidades por Maquina/ Equipo: Razón:* "+DescrMaquinaTFCP1[-1]+ " - *Tiempo:* "+TiempoMaquinaTFCP1[-1]+"min" 
        print(mensajeMaquinaTFCP1)
    else:
        mensajeMaquinaTFCP1=""
        print(mensajeMaquinaTFCP1)

#POR MANO DE OBRA COPA1::::::::
    ManoDeObraTFCP1=TesteoFinal.col_values(23)
    if ManoDeObraTFCP1[-1]=="Si":
        DescrManoDeObraTFCP1=TesteoFinal.col_values(27)
        TiempoManoDeObraTFCP1=TesteoFinal.col_values(24)
        mensajeManoDeObraTFCP1="*Hubo afectación en las unidades por Mano De Obra: Razón:* "+DescrManoDeObraTFCP1[-1]+ " - *Tiempo:* "+TiempoManoDeObraTFCP1[-1]+"min" 
        print(mensajeManoDeObraTFCP1)
    else:
        mensajeManoDeObraTFCP1=""
        print(mensajeManoDeObraTFCP1)

#MATERIA PRIMA COPA1::::

    MateriaPrimaTFCP1=TesteoFinal.col_values(28)
    if MateriaPrimaTFCP1[-1]=="Si":
        DescrMateriaPrimaTFCP1=TesteoFinal.col_values(32)
        TiempoMateriaPrimaTFCP1=TesteoFinal.col_values(29)
        mensajeMateriaPrimaTFCP1="*Hubo afectación en las unidades por Materia Prima: Razón:* "+DescrMateriaPrimaTFCP1[-1]+ " - *Tiempo:* "+TiempoMateriaPrimaTFCP1[-1]+"min" 
        print(mensajeMateriaPrimaTFCP1)
    else:
        mensajeMateriaPrimaTFCP1=""
        print(mensajeMateriaPrimaTFCP1)

#POR METODO COPA1:::
    MetodoTFCP1=TesteoFinal.col_values(33)
    if MetodoTFCP1[-1]=="Si":
        DescrMetodoTFCP1=TesteoFinal.col_values(36)
        TiempoMetodoTFCP1=TesteoFinal.col_values(34)
        mensajeMetodoTFCP1="*Hubo afectación en las unidades por Método: Razón:* "+DescrMetodoTFCP1[-1]+ "- *Tiempo:* "+TiempoMetodoTFCP1[-1]+"min" 
        print(mensajeMetodoTFCP1)
    else:
        mensajeMetodoTFCP1=""
        print(mensajeMetodoTFCP1)

#SCRAP COPA1::::::::::
    ScrapTFCP1=TesteoFinal.col_values(37)
    if ScrapTFCP1[-1]=="Si":
        DescrScrapTFCP1=TesteoFinal.col_values(39)
        CantidadScrapTFCP1=TesteoFinal.col_values(40)
        mensajeScrapTFCP1="*Se generó SCRAP: Cantidad:* "+CantidadScrapTFCP1[-1]+" - *Razón:* "+DescrScrapTFCP1[-1]
        print(mensajeScrapTFCP1)
    else:
        mensajeScrapTFCP1=""
        print(mensajeScrapTFCP1)

#REPROCESADAS COPA1::::::::
    UnidadesReprocesadasTFP1 =  TesteoFinal.col_values(41)
    if UnidadesReprocesadasTFP1[-1]=="Si":
        CantidadReprocesadasTFP1=  TesteoFinal.col_values(42)

        mensajeUnidadesReprocesadasTFP1="*Se reprocesaron unidades - Cantidad:* "+CantidadReprocesadasTFP1[-1]+""
        print (mensajeUnidadesReprocesadasTFP1)
    else:
        mensajeUnidadesReprocesadasTFP1=""
        print(mensajeUnidadesReprocesadasTFP1)

##SELECCIONA COPA 2 HACEB::::::::::::::::::::::

if SelectReferenciaTF[-1]=="Copa 2.0 Haceb":
    print("COPA 2 HACEB::::")
    #PAROS PROGRAMADOS::
    ParoProgramadoTFCP2H = TesteoFinal.col_values(43)
    if ParoProgramadoTFCP2H[-1]=="Si":
        RazonParoProgramadoTFCP2H =  TesteoFinal.col_values(44)
        TiempoParoProgramadoTFCP2H =  TesteoFinal.col_values(45)
        mensajeParoProgramadoTFCP2H="*Paro programado - Tiempo:* "+TiempoParoProgramadoTFCP2H[-1]+" min, *Razón:* "+RazonParoProgramadoTFCP2H[-1]
        print(mensajeParoProgramadoTFCP2H)
    else:
        mensajeParoProgramadoTFCP2H=""
        print(mensajeParoProgramadoTFCP2H)
    
    #INCIDENTES::
    IncidenteTFCP2H=TesteoFinal.col_values(46)
    if IncidenteTFCP2H[-1]=="Si":
        DescrIncidenteTFCP2H=TesteoFinal.col_values(48)
        ValidarParoIncidenteTFCP2H=TesteoFinal.col_values(49)
        mensajeIncidenteTFCP2H="*Incidente y/o accidente ambiental y/o SST: Razón:* "+DescrIncidenteTFCP2H[-1]+ " no se generó paro."
        print(mensajeIncidenteTFCP2H)
        if ValidarParoIncidenteTFCP2H[-1]=="Si":   
            TiempoIncidenteTFCP2H=TesteoFinal.col_values(50)
            mensajeIncidenteTFCP2H="*Incidente y/o accidente ambiental y/o SST - Tiempo:* "+TiempoIncidenteTFCP2H[-1]+" min, *Razón:* "+DescrIncidenteTFCP2H[-1]
            print (mensajeIncidenteTFCP2H)
        else:
            #DescrIncidenteEGCP1=TesteoFinal.col_values(12)
            mensajeIncidenteTFCP2H="*Incidente y/o accidente ambiental y/o SST: Razón:* "+DescrIncidenteTFCP2H[-1] + " no se generó paro."
            print (mensajeIncidenteTFCP2H)
    else:
        mensajeIncidenteTFCP2H=""
        print(mensajeIncidenteTFCP2H)

##SERVICIOS PUBLICOS COPA2::
    ServiciosPublicosTFCP2H=TesteoFinal.col_values(51)
    if ServiciosPublicosTFCP2H[-1]=="Si":
        DescrServiciosPublicosTFCP2H=TesteoFinal.col_values(53)
        TiempoServiciosPublicosTFCP2H=TesteoFinal.col_values(52)
        mensajeServiciosPublicosTFCP2H="*Hubo afectación en las unidades del hora a hora por falta de servicios públicos: Razón:* "+DescrServiciosPublicosTFCP2H[-1]+ " -*Tiempo:* :"+TiempoServiciosPublicosTFCP2H[-1]+"min"
        print(mensajeServiciosPublicosTFCP2H)
    else:
        mensajeServiciosPublicosTFCP2H=""
        print(mensajeServiciosPublicosTFCP2H)
#POR MAQUINA COPA2:::
    MaquinaTFCP2H=TesteoFinal.col_values(54)
    if MaquinaTFCP2H[-1]=="Si":
        DescrMaquinaTFCP2H=TesteoFinal.col_values(57)
        TiempoMaquinaTFCP2H=TesteoFinal.col_values(55)
        mensajeMaquinaTFCP2H="*Hubo afectación en las unidades por Maquina/ Equipo: Razón:* "+DescrMaquinaTFCP2H[-1]+ " - *Tiempo:* "+TiempoMaquinaTFCP2H[-1]+"min" 
        print(mensajeMaquinaTFCP2H)
    else:
        mensajeMaquinaTFCP2H=""
        print(mensajeMaquinaTFCP2H)

#POR MANO DE OBRA COPA2::::::::
    ManoDeObraTFCP2H=TesteoFinal.col_values(58)
    if ManoDeObraTFCP2H[-1]=="Si":
        DescrManoDeObraTFCP2H=TesteoFinal.col_values(62)
        TiempoManoDeObraTFCP2H=TesteoFinal.col_values(59)
        mensajeManoDeObraTFCP2H="*Hubo afectación en las unidades por Mano De Obra: Razón:* "+DescrManoDeObraTFCP2H[-1]+ " - *Tiempo:* "+TiempoManoDeObraTFCP2H[-1]+"min" 
        print(mensajeManoDeObraTFCP2H)
    else:
        mensajeManoDeObraTFCP2H=""
        print(mensajeManoDeObraTFCP2H)

#MATERIA PRIMA COPA2::::

    MateriaPrimaTFCP2H=TesteoFinal.col_values(63)
    if MateriaPrimaTFCP2H[-1]=="Si":
        DescrMateriaPrimaTFCP2H=TesteoFinal.col_values(67)
        TiempoMateriaPrimaTFCP2H=TesteoFinal.col_values(64)
        mensajeMateriaPrimaTFCP2H="*Hubo afectación en las unidades por Materia Prima: Razón:* "+DescrMateriaPrimaTFCP2H[-1]+ " - *Tiempo:* "+TiempoMateriaPrimaTFCP2H[-1]+"min" 
        print(mensajeMateriaPrimaTFCP2H)
    else:
        mensajeMateriaPrimaTFCP2H=""
        print(mensajeMateriaPrimaTFCP2H)

#POR METODO COPA2:::
    MetodoTFCP2H=TesteoFinal.col_values(68)
    if MetodoTFCP2H[-1]=="Si":
        DescrMetodoTFCP2H=TesteoFinal.col_values(71)
        TiempoMetodoTFCP2H=TesteoFinal.col_values(69)
        mensajeMetodoTFCP2H="*Hubo afectación en las unidades por Método: Razón:* "+DescrMetodoTFCP2H[-1]+ "- *Tiempo:* "+TiempoMetodoTFCP2H[-1]+"min" 
        print(mensajeMetodoTFCP2H)
    else:
        mensajeMetodoTFCP2H=""
        print(mensajeMetodoTFCP2H)

#SCRAP COPA2::::::::::
    ScrapTFCP2H=TesteoFinal.col_values(72)
    if ScrapTFCP2H[-1]=="Si":
        DescrScrapTFCP2H=TesteoFinal.col_values(74)
        CantidadScrapTFCP2H=TesteoFinal.col_values(75)
        mensajeScrapTFCP2H="*Se generó SCRAP: Cantidad:* "+CantidadScrapTFCP2H[-1]+" - *Razón:* "+DescrScrapTFCP2H[-1]
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
# EMSAMBLE MECANISMOS COPA 2 WHIRLPOOL__:::::
if SelectReferenciaTF[-1]=="Copa 2.0 Whirlpool":
    print("COPA 2 Whirlpool::::::")
    #PAROS PROGRAMADOS::::
    ParoProgramadoTFCP2W = TesteoFinal.col_values(78)
    if ParoProgramadoTFCP2W[-1]=="Si":
        RazonParoProgramadoTFCP2W =  TesteoFinal.col_values(79)
        TiempoParoProgramadoTFCP2W =  TesteoFinal.col_values(80)
        mensajeParoProgramadoTFCP2W="*Paro programado - Tiempo:* "+TiempoParoProgramadoTFCP2W[-1]+" min, *Razón:* "+RazonParoProgramadoTFCP2W[-1]
        print(mensajeParoProgramadoTFCP2W)
    else:
        mensajeParoProgramadoTFCP2W=""
        print(mensajeParoProgramadoTFCP2W)
    
    #INCIDENTES WHIRLPOOL COPA 2:::::
    IncidenteTFCP2W=TesteoFinal.col_values(81)
    if IncidenteTFCP2W[-1]=="Si":
        DescrIncidenteTFCP2W=TesteoFinal.col_values(83)
        ValidarParoIncidenteTFCP2W=TesteoFinal.col_values(84)
        mensajeIncidenteTFCP2W="*Incidente y/o accidente ambiental y/o SST: Razón:* "+DescrIncidenteTFCP2W[-1]+ " no se generó paro."
        print(mensajeIncidenteTFCP2W)
        if ValidarParoIncidenteTFCP2W[-1]=="Si":   
            TiempoIncidenteTFCP2W=TesteoFinal.col_values(85)
            mensajeIncidenteTFCP2W="*Incidente y/o accidente ambiental y/o SST - Tiempo:* "+TiempoIncidenteTFCP2W[-1]+" min, *Razón:* "+DescrIncidenteTFCP2W[-1]
            print (mensajeIncidenteTFCP2W)
        else:
            #DescrIncidenteEGCP1=TesteoFinal.col_values(12)
            mensajeIncidenteTFCP2W="*Incidente y/o accidente ambiental y/o SST: Razón:* "+DescrIncidenteTFCP2W[-1] + " no se generó paro."
            print (mensajeIncidenteTFCP2W)
    else:
        mensajeIncidenteTFCP2H=""
        print(mensajeIncidenteTFCP2H)

##SERVICIOS PUBLICOS COPA2 WHIRPOOL:::
    ServiciosPublicosTFCP2W=TesteoFinal.col_values(86)
    if ServiciosPublicosTFCP2W[-1]=="Si":
        DescrServiciosPublicosTFCP2W=TesteoFinal.col_values(87)
        TiempoServiciosPublicosTFCP2W=TesteoFinal.col_values(88)
        mensajeServiciosPublicosTFCP2W="*Hubo afectación en las unidades del hora a hora por falta de servicios públicos: Razón:* "+DescrServiciosPublicosTFCP2W[-1]+ " -*Tiempo:* :"+TiempoServiciosPublicosTFCP2W[-1]+"min"
        print(mensajeServiciosPublicosTFCP2W)
    else:
        mensajeServiciosPublicosTFCP2W=""
        print(mensajeServiciosPublicosTFCP2W)

#POR MAQUINA COPA2 WHIRLPOOL::::::::
    MaquinaTFCP2W=TesteoFinal.col_values(89)
    if MaquinaTFCP2W[-1]=="Si":
        DescrMaquinaTFCP2W=TesteoFinal.col_values(92)
        TiempoMaquinaTFCP2W=TesteoFinal.col_values(90)
        mensajeMaquinaTFCP2W="*Hubo afectación en las unidades por Maquina/ Equipo: Razón:* "+DescrMaquinaTFCP2W[-1]+ " - *Tiempo:* "+TiempoMaquinaTFCP2W[-1]+"min" 
        print(mensajeMaquinaTFCP2W)
    else:
        mensajeMaquinaTFCP2W=""
        print(mensajeMaquinaTFCP2W)

#POR MANO DE OBRA COPA2 WHIRLPOOL::::::::
    ManoDeObraTFCP2W=TesteoFinal.col_values(93)
    if ManoDeObraTFCP2W[-1]=="Si":
        DescrManoDeObraTFCP2W=TesteoFinal.col_values(97)
        TiempoManoDeObraTFCP2W=TesteoFinal.col_values(94)
        mensajeManoDeObraTFCP2W="*Hubo afectación en las unidades por Mano De Obra: Razón:* "+DescrManoDeObraTFCP2W[-1]+ " - *Tiempo:* "+TiempoManoDeObraTFCP2W[-1]+"min" 
        print(mensajeManoDeObraTFCP2W)
    else:
        mensajeManoDeObraTFCP2W=""
        print(mensajeManoDeObraTFCP2W)

#MATERIA PRIMA COPA2 WHIRPOOL::::

    MateriaPrimaTFCP2W=TesteoFinal.col_values(98)
    if MateriaPrimaTFCP2W[-1]=="Si":
        DescrMateriaPrimaTFCP2W=TesteoFinal.col_values(102)
        TiempoMateriaPrimaTFCP2W=TesteoFinal.col_values(99)
        mensajeMateriaPrimaTFCP2W="*Hubo afectación en las unidades por Materia Prima: Razón:* "+DescrMateriaPrimaTFCP2W[-1]+ " - *Tiempo:* "+TiempoMateriaPrimaTFCP2W[-1]+"min" 
        print(mensajeMateriaPrimaTFCP2W)
    else:
        mensajeMateriaPrimaTFCP2W=""
        print(mensajeMateriaPrimaTFCP2W)

#POR METODO COPA2 WHIRLPOOL:::
    MetodoTFCP2W=TesteoFinal.col_values(103)
    if MetodoTFCP2W[-1]=="Si":
        DescrMetodoTFCP2W=TesteoFinal.col_values(106)
        TiempoMetodoTFCP2W=TesteoFinal.col_values(104)
        mensajeMetodoTFCP2W="*Hubo afectación en las unidades por Método: Razón:* "+DescrMetodoTFCP2W[-1]+ "- *Tiempo:* "+TiempoMetodoTFCP2W[-1]+"min" 
        print(mensajeMetodoTFCP2W)
    else:
        mensajeMetodoTFCP2W=""
        print(mensajeMetodoTFCP2W)

#SCRAP COPA2 WHIRLPOOL::::::::::
    ScrapTFCP2W=TesteoFinal.col_values(107)
    if ScrapTFCP2W[-1]=="Si":
        DescrScrapTFCP2W=TesteoFinal.col_values(109)
        CantidadScrapTFCP2W=TesteoFinal.col_values(110)
        mensajeScrapTFCP2W="*Se generó SCRAP: Cantidad:* "+CantidadScrapTFCP2W[-1]+" - *Razón:* "+DescrScrapTFCP2W[-1]
        print(mensajeScrapTFCP2W)
    else:
        mensajeScrapTFCP2W=""
        print(mensajeScrapTFCP2W)

#REPROCESADAS COPA2::::::::
    UnidadesReprocesadasTFCP2W =  TesteoFinal.col_values(111)
    if UnidadesReprocesadasTFCP2W[-1]=="Si":
        CantidadReprocesadasTFCP2W=  TesteoFinal.col_values(112)

        mensajeUnidadesReprocesadasTFCP2W="*Se reprocesaron unidades - Cantidad:* "+CantidadReprocesadasTFCP2W[-1]+""
        print (mensajeUnidadesReprocesadasTFCP2W)
    else:
        mensajeUnidadesReprocesadasTFCP2W=""
        print(mensajeUnidadesReprocesadasTFCP2W)


OeeTF= TesteoFinal.col_values(118)
OeeTesteoFinal = OeeTF[-1]
print("Porcentaje OEE: " + OeeTesteoFinal)




#TAPA MOVIL:::::::::::::::::
print("TAPA MOVIL:::---------")

#SELECCION DE LA HOJA::
TapaMovil = sh.get_worksheet(4)
#SELECCIONAR LA REFERENCIA::::: COPA 1 -- COPA 2 HACEB -- COPA 2 WHIRLPOOL
UnidadesFabricadasTM=  TapaMovil.col_values(5)
print("Unidades Fabricadas: "+UnidadesFabricadasTM[-1])
MensajeUnidadesFabricadasTM="Unidades Producidas: "+ UnidadesFabricadasTM[-1]
#SELECCION DE COPA 1:::::
SelectReferenciaTM = TapaMovil.col_values(7)
if SelectReferenciaTM[-1]=="Copa 1.0":
    print("COPA 1::::")
    #PAROS PROGRAMADOS TAPA MOVIL COPA 1::::
    ParoProgramadoTMCP1 = TapaMovil.col_values(8)
    if ParoProgramadoTMCP1[-1]=="Si":
        RazonParoProgramadoTMCP1 =  TapaMovil.col_values(9)
        TiempoParoProgramadoTMCP1 =  TapaMovil.col_values(10)
        mensajeParoProgramadoTMCP1="*Paro programado - Tiempo:* "+TiempoParoProgramadoTMCP1[-1]+" min, *Razón:* "+RazonParoProgramadoTMCP1[-1]
        print(mensajeParoProgramadoTMCP1)
    else:
        mensajeParoProgramadoTMCP1=""
        print(mensajeParoProgramadoTMCP1)
    
    #INCIDENTES TAPA MOVIL COPA 1::
    IncidenteTMCP1=TapaMovil.col_values(11)
    if IncidenteTMCP1[-1]=="Si":
        DescrIncidenteTMCP1=TapaMovil.col_values(13)
        ValidarParoIncidenteTMCP1=TapaMovil.col_values(14)
        mensajeIncidenteTMCP1="*Incidente y/o accidente ambiental y/o SST: Razón:* "+DescrIncidenteTMCP1[-1] + " No se generó paro"
        print(mensajeIncidenteTMCP1)
        if ValidarParoIncidenteTMCP1[-1]=="Si":   
            TiempoIncidenteTMCP1=TapaMovil.col_values(15)
            mensajeIncidenteTMCP1="*Incidente y/o accidente ambiental y/o SST - Tiempo:* "+TiempoIncidenteTMCP1[-1]+" min, *Razón:* "+DescrIncidenteTMCP1[-1]
            print (mensajeIncidenteTMCP1)
        else:
            #DescrIncidenteTMCP1=TapaMovil.col_values(12)
            mensajeIncidenteTMCP1="*Incidente y/o accidente ambiental y/o SST: Razón:* "+DescrIncidenteTMCP1[-1] + " No se generó paro"
            print (mensajeIncidenteTMCP1)
    else:
        mensajeIncidenteTMCP1=""
        print(mensajeIncidenteTMCP1)

##SERVICIOS PUBLICOS  TAPA MOVIL COPA1:::.
    ServiciosPublicosTMCP1=TapaMovil.col_values(16)
    if ServiciosPublicosTMCP1[-1]=="Si":
        DescrServiciosPublicosTMCP1=TapaMovil.col_values(18)
        TiempoServiciosPublucosTMCP1=TapaMovil.col_values(17)
        mensajeServiciosPublicosTMCP1="*Hubo afectación en las unidades del hora a hora por falta de servicios públicos: Razón:* "+DescrServiciosPublicosTMCP1[-1]+ " -*Tiempo:* :"+TiempoServiciosPublucosTMCP1[-1]+"min"
        print(mensajeServiciosPublicosTMCP1)
    else:
        mensajeServiciosPublicosTMCP1=""
        print(mensajeServiciosPublicosTMCP1)
#POR MAQUINA COPA1 TAPA MOVIL:::
    MaquinaTMCP1=TapaMovil.col_values(19)
    if MaquinaTMCP1[-1]=="Si":
        DescrMaquinaTMCP1=TapaMovil.col_values(22)
        TiempoMaquinaTMCP1=TapaMovil.col_values(20)
        mensajeMaquinaTMCP1="*Hubo afectación en las unidades por Maquina/ Equipo: Razón:* "+DescrMaquinaTMCP1[-1]+ " - *Tiempo:* "+TiempoMaquinaTMCP1[-1]+"min" 
        print(mensajeMaquinaTMCP1)
    else:
        mensajeMaquinaTMCP1=""
        print(mensajeMaquinaTMCP1)

#POR MANO DE OBRA COPA1 TAPA MOVIL::::::::
    ManoDeObraTMCP1=TapaMovil.col_values(23)
    if ManoDeObraTMCP1[-1]=="Si":
        DescrManoDeObraTMCP1=TapaMovil.col_values(27)
        TiempoManoDeObraTMCP1=TapaMovil.col_values(24)
        mensajeManoDeObraTMCP1="*Hubo afectación en las unidades por Mano De Obra: Razón:* "+DescrManoDeObraTMCP1[-1]+ " - *Tiempo:* "+TiempoManoDeObraTMCP1[-1]+"min" 
        print(mensajeManoDeObraTMCP1)
    else:
        mensajeManoDeObraTMCP1=""
        print(mensajeManoDeObraTMCP1)

#MATERIA PRIMA COPA1 TAPA MOVIL::::

    MateriaPrimaTMCP1=TapaMovil.col_values(28)
    if MateriaPrimaTMCP1[-1]=="Si":
        DescrMateriaPrimaTMCP1=TapaMovil.col_values(32)
        TiempoMateriaPrimaTMCP1=TapaMovil.col_values(29)
        mensajeMateriaPrimaTMCP1="*Hubo afectación en las unidades por Materia Prima: Razón:* "+DescrMateriaPrimaTMCP1[-1]+ " - *Tiempo:* "+TiempoMateriaPrimaTMCP1[-1]+"min" 
        print(mensajeMateriaPrimaTMCP1)
    else:
        mensajeMateriaPrimaTMCP1=""
        print(mensajeMateriaPrimaTMCP1)

#POR METODO COPA1 TAPA MOVIL:::
    MetodoTMCP1=TapaMovil.col_values(33)
    if MetodoTMCP1[-1]=="Si":
        DescrMetodoTMCP1=TapaMovil.col_values(36)
        TiempoMetodoTMCP1=TapaMovil.col_values(34)
        mensajeMetodoTMCP1="*Hubo afectación en las unidades por Método: Razón:* "+DescrMetodoTMCP1[-1]+ "- *Tiempo:* "+TiempoMetodoTMCP1[-1]+"min" 
        print(mensajeMetodoTMCP1)
    else:
        mensajeMetodoTMCP1=""
        print(mensajeMetodoTMCP1)

#SCRAP COPA1 TAPA MOVIL::::::::::
    ScrapTMCP1=TapaMovil.col_values(37)
    if ScrapTMCP1[-1]=="Si":
        DescrScrapTMCP1=TapaMovil.col_values(39)
        CantidadScrapTMCP1=TapaMovil.col_values(40)
        mensajeScrapTMCP1="*Se generó SCRAP: Cantidad:* "+CantidadScrapTMCP1[-1]+" - *Razón:* "+DescrScrapTMCP1[-1]
        print(mensajeScrapTMCP1)
    else:
        mensajeScrapTMCP1=""
        print(mensajeScrapTMCP1)

#REPROCESADAS COPA1 TAPA MOVIL::::::::
    UnidadesReprocesadasEGP1 =  TapaMovil.col_values(41)
    if UnidadesReprocesadasEGP1[-1]=="Si":
        CantidadReprocesadasEGP1=  TapaMovil.col_values(42)

        mensajeUnidadesReprocesadasEGP1="*Se reprocesaron unidades - Cantidad:* "+CantidadReprocesadasEGP1[-1]+""
        print (mensajeUnidadesReprocesadasEGP1)
    else:
        mensajeUnidadesReprocesadasEGP1=""
        print(mensajeUnidadesReprocesadasEGP1)

##SELECCIONA COPA 2 HACEB TAPA MOVIL::::::::::::::::::::::

if SelectReferenciaTM[-1]=="Copa 2.0 Haceb":
    print("COPA 2 HACEB::::")
    #PAROS PROGRAMADOS TAPA MOVIL COPA 2 HACEB::
    ParoProgramadoTMCP2H = TapaMovil.col_values(43)
    if ParoProgramadoTMCP2H[-1]=="Si":
        RazonParoProgramadoTMCP2H =  TapaMovil.col_values(44)
        TiempoParoProgramadoTMCP2H =  TapaMovil.col_values(45)
        mensajeParoProgramadoTMCP2H="*Paro programado - Tiempo:* "+TiempoParoProgramadoTMCP2H[-1]+" min, *Razón:* "+RazonParoProgramadoTMCP2H[-1]
        print(mensajeParoProgramadoTMCP2H)
    else:
        mensajeParoProgramadoTMCP2H=""
        print(mensajeParoProgramadoTMCP2H)
    
    #INCIDENTES TAPA MOVIL COPA 2 HACEB::
    IncidenteTMCP2H=TapaMovil.col_values(46)
    if IncidenteTMCP2H[-1]=="Si":
        DescrIncidenteTMCP2H=TapaMovil.col_values(48)
        ValidarParoIncidenteEMCP2H=TapaMovil.col_values(49)
        mensajeIncidenteTMCP2H="*Incidente y/o accidente ambiental y/o SST: Razón:* "+DescrIncidenteTMCP2H[-1]+ " no se generó paro."
        print(mensajeIncidenteTMCP2H)
        if ValidarParoIncidenteEMCP2H[-1]=="Si":   
            TiempoIncidenteTMCP2H=TapaMovil.col_values(50)
            mensajeIncidenteTMCP2H="*Incidente y/o accidente ambiental y/o SST - Tiempo:* "+TiempoIncidenteTMCP2H[-1]+" min, *Razón:* "+DescrIncidenteTMCP2H[-1]
            print (mensajeIncidenteTMCP2H)
        else:
            #DescrIncidenteTMCP1=TapaMovil.col_values(12)
            mensajeIncidenteTMCP2H="*Incidente y/o accidente ambiental y/o SST: Razón:* "+DescrIncidenteTMCP2H[-1] + " no se generó paro."
            print (mensajeIncidenteTMCP2H)
    else:
        mensajeIncidenteTMCP2H=""
        print(mensajeIncidenteTMCP2H)
 
##SERVICIOS PUBLICOS COPA2 TAPA MOVIL::
    ServiciosPublicosTMCP2H=TapaMovil.col_values(51)
    if ServiciosPublicosTMCP2H[-1]=="Si":
        DescrServiciosPublicosTMCP2H=TapaMovil.col_values(53)
        TiempoServiciosPublicosTMCP2H=TapaMovil.col_values(52)
        mensajeServiciosPublicosTMCP2H="*Hubo afectación en las unidades del hora a hora por falta de servicios públicos: Razón:* "+DescrServiciosPublicosTMCP2H[-1]+ " -*Tiempo:* :"+TiempoServiciosPublicosTMCP2H[-1]+"min"
        print(mensajeServiciosPublicosTMCP2H)
    else:
        mensajeServiciosPublicosTMCP2H=""
        print(mensajeServiciosPublicosTMCP2H)
#POR MAQUINA COPA2 TAPA MOVIL:::
    MaquinaTMCP2H=TapaMovil.col_values(54)
    if MaquinaTMCP2H[-1]=="Si":
        DescrMaquinaTMCP2H=TapaMovil.col_values(57)
        TiempoMaquinaTMCP2H=TapaMovil.col_values(55)
        mensajeMaquinaTMCP2H="*Hubo afectación en las unidades por Maquina/ Equipo: Razón:* "+DescrMaquinaTMCP2H[-1]+ " - *Tiempo:* "+TiempoMaquinaTMCP2H[-1]+"min" 
        print(mensajeMaquinaTMCP2H)
    else:
        mensajeMaquinaTMCP2H=""
        print(mensajeMaquinaTMCP2H)

#POR MANO DE OBRA COPA2 TAPA MOVIL::::::::
    ManoDeObraTMCP2H=TapaMovil.col_values(58)
    if ManoDeObraTMCP2H[-1]=="Si":
        DescrManoDeObraTMCP2H=TapaMovil.col_values(62)
        TiempoManoDeObraTMCP2H=TapaMovil.col_values(59)
        mensajeManoDeObraTMCP2H="*Hubo afectación en las unidades por Mano De Obra: Razón:* "+DescrManoDeObraTMCP2H[-1]+ " - *Tiempo:* "+TiempoManoDeObraTMCP2H[-1]+"min" 
        print(mensajeManoDeObraTMCP2H)
    else:
        mensajeManoDeObraTMCP2H=""
        print(mensajeManoDeObraTMCP2H)

#MATERIA PRIMA COPA2 TAPA MOVIL::::

    MateriaPrimaTMCP2H=TapaMovil.col_values(63)
    if MateriaPrimaTMCP2H[-1]=="Si":
        DescrMateriaPrimaTMCP2H=TapaMovil.col_values(67)
        TiempoMateriaPrimaTMCP2H=TapaMovil.col_values(64)
        mensajeMateriaPrimaTMCP2H="*Hubo afectación en las unidades por Materia Prima: Razón:* "+DescrMateriaPrimaTMCP2H[-1]+ " - *Tiempo:* "+TiempoMateriaPrimaTMCP2H[-1]+"min" 
        print(mensajeMateriaPrimaTMCP2H)
    else:
        mensajeMateriaPrimaTMCP2H=""
        print(mensajeMateriaPrimaTMCP2H)

#POR METODO COPA2 TAPA MOVIL:::
    MetodoTMCP2H=TapaMovil.col_values(68)
    if MetodoTMCP2H[-1]=="Si":
        DescrMetodoTMCP2H=TapaMovil.col_values(71)
        TiempoMetodoTMCP2H=TapaMovil.col_values(69)
        mensajeMetodoTMCP2H="*Hubo afectación en las unidades por Método: Razón:* "+DescrMetodoTMCP2H[-1]+ "- *Tiempo:* "+TiempoMetodoTMCP2H[-1]+"min" 
        print(mensajeMetodoTMCP2H)
    else:
        mensajeMetodoTMCP2H=""
        print(mensajeMetodoTMCP2H)

#SCRAP COPA2 TAPA MOVIL::::::::::
    ScrapTMCP2H=TapaMovil.col_values(72)
    if ScrapTMCP2H[-1]=="Si":
        DescrScrapTMCP2H=TapaMovil.col_values(74)
        CantidadScrapTMCP2H=TapaMovil.col_values(75)
        mensajeScrapTMCP2H="*Se generó SCRAP: Cantidad:* "+CantidadScrapTMCP2H[-1]+" - *Razón:* "+DescrScrapTMCP2H[-1]
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
        mensajeParoProgramadoTMCP2W="*Paro programado - Tiempo:* "+TiempoParoProgramadoTMCP2W[-1]+" min, *Razón:* "+RazonParoProgramadoTMCP2W[-1]
        print(mensajeParoProgramadoTMCP2W)
    else:
        mensajeParoProgramadoTMCP2W=""
        print(mensajeParoProgramadoTMCP2W)
    
    #INCIDENTES WHIRLPOOL COPA 2 TAPA MOVIL:::::
    IncidenteTMCP2W=TapaMovil.col_values(81)
    if IncidenteTMCP2W[-1]=="Si":
        DescrIncidenteTMCP2W=TapaMovil.col_values(83)
        ValidarParoIncidenteEMCP2W=TapaMovil.col_values(84)
        mensajeIncidenteTMCP2W="*Incidente y/o accidente ambiental y/o SST: Razón:* "+DescrIncidenteTMCP2W[-1]+ " no se generó paro."
        print(mensajeIncidenteTMCP2W)
        if ValidarParoIncidenteEMCP2W[-1]=="Si":   
            TiempoIncidenteTMCP2W=TapaMovil.col_values(85)
            mensajeIncidenteTMCP2W="*Incidente y/o accidente ambiental y/o SST - Tiempo:* "+TiempoIncidenteTMCP2W[-1]+" min, *Razón:* "+DescrIncidenteTMCP2W[-1]
            print (mensajeIncidenteTMCP2W)
        else:
            #DescrIncidenteTMCP1=TapaMovil.col_values(12)
            mensajeIncidenteTMCP2W="*Incidente y/o accidente ambiental y/o SST: Razón:* "+DescrIncidenteTMCP2W[-1] + " no se generó paro."
            print (mensajeIncidenteTMCP2W)
    else:
        mensajeIncidenteTMCP2W=""
        print(mensajeIncidenteTMCP2W)

##SERVICIOS PUBLICOS COPA2 WHIRPOOL TAPA MOVIL:::
    ServiciosPublicosTMCP2W=TapaMovil.col_values(86)
    if ServiciosPublicosTMCP2W[-1]=="Si":
        DescrServiciosPublicosTMCP2W=TapaMovil.col_values(87)
        TiempoServiciosPublicosTMCP2W=TapaMovil.col_values(88)
        mensajeServiciosPublicosTMCP2W="*Hubo afectación en las unidades del hora a hora por falta de servicios públicos: Razón:* "+DescrServiciosPublicosTMCP2W[-1]+ " -*Tiempo:* :"+TiempoServiciosPublicosTMCP2W[-1]+"min"
        print(mensajeServiciosPublicosTMCP2W)
    else:
        mensajeServiciosPublicosTMCP2W=""
        print(mensajeServiciosPublicosTMCP2W)

#POR MAQUINA COPA2 WHIRLPOOL TAPA MOVIL::::::::
    MaquinaTMCP2W=TapaMovil.col_values(89)
    if MaquinaTMCP2W[-1]=="Si":
        DescrMaquinaTMCP2W=TapaMovil.col_values(92)
        TiempoMaquinaTMCP2W=TapaMovil.col_values(90)
        mensajeMaquinaTMCP2W="*Hubo afectación en las unidades por Maquina/ Equipo: Razón:* "+DescrMaquinaTMCP2W[-1]+ " - *Tiempo:* "+TiempoMaquinaTMCP2W[-1]+"min" 
        print(mensajeMaquinaTMCP2W)
    else:
        mensajeMaquinaTMCP2W=""
        print(mensajeMaquinaTMCP2W)

#POR MANO DE OBRA COPA2 WHIRLPOOL TAPA MOVIL::::::::
    ManoDeObraTMCP2W=TapaMovil.col_values(93)
    if ManoDeObraTMCP2W[-1]=="Si":
        DescrManoDeObraTMCP2W=TapaMovil.col_values(97)
        TiempoManoDeObraTMCP2W=TapaMovil.col_values(94)
        mensajeManoDeObraTMCP2W="*Hubo afectación en las unidades por Mano De Obra: Razón:* "+DescrManoDeObraTMCP2W[-1]+ " - *Tiempo:* "+TiempoManoDeObraTMCP2W[-1]+"min" 
        print(mensajeManoDeObraTMCP2W)
    else:
        mensajeManoDeObraTMCP2W=""
        print(mensajeManoDeObraTMCP2W)

#MATERIA PRIMA COPA2 WHIRPOOL TAPA MOVIL::::

    MateriaPrimaTMCP2W=TapaMovil.col_values(98)
    if MateriaPrimaTMCP2W[-1]=="Si":
        DescrMateriaPrimaTMCP2W=TapaMovil.col_values(102)
        TiempoMateriaPrimaTMCP2W=TapaMovil.col_values(99)
        mensajeMateriaPrimaTMCP2W="*Hubo afectación en las unidades por Materia Prima: Razón:* "+DescrMateriaPrimaTMCP2W[-1]+ " - *Tiempo:* "+TiempoMateriaPrimaTMCP2W[-1]+"min" 
        print(mensajeMateriaPrimaTMCP2W)
    else:
        mensajeMateriaPrimaTMCP2W=""
        print(mensajeMateriaPrimaTMCP2W)

#POR METODO COPA2 WHIRLPOOL TAPA MOVIL:::
    MetodoTMCP2W=TapaMovil.col_values(103)
    if MetodoTMCP2W[-1]=="Si":
        DescrMetodoTMCP2W=TapaMovil.col_values(106)
        TiempoMetodoTMCP2W=TapaMovil.col_values(104)
        mensajeMetodoTMCP2W="*Hubo afectación en las unidades por Método: Razón:* "+DescrMetodoTMCP2W[-1]+ "- *Tiempo:* "+TiempoMetodoTMCP2W[-1]+"min" 
        print(mensajeMetodoTMCP2W)
    else:
        mensajeMetodoTMCP2W=""
        print(mensajeMetodoTMCP2W)

#SCRAP COPA2 WHIRLPOOL TAPA MOVIL::::::::::
    ScrapTMCP2W=TapaMovil.col_values(107)
    if ScrapTMCP2W[-1]=="Si":
        DescrScrapTMCP2W=TapaMovil.col_values(109)
        CantidadScrapTMCP2W=TapaMovil.col_values(110)
        mensajeScrapTMCP2W="*Se generó SCRAP: Cantidad:* "+CantidadScrapTMCP2W[-1]+" - *Razón:* "+DescrScrapTMCP2W[-1]
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
MensajeOeeTM="OEE: "+OeeTM[-1]