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
    filename='named-haven-340115-ac55768dd57c.json')
gc4 = gspread.service_account(
    filename='named-haven-340115-8092bf8dd87f.json'
)


    
sh = gc3.open("Gestión Célula Hora a Hora Células")
""" sh2 = gc4.open("Gestión Célula Hora a Hora Células") """

#TAPA FIJA:::::::::::::::::
print("TAPA FIJA:::---------")

#SELECCION DE LA HOJA::
TapaFija = sh.get_worksheet(5)
#SELECCIONAR LA REFERENCIA:::::Back Panel -- COPA 2 WHIRLPOOL -- AGIPELER --- BACK PANEL --- IMPELER --- QUASAR
UnidadesFabricadasTF=TapaFija.col_values(5)
print("Unidades Fabricadas: "+UnidadesFabricadasTF[-1])
MensajeUnidadesFabricadasTF="UnidadesProducidas: "+UnidadesFabricadasTF[-1]
#SELECCION DE Agipeller:::::
SelectReferenciaTF = TapaFija.col_values(7)
if SelectReferenciaTF[-1]=="Agipeller":
    print("Agipeller::::")
    #PAROS PROGRAMADOS TAPA FIJA Agipeller::::
    ParoProgramadoTFAGI = TapaFija.col_values(8)
    if ParoProgramadoTFAGI[-1]=="Si":
        RazonParoProgramadoTFAGI =  TapaFija.col_values(9)
        TiempoParoProgramadoTFAGI =  TapaFija.col_values(10)
        mensajeParoProgramadoTFAGI="*Paro programado - Tiempo:* "+TiempoParoProgramadoTFAGI[-1]+" min, *Razón:* "+RazonParoProgramadoTFAGI[-1]
        print(mensajeParoProgramadoTFAGI)
    else:
        mensajeParoProgramadoTFAGI=""
        print(mensajeParoProgramadoTFAGI)
    
    #INCIDENTES TAPA FIJA Agipeller::
    IncidenteTFAGI=TapaFija.col_values(11)
    if IncidenteTFAGI[-1]=="Si":
        DescrIncidenteTFAGI=TapaFija.col_values(13)
        ValidarParoIncidenteTFAGI=TapaFija.col_values(14)
        mensajeIncidenteTFAGI="*Incidente y/o accidente ambiental y/o SST: Razón:* "+DescrIncidenteTFAGI[-1] + " No se generó paro"
        print(mensajeIncidenteTFAGI)
        if ValidarParoIncidenteTFAGI[-1]=="Si":   
            TiempoIncidenteTFAGI=TapaFija.col_values(15)
            mensajeIncidenteTFAGI="*Incidente y/o accidente ambiental y/o SST - Tiempo:* "+TiempoIncidenteTFAGI[-1]+" min, *Razón:* "+DescrIncidenteTFAGI[-1]
            print (mensajeIncidenteTFAGI)
        else:
            #DescrIncidenteTFAGI=TapaFija.col_values(12)
            mensajeIncidenteTFAGI="*Incidente y/o accidente ambiental y/o SST: Razón:* "+DescrIncidenteTFAGI[-1] + " No se generó paro"
            print (mensajeIncidenteTFAGI)
    else:
        mensajeIncidenteTFAGI=""
        print(mensajeIncidenteTFAGI)

##SERVICIOS PUBLICOS  TAPA FIJA Agipeller:::.
    ServiciosPublicosTFAGI=TapaFija.col_values(16)
    if ServiciosPublicosTFAGI[-1]=="Si":
        DescrServiciosPublicosTFAGI=TapaFija.col_values(18)
        TiempoServiciosPublucosTFAGI=TapaFija.col_values(17)
        mensajeServiciosPublicosTFAGI="*Hubo afectación en las unidades del hora a hora por falta de servicios públicos: Razón:* "+DescrServiciosPublicosTFAGI[-1]+ " -*Tiempo:* :"+TiempoServiciosPublucosTFAGI[-1]+"min"
        print(mensajeServiciosPublicosTFAGI)
    else:
        mensajeServiciosPublicosTFAGI=""
        print(mensajeServiciosPublicosTFAGI)
#POR MAQUINA Agipeller TAPA FIJA:::
    MaquinaTFAGI=TapaFija.col_values(19)
    if MaquinaTFAGI[-1]=="Si":
        DescrMaquinaTFAGI=TapaFija.col_values(22)
        TiempoMaquinaTFAGI=TapaFija.col_values(20)
        mensajeMaquinaTFAGI="*Hubo afectación en las unidades por Maquina/ Equipo: Razón:* "+DescrMaquinaTFAGI[-1]+ " - *Tiempo:* "+TiempoMaquinaTFAGI[-1]+"min" 
        print(mensajeMaquinaTFAGI)
    else:
        mensajeMaquinaTFAGI=""
        print(mensajeMaquinaTFAGI)

#POR MANO DE OBRA Agipeller TAPA FIJA::::::::
    ManoDeObraTFAGI=TapaFija.col_values(23)
    if ManoDeObraTFAGI[-1]=="Si":
        DescrManoDeObraTFAGI=TapaFija.col_values(27)
        TiempoManoDeObraTFAGI=TapaFija.col_values(24)
        mensajeManoDeObraTFAGI="*Hubo afectación en las unidades por Mano De Obra: Razón:* "+DescrManoDeObraTFAGI[-1]+ " - *Tiempo:* "+TiempoManoDeObraTFAGI[-1]+"min" 
        print(mensajeManoDeObraTFAGI)
    else:
        mensajeManoDeObraTFAGI=""
        print(mensajeManoDeObraTFAGI)

#MATERIA PRIMA Agipeller TAPA FIJA::::

    MateriaPrimaTFAGI=TapaFija.col_values(28)
    if MateriaPrimaTFAGI[-1]=="Si":
        DescrMateriaPrimaTFAGI=TapaFija.col_values(32)
        TiempoMateriaPrimaTFAGI=TapaFija.col_values(29)
        mensajeMateriaPrimaTFAGI="*Hubo afectación en las unidades por Materia Prima: Razón:* "+DescrMateriaPrimaTFAGI[-1]+ " - *Tiempo:* "+TiempoMateriaPrimaTFAGI[-1]+"min" 
        print(mensajeMateriaPrimaTFAGI)
    else:
        mensajeMateriaPrimaTFAGI=""
        print(mensajeMateriaPrimaTFAGI)

#POR METODO Agipeller TAPA FIJA:::
    MetodoTFAGI=TapaFija.col_values(33)
    if MetodoTFAGI[-1]=="Si":
        DescrMetodoTFAGI=TapaFija.col_values(36)
        TiempoMetodoTFAGI=TapaFija.col_values(34)
        mensajeMetodoTFAGI="*Hubo afectación en las unidades por Método: Razón:* "+DescrMetodoTFAGI[-1]+ "- *Tiempo:* "+TiempoMetodoTFAGI[-1]+"min" 
        print(mensajeMetodoTFAGI)
    else:
        mensajeMetodoTFAGI=""
        print(mensajeMetodoTFAGI)

#SCRAP Agipeller TAPA FIJA::::::::::
    ScrapTFAGI=TapaFija.col_values(37)
    if ScrapTFAGI[-1]=="Si":
        DescrScrapTFAGI=TapaFija.col_values(39)
        CantidadScrapTFAGI=TapaFija.col_values(40)
        mensajeScrapTFAGI="*Se generó SCRAP: Cantidad:* "+CantidadScrapTFAGI[-1]+" - *Razón:* "+DescrScrapTFAGI[-1]
        print(mensajeScrapTFAGI)
    else:
        mensajeScrapTFAGI=""
        print(mensajeScrapTFAGI)

#REPROCESADAS Agipeller TAPA FIJA::::::::
    UnidadesReprocesadasTFAGI =  TapaFija.col_values(41)
    if UnidadesReprocesadasTFAGI[-1]=="Si":
        CantidadReprocesadasTFAGI=  TapaFija.col_values(42)

        mensajeUnidadesReprocesadasTFAGI="*Se reprocesaron unidades - Cantidad:* "+CantidadReprocesadasTFAGI[-1]+""
        print (mensajeUnidadesReprocesadasTFAGI)
    else:
        mensajeUnidadesReprocesadasTFAGI=""
        print(mensajeUnidadesReprocesadasTFAGI)

##SELECCIONA Back Panel TAPA FIJA::::::::::::::::::::::________________________________________
#_____________________________________________________

if SelectReferenciaTF[-1]=="Back Panel":
    print("Back Panel:::")
    #PAROS PROGRAMADOS TAPA FIJA Back Panel::
    ParoProgramadoTFBP = TapaFija.col_values(43)
    if ParoProgramadoTFBP[-1]=="Si":
        RazonParoProgramadoTFBP =  TapaFija.col_values(44)
        TiempoParoProgramadoTFBP =  TapaFija.col_values(45)
        mensajeParoProgramadoTFBP="*Paro programado - Tiempo:* "+TiempoParoProgramadoTFBP[-1]+" min, *Razón:* "+RazonParoProgramadoTFBP[-1]
        print(mensajeParoProgramadoTFBP)
    else:
        mensajeParoProgramadoTFBP=""
        print(mensajeParoProgramadoTFBP)
    
    #INCIDENTES TAPA FIJA Back Panel::
    IncidenteTFBP=TapaFija.col_values(46)
    if IncidenteTFBP[-1]=="Si":
        DescrIncidenteTFBP=TapaFija.col_values(48)
        ValidarParoIncidenteEMCP2H=TapaFija.col_values(49)
        mensajeIncidenteTFBP="*Incidente y/o accidente ambiental y/o SST: Razón:* "+DescrIncidenteTFBP[-1]+ " no se generó paro."
        print(mensajeIncidenteTFBP)
        if ValidarParoIncidenteEMCP2H[-1]=="Si":   
            TiempoIncidenteTFBP=TapaFija.col_values(50)
            mensajeIncidenteTFBP="*Incidente y/o accidente ambiental y/o SST - Tiempo:* "+TiempoIncidenteTFBP[-1]+" min, *Razón:* "+DescrIncidenteTFBP[-1]
            print (mensajeIncidenteTFBP)
        else:
            #DescrIncidenteTFAGI=TapaFija.col_values(12)
            mensajeIncidenteTFBP="*Incidente y/o accidente ambiental y/o SST: Razón:* "+DescrIncidenteTFBP[-1] + " no se generó paro."
            print (mensajeIncidenteTFBP)
    else:
        mensajeIncidenteTFBP=""
        print(mensajeIncidenteTFBP)
 
##SERVICIOS PUBLICOS backpanel TAPA FIJA::
    ServiciosPublicosTFBP=TapaFija.col_values(51)
    if ServiciosPublicosTFBP[-1]=="Si":
        DescrServiciosPublicosTFBP=TapaFija.col_values(53)
        TiempoServiciosPublicosTFBP=TapaFija.col_values(52)
        mensajeServiciosPublicosTFBP="*Hubo afectación en las unidades del hora a hora por falta de servicios públicos: Razón:* "+DescrServiciosPublicosTFBP[-1]+ " -*Tiempo:* :"+TiempoServiciosPublicosTFBP[-1]+"min"
        print(mensajeServiciosPublicosTFBP)
    else:
        mensajeServiciosPublicosTFBP=""
        print(mensajeServiciosPublicosTFBP)
#POR MAQUINA backpanel TAPA FIJA:::
    MaquinaTFBP=TapaFija.col_values(54)
    if MaquinaTFBP[-1]=="Si":
        DescrMaquinaTFBP=TapaFija.col_values(57)
        TiempoMaquinaTFBP=TapaFija.col_values(55)
        mensajeMaquinaTFBP="*Hubo afectación en las unidades por Maquina/ Equipo: Razón:* "+DescrMaquinaTFBP[-1]+ " - *Tiempo:* "+TiempoMaquinaTFBP[-1]+"min" 
        print(mensajeMaquinaTFBP)
    else:
        mensajeMaquinaTFBP=""
        print(mensajeMaquinaTFBP)

#POR MANO DE OBRA backpanel TAPA FIJA::::::::
    ManoDeObraTFBP=TapaFija.col_values(58)
    if ManoDeObraTFBP[-1]=="Si":
        DescrManoDeObraTFBP=TapaFija.col_values(62)
        TiempoManoDeObraTFBP=TapaFija.col_values(59)
        mensajeManoDeObraTFBP="*Hubo afectación en las unidades por Mano De Obra: Razón:* "+DescrManoDeObraTFBP[-1]+ " - *Tiempo:* "+TiempoManoDeObraTFBP[-1]+"min" 
        print(mensajeManoDeObraTFBP)
    else:
        mensajeManoDeObraTFBP=""
        print(mensajeManoDeObraTFBP)

#MATERIA PRIMA COPA2 TAPA FIJA::::

    MateriaPrimaTFBP=TapaFija.col_values(63)
    if MateriaPrimaTFBP[-1]=="Si":
        DescrMateriaPrimaTFBP=TapaFija.col_values(67)
        TiempoMateriaPrimaTFBP=TapaFija.col_values(64)
        mensajeMateriaPrimaTFBP="*Hubo afectación en las unidades por Materia Prima: Razón:* "+DescrMateriaPrimaTFBP[-1]+ " - *Tiempo:* "+TiempoMateriaPrimaTFBP[-1]+"min" 
        print(mensajeMateriaPrimaTFBP)
    else:
        mensajeMateriaPrimaTFBP=""
        print(mensajeMateriaPrimaTFBP)

#POR METODO COPA2 TAPA FIJA:::
    MetodoTFBP=TapaFija.col_values(68)
    if MetodoTFBP[-1]=="Si":
        DescrMetodoTFBP=TapaFija.col_values(71)
        TiempoMetodoTFBP=TapaFija.col_values(69)
        mensajeMetodoTFBP="*Hubo afectación en las unidades por Método: Razón:* "+DescrMetodoTFBP[-1]+ "- *Tiempo:* "+TiempoMetodoTFBP[-1]+"min" 
        print(mensajeMetodoTFBP)
    else:
        mensajeMetodoTFBP=""
        print(mensajeMetodoTFBP)

#SCRAP COPA2 TAPA FIJA::::::::::
    ScrapTFBP=TapaFija.col_values(72)
    if ScrapTFBP[-1]=="Si":
        DescrScrapTFBP=TapaFija.col_values(74)
        CantidadScrapTFBP=TapaFija.col_values(75)
        mensajeScrapTFBP="*Se generó SCRAP: Cantidad:* "+CantidadScrapTFBP[-1]+" - *Razón:* "+DescrScrapTFBP[-1]
        print(mensajeScrapTFBP)
    else:
        mensajeScrapTFBP=""
        print(mensajeScrapTFBP)

#REPROCESADAS COPA2 TAPA FIJA::::::::
    UnidadesReprocesadasTFBP =  TapaFija.col_values(76)
    if UnidadesReprocesadasTFBP[-1]=="Si":
        CantidadReprocesadasTFBP=  TapaFija.col_values(77)

        mensajeUnidadesReprocesadasTFBP="*Se reprocesaron unidades - Cantidad:* "+CantidadReprocesadasTFBP[-1]+""
        print (mensajeUnidadesReprocesadasTFBP)
    else:
        mensajeUnidadesReprocesadasTFBP=""
        print(mensajeUnidadesReprocesadasTFBP)

##Copa 2.0 Haceb TAPA FIJA::::::::::::::
# TAPA FIJA Copa 2.0 Haceb__:::::
if SelectReferenciaTF[-1]=="Copa 2.0 Haceb":
    print("Copa 2.0 Haceb::::::")
    #PAROS PROGRAMADOS::::
    ParoProgramadoTFCP2H = TapaFija.col_values(78)
    if ParoProgramadoTFCP2H[-1]=="Si":
        RazonParoProgramadoTFCP2H =  TapaFija.col_values(79)
        TiempoParoProgramadoTFCP2H =  TapaFija.col_values(80)
        mensajeParoProgramadoTFCP2H="*Paro programado - Tiempo:* "+TiempoParoProgramadoTFCP2H[-1]+" min, *Razón:* "+RazonParoProgramadoTFCP2H[-1]
        print(mensajeParoProgramadoTFCP2H)
    else:
        mensajeParoProgramadoTFCP2H=""
        print(mensajeParoProgramadoTFCP2H)
    
    #INCIDENTES Copa 2.0 Haceb TAPA FIJA:::::
    IncidenteTFCP2H=TapaFija.col_values(81)
    if IncidenteTFCP2H[-1]=="Si":
        DescrIncidenteTFCP2H=TapaFija.col_values(83)
        ValidarParoIncidenteEMCP2W=TapaFija.col_values(84)
        mensajeIncidenteTFCP2H="*Incidente y/o accidente ambiental y/o SST: Razón:* "+DescrIncidenteTFCP2H[-1]+ " no se generó paro."
        print(mensajeIncidenteTFCP2H)
        if ValidarParoIncidenteEMCP2W[-1]=="Si":   
            TiempoIncidenteTFCP2H=TapaFija.col_values(85)
            mensajeIncidenteTFCP2H="*Incidente y/o accidente ambiental y/o SST - Tiempo:* "+TiempoIncidenteTFCP2H[-1]+" min, *Razón:* "+DescrIncidenteTFCP2H[-1]
            print (mensajeIncidenteTFCP2H)
        else:
            #DescrIncidenteTFAGI=TapaFija.col_values(12)
            mensajeIncidenteTFCP2H="*Incidente y/o accidente ambiental y/o SST: Razón:* "+DescrIncidenteTFCP2H[-1] + " no se generó paro."
            print (mensajeIncidenteTFCP2H)
    else:
        mensajeIncidenteTFCP2H=""
        print(mensajeIncidenteTFCP2H)

##SERVICIOS PUBLICOS Copa 2.0 Haceb TAPA FIJA:::
    ServiciosPublicosTFCP2H=TapaFija.col_values(86)
    if ServiciosPublicosTFCP2H[-1]=="Si":
        DescrServiciosPublicosTFCP2H=TapaFija.col_values(88)
        TiempoServiciosPublicosTFCP2H=TapaFija.col_values(87)
        mensajeServiciosPublicosTFCP2H="*Hubo afectación en las unidades del hora a hora por falta de servicios públicos: Razón:* "+DescrServiciosPublicosTFCP2H[-1]+ " -*Tiempo:* :"+TiempoServiciosPublicosTFCP2H[-1]+"min"
        print(mensajeServiciosPublicosTFCP2H)
    else:
        mensajeServiciosPublicosTFCP2H=""
        print(mensajeServiciosPublicosTFCP2H)

#POR MAQUINA Copa 2.0 Haceb TAPA FIJA::::::::
    MaquinaTFCP2H=TapaFija.col_values(89)
    if MaquinaTFCP2H[-1]=="Si":
        DescrMaquinaTFCP2H=TapaFija.col_values(92)
        TiempoMaquinaTFCP2H=TapaFija.col_values(90)
        mensajeMaquinaTFCP2H="*Hubo afectación en las unidades por Maquina/ Equipo: Razón:* "+DescrMaquinaTFCP2H[-1]+ " - *Tiempo:* "+TiempoMaquinaTFCP2H[-1]+"min" 
        print(mensajeMaquinaTFCP2H)
    else:
        mensajeMaquinaTFCP2H=""
        print(mensajeMaquinaTFCP2H)

#POR MANO DE OBRA Copa 2.0 Haceb TAPA FIJA::::::::
    ManoDeObraTFCP2H=TapaFija.col_values(93)
    if ManoDeObraTFCP2H[-1]=="Si":
        DescrManoDeObraTFCP2H=TapaFija.col_values(97)
        TiempoManoDeObraTFCP2H=TapaFija.col_values(94)
        mensajeManoDeObraTFCP2H="*Hubo afectación en las unidades por Mano De Obra: Razón:* "+DescrManoDeObraTFCP2H[-1]+ " - *Tiempo:* "+TiempoManoDeObraTFCP2H[-1]+"min" 
        print(mensajeManoDeObraTFCP2H)
    else:
        mensajeManoDeObraTFCP2H=""
        print(mensajeManoDeObraTFCP2H)

#MATERIA PRIMA Copa 2.0 Haceb TAPA FIJA::::

    MateriaPrimaTFCP2H=TapaFija.col_values(98)
    if MateriaPrimaTFCP2H[-1]=="Si":
        DescrMateriaPrimaTFCP2H=TapaFija.col_values(102)
        TiempoMateriaPrimaTFCP2H=TapaFija.col_values(99)
        mensajeMateriaPrimaTFCP2H="*Hubo afectación en las unidades por Materia Prima: Razón:* "+DescrMateriaPrimaTFCP2H[-1]+ " - *Tiempo:* "+TiempoMateriaPrimaTFCP2H[-1]+"min" 
        print(mensajeMateriaPrimaTFCP2H)
    else:
        mensajeMateriaPrimaTFCP2H=""
        print(mensajeMateriaPrimaTFCP2H)

#POR METODO Copa 2.0 Haceb TAPA FIJA:::
    MetodoTFCP2H=TapaFija.col_values(103)
    if MetodoTFCP2H[-1]=="Si":
        DescrMetodoTFCP2H=TapaFija.col_values(106)
        TiempoMetodoTFCP2H=TapaFija.col_values(104)
        mensajeMetodoTFCP2H="*Hubo afectación en las unidades por Método: Razón:* "+DescrMetodoTFCP2H[-1]+ "- *Tiempo:* "+TiempoMetodoTFCP2H[-1]+"min" 
        print(mensajeMetodoTFCP2H)
    else:
        mensajeMetodoTFCP2H=""
        print(mensajeMetodoTFCP2H)

#SCRAP Copa 2.0 Haceb TAPA FIJA::::::::::
    ScrapTFCP2H=TapaFija.col_values(107)
    if ScrapTFCP2H[-1]=="Si":
        DescrScrapTFCP2H=TapaFija.col_values(109)
        CantidadScrapTFCP2H=TapaFija.col_values(110)
        mensajeScrapTFCP2H="*Se generó SCRAP: Cantidad:* "+CantidadScrapTFCP2H[-1]+" - *Razón:* "+DescrScrapTFCP2H[-1]
        print(mensajeScrapTFCP2H)
    else:
        mensajeScrapTFCP2H=""
        print(mensajeScrapTFCP2H)

#REPROCESADAS Copa 2.0 Haceb::::::::
    UnidadesReprocesadasTFCP2H =  TapaFija.col_values(111)
    if UnidadesReprocesadasTFCP2H[-1]=="Si":
        CantidadReprocesadasTFCP2H=  TapaFija.col_values(112)

        mensajeUnidadesReprocesadasTFCP2H="*Se reprocesaron unidades - Cantidad:* "+CantidadReprocesadasTFCP2H[-1]+""
        print (mensajeUnidadesReprocesadasTFCP2H)
    else:
        mensajeUnidadesReprocesadasTFCP2H=""
        print(mensajeUnidadesReprocesadasTFCP2H)


##COPA 2WHIRLPOOL::::::::::::__________________________________________________________________

##COPA 2WHIRLPOOL Haceb TAPA FIJA::::::::::::::
# TAPA FIJA COPA 2WHIRLPOOL Haceb__:::::
if SelectReferenciaTF[-1]=="Copa 2.0 Whirlpool":
    print("COPA 2 WHIRLPOOL::::::")
    #PAROS PROGRAMADOS::::
    ParoProgramadoTFCP2W = TapaFija.col_values(113)
    if ParoProgramadoTFCP2W[-1]=="Si":
        RazonParoProgramadoTFCP2W =  TapaFija.col_values(114)
        TiempoParoProgramadoTFCP2W =  TapaFija.col_values(115)
        mensajeParoProgramadoTFCP2W="*Paro programado - Tiempo:* "+TiempoParoProgramadoTFCP2W[-1]+" min, *Razón:* "+RazonParoProgramadoTFCP2W[-1]
        print(mensajeParoProgramadoTFCP2W)
    else:
        mensajeParoProgramadoTFCP2W=""
        print(mensajeParoProgramadoTFCP2W)
    
    #INCIDENTES Copa 2.0 Whirlpool TAPA FIJA:::::
    IncidenteTFCP2W=TapaFija.col_values(116)
    if IncidenteTFCP2W[-1]=="Si":
        DescrIncidenteTFCP2W=TapaFija.col_values(118)
        ValidarParoIncidenteEMCP2W=TapaFija.col_values(119)
        mensajeIncidenteTFCP2W="*Incidente y/o accidente ambiental y/o SST: Razón:* "+DescrIncidenteTFCP2W[-1]+ " no se generó paro."
        print(mensajeIncidenteTFCP2W)
        if ValidarParoIncidenteEMCP2W[-1]=="Si":   
            TiempoIncidenteTFCP2W=TapaFija.col_values(120)
            mensajeIncidenteTFCP2W="*Incidente y/o accidente ambiental y/o SST - Tiempo:* "+TiempoIncidenteTFCP2W[-1]+" min, *Razón:* "+DescrIncidenteTFCP2W[-1]
            print (mensajeIncidenteTFCP2W)
        else:
            #DescrIncidenteTFAGI=TapaFija.col_values(12)
            mensajeIncidenteTFCP2W="*Incidente y/o accidente ambiental y/o SST: Razón:* "+DescrIncidenteTFCP2W[-1] + " no se generó paro."
            print (mensajeIncidenteTFCP2W)
    else:
        mensajeIncidenteTFCP2W=""
        print(mensajeIncidenteTFCP2W)

##SERVICIOS PUBLICOS Copa 2.0 Whirlpool TAPA FIJA:::
    ServiciosPublicosTFCP2W=TapaFija.col_values(121)
    if ServiciosPublicosTFCP2W[-1]=="Si":
        DescrServiciosPublicosTFCP2W=TapaFija.col_values(123)
        TiempoServiciosPublicosTFCP2W=TapaFija.col_values(122)
        mensajeServiciosPublicosTFCP2W="*Hubo afectación en las unidades del hora a hora por falta de servicios públicos: Razón:* "+DescrServiciosPublicosTFCP2W[-1]+ " -*Tiempo:* :"+TiempoServiciosPublicosTFCP2W[-1]+"min"
        print(mensajeServiciosPublicosTFCP2W)
    else:
        mensajeServiciosPublicosTFCP2W=""
        print(mensajeServiciosPublicosTFCP2W)

#POR MAQUINA Copa 2.0 Haceb TAPA FIJA::::::::
    MaquinaTFCP2W=TapaFija.col_values(124)
    if MaquinaTFCP2W[-1]=="Si":
        DescrMaquinaTFCP2W=TapaFija.col_values(127)
        TiempoMaquinaTFCP2W=TapaFija.col_values(125)
        mensajeMaquinaTFCP2W="*Hubo afectación en las unidades por Maquina/ Equipo: Razón:* "+DescrMaquinaTFCP2W[-1]+ " - *Tiempo:* "+TiempoMaquinaTFCP2W[-1]+"min" 
        print(mensajeMaquinaTFCP2W)
    else:
        mensajeMaquinaTFCP2W=""
        print(mensajeMaquinaTFCP2W)

#POR MANO DE OBRA Copa 2.0 Whirlpool TAPA FIJA::::::::
    ManoDeObraTFCP2W=TapaFija.col_values(128)
    if ManoDeObraTFCP2W[-1]=="Si":
        DescrManoDeObraTFCP2W=TapaFija.col_values(132)
        TiempoManoDeObraTFCP2W=TapaFija.col_values(129)
        mensajeManoDeObraTFCP2W="*Hubo afectación en las unidades por Mano De Obra: Razón:* "+DescrManoDeObraTFCP2W[-1]+ " - *Tiempo:* "+TiempoManoDeObraTFCP2W[-1]+"min" 
        print(mensajeManoDeObraTFCP2W)
    else:
        mensajeManoDeObraTFCP2W=""
        print(mensajeManoDeObraTFCP2W)

#MATERIA PRIMA Copa 2.0 Whirlpool TAPA FIJA::::

    MateriaPrimaTFCP2W=TapaFija.col_values(133)
    if MateriaPrimaTFCP2W[-1]=="Si":
        DescrMateriaPrimaTFCP2W=TapaFija.col_values(137)
        TiempoMateriaPrimaTFCP2W=TapaFija.col_values(174)
        mensajeMateriaPrimaTFCP2W="*Hubo afectación en las unidades por Materia Prima: Razón:* "+DescrMateriaPrimaTFCP2W[-1]+ " - *Tiempo:* "+TiempoMateriaPrimaTFCP2W[-1]+"min" 
        print(mensajeMateriaPrimaTFCP2W)
    else:
        mensajeMateriaPrimaTFCP2W=""
        print(mensajeMateriaPrimaTFCP2W)

#POR METODO Copa 2.0 Whirlpool TAPA FIJA:::
    MetodoTFCP2W=TapaFija.col_values(138)
    if MetodoTFCP2W[-1]=="Si":
        DescrMetodoTFCP2W=TapaFija.col_values(141)
        TiempoMetodoTFCP2W=TapaFija.col_values(139)
        mensajeMetodoTFCP2W="*Hubo afectación en las unidades por Método: Razón:* "+DescrMetodoTFCP2W[-1]+ "- *Tiempo:* "+TiempoMetodoTFCP2W[-1]+"min" 
        print(mensajeMetodoTFCP2W)
    else:
        mensajeMetodoTFCP2W=""
        print(mensajeMetodoTFCP2W)

#SCRAP Copa 2.0 Haceb TAPA FIJA::::::::::
    ScrapTFCP2W=TapaFija.col_values(142)
    if ScrapTFCP2W[-1]=="Si":
        DescrScrapTFCP2W=TapaFija.col_values(144)
        CantidadScrapTFCP2W=TapaFija.col_values(145)
        mensajeScrapTFCP2W="*Se generó SCRAP: Cantidad:* "+CantidadScrapTFCP2W[-1]+" - *Razón:* "+DescrScrapTFCP2W[-1]
        print(mensajeScrapTFCP2W)
    else:
        mensajeScrapTFCP2W=""
        print(mensajeScrapTFCP2W)

#REPROCESADAS Copa 2.0 Haceb::::::::
    UnidadesReprocesadasTFCP2W =  TapaFija.col_values(146)
    if UnidadesReprocesadasTFCP2W[-1]=="Si":
        CantidadReprocesadasTFCP2W = TapaFija.col_values(147)

        mensajeUnidadesReprocesadasTFCP2W="*Se reprocesaron unidades - Cantidad:* "+CantidadReprocesadasTFCP2W[-1]+""
        print (mensajeUnidadesReprocesadasTFCP2W)
    else:
        mensajeUnidadesReprocesadasTFCP2W=""
        print(mensajeUnidadesReprocesadasTFCP2W)


##IMPELLER TAPA FIJA:::::::::::::::::::::::::::::::::::::::::::::::::
# ___________________________________________________


if SelectReferenciaTF[-1]=="Impeller":
    print("Impeller::::::")
    #PAROS PROGRAMADOS::::
    ParoProgramadoTFIMPELLER = TapaFija.col_values(148)
    if ParoProgramadoTFIMPELLER[-1]=="Si":
        RazonParoProgramadoTFIMPELLER =  TapaFija.col_values(149)
        TiempoParoProgramadoTFIMPELLER =  TapaFija.col_values(150)
        mensajeParoProgramadoTFIMPELLER="*Paro programado - Tiempo:* "+TiempoParoProgramadoTFIMPELLER[-1]+" min, *Razón:* "+RazonParoProgramadoTFIMPELLER[-1]
        print(mensajeParoProgramadoTFIMPELLER)
    else:
        mensajeParoProgramadoTFIMPELLER=""
        print(mensajeParoProgramadoTFIMPELLER)
    
    #INCIDENTES Impeller TAPA FIJA:::::
    IncidenteTFIMPELLER=TapaFija.col_values(151)
    if IncidenteTFIMPELLER[-1]=="Si":
        DescrIncidenteTFIMPELLER=TapaFija.col_values(153)
        ValidarParoIncidenteEMCP2W=TapaFija.col_values(154)
        mensajeIncidenteTFIMPELLER="*Incidente y/o accidente ambiental y/o SST: Razón:* "+DescrIncidenteTFIMPELLER[-1]+ " no se generó paro."
        print(mensajeIncidenteTFIMPELLER)
        if ValidarParoIncidenteEMCP2W[-1]=="Si":   
            TiempoIncidenteTFIMPELLER=TapaFija.col_values(155)
            mensajeIncidenteTFIMPELLER="*Incidente y/o accidente ambiental y/o SST - Tiempo:* "+TiempoIncidenteTFIMPELLER[-1]+" min, *Razón:* "+DescrIncidenteTFIMPELLER[-1]
            print (mensajeIncidenteTFIMPELLER)
        else:
            #DescrIncidenteTFAGI=TapaFija.col_values(12)
            mensajeIncidenteTFIMPELLER="*Incidente y/o accidente ambiental y/o SST: Razón:* "+DescrIncidenteTFIMPELLER[-1] + " no se generó paro."
            print (mensajeIncidenteTFIMPELLER)
    else:
        mensajeIncidenteTFIMPELLER=""
        print(mensajeIncidenteTFIMPELLER)

##SERVICIOS PUBLICOS Copa 2.0 Whirlpool TAPA FIJA:::
    ServiciosPublicosTFIMPELLER=TapaFija.col_values(156)
    if ServiciosPublicosTFIMPELLER[-1]=="Si":
        DescrServiciosPublicosTFIMPELLER=TapaFija.col_values(158)
        TiempoServiciosPublicosTFIMPELLER=TapaFija.col_values(157)
        mensajeServiciosPublicosTFIMPELLER="*Hubo afectación en las unidades del hora a hora por falta de servicios públicos: Razón:* "+DescrServiciosPublicosTFIMPELLER[-1]+ " -*Tiempo:* :"+TiempoServiciosPublicosTFIMPELLER[-1]+"min"
        print(mensajeServiciosPublicosTFIMPELLER)
    else:
        mensajeServiciosPublicosTFIMPELLER=""
        print(mensajeServiciosPublicosTFIMPELLER)

#POR MAQUINA Impeller TAPA FIJA::::::::
    MaquinaTFIMPELLER=TapaFija.col_values(159)
    if MaquinaTFIMPELLER[-1]=="Si":
        DescrMaquinaTFIMPELLER=TapaFija.col_values(162)
        TiempoMaquinaTFIMPELLER=TapaFija.col_values(160)
        mensajeMaquinaTFIMPELLER="*Hubo afectación en las unidades por Maquina/ Equipo: Razón:* "+DescrMaquinaTFIMPELLER[-1]+ " - *Tiempo:* "+TiempoMaquinaTFIMPELLER[-1]+"min" 
        print(mensajeMaquinaTFIMPELLER)
    else:
        mensajeMaquinaTFIMPELLER=""
        print(mensajeMaquinaTFIMPELLER)

#POR MANO DE OBRA Impeller TAPA FIJA::::::::
    ManoDeObraTFIMPELLER=TapaFija.col_values(163)
    if ManoDeObraTFIMPELLER[-1]=="Si":
        DescrManoDeObraTFIMPELLER=TapaFija.col_values(167)
        TiempoManoDeObraTFIMPELLER=TapaFija.col_values(164)
        mensajeManoDeObraTFIMPELLER="*Hubo afectación en las unidades por Mano De Obra: Razón:* "+DescrManoDeObraTFIMPELLER[-1]+ " - *Tiempo:* "+TiempoManoDeObraTFIMPELLER[-1]+"min" 
        print(mensajeManoDeObraTFIMPELLER)
    else:
        mensajeManoDeObraTFIMPELLER=""
        print(mensajeManoDeObraTFIMPELLER)

#MATERIA PRIMA Impeller TAPA FIJA::::

    MateriaPrimaTFIMPELLER=TapaFija.col_values(168)
    if MateriaPrimaTFIMPELLER[-1]=="Si":
        DescrMateriaPrimaTFIMPELLER=TapaFija.col_values(172)
        TiempoMateriaPrimaTFIMPELLER=TapaFija.col_values(169)
        mensajeMateriaPrimaTFIMPELLER="*Hubo afectación en las unidades por Materia Prima: Razón:* "+DescrMateriaPrimaTFIMPELLER[-1]+ " - *Tiempo:* "+TiempoMateriaPrimaTFIMPELLER[-1]+"min" 
        print(mensajeMateriaPrimaTFIMPELLER)
    else:
        mensajeMateriaPrimaTFIMPELLER=""
        print(mensajeMateriaPrimaTFIMPELLER)

#POR METODO Impeller TAPA FIJA:::
    MetodoTFIMPELLER=TapaFija.col_values(173)
    if MetodoTFIMPELLER[-1]=="Si":
        DescrMetodoTFIMPELLER=TapaFija.col_values(176)
        TiempoMetodoTFIMPELLER=TapaFija.col_values(174)
        mensajeMetodoTFIMPELLER="*Hubo afectación en las unidades por Método: Razón:* "+DescrMetodoTFIMPELLER[-1]+ "- *Tiempo:* "+TiempoMetodoTFIMPELLER[-1]+"min" 
        print(mensajeMetodoTFIMPELLER)
    else:
        mensajeMetodoTFIMPELLER=""
        print(mensajeMetodoTFIMPELLER)

#SCRAP Impeller TAPA FIJA::::::::::
    ScrapTFIMPELLER=TapaFija.col_values(177)
    if ScrapTFIMPELLER[-1]=="Si":
        DescrScrapTFIMPELLER=TapaFija.col_values(179)
        CantidadScrapTFIMPELLER=TapaFija.col_values(180)
        mensajeScrapTFIMPELLER="*Se generó SCRAP: Cantidad:* "+CantidadScrapTFIMPELLER[-1]+" - *Razón:* "+DescrScrapTFIMPELLER[-1]
        print(mensajeScrapTFIMPELLER)
    else:
        mensajeScrapTFIMPELLER=""
        print(mensajeScrapTFIMPELLER)

#REPROCESADAS Impeller Haceb::::::::
    UnidadesReprocesadasTFIMPELLER =  TapaFija.col_values(181)
    if UnidadesReprocesadasTFIMPELLER[-1]=="Si":
        CantidadReprocesadasTFIMPELLER = TapaFija.col_values(182)

        mensajeUnidadesReprocesadasTFIMPELLER="*Se reprocesaron unidades - Cantidad:* "+CantidadReprocesadasTFIMPELLER[-1]+""
        print (mensajeUnidadesReprocesadasTFIMPELLER)
    else:
        mensajeUnidadesReprocesadasTFIMPELLER=""
        print(mensajeUnidadesReprocesadasTFIMPELLER)


##QUASAR TAPA FIJA:::::::::::::::::::::::::::::::::::::::::::::::::
# ___________________________________________________

if SelectReferenciaTF[-1]=="Quasar":
    print("Quasar::::::::::")
    #PAROS PROGRAMADOS:::::::::::
    #________
    ParoProgramadoTFQUASAR = TapaFija.col_values(183)
    if ParoProgramadoTFQUASAR[-1]=="Si":
        RazonParoProgramadoTFQUASAR =  TapaFija.col_values(184)
        TiempoParoProgramadoTFQUASAR =  TapaFija.col_values(185)
        mensajeParoProgramadoTFQUASAR="*Paro programado - Tiempo:* "+TiempoParoProgramadoTFQUASAR[-1]+" min, *Razón:* "+RazonParoProgramadoTFQUASAR[-1]
        print(mensajeParoProgramadoTFQUASAR)
    else:
        mensajeParoProgramadoTFQUASAR=""
        print(mensajeParoProgramadoTFQUASAR)
    
    #INCIDENTES Quasar TAPA FIJA:::::
    IncidenteTFQUASAR=TapaFija.col_values(186)
    if IncidenteTFQUASAR[-1]=="Si":
        DescrIncidenteTFQUASAR=TapaFija.col_values(188)
        ValidarParoIncidenteEMCP2W=TapaFija.col_values(189)
        mensajeIncidenteTFQUASAR="*Incidente y/o accidente ambiental y/o SST: Razón:* "+DescrIncidenteTFQUASAR[-1]+ " no se generó paro."
        print(mensajeIncidenteTFQUASAR)
        if ValidarParoIncidenteEMCP2W[-1]=="Si":   
            TiempoIncidenteTFQUASAR=TapaFija.col_values(190)
            mensajeIncidenteTFQUASAR="*Incidente y/o accidente ambiental y/o SST - Tiempo:* "+TiempoIncidenteTFQUASAR[-1]+" min, *Razón:* "+DescrIncidenteTFQUASAR[-1]
            print (mensajeIncidenteTFQUASAR)
        else:
            #DescrIncidenteTFAGI=TapaFija.col_values(12)
            mensajeIncidenteTFQUASAR="*Incidente y/o accidente ambiental y/o SST: Razón:* "+DescrIncidenteTFQUASAR[-1] + " no se generó paro."
            print (mensajeIncidenteTFQUASAR)
    else:
        mensajeIncidenteTFQUASAR=""
        print(mensajeIncidenteTFQUASAR)

##SERVICIOS PUBLICOS Quasar TAPA FIJA:::
    ServiciosPublicosTFQUASAR=TapaFija.col_values(191)
    if ServiciosPublicosTFQUASAR[-1]=="Si":
        DescrServiciosPublicosTFQUASAR=TapaFija.col_values(193)
        TiempoServiciosPublicosTFQUASAR=TapaFija.col_values(192)
        mensajeServiciosPublicosTFQUASAR="*Hubo afectación en las unidades del hora a hora por falta de servicios públicos: Razón:* "+DescrServiciosPublicosTFQUASAR[-1]+ " -*Tiempo:* :"+TiempoServiciosPublicosTFQUASAR[-1]+"min"
        print(mensajeServiciosPublicosTFQUASAR)
    else:
        mensajeServiciosPublicosTFQUASAR=""
        print(mensajeServiciosPublicosTFQUASAR)

#POR MAQUINA Quasar TAPA FIJA::::::::
    MaquinaTFQUASAR=TapaFija.col_values(194)
    if MaquinaTFQUASAR[-1]=="Si":
        DescrMaquinaTFQUASAR=TapaFija.col_values(197)
        TiempoMaquinaTFQUASAR=TapaFija.col_values(195)
        mensajeMaquinaTFQUASAR="*Hubo afectación en las unidades por Maquina/ Equipo: Razón:* "+DescrMaquinaTFQUASAR[-1]+ " - *Tiempo:* "+TiempoMaquinaTFQUASAR[-1]+"min" 
        print(mensajeMaquinaTFQUASAR)
    else:
        mensajeMaquinaTFQUASAR=""
        print(mensajeMaquinaTFQUASAR)

#POR MANO DE OBRA Quasar TAPA FIJA::::::::
    ManoDeObraTFQUASAR=TapaFija.col_values(198)
    if ManoDeObraTFQUASAR[-1]=="Si":
        DescrManoDeObraTFQUASAR=TapaFija.col_values(202)
        TiempoManoDeObraTFQUASAR=TapaFija.col_values(199)
        mensajeManoDeObraTFQUASAR="*Hubo afectación en las unidades por Mano De Obra: Razón:* "+DescrManoDeObraTFQUASAR[-1]+ " - *Tiempo:* "+TiempoManoDeObraTFQUASAR[-1]+"min" 
        print(mensajeManoDeObraTFQUASAR)
    else:
        mensajeManoDeObraTFQUASAR=""
        print(mensajeManoDeObraTFQUASAR)

#MATERIA PRIMA Quasar TAPA FIJA::::

    MateriaPrimaTFQUASAR=TapaFija.col_values(203)
    if MateriaPrimaTFQUASAR[-1]=="Si":
        DescrMateriaPrimaTFQUASAR=TapaFija.col_values(207)
        TiempoMateriaPrimaTFQUASAR=TapaFija.col_values(204)
        mensajeMateriaPrimaTFQUASAR="*Hubo afectación en las unidades por Materia Prima: Razón:* "+DescrMateriaPrimaTFQUASAR[-1]+ " - *Tiempo:* "+TiempoMateriaPrimaTFQUASAR[-1]+"min" 
        print(mensajeMateriaPrimaTFQUASAR)
    else:
        mensajeMateriaPrimaTFQUASAR=""
        print(mensajeMateriaPrimaTFQUASAR)

#POR METODO Quasar TAPA FIJA:::
    MetodoTFQUASAR=TapaFija.col_values(208)
    if MetodoTFQUASAR[-1]=="Si":
        DescrMetodoTFQUASAR=TapaFija.col_values(211)
        TiempoMetodoTFQUASAR=TapaFija.col_values(209)
        mensajeMetodoTFQUASAR="*Hubo afectación en las unidades por Método: Razón:* "+DescrMetodoTFQUASAR[-1]+ "- *Tiempo:* "+TiempoMetodoTFQUASAR[-1]+"min" 
        print(mensajeMetodoTFQUASAR)
    else:
        mensajeMetodoTFQUASAR=""
        print(mensajeMetodoTFQUASAR)

#SCRAP Quasar TAPA FIJA::::::::::
    ScrapTFQUASAR=TapaFija.col_values(212)
    if ScrapTFQUASAR[-1]=="Si":
        DescrScrapTFQUASAR=TapaFija.col_values(214)
        CantidadScrapTFQUASAR=TapaFija.col_values(215)
        mensajeScrapTFQUASAR="*Se generó SCRAP: Cantidad:* "+CantidadScrapTFQUASAR[-1]+" - *Razón:* "+DescrScrapTFQUASAR[-1]
        print(mensajeScrapTFQUASAR)
    else:
        mensajeScrapTFQUASAR=""
        print(mensajeScrapTFQUASAR)

#REPROCESADAS Quasar Haceb::::::::
    UnidadesReprocesadasTFQUASAR =  TapaFija.col_values(216)
    if UnidadesReprocesadasTFQUASAR[-1]=="Si":
        CantidadReprocesadasTFQUASAR = TapaFija.col_values(217)

        mensajeUnidadesReprocesadasTFQUASAR="*Se reprocesaron unidades - Cantidad:* "+CantidadReprocesadasTFQUASAR[-1]+""
        print (mensajeUnidadesReprocesadasTFQUASAR)
    else:
        mensajeUnidadesReprocesadasTFQUASAR=""
        print(mensajeUnidadesReprocesadasTFQUASAR)



OeeTF= TapaFija.col_values(223)
OeeTapaFija = OeeTF[-1]
print("Porcentaje OEE: "+ OeeTapaFija)
MensajeOeeTF="OEE: "+ OeeTF[-1]