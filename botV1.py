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

    #EMSAMBLE MECANISMOS:::::::::::::::
print("EMSAMBLE MECANISMOS:--------")
EmsambleMecanismos = sh.get_worksheet(0)
UnidadesFabricadasEM =  EmsambleMecanismos.col_values(5)

print(UnidadesFabricadasEM[-1])


cadencia = gc.open("Cadencia")
cadencia = cadencia.get_worksheet(0)
cadenciaList = cadencia.col_values(3)
print(cadenciaList[-1])


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
    mensajeIncidenteEM=""
    print(mensajeIncidenteEM)

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
    
OeeEM= EmsambleMecanismos.col_values(49)
OeeEmsambleMecanismos = OeeEM[-1]
print(OeeEmsambleMecanismos)


#-----------------------------------------------------------------------------------------------------------



#CONJUNTO SUSPENCIÓN:::::::::::::::
print("CONJUNTO SUSPENCIÓN:---------")

#SELECCION DE LA HOJA::
ConjuntoSuspencion = sh.get_worksheet(1)

UnidadesFabricadasCJ=  ConjuntoSuspencion.col_values(5)
print(UnidadesFabricadasCJ[-1])
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
    mensajeIncidenteCS=""
    print(mensajeIncidenteCS)

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
OeeCS= ConjuntoSuspencion.col_values(48)
OeeConjuntoSuspencion = OeeEM[-1]
print(OeeConjuntoSuspencion)

#--------------------------------------------------------------------------------------------------
#EMSAMBLE GABINETE:::::::::::::::
print("EMSAMBLE GABINETE:---------")

#SELECCION DE LA HOJA::
EmsambleGabinete = sh.get_worksheet(2)
#SELECCIONAR LA REFERENCIA::::: COPA 1 -- COPA 2 HACEB -- COPA 2 WHIRLPOOL
UnidadesFabricadasEG=  EmsambleGabinete.col_values(5)
print(UnidadesFabricadasEG[-1])
#SELECCION DE COPA 1:::::
SelectReferenciaEG = EmsambleGabinete.col_values(7)
if SelectReferenciaEG[-1]=="Copa 1.0":
    print("COPA 1::::")
    #PAROS PROGRAMADOS::
    ParoProgramadoEGCP1 = EmsambleGabinete.col_values(8)
    if ParoProgramadoEGCP1[-1]=="Si":
        RazonParoProgramadoEGCP1 =  EmsambleGabinete.col_values(9)
        TiempoParoProgramadoEGCP1 =  EmsambleGabinete.col_values(10)
        mensajeParoProgramadoEGCP1="*Paro programado - Tiempo:* "+TiempoParoProgramadoEGCP1[-1]+" min, *Razón:* "+RazonParoProgramadoEGCP1[-1]
        print(mensajeParoProgramadoEGCP1)
    else:
        mensajeParoProgramadoEGCP1=""
        print(mensajeParoProgramadoEGCP1)
    
    #INCIDENTES::
    IncidenteEGCP1=EmsambleGabinete.col_values(11)
    if IncidenteEGCP1[-1]=="Si":
        DescrIncidenteEGCP1=EmsambleGabinete.col_values(13)
        ValidarParoIncidenteEGCP1=EmsambleGabinete.col_values(14)
        mensajeIncidenteEGCP1="*Incidente y/o accidente ambiental y/o SST: Razón:* "+DescrIncidenteEGCP1[-1] + " No se generó paro"
        print(mensajeIncidenteEGCP1)
        if ValidarParoIncidenteEGCP1[-1]=="Si":   
            TiempoIncidenteEGCP1=EmsambleGabinete.col_values(15)
            mensajeIncidenteEGCP1="*Incidente y/o accidente ambiental y/o SST - Tiempo:* "+TiempoIncidenteEGCP1[-1]+" min, *Razón:* "+DescrIncidenteEGCP1[-1]
            print (mensajeIncidenteEGCP1)
        else:
            #DescrIncidenteEGCP1=EmsambleGabinete.col_values(12)
            mensajeIncidenteEGCP1="*Incidente y/o accidente ambiental y/o SST: Razón:* "+DescrIncidenteEGCP1[-1] + " No se generó paro"
            print (mensajeIncidenteEGCP1)
    else:
        mensajeIncidenteEGCP1=""
        print(mensajeIncidenteEGCP1)

##SERVICIOS PUBLICOS COPA1::
    ServiciosPublicosEGCP1=EmsambleGabinete.col_values(16)
    if ServiciosPublicosEGCP1[-1]=="Si":
        DescrServiciosPublicosEGCP1=EmsambleGabinete.col_values(18)
        TiempoServiciosPublucosEGCP1=EmsambleGabinete.col_values(17)
        mensajeServiciosPublicosEGCP1="*Hubo afectación en las unidades del hora a hora por falta de servicios públicos: Razón:* "+DescrServiciosPublicosEGCP1[-1]+ " -*Tiempo:* :"+TiempoServiciosPublucosEGCP1[-1]+"min"
        print(mensajeServiciosPublicosEGCP1)
    else:
        mensajeServiciosPublicosEGCP1=""
        print(mensajeServiciosPublicosEGCP1)
#POR MAQUINA COPA1:::
    MaquinaEGCP1=EmsambleGabinete.col_values(19)
    if MaquinaEGCP1[-1]=="Si":
        DescrMaquinaEGCP1=EmsambleGabinete.col_values(22)
        TiempoMaquinaEGCP1=EmsambleGabinete.col_values(20)
        mensajeMaquinaEGCP1="*Hubo afectación en las unidades por Maquina/ Equipo: Razón:* "+DescrMaquinaEGCP1[-1]+ " - *Tiempo:* "+TiempoMaquinaEGCP1[-1]+"min" 
        print(mensajeMaquinaEGCP1)
    else:
        mensajeMaquinaEGCP1=""
        print(mensajeMaquinaEGCP1)

#POR MANO DE OBRA COPA1::::::::
    ManoDeObraEGCP1=EmsambleGabinete.col_values(23)
    if ManoDeObraEGCP1[-1]=="Si":
        DescrManoDeObraEGCP1=EmsambleGabinete.col_values(27)
        TiempoManoDeObraEGCP1=EmsambleGabinete.col_values(24)
        mensajeManoDeObraEGCP1="*Hubo afectación en las unidades por Mano De Obra: Razón:* "+DescrManoDeObraEGCP1[-1]+ " - *Tiempo:* "+TiempoManoDeObraEGCP1[-1]+"min" 
        print(mensajeManoDeObraEGCP1)
    else:
        mensajeManoDeObraEGCP1=""
        print(mensajeManoDeObraEGCP1)

#MATERIA PRIMA COPA1::::

    MateriaPrimaEGCP1=EmsambleGabinete.col_values(28)
    if MateriaPrimaEGCP1[-1]=="Si":
        DescrMateriaPrimaEGCP1=EmsambleGabinete.col_values(32)
        TiempoMateriaPrimaEGCP1=EmsambleGabinete.col_values(29)
        mensajeMateriaPrimaEGCP1="*Hubo afectación en las unidades por Materia Prima: Razón:* "+DescrMateriaPrimaEGCP1[-1]+ " - *Tiempo:* "+TiempoMateriaPrimaEGCP1[-1]+"min" 
        print(mensajeMateriaPrimaEGCP1)
    else:
        mensajeMateriaPrimaEGCP1=""
        print(mensajeMateriaPrimaEGCP1)

#POR METODO COPA1:::
    MetodoEGCP1=EmsambleGabinete.col_values(33)
    if MetodoEGCP1[-1]=="Si":
        DescrMetodoEGCP1=EmsambleGabinete.col_values(36)
        TiempoMetodoEGCP1=EmsambleGabinete.col_values(34)
        mensajeMetodoEGCP1="*Hubo afectación en las unidades por Método: Razón:* "+DescrMetodoEGCP1[-1]+ "- *Tiempo:* "+TiempoMetodoEGCP1[-1]+"min" 
        print(mensajeMetodoEGCP1)
    else:
        mensajeMetodoEGCP1=""
        print(mensajeMetodoEGCP1)

#SCRAP COPA1::::::::::
    ScrapEGCP1=EmsambleGabinete.col_values(37)
    if ScrapEGCP1[-1]=="Si":
        DescrScrapEGCP1=EmsambleGabinete.col_values(39)
        CantidadScrapEGCP1=EmsambleGabinete.col_values(40)
        mensajeScrapEGCP1="*Se generó SCRAP: Cantidad:* "+CantidadScrapEGCP1[-1]+" - *Razón:* "+DescrScrapEGCP1[-1]
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
        mensajeParoProgramadoEGCP2H="*Paro programado - Tiempo:* "+TiempoParoProgramadoEGCP2H[-1]+" min, *Razón:* "+RazonParoProgramadoEGCP2H[-1]
        print(mensajeParoProgramadoEGCP2H)
    else:
        mensajeParoProgramadoEGCP2H=""
        print(mensajeParoProgramadoEGCP2H)
    
    #INCIDENTES::
    IncidenteEGCP2H=EmsambleGabinete.col_values(46)
    if IncidenteEGCP2H[-1]=="Si":
        DescrIncidenteEGCP2H=EmsambleGabinete.col_values(48)
        ValidarParoIncidenteEMCP1=EmsambleGabinete.col_values(49)
        mensajeIncidenteEGCP2H="*Incidente y/o accidente ambiental y/o SST: Razón:* "+DescrIncidenteEGCP2H[-1]+ " no se generó paro."
        print(mensajeIncidenteEGCP2H)
        if ValidarParoIncidenteEMCP1[-1]=="Si":   
            TiempoIncidenteEGCP2H=EmsambleGabinete.col_values(50)
            mensajeIncidenteEGCP2H="*Incidente y/o accidente ambiental y/o SST - Tiempo:* "+TiempoIncidenteEGCP2H[-1]+" min, *Razón:* "+DescrIncidenteEGCP2H[-1]
            print (mensajeIncidenteEGCP2H)
        else:
            #DescrIncidenteEGCP1=EmsambleGabinete.col_values(12)
            mensajeIncidenteEGCP2H="*Incidente y/o accidente ambiental y/o SST: Razón:* "+DescrIncidenteEGCP2H[-1] + " no se generó paro."
            print (mensajeIncidenteEGCP2H)
    else:
        mensajeIncidenteEGCP2H=""
        print(mensajeIncidenteEGCP2H)

##SERVICIOS PUBLICOS COPA2::
    ServiciosPublicosEGCP2H=EmsambleGabinete.col_values(51)
    if ServiciosPublicosEGCP2H[-1]=="Si":
        DescrServiciosPublicosEGCP2H=EmsambleGabinete.col_values(53)
        TiempoServiciosPublicosEGCP2H=EmsambleGabinete.col_values(52)
        mensajeServiciosPublicosEGCP2H="*Hubo afectación en las unidades del hora a hora por falta de servicios públicos: Razón:* "+DescrServiciosPublicosEGCP2H[-1]+ " -*Tiempo:* :"+TiempoServiciosPublicosEGCP2H[-1]+"min"
        print(mensajeServiciosPublicosEGCP2H)
    else:
        mensajeServiciosPublicosEGCP2H=""
        print(mensajeServiciosPublicosEGCP2H)
#POR MAQUINA COPA2:::
    MaquinaEGCP2H=EmsambleGabinete.col_values(54)
    if MaquinaEGCP2H[-1]=="Si":
        DescrMaquinaEGCP2H=EmsambleGabinete.col_values(57)
        TiempoMaquinaEGCP2H=EmsambleGabinete.col_values(55)
        mensajeMaquinaEGCP2H="*Hubo afectación en las unidades por Maquina/ Equipo: Razón:* "+DescrMaquinaEGCP2H[-1]+ " - *Tiempo:* "+TiempoMaquinaEGCP2H[-1]+"min" 
        print(mensajeMaquinaEGCP2H)
    else:
        mensajeMaquinaEGCP2H=""
        print(mensajeMaquinaEGCP2H)

#POR MANO DE OBRA COPA2::::::::
    ManoDeObraEGCP2H=EmsambleGabinete.col_values(58)
    if ManoDeObraEGCP2H[-1]=="Si":
        DescrManoDeObraEGCP2H=EmsambleGabinete.col_values(62)
        TiempoManoDeObraEGCP2H=EmsambleGabinete.col_values(59)
        mensajeManoDeObraEGCP2H="*Hubo afectación en las unidades por Mano De Obra: Razón:* "+DescrManoDeObraEGCP2H[-1]+ " - *Tiempo:* "+TiempoManoDeObraEGCP2H[-1]+"min" 
        print(mensajeManoDeObraEGCP2H)
    else:
        mensajeManoDeObraEGCP2H=""
        print(mensajeManoDeObraEGCP2H)

#MATERIA PRIMA COPA2::::

    MateriaPrimaEGCP2H=EmsambleGabinete.col_values(63)
    if MateriaPrimaEGCP2H[-1]=="Si":
        DescrMateriaPrimaEGCP2H=EmsambleGabinete.col_values(67)
        TiempoMateriaPrimaEGCP2H=EmsambleGabinete.col_values(64)
        mensajeMateriaPrimaEGCP2H="*Hubo afectación en las unidades por Materia Prima: Razón:* "+DescrMateriaPrimaEGCP2H[-1]+ " - *Tiempo:* "+TiempoMateriaPrimaEGCP2H[-1]+"min" 
        print(mensajeMateriaPrimaEGCP2H)
    else:
        mensajeMateriaPrimaEGCP2H=""
        print(mensajeMateriaPrimaEGCP2H)

#POR METODO COPA2:::
    MetodoEGCP2H=EmsambleGabinete.col_values(68)
    if MetodoEGCP2H[-1]=="Si":
        DescrMetodoEGCP2H=EmsambleGabinete.col_values(71)
        TiempoMetodoEGCP2H=EmsambleGabinete.col_values(69)
        mensajeMetodoEGCP2H="*Hubo afectación en las unidades por Método: Razón:* "+DescrMetodoEGCP2H[-1]+ "- *Tiempo:* "+TiempoMetodoEGCP2H[-1]+"min" 
        print(mensajeMetodoEGCP2H)
    else:
        mensajeMetodoEGCP2H=""
        print(mensajeMetodoEGCP2H)

#SCRAP COPA2::::::::::
    ScrapEGCP2H=EmsambleGabinete.col_values(72)
    if ScrapEGCP2H[-1]=="Si":
        DescrScrapEGCP2H=EmsambleGabinete.col_values(74)
        CantidadScrapEGCP2H=EmsambleGabinete.col_values(75)
        mensajeScrapEGCP2H="*Se generó SCRAP: Cantidad:* "+CantidadScrapEGCP2H[-1]+" - *Razón:* "+DescrScrapEGCP2H[-1]
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
        mensajeParoProgramadoEGCP2W="*Paro programado - Tiempo:* "+TiempoParoProgramadoEGCP2W[-1]+" min, *Razón:* "+RazonParoProgramadoEGCP2W[-1]
        print(mensajeParoProgramadoEGCP2W)
    else:
        mensajeParoProgramadoEGCP2W=""
        print(mensajeParoProgramadoEGCP2W)
    
    #INCIDENTES WHIRLPOOL COPA 2:::::
    IncidenteEGCP2W=EmsambleGabinete.col_values(81)
    if IncidenteEGCP2W[-1]=="Si":
        DescrIncidenteEGCP2W=EmsambleGabinete.col_values(83)
        ValidarParoIncidenteEMCP1=EmsambleGabinete.col_values(84)
        mensajeIncidenteEGCP2W="*Incidente y/o accidente ambiental y/o SST: Razón:* "+DescrIncidenteEGCP2W[-1]+ " no se generó paro."
        print(mensajeIncidenteEGCP2W)
        if ValidarParoIncidenteEMCP1[-1]=="Si":   
            TiempoIncidenteEGCP2W=EmsambleGabinete.col_values(85)
            mensajeIncidenteEGCP2W="*Incidente y/o accidente ambiental y/o SST - Tiempo:* "+TiempoIncidenteEGCP2W[-1]+" min, *Razón:* "+DescrIncidenteEGCP2W[-1]
            print (mensajeIncidenteEGCP2W)
        else:
            #DescrIncidenteEGCP1=EmsambleGabinete.col_values(12)
            mensajeIncidenteEGCP2W="*Incidente y/o accidente ambiental y/o SST: Razón:* "+DescrIncidenteEGCP2W[-1] + " no se generó paro."
            print (mensajeIncidenteEGCP2W)
    else:
        mensajeIncidenteEGCP2H=""
        print(mensajeIncidenteEGCP2H)

##SERVICIOS PUBLICOS COPA2::
    ServiciosPublicosEGCP2H=EmsambleGabinete.col_values(51)
    if ServiciosPublicosEGCP2H[-1]=="Si":
        DescrServiciosPublicosEGCP2H=EmsambleGabinete.col_values(53)
        TiempoServiciosPublicosEGCP2H=EmsambleGabinete.col_values(52)
        mensajeServiciosPublicosEGCP2H="*Hubo afectación en las unidades del hora a hora por falta de servicios públicos: Razón:* "+DescrServiciosPublicosEGCP2H[-1]+ " -*Tiempo:* :"+TiempoServiciosPublicosEGCP2H[-1]+"min"
        print(mensajeServiciosPublicosEGCP2H)
    else:
        mensajeServiciosPublicosEGCP2H=""
        print(mensajeServiciosPublicosEGCP2H)
#POR MAQUINA COPA2:::
    MaquinaEGCP2H=EmsambleGabinete.col_values(54)
    if MaquinaEGCP2H[-1]=="Si":
        DescrMaquinaEGCP2H=EmsambleGabinete.col_values(57)
        TiempoMaquinaEGCP2H=EmsambleGabinete.col_values(55)
        mensajeMaquinaEGCP2H="*Hubo afectación en las unidades por Maquina/ Equipo: Razón:* "+DescrMaquinaEGCP2H[-1]+ " - *Tiempo:* "+TiempoMaquinaEGCP2H[-1]+"min" 
        print(mensajeMaquinaEGCP2H)
    else:
        mensajeMaquinaEGCP2H=""
        print(mensajeMaquinaEGCP2H)

#POR MANO DE OBRA COPA2::::::::
    ManoDeObraEGCP2H=EmsambleGabinete.col_values(58)
    if ManoDeObraEGCP2H[-1]=="Si":
        DescrManoDeObraEGCP2H=EmsambleGabinete.col_values(62)
        TiempoManoDeObraEGCP2H=EmsambleGabinete.col_values(59)
        mensajeManoDeObraEGCP2H="*Hubo afectación en las unidades por Mano De Obra: Razón:* "+DescrManoDeObraEGCP2H[-1]+ " - *Tiempo:* "+TiempoManoDeObraEGCP2H[-1]+"min" 
        print(mensajeManoDeObraEGCP2H)
    else:
        mensajeManoDeObraEGCP2H=""
        print(mensajeManoDeObraEGCP2H)

#MATERIA PRIMA COPA2::::

    MateriaPrimaEGCP2H=EmsambleGabinete.col_values(63)
    if MateriaPrimaEGCP2H[-1]=="Si":
        DescrMateriaPrimaEGCP2H=EmsambleGabinete.col_values(67)
        TiempoMateriaPrimaEGCP2H=EmsambleGabinete.col_values(64)
        mensajeMateriaPrimaEGCP2H="*Hubo afectación en las unidades por Materia Prima: Razón:* "+DescrMateriaPrimaEGCP2H[-1]+ " - *Tiempo:* "+TiempoMateriaPrimaEGCP2H[-1]+"min" 
        print(mensajeMateriaPrimaEGCP2H)
    else:
        mensajeMateriaPrimaEGCP2H=""
        print(mensajeMateriaPrimaEGCP2H)

#POR METODO COPA2:::
    MetodoEGCP2H=EmsambleGabinete.col_values(68)
    if MetodoEGCP2H[-1]=="Si":
        DescrMetodoEGCP2H=EmsambleGabinete.col_values(71)
        TiempoMetodoEGCP2H=EmsambleGabinete.col_values(69)
        mensajeMetodoEGCP2H="*Hubo afectación en las unidades por Método: Razón:* "+DescrMetodoEGCP2H[-1]+ "- *Tiempo:* "+TiempoMetodoEGCP2H[-1]+"min" 
        print(mensajeMetodoEGCP2H)
    else:
        mensajeMetodoEGCP2H=""
        print(mensajeMetodoEGCP2H)

#SCRAP COPA2::::::::::
    ScrapEGCP2H=EmsambleGabinete.col_values(72)
    if ScrapEGCP2H[-1]=="Si":
        DescrScrapEGCP2H=EmsambleGabinete.col_values(74)
        CantidadScrapEGCP2H=EmsambleGabinete.col_values(75)
        mensajeScrapEGCP2H="*Se generó SCRAP: Cantidad:* "+CantidadScrapEGCP2H[-1]+" - *Razón:* "+DescrScrapEGCP2H[-1]
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
