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
        ValidarParoIncidenteEMCP2H=EmsambleGabinete.col_values(49)
        mensajeIncidenteEGCP2H="*Incidente y/o accidente ambiental y/o SST: Razón:* "+DescrIncidenteEGCP2H[-1]+ " no se generó paro."
        print(mensajeIncidenteEGCP2H)
        if ValidarParoIncidenteEMCP2H[-1]=="Si":   
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
        ValidarParoIncidenteEMCP2W=EmsambleGabinete.col_values(84)
        mensajeIncidenteEGCP2W="*Incidente y/o accidente ambiental y/o SST: Razón:* "+DescrIncidenteEGCP2W[-1]+ " no se generó paro."
        print(mensajeIncidenteEGCP2W)
        if ValidarParoIncidenteEMCP2W[-1]=="Si":   
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

##SERVICIOS PUBLICOS COPA2 WHIRPOOL:::
    ServiciosPublicosEGCP2W=EmsambleGabinete.col_values(86)
    if ServiciosPublicosEGCP2W[-1]=="Si":
        DescrServiciosPublicosEGCP2W=EmsambleGabinete.col_values(87)
        TiempoServiciosPublicosEGCP2W=EmsambleGabinete.col_values(88)
        mensajeServiciosPublicosEGCP2W="*Hubo afectación en las unidades del hora a hora por falta de servicios públicos: Razón:* "+DescrServiciosPublicosEGCP2W[-1]+ " -*Tiempo:* :"+TiempoServiciosPublicosEGCP2W[-1]+"min"
        print(mensajeServiciosPublicosEGCP2W)
    else:
        mensajeServiciosPublicosEGCP2W=""
        print(mensajeServiciosPublicosEGCP2W)

#POR MAQUINA COPA2 WHIRLPOOL::::::::
    MaquinaEGCP2W=EmsambleGabinete.col_values(89)
    if MaquinaEGCP2W[-1]=="Si":
        DescrMaquinaEGCP2W=EmsambleGabinete.col_values(92)
        TiempoMaquinaEGCP2W=EmsambleGabinete.col_values(90)
        mensajeMaquinaEGCP2W="*Hubo afectación en las unidades por Maquina/ Equipo: Razón:* "+DescrMaquinaEGCP2W[-1]+ " - *Tiempo:* "+TiempoMaquinaEGCP2W[-1]+"min" 
        print(mensajeMaquinaEGCP2W)
    else:
        mensajeMaquinaEGCP2W=""
        print(mensajeMaquinaEGCP2W)

#POR MANO DE OBRA COPA2 WHIRLPOOL::::::::
    ManoDeObraEGCP2W=EmsambleGabinete.col_values(93)
    if ManoDeObraEGCP2W[-1]=="Si":
        DescrManoDeObraEGCP2W=EmsambleGabinete.col_values(97)
        TiempoManoDeObraEGCP2W=EmsambleGabinete.col_values(94)
        mensajeManoDeObraEGCP2W="*Hubo afectación en las unidades por Mano De Obra: Razón:* "+DescrManoDeObraEGCP2W[-1]+ " - *Tiempo:* "+TiempoManoDeObraEGCP2W[-1]+"min" 
        print(mensajeManoDeObraEGCP2W)
    else:
        mensajeManoDeObraEGCP2W=""
        print(mensajeManoDeObraEGCP2W)

#MATERIA PRIMA COPA2 WHIRPOOL::::

    MateriaPrimaEGCP2W=EmsambleGabinete.col_values(98)
    if MateriaPrimaEGCP2W[-1]=="Si":
        DescrMateriaPrimaEGCP2W=EmsambleGabinete.col_values(102)
        TiempoMateriaPrimaEGCP2W=EmsambleGabinete.col_values(99)
        mensajeMateriaPrimaEGCP2W="*Hubo afectación en las unidades por Materia Prima: Razón:* "+DescrMateriaPrimaEGCP2W[-1]+ " - *Tiempo:* "+TiempoMateriaPrimaEGCP2W[-1]+"min" 
        print(mensajeMateriaPrimaEGCP2W)
    else:
        mensajeMateriaPrimaEGCP2W=""
        print(mensajeMateriaPrimaEGCP2W)

#POR METODO COPA2 WHIRLPOOL:::
    MetodoEGCP2W=EmsambleGabinete.col_values(103)
    if MetodoEGCP2W[-1]=="Si":
        DescrMetodoEGCP2W=EmsambleGabinete.col_values(106)
        TiempoMetodoEGCP2W=EmsambleGabinete.col_values(104)
        mensajeMetodoEGCP2W="*Hubo afectación en las unidades por Método: Razón:* "+DescrMetodoEGCP2W[-1]+ "- *Tiempo:* "+TiempoMetodoEGCP2W[-1]+"min" 
        print(mensajeMetodoEGCP2W)
    else:
        mensajeMetodoEGCP2W=""
        print(mensajeMetodoEGCP2W)

#SCRAP COPA2 WHIRLPOOL::::::::::
    ScrapEGCP2W=EmsambleGabinete.col_values(107)
    if ScrapEGCP2W[-1]=="Si":
        DescrScrapEGCP2W=EmsambleGabinete.col_values(109)
        CantidadScrapEGCP2W=EmsambleGabinete.col_values(110)
        mensajeScrapEGCP2W="*Se generó SCRAP: Cantidad:* "+CantidadScrapEGCP2W[-1]+" - *Razón:* "+DescrScrapEGCP2W[-1]
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
print(OeeEmsableGabinete)


#TESTEO FINAL:::::::::::::::
print("TESTEO FINAL::---------")

#SELECCION DE LA HOJA::
TesteoFinal = sh.get_worksheet(3)
#SELECCIONAR LA REFERENCIA::::: COPA 1 -- COPA 2 HACEB -- COPA 2 WHIRLPOOL
UnidadesFabricadasTF=  TesteoFinal.col_values(5)
print(UnidadesFabricadasTF[-1])
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
print(UnidadesFabricadasTM[-1])
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


OeeTM= TapaMovil.col_values(118)
OeeTapaMovil = OeeTM[-1]
print("Porcentaje OEE: "+ OeeTapaMovil)



#TAPA FIJA:::::::::::::::::
print("TAPA FIJA:::---------")

#SELECCION DE LA HOJA::
TapaFija = sh.get_worksheet(5)
#SELECCIONAR LA REFERENCIA:::::Back Panel -- COPA 2 WHIRLPOOL -- AGIPELER --- BACK PANEL --- IMPELER --- QUASAR
UnidadesFabricadasTF=  TapaFija.col_values(5)
print(UnidadesFabricadasTF[-1])
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

