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
        ValidarParoIncidenteEMCP1=EmsambleGabinete.col_values(14)
        
        if ValidarParoIncidenteEMCP1[-1]=="Si":   
            TiempoIncidenteEGCP1=EmsambleGabinete.col_values(15)
            mensajeIncidenteEGCP1="*Incidente y/o accidente ambiental y/o SST - Tiempo:* "+TiempoIncidenteEGCP1[-1]+" min, *Razón:* "+DescrIncidenteEGCP1[-1]
            print (mensajeIncidenteEGCP1)
    else:
        DescrIncidenteEGCP1=EmsambleGabinete.col_values(12)
        mensajeIncidenteEGCP1="*Incidente y/o accidente ambiental y/o SST - Razón:* "+DescrIncidenteEGCP1[-1] + ", No se generó paro."
        print (mensajeIncidenteEGCP1)

 #   if SelectReferencia[-1]=="Copa 2.0 Whirlpool":


    

