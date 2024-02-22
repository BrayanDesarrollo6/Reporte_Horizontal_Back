import sys
import pandas as pd
import numpy as np
import json
import xlsxwriter
from datetime import datetime
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image
import requests
from io import BytesIO
from Directories.Directory import DirectoryReporteHorizontal
from Prenomina.accessToken import funcionesGenerales

# reemplazar acentos
def normalize(s):
    replacements = (
        ("á", "a"),("é", "e"),("í", "i"),("ó", "o"),("ú", "u"),
        ("Á", "A"),("É", "E"),("Í", "I"),("Ó", "O"),("Ú", "U")
    )
    for a, b in replacements:
        s = s.replace(a, b).replace(a.upper(), b.upper())
    return s
# reemplazar caracteres
def replacement(name_company):
    name_company = name_company.replace(".", "")
    name_company = name_company.replace("-", "_")
    name_company = name_company.replace("–", "_")
    name_company = name_company.replace("—", "_")
    return name_company
# Mes name
def mesName(mes):
    mes_ = "Enero"
    if(mes == "2"):
        mes_ = "Febrero"
    elif (mes == "3"):
        mes_ = "Marzo"
    elif (mes == "4"):
        mes_ = "Abril"
    elif (mes == "5"):
        mes_ = "Mayo"
    elif (mes == "6"):
        mes_ = "Junio"
    elif (mes == "7"):
        mes_ = "Julio"
    elif (mes == "8"):
        mes_ = "Agosto"
    elif (mes == "9"):
        mes_ = "Septiembre"
    elif (mes == "10"):
        mes_ = "Octubre"
    elif (mes == "11"):
        mes_ = "Noviembre"
    elif (mes == "12"):
        mes_ = "Diciembre"
    return mes_   
def custom_xl_col_to_name(col):
    return str(col + 1)  # Sumar 1 ya que los índices de las columnas comienzan desde 1 en Excel
def hour_to_day(horas,horas_mes):
    return (horas /(horas_mes))
def horas_df(df,columna):
    return abs(df.loc[ (df['Parametrizacion reportes especiales'] == columna) & (df['Neto'] > 0), 'Horas'].sum())
def dias_df(df,columna,horasDia_):
    return hour_to_day(abs(df.loc[ (df['Parametrizacion reportes especiales'] == columna) & (df['Neto'] > 0), 'Horas'].sum()),horasDia_)
def valor_df(df,columna):
    return abs(df.loc[ (df['Parametrizacion reportes especiales'] == columna) & (df['Neto'] > 0), 'Neto'].sum())
def valor_negativo_df(df,columna):
    return df.loc[ (df['Parametrizacion reportes especiales'] == columna) & (df['Neto'] > 0), 'Neto'].sum()
def separar_texto(texto):
    if texto and texto != "nan":
        partes = texto.split(" - ")
        primera_palabra = partes[0]
        segunda_palabra = partes[1]
    else:
        primera_palabra = ""
        segunda_palabra = ""
    return primera_palabra,segunda_palabra
# Función secundaria 
def generar_dataframe_horizontal(ContratoPos, Horizontal):
    FilaAgregar = {}
    horasDia_ = 7.83
    if(str(ContratoPos.iloc[0]['Tipo de Jornada']) == "Jornada laboral medio tiempo"):
        horasDia_ = 4
    if(str(ContratoPos.iloc[0]['Tipo de Jornada']) == "Jornada laboral por días"):
        horasDia_ = 8
    if(str(ContratoPos.iloc[0]['Tipo de Jornada']) == "Jornada laboral 180 horas"):
        horasDia_ = 6
    if(str(ContratoPos.iloc[0]['Tipo de Jornada']) == "Jornada laboral 150 horas"):
        horasDia_ = 5
    if(str(ContratoPos.iloc[0]['Tipo de Jornada']) == "Destajo"):
        horasDia_ = 8
    if(str(ContratoPos.iloc[0]['Tipo de Jornada']) == "Jornada Laboral 190 Horas"):
        horasDia_ = 6.33
        
    ##Informacion general inicial
    FilaAgregar["Temporal"] = ContratoPos.iloc[0]['Temporal']
    tipoPeriodo_ = "1Q"
    if(ContratoPos.iloc[0]['Tipo de Perido'] == "2"):
        tipoPeriodo_ = "2Q"
    if(ContratoPos.iloc[0]['Tipo de Perido'] == "3"):
        tipoPeriodo_ = "M"
    mes = mesName( ContratoPos.iloc[0]['Mes'])
    FilaAgregar["Mes"] = f"{tipoPeriodo_} {mes} {ContratoPos.iloc[0]['Año']}"
    FilaAgregar["No. factura"] = "0"
    FilaAgregar["Codigo compañía"] = ContratoPos.iloc[0]['Código del cliente']
    FilaAgregar["Empresa a la que se le factura"] = "SUPPLA"
    cost_name , cost = separar_texto(str(ContratoPos.iloc[0]['Sub centro de costo - Cost center']))
    FilaAgregar["Cost center"] = cost
    FilaAgregar["Cost center name"] = cost_name
    FilaAgregar["City - población"] = str(ContratoPos.iloc[0]['Ciudad'])
    cu_name , cu = separar_texto(str(ContratoPos.iloc[0]['Proyecto - CU']))
    FilaAgregar["Cu (customer -cliente)"] = cu
    FilaAgregar["Cu name"] = cu_name
    at_name , at = separar_texto(str(ContratoPos.iloc[0]['Naturaleza Centro Costo - AT']))
    FilaAgregar["At (actividad)"] = at
    FilaAgregar["At name"] = at_name
    FilaAgregar["Cuenta"] = " "
    FilaAgregar["Cedula"] = ContratoPos.iloc[0]['Numero de Identificación']
    FilaAgregar["Nombre empleado"] = ContratoPos.iloc[0]['Nombres y Apellidos']
    FilaAgregar["Cargo"] = ContratoPos.iloc[0]['Cargo']
    FilaAgregar["Fechaingreso"] = pd.to_datetime(ContratoPos.iloc[0]['Fecha Ingreso']).date()
    FilaAgregar["Fecharetiro"] = pd.to_datetime(ContratoPos.iloc[0]['Fecha Retiro']).date()
    FilaAgregar["Estado"] = ContratoPos.iloc[0]['Estado Trabajador']
    FilaAgregar["Tipo contrato"] = ContratoPos.iloc[0]['Tipo de Contrato']
    FilaAgregar["Salario basico"] = (float(ContratoPos.iloc[0]['Salario Base']))
    FilaAgregar["Dias Salario (pagos nómina)"] = ContratoPos.iloc[0]['Días Trabajados']
    #HASTA AQUI NARANJA
    FilaAgregar["Grupo # 1\nDias ausencias justificadas con reconocimiento $ (calamidad, permisos justificados, lic, remunerada, incapacidad dia 1 y 2)"] = 0
    FilaAgregar["Grupo # 2\nDias ausencias justificadas sin cobro (vac. habiles, incapaidad del dia 3 en adelante, lic, maternidad y paternidad)"] = 0
    FilaAgregar["Grupo # 3\nDias ausencias injustificadas, sanciones, dominical, Licencia No Remunerada"] = 0
    FilaAgregar["Total días de liquidación (mes o quincena)"] = 0
    FilaAgregar["Valor Salario"] = valor_df(ContratoPos, "Valor Salario")
    FilaAgregar["Reajuste salario"] = valor_df(ContratoPos, "Reajuste salario")
    FilaAgregar["Aux.transporte"] = valor_df(ContratoPos, "Aux.transporte")
    FilaAgregar["Reajuste aux. transporte"] = valor_df(ContratoPos, "Reajuste aux. transporte")
    FilaAgregar["Auxilio conectividad"] = valor_df(ContratoPos, "Auxilio conectividad")
    FilaAgregar["Valor hora ordinaria"] = FilaAgregar["Salario basico"] / 235
    # Horas extras
    columnas_horas = [
        "HED hora extra diurna 1,25",
        "HEN hora extra nocturna 1,75",
        "HEDN Hora extra dominical nocturna 2.50",
        "HEFN Hora extra festiva nocturna 2.50",
        "HEDD hora extra diurna dominical 2.00",
        "HEFD hora extra diurna Festiva 2.00",
    ]
    for tipo in columnas_horas:
        FilaAgregar[f"Cantidad. {tipo}%"] = horas_df(ContratoPos, tipo)
        FilaAgregar[f"Valor. {tipo}%"] = valor_df(ContratoPos, tipo)
    # AGREGAR REAJUSTE
    FilaAgregar["Reajuste Cantidad horas extras"] = horas_df(ContratoPos,"Reajuste horas extras")
    FilaAgregar["Reajuste Valor horas extras"] = valor_df(ContratoPos,"Reajuste horas extras")
    # HORAS
    horasExtras_ = sum(FilaAgregar[f"Cantidad. {columna}%"] for columna in columnas_horas)
    horasExtras_ += FilaAgregar["Reajuste Cantidad horas extras"]
    # VALOR
    valorExtras_ = sum(FilaAgregar[f"Valor. {columna}%"] for columna in columnas_horas)
    valorExtras_ += FilaAgregar["Reajuste Valor horas extras"]
    # ASIGNAR TOTALES
    FilaAgregar["Total cantidad horas extras sin R.N."] = horasExtras_
    FilaAgregar["Total Valor  extras sin R.N."] = valorExtras_
    # RECARGOS
    columnas_recargos = [
        "HDD hora ordinaria dominical Diurna 1.75",
        "HFD hora ordinaria festiva Diurna 1.75",
        "RN recargo nocturna 0,35",
        "HDN Hora dominical nocturno 2.10",
        "HFN Hora festivo nocturno 2.10",
        "HDNC hora dominical nocturno Compensado 1.10",
        "HDDC hora dominical diurno Compensado 0.75",
        "Reajuste recargos"
    ]
    for tipo in columnas_recargos:
        FilaAgregar[f"Cantidad. {tipo}"] = horas_df(ContratoPos, tipo)
        FilaAgregar[f"Valor. {tipo}"] = valor_df(ContratoPos, tipo)
    horasRecargos_ = sum(FilaAgregar[f"Cantidad. {columna}"] for columna in columnas_recargos)
    valorRecargos_ = sum(FilaAgregar[f"Valor. {columna}"] for columna in columnas_recargos)
    # totales recargos
    FilaAgregar["Total cantidad. Recargo"] = horasRecargos_
    FilaAgregar["Total Valor. Recargo"] = valorRecargos_
    # TOTAL EXTRAS Y RECARGOS
    FilaAgregar["Total cantidad HE + Recargos"] = horasRecargos_ + horasExtras_
    FilaAgregar["Total valor HE + Recargos"] = valorRecargos_ + valorExtras_
    # Incapacidad empresa
    FilaAgregar["Días incapacidad enfermedad general (Días 1 y 2)"] = dias_df(ContratoPos,"incapacidad enfermedad general (Días 1 y 2)",horasDia_)
    FilaAgregar["Valor Días incapacidad enfermedad general (Días 1 y 2)"] = valor_df(ContratoPos,"incapacidad enfermedad general (Días 1 y 2)")
    # Ausencias justificadas sin cobro
    # Cálculo de días y valor para cada tipo de ausencia justificada sin cobro
    tipos_ausencias = [
        "incapacidad accidente de trabajo",
        "incapacidad enfermedad general 3 Días en adelante hasta el día 90",
        "incapacidad permanente entre día 91 al de 180 Días",
        "incapacidad permanente + de 180 Días",
        "licencia de maternidad",
        "Licencia de Paternidad",
        "Vacaciones en Disfrute Causadas"
    ]
    for tipo in tipos_ausencias:
        FilaAgregar[f"Días {tipo}"] = dias_df(ContratoPos, tipo, horasDia_)
        FilaAgregar[f"Valor {tipo}"] = valor_df(ContratoPos, tipo)
    # TOTALIZAR GRUPO 2
    # Suma total de días y valor para el grupo 2
    diasGrupo2_ = sum(FilaAgregar[f"Días {tipo}"] for tipo in tipos_ausencias)
    valorGrupo2_ = sum(FilaAgregar[f"Valor {tipo}"] for tipo in tipos_ausencias)
    FilaAgregar["Grupo # 2 Total días ausencias justificadas sin cobro"] = diasGrupo2_
    FilaAgregar["Grupo # 2 Valor ausencias justificadas sin cobro"] = valorGrupo2_
    # EMPIEZA GRUPO 1
    FilaAgregar["Días Vacaciones en Disfrute Anticipadas"] = dias_df(ContratoPos,"Vacaciones en Disfrute Anticipadas",horasDia_)
    FilaAgregar["Valor Vacaciones en Disfrute Anticipadas"] = valor_df(ContratoPos,"Vacaciones en Disfrute Anticipadas")
    # Realizar suma para licencia remunerada
    diasLicencia_ = dias_df(ContratoPos,"Permiso Personal sin Reposición de tiempo por (de 1 o 2 Días)",horasDia_)
    valorLicencia_ = valor_df(ContratoPos,"Permiso Personal sin Reposición de tiempo por (de 1 o 2 Días)")
    if(diasLicencia_ > 2):
        # valor dia
        valordiaLicencia_ = valorLicencia_ / diasLicencia_
        # Primeros dos dias
        diasLicenca2_ = 2
        valorLicencia2_ = valordiaLicencia_ * 2
        # Despues de los dos dias
        diasLicenca3_ = diasLicencia_ - 2
        valorLicencia3_ = valordiaLicencia_ * diasLicenca3_
    else:
        # Primeros dos dias
        diasLicenca2_ = diasLicencia_
        valorLicencia2_ = valorLicencia_
        # Despues de los dos dias
        diasLicenca3_ = 0
        valorLicencia3_ = 0
    FilaAgregar["Días Permiso Personal sin Reposición de tiempo por (de 1 o 2 Días)"] = diasLicenca2_
    FilaAgregar["Valor Permiso Personal sin Reposición de tiempo por (de 1 o 2 Días)"] = valorLicencia2_
    FilaAgregar["Días Permiso Justificado - Covid prevención aislamiento"] = dias_df(ContratoPos,"Permiso Justificado - Covid prevención aislamiento",horasDia_)
    FilaAgregar["Valor Permiso Justificado - Covid prevención aislamiento"] = valor_df(ContratoPos,"Permiso Justificado - Covid prevención aislamiento")
    FilaAgregar["Día Familiar"] = dias_df(ContratoPos,"Día Familiar",horasDia_)
    FilaAgregar["Valor Día Familiar"] = valor_df(ContratoPos,"Día Familiar")
    FilaAgregar["Día Compensación por desempeño"] = dias_df(ContratoPos,"	Compensación por desempeño",horasDia_) 
    FilaAgregar["Valor Día Compensación por desempeño"] = valor_df(ContratoPos,"	Compensación por desempeño")
    FilaAgregar["Días Calamidad Domestica (de 1 o 2 Días)"] = dias_df(ContratoPos,"Calamidad Domestica (de 1 o 2 Días)",horasDia_)
    FilaAgregar["Valor Calamidad Domestica (de 1 o 2 Días)"] = valor_df(ContratoPos,"Calamidad Domestica (de 1 o 2 Días)")
    FilaAgregar["Dia Libre Jurado Votaciones"] = dias_df(ContratoPos,"Dia Libre Jurado Votaciones",horasDia_)
    FilaAgregar["Valor Libre Jurado Votaciones"] = valor_df(ContratoPos,"Dia Libre Jurado Votaciones")
    FilaAgregar["Días Licencia Remunerado (mayor a 2 Días) Aprobación RH"] = diasLicenca3_
    FilaAgregar["Valor Licencia Remunerado (mayor a 2 Días) Aprobación RH"] = valorLicencia3_
    FilaAgregar["Días Licencia Remunerada - covid casos vulnerables"] = dias_df(ContratoPos,"Días Licencia Remunerada - covid casos vulnerables",horasDia_) 
    FilaAgregar["Valor Licencia Remunerada - covid casos vulnerables"] = valor_df(ContratoPos,"Días Licencia Remunerada - covid casos vulnerables")
    FilaAgregar["Días Licencia de Luto 5 Días hab (muerte de un familiar *1er grado de consanguinidad 5)"] = dias_df(ContratoPos,"Licencia de Luto 5 Días hab",horasDia_)
    FilaAgregar["Valor Licencia de Luto 5 Días hab (muerte de un familiar *1er grado de consanguinidad 5)"] = valor_df(ContratoPos,"Licencia de Luto 5 Días hab")
    FilaAgregar["Días Licencia de Matrimonio"] = dias_df(ContratoPos,"Licencia de Matrimonio",horasDia_)
    FilaAgregar["Valor Licencia de Matrimonio"] = valor_df(ContratoPos,"Licencia de Matrimonio")
    FilaAgregar["Dias Inasistencia por Alteración del Orden Publico 14708"] = dias_df(ContratoPos,"Inasistencia por Alteración",horasDia_)
    FilaAgregar["Valor Inasistencia por Alteración del Orden Publico 14708"] = valor_df(ContratoPos,"Inasistencia por Alteración")
    FilaAgregar["Dias Hospitalizacion"] = dias_df(ContratoPos,"Hospitalizacion",horasDia_)
    FilaAgregar["Valor Hospitalizacion"] = valor_df(ContratoPos,"Hospitalizacion")
    # TOTALIZAR GRUPO 1
    # DIAS
    columnas_dias_grupo1 = [
        "Dias Hospitalizacion",
        "Dias Inasistencia por Alteración del Orden Publico 14708",
        "Días Licencia de Matrimonio",
        "Días Licencia de Luto 5 Días hab (muerte de un familiar *1er grado de consanguinidad 5)",
        "Días Licencia Remunerada - covid casos vulnerables",
        "Días Licencia Remunerado (mayor a 2 Días) Aprobación RH",
        "Dia Libre Jurado Votaciones",
        "Días Calamidad Domestica (de 1 o 2 Días)",
        "Día Compensación por desempeño",
        "Día Familiar",
        "Días Permiso Justificado - Covid prevención aislamiento",
        "Días Permiso Personal sin Reposición de tiempo por (de 1 o 2 Días)",
        "Días Vacaciones en Disfrute Anticipadas"
    ]
    columnas_valor_grupo1 = [
        "Valor Hospitalizacion",
        "Valor Inasistencia por Alteración del Orden Publico 14708",
        "Valor Licencia de Matrimonio",
        "Valor Licencia de Luto 5 Días hab (muerte de un familiar *1er grado de consanguinidad 5)",
        "Valor Licencia Remunerada - covid casos vulnerables",
        "Valor Licencia Remunerado (mayor a 2 Días) Aprobación RH",
        "Valor Libre Jurado Votaciones",
        "Valor Calamidad Domestica (de 1 o 2 Días)",
        "Valor Día Compensación por desempeño",
        "Valor Día Familiar",
        "Valor Permiso Justificado - Covid prevención aislamiento",
        "Valor Permiso Personal sin Reposición de tiempo por (de 1 o 2 Días)",
        "Valor Vacaciones en Disfrute Anticipadas"
    ]
    # Inicializar sumas de días y valores
    diasGrupo1_ = 0
    valorGrupo1_ = 0

    # Calcular sumas de días y valores
    for columna in columnas_dias_grupo1:
        diasGrupo1_ += FilaAgregar[columna]
    for columna in columnas_valor_grupo1:
        valorGrupo1_ += FilaAgregar[columna]
    # TOTAL
    FilaAgregar["Grupo #1 Total días ausencias justificadas con reconocimiento"] = diasGrupo1_
    FilaAgregar["Grupo # 1 Valor total ausencias justificadas con reconocimiento"] = valorGrupo1_
    # EMPIEZA GRUPO 3
    FilaAgregar["Días Licencia No Remunerada (mayor a 2 Días) Aprobación RH"] = dias_df(ContratoPos,"Licencia No Remunerada (mayor a 2 Días) Aprobación RH",horasDia_)
    FilaAgregar["Valor Licencia No Remunerada (mayor a 2 Días) Aprobación RH (valor negativo)"] = valor_negativo_df(ContratoPos,"Licencia No Remunerada (mayor a 2 Días) Aprobación RH")
    FilaAgregar["Días Suspensión (originada Sanción)"] = dias_df(ContratoPos,"Suspensión (originada Sanción)",horasDia_)
    FilaAgregar["Valor Días Suspensión (originada Sanción) valor negativo"] = valor_negativo_df(ContratoPos,"Suspensión (originada Sanción)")
    FilaAgregar["Días Dominical por Suspensión (Inasistencia)"] = dias_df(ContratoPos,"Dominical por Suspensión (Inasistencia)",horasDia_)
    FilaAgregar["Valor Dominical por Suspensión (Inasistencia) - valor negativo"] = valor_negativo_df(ContratoPos,"Dominical por Suspensión (Inasistencia)")
    FilaAgregar["Días Inasistencia injustificada"] = dias_df(ContratoPos,"Inasistencia injustificada",horasDia_)
    FilaAgregar["Valor Inasistencia injustificada (este valor debe ser negativo)"] = valor_negativo_df(ContratoPos,"Inasistencia injustificada")
    # TOTALIZAR GRUPO 3
    diasGrupo3_ = FilaAgregar["Días Licencia No Remunerada (mayor a 2 Días) Aprobación RH"] + FilaAgregar["Días Suspensión (originada Sanción)"] + FilaAgregar["Días Dominical por Suspensión (Inasistencia)"] + FilaAgregar["Días Inasistencia injustificada"]
    valorGrupo3_ = FilaAgregar["Valor Licencia No Remunerada (mayor a 2 Días) Aprobación RH (valor negativo)"] + FilaAgregar["Valor Días Suspensión (originada Sanción) valor negativo"] + FilaAgregar["Valor Dominical por Suspensión (Inasistencia) - valor negativo"] + FilaAgregar["Valor Inasistencia injustificada (este valor debe ser negativo)"]
    FilaAgregar["Grupo # 3 Total Dias ausencias injustificadas, sanciones, dominical"] = diasGrupo3_
    FilaAgregar["Grupo # 3 Valor total ausencias injustificadas, sanciones, dominical"] = valorGrupo3_
    # OTROS DEVENGOS
    FilaAgregar["Indemnización, Bono por Retiro o Suma transaccional"] = valor_df(ContratoPos,"Indemnización, Bono por Retiro o Suma transaccional")
    FilaAgregar["Auxilio desplazamiento (parametrizado en el sistema por días laborados depende lugar trabajo y lugar residencia)"] = valor_df(ContratoPos,"Auxilio desplazamiento")
    FilaAgregar["Gastos de transporte fijo (monto mensual asignado para desempeño de sus funciones ej. Comerciales, mantenimiento, gerentes, area seguridad etc)"] = valor_df(ContratoPos,"Gastos de transporte fijo")
    FilaAgregar["Gastos de transporte ocasional (reportado por la operación quincenalmente)"] = valor_df(ContratoPos,"Gastos de transporte ocasional")
    FilaAgregar["Bonificacion no constitutiva de salario  BNCS"] = valor_df(ContratoPos,"	Bonificacion no constitutiva de salario BNCS")
    FilaAgregar["Bonificacion salarial"] = valor_df(ContratoPos,"Bonificacion salarial")
    # TOTAL NOMINA 
    otrosDevengos_ = FilaAgregar["Indemnización, Bono por Retiro o Suma transaccional"] + FilaAgregar["Auxilio desplazamiento (parametrizado en el sistema por días laborados depende lugar trabajo y lugar residencia)"] + FilaAgregar["Gastos de transporte fijo (monto mensual asignado para desempeño de sus funciones ej. Comerciales, mantenimiento, gerentes, area seguridad etc)"] + FilaAgregar["Gastos de transporte ocasional (reportado por la operación quincenalmente)"] + FilaAgregar["Bonificacion no constitutiva de salario  BNCS"] + FilaAgregar["Bonificacion salarial"]
    totalesSalarioEXRN_ = FilaAgregar["Valor Salario"] + FilaAgregar["Reajuste salario"] + FilaAgregar["Aux.transporte"] + FilaAgregar["Reajuste aux. transporte"] + FilaAgregar["Auxilio conectividad"] + FilaAgregar["Total valor HE + Recargos"] 
    totalNomina_ = totalesSalarioEXRN_ + FilaAgregar["Valor Días incapacidad enfermedad general (Días 1 y 2)"] + FilaAgregar["Grupo # 1 Valor total ausencias justificadas con reconocimiento"] + FilaAgregar["Grupo # 3 Valor total ausencias injustificadas, sanciones, dominical"] + otrosDevengos_
    FilaAgregar["Total conceptos nómina"] = totalNomina_
    FilaAgregar["% Porcentaje Arl"] = str(ContratoPos.iloc[0]['Riesgo ARL']) + "%"
    # SEGURIDAD SOCIAL
    arl_ = float(ContratoPos.iloc[0]['ARL'])
    salud_ = float(ContratoPos.iloc[0]['EPS'])
    pension_ = float(ContratoPos.iloc[0]['AFP'])
    FilaAgregar["Arl"] = (arl_)
    FilaAgregar["Salud"] = (salud_)
    FilaAgregar["Pension 12%"] = (pension_)
    seguridadSocial_ = arl_ + salud_ + pension_
    FilaAgregar["Total Seguridad Social"] = ( seguridadSocial_ )
    # PARAFISCALES
    ccf_ = float(ContratoPos.iloc[0]['CCF'])
    sena_ = float(ContratoPos.iloc[0]['SENA'])
    icbf_ =  float(ContratoPos.iloc[0]['ICBF'])
    FilaAgregar["Caja de comp 4%"] = (ccf_)
    FilaAgregar["Sena 3%"] = (sena_)
    FilaAgregar["Icbf 2%"] = (icbf_)
    parafiscales_ = ccf_ + sena_ + icbf_
    FilaAgregar["Valor parafiscales"] = ( parafiscales_ )
    # PROVISIONES
    cesantias_ = float(ContratoPos.iloc[0]['Cesantías'])
    interes_ = float(ContratoPos.iloc[0]['Interés cesantías'])
    prima_ = float(ContratoPos.iloc[0]['Prima'])
    vacaciones_ = float(ContratoPos.iloc[0]['Vacaciones tiempo'])
    # ASIGNAR PROVISIONES
    FilaAgregar["Cesantias 8,33%"] = (cesantias_)
    FilaAgregar["Int. cesantias 1%"] = (interes_)
    FilaAgregar["Prima 8,33%"] = (vacaciones_)
    FilaAgregar["Vacaciones 4.1667%"] = (vacaciones_)
    prestacionesSociales_ = cesantias_ + interes_ + prima_ + vacaciones_
    FilaAgregar["Valor prestaciones sociales"] = ( prestacionesSociales_ )
    # TOTAL NOMINA
    totalNominaSS_ = totalNomina_ + prestacionesSociales_ + parafiscales_ + seguridadSocial_
    FilaAgregar["Total nomina + Seguridad Social + parafiscales + prestaciones"] = (totalNominaSS_)
    # FACTURACION
    FilaAgregar["Administración temporales (el % que tenga cada temporal)"] = 0
    FilaAgregar["Examenes medicos  servicios"] = 0
    FilaAgregar["Menos servicio de alimentacion"] = 0
    FilaAgregar["Subtotal factura suppla"] = 0
    FilaAgregar["Iva del 19%"] = 0
    FilaAgregar["Total Neto Factura"] = 0
    FilaAgregar["Justificación (para casos puntuales que se requieran detallar)"] = ""
    FilaAgregar["Dcto my v/r pagado salario/Saldo en rojo"] = valor_negativo_df(ContratoPos,"Dcto my v/r pagado salario/Saldo en rojo")
    FilaAgregar["Saldo reconocido por el cliente"] = valor_negativo_df(ContratoPos,"Saldo reconocido por el cliente")
    FilaAgregar["DEDUCCIONES VARIAS - NC"] = valor_negativo_df(ContratoPos,"DEDUCCIONES VARIAS - NC")
    FilaAgregar["Deducicon Casino"] = valor_negativo_df(ContratoPos,"	Deducicon Casino")
    FilaAgregar["EXCEDENTE DE SEGURIDAD SOCIAL"] = float(ContratoPos.iloc[0]['Excedente seguridad social'])
    FilaAgregar["Seguridad Social Ley 1393 del 2010"] = float(ContratoPos.iloc[0]['Seguridad Social Ley 1393 del 2010'])
    # ASIGNAR DIAS DE NOVEDADES INICIALES
    FilaAgregar["Grupo # 1\nDias ausencias justificadas con reconocimiento $ (calamidad, permisos justificados, lic, remunerada, incapacidad dia 1 y 2)"] = diasGrupo1_
    FilaAgregar["Grupo # 2\nDias ausencias justificadas sin cobro (vac. habiles, incapaidad del dia 3 en adelante, lic, maternidad y paternidad)"] = diasGrupo2_
    FilaAgregar["Grupo # 3\nDias ausencias injustificadas, sanciones, dominical, Licencia No Remunerada"] = diasGrupo3_
    # ASIGNAR VALORES FINALES 
    FilaAgregar["Examenes medicos  servicios"] = 0
    FilaAgregar["Menos servicio de alimentacion"] = 0
    administracion_ = totalNominaSS_ + FilaAgregar["Examenes medicos  servicios"] + FilaAgregar["EXCEDENTE DE SEGURIDAD SOCIAL"] + FilaAgregar["Seguridad Social Ley 1393 del 2010"]
    FilaAgregar["Administración temporales (el % que tenga cada temporal)"] = administracion_ * aiu
    subtotal_ = totalNominaSS_ + FilaAgregar["Administración temporales (el % que tenga cada temporal)"] + FilaAgregar["Examenes medicos  servicios"] + FilaAgregar["EXCEDENTE DE SEGURIDAD SOCIAL"] + FilaAgregar["Seguridad Social Ley 1393 del 2010"]
    FilaAgregar["Subtotal factura suppla"] = subtotal_
    FilaAgregar["Iva del 19%"] = subtotal_ * 0.19
    FilaAgregar["Total Neto Factura"] = FilaAgregar["Subtotal factura suppla"] + FilaAgregar["Iva del 19%"]
    Horizontal = pd.concat([Horizontal,pd.DataFrame.from_records([FilaAgregar])],ignore_index=True)
    return Horizontal  
# Funcion principal
def procesar(df1, IdProceso):
    global aiu
    aiu = funcionesGenerales().getAIU("DHL SUPPLY CHAIN COLOMBIA")
    Horizontal = pd.DataFrame()
    Contrato  = df1['Numero de Contrato'].unique().tolist()
    IDPeriodo = df1['ID Periodo'].tolist()
    IDPeriodo = np.unique(IDPeriodo)
    #Obtener información para agregar al nuevo data frame
    for j in IDPeriodo:
        for i in Contrato:
            Valores = df1['Numero de Contrato'] == str(i)
            ContratoPos = df1[Valores]
            Valores2 = ContratoPos['ID Periodo'] == int(j)
            if IdProceso == "Agrupar ID proceso":
                ContratoPos = ContratoPos[Valores2]
                if ContratoPos.empty == False:
                    Horizontal = generar_dataframe_horizontal(ContratoPos, Horizontal)
            else:
                ContratoPos2 = ContratoPos[Valores2]
                ID_Procesos = ContratoPos2['Id Proceso'].unique().tolist()
                for k in ID_Procesos:
                    # Valido los contrato pos que tiene mi id proceso
                    Valores3 = ContratoPos2['Id Proceso'] == int(k)
                    # traigo de nuevo dataframe la fila o registro que contiene el proceso analizado y lo almaceno en contratopos
                    ContratoPos = ContratoPos2[Valores3]
                    if ContratoPos.empty == False:
                        Horizontal = generar_dataframe_horizontal(ContratoPos, Horizontal)
            
    # Dataframe final para obtener los indices de las primeras columnas 
    Horizontal_heads_end = pd.DataFrame()
    Horizontal_heads_end = Horizontal
    NombreDocumento = "Horizontal " + str(Horizontal.iloc[0]['Empresa a la que se le factura']) +"-"+ str(Horizontal.iloc[0]['Mes'])
    # Normalizar nombre del documento
    NombreDocumento = normalize(NombreDocumento)
    NombreDocumento = replacement(NombreDocumento)
    heads = Horizontal.columns.values
    FilaAgregar = {}
    Validador = False
    
    for k in heads:
        if(str(k).__contains__("Salario Base")):
            Validador = True
        if(Validador):
            Horizontal[k] = Horizontal[k].astype('float')
            FilaAgregar[k] = sum(Horizontal[k])
    Horizontal = pd.concat([Horizontal,pd.DataFrame.from_records([FilaAgregar])],ignore_index=True)

    
    wb = Workbook()
    ws = wb.active

    for r in dataframe_to_rows(Horizontal, index=False, header=True):
        ws.append(r)


    Horizontal = pd.DataFrame(ws.values)
    writer = pd.ExcelWriter(DirectoryReporteHorizontal+NombreDocumento+".xlsx", engine='xlsxwriter')
    Horizontal.to_excel(writer, sheet_name='PRENOMINA',index = False, header = False)
    workbook = writer.book
    worksheet = writer.sheets["PRENOMINA"]    
    # default_format = workbook.add_format({'font_name': 'Calibri Light'})
    # worksheet.set_column(0, len(heads) - 1, None, default_format) 
    MaxFilas = len(Horizontal.axes[0])
    Totales = Horizontal.loc[MaxFilas -1]
    
    #Rango de colores para el estilo
    rango_colores = {
    (1, 22): '#FFC22C',
    (23, 23): '#38C100',
    (24, 24): '#FF8F68',
    (25, 25): '#D5CECE',
    (26, 32): '#FFC22C',
    (33, 46): '#FDF2AF',
    (47, 48): '#F2FE2A',
    (49, 64): '#7FCADA',
    (65, 68): '#F2FE2A',
    (69, 70): '#38C100',
    (71, 86): '#FF8F68',
    (87, 114): '#38C100',
    (115, 124): '#CDDCDF',
    (125, 154): '#FFC22C',
    (155, 158): '#D7F5FA',
    (159, 159): '#FAE9D7'
    }
    # Escribir encabezados en la primera fila
    contador = 0
    for k in heads:
        for (inicio, fin), color in rango_colores.items():
            if inicio <= contador + 1 <= fin:
                # Configurar el formato para la primera fila (encabezado)
                header_format = workbook.add_format()
                # header_format.set_pattern(1)
                header_format.set_align('center')
                header_format.set_align('vcenter')
                header_format.set_border(1)
                header_format.set_text_wrap()
                # header_format.set_text_wrap()
                header_format.set_bg_color(color)
                worksheet.write_string(0, contador, str(k), header_format)
                break
        contador += 1
    worksheet.set_row(0, 80)
    # Ajustar automáticamente el ancho de las columnas según el contenido
    for i, heading in enumerate(heads):
        column_values = Horizontal.iloc[:, i].astype(str)
        max_value = max(column_values, key=len)
        column_width = max(len(str(heading)), len(max_value))
        worksheet.set_column(i, i, column_width + 2)
    contador = 0
    for k in Totales:
        Dato = ""
        if(str(k) != "nan"):
            Dato = str(k)
        worksheet.write_string(MaxFilas, contador,Dato)
        contador += 1
    writer.close()
    return NombreDocumento+".xlsx"

# Validar que tenga contenido los ID 
def validar_contenido_id():  
    if(estado == '1'):
        IdProceso = sys.argv[2]
        IdPeriodo1 = sys.argv[3]
        IdPeriodo = '['+IdPeriodo1+']'
    elif(estado == '2'):
        IdProceso = sys.argv[2]
        IdPeriodo1 = sys.argv[3]
        IdPeriodo2 = sys.argv[4]
        IdPeriodo = '['+IdPeriodo1+','+IdPeriodo2+']'
    else:
        IdProceso = sys.argv[2]
        IdPeriodo1 = sys.argv[3]
        IdPeriodo2 = sys.argv[4]
        IdPeriodo3 = sys.argv[5]
        IdPeriodo = '['+IdPeriodo1+','+IdPeriodo2+','+IdPeriodo3+']'
    URL = "https://creatorapp.zohopublic.com/hq5colombia/compensacionhq5/xls/Conceptos_Nomina_Desarrollo/jwhRFUOR47TqCS9AAT82eCybwgdmgeArEtKG7U8H9s3hSjTzBd3G8bPdg37PHVygvxurxwCQvMCgHRG68dOCWKTmMWaQJU2TMwnr?ID_Periodo="+IdPeriodo
    # url_ = "https://creatorapp.zohopublic.com/hq5colombia/compensacionhq5/xls/Prenomina/WWjRAOJ2MGyyNGd5BxdvwApYGzgq5A9AQ5Q6bUmpsTQvWTMJE4qE5MyKnY4KKPXneurq8RnTZ2O698AO8N2KQ7Fa7qt4hpwSet0K?Periodo=" + _idPeriodo + "&zc_FileName=PreNomina_" + _idPeriodo;
    df = pd.read_excel(URL)
    df1 = pd.DataFrame(df)
    if(df1.empty):
        print("No existe registro")
    else:
        Documento_one = procesar(df1, IdProceso)
        print(Documento_one)  
# Si llegó mas de un ID primero válide la empresa usuaria
def validar_empresa():
    if(estado == '2'):
        IdProceso = sys.argv[2]
        IdPeriodo1 = sys.argv[3]
        IdPeriodo2 = sys.argv[4]
        URL = "https://creatorapp.zohopublic.com/hq5colombia/compensacionhq5/xls/Conceptos_Nomina_Desarrollo/jwhRFUOR47TqCS9AAT82eCybwgdmgeArEtKG7U8H9s3hSjTzBd3G8bPdg37PHVygvxurxwCQvMCgHRG68dOCWKTmMWaQJU2TMwnr?ID_Periodo="+IdPeriodo1
        URL2 = "https://creatorapp.zohopublic.com/hq5colombia/compensacionhq5/xls/Conceptos_Nomina_Desarrollo/jwhRFUOR47TqCS9AAT82eCybwgdmgeArEtKG7U8H9s3hSjTzBd3G8bPdg37PHVygvxurxwCQvMCgHRG68dOCWKTmMWaQJU2TMwnr?ID_Periodo="+IdPeriodo2
        dataf = pd.read_excel(URL)
        dataf2 = pd.read_excel(URL2)
        dataf1 = pd.DataFrame(dataf)
        dataf2 = pd.DataFrame(dataf2)
        if(dataf1.empty):
            print("No existe registro")
        elif(dataf2.empty):
            print("No existe registro")
        else:
            empresa_one = dataf1['Empresa'].unique().tolist()
            empresa_two = dataf2['Empresa'].unique().tolist()
            empresa_one = str(empresa_one[0]).strip()
            empresa_two = str(empresa_two[0]).strip()
            empresa_one = normalize(empresa_one)
            empresa_one = replacement(empresa_one)
            empresa_two = normalize(empresa_two)
            empresa_two = replacement(empresa_two)
            if(empresa_one == empresa_two):
                validar_contenido_id()
            else:
                Documento_one = procesar(dataf1, IdProceso)
                Documento_two = procesar(dataf2, IdProceso)
                print(Documento_one + ',' + Documento_two)
    if(estado == '3'):
        IdProceso = sys.argv[2]
        IdPeriodo1 = sys.argv[3]
        IdPeriodo2 = sys.argv[4]
        IdPeriodo3 = sys.argv[5]
        URL = "https://creatorapp.zohopublic.com/hq5colombia/compensacionhq5/xls/Conceptos_Nomina_Desarrollo/jwhRFUOR47TqCS9AAT82eCybwgdmgeArEtKG7U8H9s3hSjTzBd3G8bPdg37PHVygvxurxwCQvMCgHRG68dOCWKTmMWaQJU2TMwnr?ID_Periodo="+IdPeriodo1
        URL2 = "https://creatorapp.zohopublic.com/hq5colombia/compensacionhq5/xls/Conceptos_Nomina_Desarrollo/jwhRFUOR47TqCS9AAT82eCybwgdmgeArEtKG7U8H9s3hSjTzBd3G8bPdg37PHVygvxurxwCQvMCgHRG68dOCWKTmMWaQJU2TMwnr?ID_Periodo="+IdPeriodo2
        URL3 = "https://creatorapp.zohopublic.com/hq5colombia/compensacionhq5/xls/Conceptos_Nomina_Desarrollo/jwhRFUOR47TqCS9AAT82eCybwgdmgeArEtKG7U8H9s3hSjTzBd3G8bPdg37PHVygvxurxwCQvMCgHRG68dOCWKTmMWaQJU2TMwnr?ID_Periodo="+IdPeriodo3
        dataf = pd.read_excel(URL)
        dataf2 = pd.read_excel(URL2)
        dataf3 = pd.read_excel(URL3)
        dataf1 = pd.DataFrame(dataf)
        dataf2 = pd.DataFrame(dataf2)
        dataf3 = pd.DataFrame(dataf3)
        if(dataf1.empty):
            print("No existe registro")
        elif(dataf2.empty):
            print("No existe registro")
        elif(dataf3.empty):
            print("No existe registro")
        else:
            empresa_one = dataf1['Empresa'].unique().tolist()
            empresa_two = dataf2['Empresa'].unique().tolist()
            empresa_three = dataf3['Empresa'].unique().tolist()
            empresa_one = str(empresa_one[0]).strip()
            empresa_two = str(empresa_two[0]).strip()
            empresa_three = str(empresa_three[0]).strip()
            empresa_one = normalize(empresa_one)
            empresa_one = replacement(empresa_one)
            empresa_two = normalize(empresa_two)
            empresa_two = replacement(empresa_two)
            empresa_three = normalize(empresa_three)
            empresa_three = replacement(empresa_three)
            if(empresa_one == empresa_two == empresa_three):
                validar_contenido_id()
            else:
                Documento_one = procesar(dataf1,IdProceso)
                Documento_two = procesar(dataf2,IdProceso)
                Documento_three = procesar(dataf3,IdProceso)
                print(Documento_one + ',' + Documento_two + ',' + Documento_three)

# validar cuantos ID llegaron 1 o más
def estados():
    if(estado == '1'):
        validar_contenido_id()
    else:
        validar_empresa()
    
# Inicio del programa 
global estado
estado = sys.argv[1]
estados()