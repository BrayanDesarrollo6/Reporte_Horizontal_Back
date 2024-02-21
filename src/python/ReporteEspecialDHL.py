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
# Función secundaria 
def generar_dataframe_horizontal(ContratoPos, ConceptosDev, ConceptosDed, Horizontal):
    FilaAgregar = {}
    
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
    FilaAgregar["Cost center"] = " "
    FilaAgregar["Cost center name"] = " "
    FilaAgregar["City -población"] = " "
    FilaAgregar["Cu (customer -cliente)"] = " "
    FilaAgregar["Cu name"] = " "
    FilaAgregar["At (actividad)"] = " "
    FilaAgregar["At name"] = " "
    FilaAgregar["Cuenta"] = " "
    FilaAgregar["Cedula"] = ContratoPos.iloc[0]['Numero de Identificación']
    FilaAgregar["Nombre empleado"] = ContratoPos.iloc[0]['Nombres y Apellidos']
    FilaAgregar["Cargo"] = ContratoPos.iloc[0]['Cargo']
    FilaAgregar["Fechaingreso"] = pd.to_datetime(ContratoPos.iloc[0]['Fecha Ingreso']).date()
    FilaAgregar["Fecharetiro"] = pd.to_datetime(ContratoPos.iloc[0]['Fecha Retiro']).date()
    FilaAgregar["Estado"] = ContratoPos.iloc[0]['Estado Trabajador']
    FilaAgregar["Tipo contrato"] = ContratoPos.iloc[0]['Tipo de Contrato']
    FilaAgregar["Salario basico"] = ContratoPos.iloc[0]['Salario Base']
    FilaAgregar["Dias Salario (pagos nómina)"] = ContratoPos.iloc[0]['Días Trabajados']
    #HASTA AQUI NARANJA
    # abs(ContratoPos.loc[ (ContratoPos['Concepto'] == "DevOtroConceptoNS")  & (dfAcumuladoEmpleado_['Neto'] > 0), 'Neto'].sum())
    FilaAgregar["Grupo # 1 Dias ausencias justificadas con reconocimiento $ (calamidad, permisos justificados, lic, remunerada, incapacidad dia 1 y 2)"] = ContratoPos.iloc[0]['Días Trabajados']
    FilaAgregar["Grupo # 2 Dias ausencias justificadas sin cobro (vac. habiles, incapaidad del dia 3 en adelante, lic, maternidad y paternidad)"] = ContratoPos.iloc[0]['Días Trabajados']
    FilaAgregar["Grupo # 3 Dias ausencias injustificadas, sanciones, dominical, Licencia No Remunerada"] = ContratoPos.iloc[0]['Días Trabajados']
    FilaAgregar["Total días de liquidación (mes o quincena)"] = ContratoPos.iloc[0]['Días Trabajados']
    FilaAgregar["Valor Salario"] = ContratoPos.iloc[0]['Días Trabajados']
    FilaAgregar["Reajuste salario"] = ContratoPos.iloc[0]['Días Trabajados']
    FilaAgregar["Aux.transporte"] = ContratoPos.iloc[0]['Días Trabajados']
    FilaAgregar["Reajuste aux. transporte"] = ContratoPos.iloc[0]['Días Trabajados']
    FilaAgregar["Auxilio conectividad"] = ContratoPos.iloc[0]['Días Trabajados']
    FilaAgregar["Valor hora ordinaria"] = ContratoPos.iloc[0]['Días Trabajados']
    FilaAgregar["Cantidad. HED hora extra diurna 1,25"] = ContratoPos.iloc[0]['Días Trabajados']
    FilaAgregar["Valor. HED hora extra diurna 1,25"] = ContratoPos.iloc[0]['Días Trabajados']
    FilaAgregar["Cantidad. HEN hora extra nocturna 1,75"] = ContratoPos.iloc[0]['Días Trabajados']
    FilaAgregar["Valor. HEN hora extra nocturna 1,75"] = ContratoPos.iloc[0]['Días Trabajados']
    FilaAgregar["Cantidad. HEDN Hora extra dominical nocturna 2.50%"] = ContratoPos.iloc[0]['Días Trabajados']
    FilaAgregar["Valor HEDN Hora extra dominical nocturna 2.50%"] = ContratoPos.iloc[0]['Días Trabajados']
    FilaAgregar["Cantidad. HEFN Hora extra festiva nocturna 2.50%"] = ContratoPos.iloc[0]['Días Trabajados']
    FilaAgregar["Valor. HEFN Hora extra festiva nocturna 2.50%"] = ContratoPos.iloc[0]['Días Trabajados']
    FilaAgregar["Cantidad. HEDD hora extra diurna dominical 2.00%"] = ""
    FilaAgregar["Valor. HEDD hora extra diurna dominical 2.00%"] = ""
    FilaAgregar["Cantidad. HEFD hora extra diurna Festiva 2.00%"] = ""
    FilaAgregar["Valor. HEFD hora extra diurna Festiva 2.00%"] = ""
    FilaAgregar["Reajuste Cantidad horas extras"] = ""
    FilaAgregar["Reajuste Valor horas extras"] = ""
    FilaAgregar["Total cantidad horas extras sin R.N."] = ""
    FilaAgregar["Total Valor  extras sin R.N."] = ""
    FilaAgregar["Cantidad  HDD hora ordinaria dominical Diurna 1.75"] = ""
    FilaAgregar["Valor HDD hora ordinaria dominical Diurna 1.75"] = ""
    FilaAgregar["Cantidad  HFD hora ordinaria festiva Diurna  1.75"] = ""
    FilaAgregar["Valor  HFD hora ordinaria festiva Diurna  1.75"] = ""
    FilaAgregar["Cantidad. RN recargo nocturna 0,35%"] = ""
    FilaAgregar["Valor. RN recargo nocturna 0,35%"] = ""
    FilaAgregar["Cantidad. HDN Hora dominical nocturno 2.10%"] = ""
    FilaAgregar["Valor. HDN Hora dominical nocturno 2.10%"] = ""
    FilaAgregar["Cantidad. HFN Hora festivo nocturno 2.10%"] = ""
    FilaAgregar["Valor. HDN Hora festivo nocturno 2.10%"] = ""
    FilaAgregar["Cantidad. HDNC hora dominical nocturno Compensado 1.10%"] = ""
    FilaAgregar["Valor. HDNC hora dominical nocturno Compensado 1.10%"] = ""
    FilaAgregar["Cantidad. HDNC hora dominical diurno Compensado 0.75%"] = ""
    FilaAgregar["Valor. HDNC hora dominical nocturno Compensado 0.75%"] = ""
    FilaAgregar["Cantidad. Reajuste recargos"] = ""
    FilaAgregar["Valor. Reajuste recargos"] = ""
    FilaAgregar["Total cantidad. Recargo"] = ""
    FilaAgregar["Total Valor  Recargo"] = ""
    FilaAgregar["Total cantidad HE + Recargos"] = ""
    FilaAgregar["Total valor HE + Recargos"] = ""
    FilaAgregar["Días incapacidad enfermedad general (Días 1 y 2)"] = ""
    FilaAgregar["Valor Días incapacidad enfermedad general (Días 1 y 2)"] = ""
    FilaAgregar["Días incapacidad accidente de trabajo"] = ""
    FilaAgregar["Valor incapacidad accidente de trabajo"] = ""
    FilaAgregar["Días incapacidad enfermedad general 3 Días en adelante hasta el día 90"] = ""
    FilaAgregar["Valor incapacidad enfermedad general 3 Días en adelante hasta el día 90"] = ""
    FilaAgregar["Días incapacidad permanente entre día 91 al de 180 Días     (se paga al 50%)"] = ""
    FilaAgregar["Valor incapacidad permanente entre día 91 al de 180 Días                (se paga al 50%)"] = ""
    FilaAgregar["Días incapacidad permanente + de 180 Días"] = ""
    FilaAgregar["Valor Días incapacidad permanente + de 180 Días"] = ""
    FilaAgregar["Días licencia de maternidad"] = ""
    FilaAgregar["Valor Licencia de Maternidad"] = ""
    FilaAgregar["Días Licencia de Paternidad"] = ""
    FilaAgregar["Valor Licencia de Paternidad"] = ""
    FilaAgregar["Días Vacaciones en Disfrute Causadas"] = ""
    FilaAgregar["Valor Vacaciones en Disfrute Causadas"] = ""
    FilaAgregar["Grupo # 2 Total días ausencias justificadas sin cobro"] = ""
    FilaAgregar["Grupo # 2 Valor ausencias justificadas sin cobro"] = ""
    FilaAgregar["Días Vacaciones en Disfrute Anticipadas"] = ""
    FilaAgregar["Valor Vacaciones en Disfrute Anticipadas"] = ""
    FilaAgregar["Días Permiso Personal sin Reposición de tiempo por (de 1 o 2 Días)"] = ""
    FilaAgregar["Valor Permiso Personal sin Reposición de tiempo por (de 1 o 2 Días)"] = ""
    FilaAgregar["Días Permiso Justificado - Covid prevención aislamiento"] = ""
    FilaAgregar["Valor Permiso Justificado - Covid prevención aislamiento"] = ""
    FilaAgregar["Día Familiar"] = ""
    FilaAgregar["Valor Día Familiar"] = ""
    FilaAgregar["Día Compensación por desempeño"] = ""
    FilaAgregar["Valor Día Compensación por desempeño"] = ""
    FilaAgregar["Días Calamidad Domestica (de 1 o 2 Días)"] = ""
    FilaAgregar["Valor Calamidad Domestica (de 1 o 2 Días)"] = ""
    FilaAgregar["Dia Libre Jurado Votaciones"] = ""
    FilaAgregar["Valor Libre Jurado Votaciones"] = ""
    FilaAgregar["Días Licencia Remunerado (mayor a 2 Días) Aprobación RH"] = ""
    FilaAgregar["Valor Licencia Remunerado (mayor a 2 Días) Aprobación RH"] = ""
    FilaAgregar["Días Licencia Remunerada - covid casos vulnerables"] = ""
    FilaAgregar["Valor Licencia Remunerada - covid casos vulnerables"] = ""
    FilaAgregar["Días Licencia de Luto 5 Días hab (muerte de un familiar *1er grado de consanguinidad 5)"] = ""
    FilaAgregar["Valor Licencia de Luto 5 Días hab (muerte de un familiar *1er grado de consanguinidad 5)"] = ""
    FilaAgregar["Días Licencia de Matrimonio"] = ""
    FilaAgregar["Valor Licencia de Matrimonio"] = ""
    FilaAgregar["Dias Inasistencia por Alteración del Orden Publico 14708"] = ""
    FilaAgregar["Valor Inasistencia por Alteración del Orden Publico 14708"] = ""
    FilaAgregar["Diaz Hospitalizacion"] = ""
    FilaAgregar["Valor Hospitalizacion"] = ""
    FilaAgregar["Grupo #1 Total días ausencias justificadas con reconocimiento"] = ""
    FilaAgregar["Grupo # 1 Valor total ausencias justificadas con reconocimiento"] = ""
    FilaAgregar["Días Licencia No Remunerada (mayor a 2 Días) Aprobación RH"] = ""
    FilaAgregar["Valor Licencia No Remunerada (mayor a 2 Días) Aprobación RH (valor negativo)"] = ""
    FilaAgregar["Días Suspensión (originada Sanción)"] = ""
    FilaAgregar["Valor Días Suspensión (originada Sanción) valor negativo"] = ""
    FilaAgregar["Días Dominical por Suspensión (Inasistencia)"] = ""
    FilaAgregar["Valor Dominical por Suspensión (Inasistencia) - valor negativo"] = ""
    FilaAgregar["Días Inasistencia injustificada"] = ""
    FilaAgregar["Valor Inasistencia injustificada (este valor debe ser negativo)"] = ""
    FilaAgregar["Grupo # 3 Total Dias ausencias injustificadas, sanciones, dominical"] = ""
    FilaAgregar["Grupo # 3 Valor total ausencias injustificadas, sanciones, dominical"] = ""
    FilaAgregar["Indemnización, Bono por Retiro o Suma transaccional"] = ""
    FilaAgregar["Auxilio desplazamiento (parametrizado en el sistema por días laborados depende lugar trabajo y lugar residencia)"] = ""
    FilaAgregar["Gastos de transporte fijo (monto mensual asignado para desempeño de sus funciones ej. Comerciales, mantenimiento, gerentes, area seguridad etc)"] = ""
    FilaAgregar["Gastos de transporte ocasional (reportado por la operación quincenalmente)"] = ""
    FilaAgregar["Bonificacion no constitutiva de salario  BNCS"] = ""
    FilaAgregar["Bonificacion salarial"] = ""
    FilaAgregar["Total conceptos nómina"] = ""
    FilaAgregar["% Porcentaje Arl"] = ""
    FilaAgregar["Arl"] = ""
    FilaAgregar["Salud"] = ""
    FilaAgregar["Pension 12%"] = ""
    FilaAgregar["Total Seguridad Social"] = ""
    FilaAgregar["Caja de comp 4%"] = ""
    FilaAgregar["Sena 3%"] = ""
    FilaAgregar["Icbf 2%"] = ""
    FilaAgregar["Valor parafiscales"] = ""
    FilaAgregar["Cesantias 8,33%"] = ""
    FilaAgregar["Int. cesantias 1%"] = ""
    FilaAgregar["Prima 8,33%"] = ""
    FilaAgregar["Vacaciones 4.1667%"] = ""
    FilaAgregar["Valor prestaciones sociales"] = ""
    FilaAgregar["Total nomina + Seguridad Social + parafiscales + prestaciones"] = ""
    FilaAgregar["Administración temporales (el % que tenga cada temporal)"] = ""
    FilaAgregar["Examenes medicos  servicios"] = ""
    FilaAgregar["Menos servicio de alimentacion"] = ""
    FilaAgregar["Subtotal factura suppla"] = ""
    FilaAgregar["Iva del 19%"] = ""
    FilaAgregar["Total Neto Factura"] = ""
    FilaAgregar["Justificación (para casos puntuales que se requieran detallar)"] = ""
    FilaAgregar["Dcto my v/r pagado salario/Saldo en rojo"] = ""
    FilaAgregar["Saldo reconocido por el cliente"] = ""
    FilaAgregar["DEDUCCIONES VARIAS - NC"] = ""
    FilaAgregar["Deducicon Casino"] = ""
    FilaAgregar["EXCEDENTE DE SEGURIDAD SOCIAL"] = ""
    FilaAgregar["Seguridad Social Ley 1393 del 2010"] = ""
    SumatoriaNetoDev = 0
    SumatoriaNetoDed = 0
    
    #Ciclo para tomar informacion de los conceptos
    # for elemento in ConceptosDev:
    #     de = ContratoPos["Concepto"] == str(elemento)
    #     Conce= ContratoPos[de]
    #     Unidades = 0
    #     Neto = 0
    #     if (Conce.empty == False):
    #         Unidades = Conce["Horas"].sum()
    #         Neto = Conce["Neto"].sum()
    #         SumatoriaNetoDev += Neto
    #     if (elemento + " / Neto" in FilaAgregar):
    #         FilaAgregar[elemento + " / Unidades"] += Unidades
    #         FilaAgregar[elemento + " / Neto"] += Neto 
    #     else:
    #         FilaAgregar[elemento + " / Unidades"] = Unidades
    #         FilaAgregar[elemento + " / Neto"] = Neto 
    # FilaAgregar["Total Devengo"] = SumatoriaNetoDev
    
    # for elemento in ConceptosDed:
    #     de = ContratoPos["Concepto"] == str(elemento)
    #     Conce= ContratoPos[de]
    #     Unidades = 0
    #     Neto = 0
    #     if (Conce.empty == False):
    #         Unidades = Conce["Horas"].sum()
    #         Neto = Conce["Neto"].sum()
    #         SumatoriaNetoDed += Neto
    #     if (elemento + " / Neto" in FilaAgregar):
    #         FilaAgregar[elemento + " / Unidades"] += Unidades
    #         FilaAgregar[elemento + " / Neto"] += Neto 
    #     else:
    #         FilaAgregar[elemento + " / Unidades"] = Unidades
    #         FilaAgregar[elemento + " / Neto"] = Neto 
    Horizontal = pd.concat([Horizontal,pd.DataFrame.from_records([FilaAgregar])],ignore_index=True)
    return Horizontal  
# Funcion principal
def procesar(df1, IdProceso):
    Horizontal = pd.DataFrame()
    Contrato  = df1['Numero de Contrato'].unique().tolist()
    IDPeriodo = df1['ID Periodo'].tolist()
    IDPeriodo = np.unique(IDPeriodo)
    Conceptos = df1['Concepto'].unique().tolist()
    Conceptos.sort()
    ConceptosDev = []
    ConceptosDed = []
    for conceptosx in Conceptos:
        Valores = df1['Concepto'] == str(conceptosx)
        ContratoPos = df1[Valores]
        Total = ContratoPos['Neto'].sum()
        if(Total >= 0 ):
            ConceptosDev.append(conceptosx)
        else:
            ConceptosDed.append(conceptosx)
    Conceptos.clear()
    Conceptos= ConceptosDev + ConceptosDed

    #Obtener información para agregar al nuevo data frame
    for j in IDPeriodo:
        for i in Contrato:
            Valores = df1['Numero de Contrato'] == str(i)
            ContratoPos = df1[Valores]
            Valores2 = ContratoPos['ID Periodo'] == int(j)
            if IdProceso == "Agrupar ID proceso":
                ContratoPos = ContratoPos[Valores2]
                if ContratoPos.empty == False:
                    Horizontal = generar_dataframe_horizontal(ContratoPos, ConceptosDev, ConceptosDed, Horizontal)
            else:
                ContratoPos2 = ContratoPos[Valores2]
                ID_Procesos = ContratoPos2['Id Proceso'].unique().tolist()
                for k in ID_Procesos:
                    # Valido los contrato pos que tiene mi id proceso
                    Valores3 = ContratoPos2['Id Proceso'] == int(k)
                    # traigo de nuevo dataframe la fila o registro que contiene el proceso analizado y lo almaceno en contratopos
                    ContratoPos = ContratoPos2[Valores3]
                    if ContratoPos.empty == False:
                        Horizontal = generar_dataframe_horizontal(ContratoPos, ConceptosDev, ConceptosDed, Horizontal)
            
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
        worksheet.write_string(MaxFilas-1, contador,Dato)
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