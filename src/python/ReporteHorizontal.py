from django.http import HttpResponse
from django.shortcuts import render
import pandas as pd
from xlsxwriter import Workbook
import numpy as np
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl import Workbook
from openpyxl import load_workbook
import sys

# ------------------------------------------------------------------------------------------
# reemplazar acentos
def normalize(s):
    replacements = (
        ("á", "a"),("é", "e"),("í", "i"),("ó", "o"),("ú", "u"),
        ("Á", "A"),("É", "E"),("Í", "I"),("Ó", "O"),("Ú", "U")
    )
    for a, b in replacements:
        s = s.replace(a, b).replace(a.upper(), b.upper())
    return s

# ------------------------------------------------------------------------------------------
# reemplazar caracteres
def replacement(name_company):
    name_company = name_company.replace(".", "")
    name_company = name_company.replace("-", "_")
    name_company = name_company.replace("–", "_")
    name_company = name_company.replace("—", "_")
    return name_company

# ------------------------------------------------------------------------------------------
# Función secundaria 
def generar_dataframe_horizontal(ContratoPos, ConceptosDev, ConceptosDed, Horizontal):
    # Fila agregar
    FilaAgregar = {}
    ##Informacion general inicial
    FilaAgregar["Estado Nomina"] = ContratoPos.iloc[0]['Estado Nomina']
    FilaAgregar["Temporal"] = ContratoPos.iloc[0]['Temporal']
    FilaAgregar["Empresa"] = ContratoPos.iloc[0]['Empresa']
    FilaAgregar["ID Periodo"] = ContratoPos.iloc[0]['ID Periodo']
    FilaAgregar["Tipo de Perido"] = ContratoPos.iloc[0]['Tipo de Perido']
    FilaAgregar["Id Proceso"] = ContratoPos.iloc[0]['Id Proceso']
    FilaAgregar["Mes"] = ContratoPos.iloc[0]['Mes']
    FilaAgregar["Numero de Contrato"] = ContratoPos.iloc[0]['Numero de Contrato']
    FilaAgregar["Nombres y Apellidos"] = ContratoPos.iloc[0]['Nombres y Apellidos']
    FilaAgregar["Numero de Identificación"] = ContratoPos.iloc[0]['Numero de Identificación']
    FilaAgregar["Centro de Costo"] = ContratoPos.iloc[0]['Centro de Costo']   
    if(ContratoPos.iloc[0]['Dependencia']):
        FilaAgregar["Dependencia"] = ContratoPos.iloc[0]['Dependencia']
    if(ContratoPos.iloc[0]['Proceso']):
        FilaAgregar["Proceso"] = ContratoPos.iloc[0]['Proceso']
    FilaAgregar["Fecha Ingreso"] = pd.to_datetime(ContratoPos.iloc[0]['Fecha Ingreso']).date()
    FilaAgregar["Fecha Retiro"] = pd.to_datetime(ContratoPos.iloc[0]['Fecha Retiro']).date()
    FilaAgregar["Cargo"] = ContratoPos.iloc[0]['Cargo']
    FilaAgregar["Salario Base"] = ContratoPos.iloc[0]['Salario Base']
    SumatoriaNetoDev = 0
    SumatoriaNetoDed = 0
    #Ciclo para tomar informacion de los conceptos
    for elemento in ConceptosDev:
        de = ContratoPos["Concepto"] == str(elemento)
        Conce= ContratoPos[de]
        Unidades = 0
        Neto = 0
        if (Conce.empty == False):
            Unidades = Conce["Horas"].sum()
            Neto = Conce["Neto"].sum()
            SumatoriaNetoDev += Neto
        if (elemento + " / Neto" in FilaAgregar):
            FilaAgregar[elemento + " / Unidades"] += Unidades
            FilaAgregar[elemento + " / Neto"] += Neto 
        else:
            FilaAgregar[elemento + " / Unidades"] = Unidades
            FilaAgregar[elemento + " / Neto"] = Neto 
    FilaAgregar["Total Devengo"] = SumatoriaNetoDev
    for elemento in ConceptosDed:
        de = ContratoPos["Concepto"] == str(elemento)
        Conce= ContratoPos[de]
        Unidades = 0
        Neto = 0
        if (Conce.empty == False):
            Unidades = Conce["Horas"].sum()
            Neto = Conce["Neto"].sum()
            SumatoriaNetoDed += Neto
        if (elemento + " / Neto" in FilaAgregar):
            FilaAgregar[elemento + " / Unidades"] += Unidades
            FilaAgregar[elemento + " / Neto"] += Neto 
        else:
            FilaAgregar[elemento + " / Unidades"] = Unidades
            FilaAgregar[elemento + " / Neto"] = Neto 
    FilaAgregar["Total Deduccion"] = SumatoriaNetoDed
    FilaAgregar["Neto A Pagar"] = SumatoriaNetoDev - abs(SumatoriaNetoDed)
    # Informacion general de provisiones y SS             
    FilaAgregar["EPS"] = ContratoPos['EPS'].unique().sum()
    FilaAgregar["AFP"] = ContratoPos['AFP'].unique().sum()
    FilaAgregar["ARL"] = ContratoPos['ARL'].unique().sum()
    FilaAgregar["Riesgo ARL"] = ContratoPos['Riesgo ARL'].unique().sum()
    FilaAgregar["CCF"] = ContratoPos['CCF'].unique().sum()
    FilaAgregar["SENA"] = ContratoPos['SENA'].unique().sum()
    FilaAgregar["ICBF"]  = ContratoPos['ICBF'].unique().sum()
    FilaAgregar["Total Seguridad Social"] = ContratoPos['EPS'].unique().sum() + ContratoPos['AFP'].unique().sum() + ContratoPos['ARL'].unique().sum() + ContratoPos['CCF'].unique().sum() + ContratoPos['SENA'].unique().sum() + ContratoPos['ICBF'].unique().sum()
    FilaAgregar["Vacaciones tiempo"] = ContratoPos['Vacaciones tiempo'].unique().sum()
    FilaAgregar["Prima"] = ContratoPos['Prima'].unique().sum()
    FilaAgregar["Cesantías"] = ContratoPos['Cesantías'].unique().sum()
    FilaAgregar["Interés cesantías"] = ContratoPos['Interés cesantías'].unique().sum()
    FilaAgregar["Total provisiones"] = ContratoPos['Vacaciones tiempo'].unique().sum() + ContratoPos['Prima'].unique().sum() + ContratoPos['Cesantías'].unique().sum() + ContratoPos['Interés cesantías'].unique().sum()
    Horizontal = pd.concat([Horizontal,pd.DataFrame.from_records([FilaAgregar])],ignore_index=True)
    return Horizontal  

# ------------------------------------------------------------------------------------------
# Funcion principal
def procesar(df1, IdProceso):
    
    # Dataframe final
    Horizontal = pd.DataFrame()
    # Datos generales
    Contrato  = df1['Numero de Contrato'].unique().tolist()
    IDPeriodo = df1['ID Periodo'].tolist()
    IDPeriodo = np.unique(IDPeriodo)
    #Filtrar cada concepto unico que existe en ese reporte
    Conceptos = df1['Concepto'].unique().tolist()
    Conceptos.sort()
    ConceptosDev = []
    ConceptosDed = []
    for conceptosx in Conceptos:
        Valores = df1['Concepto'] == str(conceptosx)
        ContratoPos = df1[Valores]
        #Sumatoria
        Total = ContratoPos['Neto'].sum()
        if(Total >= 0 ):
            ConceptosDev.append(conceptosx)
        else:
            ConceptosDed.append(conceptosx)
    Conceptos.clear()
    Conceptos= ConceptosDev + ConceptosDed
    #Se obtienen los datos dependiendo del empleado

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
            
    NombreDocumento = "Horizontal " + Horizontal.iloc[0]['Empresa'] +"-"+ str(Horizontal.iloc[0]['Mes'])+ "-" + str(Horizontal.iloc[0]['Tipo de Perido'])
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

    ws.insert_rows(1)
    ws.insert_rows(1)
    ws.insert_rows(1)
    
    Horizontal = pd.DataFrame(ws.values)
    
    writer = pd.ExcelWriter("./src/database/"+NombreDocumento+".xlsx", engine='xlsxwriter')
    Horizontal.to_excel(writer, sheet_name='Sheet1',index = False, header = False)
    workbook = writer.book
    worksheet = writer.sheets["Sheet1"]
    format = workbook.add_format()
    format.set_pattern(1)
    format.set_bg_color('#AFAFAF')
    format.set_bold(True) 
        
    worksheet.write_string(1, 1, str(Horizontal_heads_end.iloc[0]['Temporal']),format)
    worksheet.write_string(1, 2,str(Horizontal_heads_end.iloc[0]['Empresa']),format)
    worksheet.write_string(1, 3,str(Horizontal_heads_end.iloc[0]['ID Periodo']),format)
    worksheet.write_string(1, 4,str(Horizontal_heads_end.iloc[0]['Tipo de Perido']),format)
    worksheet.write_string(1, 5,str(Horizontal_heads_end.iloc[0]['Mes']),format)
    
    contador = 0
    MaxFilas = len(Horizontal.axes[0])
    Totales = Horizontal.loc[MaxFilas -1]
    for k in heads:
        
        worksheet.write_string(3, contador,str(k),format)
        contador += 1
        
    contador = 0
    for k in Totales:
        Dato = ""
        if(str(k) != "nan"):
            Dato = str(k)
        worksheet.write_string(MaxFilas-1, contador,Dato ,format)
        contador += 1
        
    writer.close()
    return NombreDocumento+".xlsx"

# ------------------------------------------------------------------------------------------
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

    #Traer información
    URL = "https://creatorapp.zohopublic.com/hq5colombia/compensacionhq5/xls/Conceptos_Nomina_Desarrollo/jwhRFUOR47TqCS9AAT82eCybwgdmgeArEtKG7U8H9s3hSjTzBd3G8bPdg37PHVygvxurxwCQvMCgHRG68dOCWKTmMWaQJU2TMwnr?ID_Periodo="+IdPeriodo
    # url_ = "https://creatorapp.zohopublic.com/hq5colombia/compensacionhq5/xls/Prenomina/WWjRAOJ2MGyyNGd5BxdvwApYGzgq5A9AQ5Q6bUmpsTQvWTMJE4qE5MyKnY4KKPXneurq8RnTZ2O698AO8N2KQ7Fa7qt4hpwSet0K?Periodo=" + _idPeriodo + "&zc_FileName=PreNomina_" + _idPeriodo;
    df = pd.read_excel(URL)
    df1 = pd.DataFrame(df)
    if(df1.empty):
        print("No existe registro")
    else:
        Documento_one = procesar(df1, IdProceso)
        print(Documento_one)
      
# ------------------------------------------------------------------------------------------
# Si llegó mas de un ID primero válide la empresa usuaria
def validar_empresa():
    if(estado == '2'):
        IdProceso = sys.argv[2]
        IdPeriodo1 = sys.argv[3]
        IdPeriodo2 = sys.argv[4]
        #Traer información
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
        #Traer información
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

# ------------------------------------------------------------------------------------------
# validar cuantos ID llegaron 1 o más
def estados():
    if(estado == '1'):
        validar_contenido_id()
    else:
        validar_empresa()
    
# ------------------------------------------------------------------------------------------
# Inicio del programa 
global estado
estado = sys.argv[1]
estados()
# validar_empresa()
# validar_contenido_id()