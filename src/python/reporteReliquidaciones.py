from django.http import HttpResponse
from django.shortcuts import render
import pandas as pd
from xlsxwriter import Workbook
import numpy as np
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl import Workbook
from openpyxl import load_workbook
import sys

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

def procesar(df,df3):
     # Dataframe final
    Horizontal = pd.DataFrame()
    Contrato  = df['Numero de Contrato'].unique().tolist()
    IDLiquidacion = df['Id Proceso'].tolist()
    IDLiquidacion = np.unique(IDLiquidacion)
    #Filtrar cada concepto unico que existe en ese reporte
    Conceptos = df['Concepto'].unique().tolist()
    Conceptos.sort()
    ConceptosDev = []
    ConceptosDed = []
    for conceptosx in Conceptos:
        Valores = df['Concepto'] == str(conceptosx)
        ContratoPos = df[Valores]
        #Sumatoria
        Total = ContratoPos['Neto'].sum()
        if(Total >= 0 ):
            ConceptosDev.append(conceptosx)
        else:
            ConceptosDed.append(conceptosx)
    Conceptos.clear()
    Conceptos= ConceptosDev + ConceptosDed
    
    for j in IDLiquidacion:
        Valores = df['Id Proceso'] == j
        ContratoPos = df[Valores]
        if ContratoPos.empty == False:
            FilaAgregar = {}
            SumatoriaNetoprestaciones = 0
            Subtotal = 0
            ##Informacion general inicial
            FilaAgregar["Id Proceso"] = ContratoPos.iloc[0]['Id Proceso']
            FilaAgregar["Estado"] = ContratoPos.iloc[0]['Estado']
            FilaAgregar["Temporal"] = ContratoPos.iloc[0]['Temporal']
            FilaAgregar["Empresa"] = ContratoPos.iloc[0]['Empresa Usuaria']
            FilaAgregar["Periodo"] = ContratoPos.iloc[0]['Periodo']
            FilaAgregar["Numero de Contrato"] = ContratoPos.iloc[0]['Numero de Contrato']
            FilaAgregar["Nombres y Apellidos"] = ContratoPos.iloc[0]['Nombres y Apellidos']
            FilaAgregar["Numero de Identificación"] = ContratoPos.iloc[0]['Numero de Identificación']
            FilaAgregar["Fecha Ingreso"] = pd.to_datetime(ContratoPos.iloc[0]['Fecha Ingreso']).date()
            FilaAgregar["Fecha Retiro"] = pd.to_datetime(ContratoPos.iloc[0]['Fecha Retiro']).date()
            FilaAgregar["Cargo"] = ContratoPos.iloc[0]['Cargo']
            FilaAgregar["Salario Base"] = ContratoPos.iloc[0]['Salario Base']
            #SE REVISA LOS CONCEPTOS DE LIQUIDACION PARA CADA UNA 
            # LEYENDO PARA CADA 1
            SumatoriaNetoDev = 0
            SumatoriaNetoDed = 0
            #Ciclo para tomar informacion de los conceptos
            for elemento in ConceptosDev:
                de = ContratoPos["Concepto"] == str(elemento)
                Conce= ContratoPos[de]
                Unidades = 0
                Neto = 0
                if (Conce.empty == False):
                    Unidades = Conce["Unidades"].sum()
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
                    Unidades = Conce["Unidades"].sum()
                    Neto = Conce["Neto"].sum()
                    SumatoriaNetoDed += Neto
                if (elemento + " / Neto" in FilaAgregar):
                    FilaAgregar[elemento + " / Unidades"] += Unidades
                    FilaAgregar[elemento + " / Neto"] += Neto 
                else:
                    FilaAgregar[elemento + " / Unidades"] = Unidades
                    FilaAgregar[elemento + " / Neto"] = Neto 
            FilaAgregar["Total Deduccion"] = SumatoriaNetoDed
            FilaAgregar["Subtotal a pagar"] = SumatoriaNetoDev - abs(SumatoriaNetoDed)
            Subtotal = SumatoriaNetoDev - abs(SumatoriaNetoDed)

            if(df3.empty):
                print("No existe registro")
            else:
                Valores = df3['Id Proceso'] == j
                ConceptosPrestaciones = df3[Valores]
                Conceptos = df3['Concepto'].unique().tolist()
                for elemento in Conceptos:
                    de = ConceptosPrestaciones["Concepto"] == str(elemento)
                    Conce= ConceptosPrestaciones[de]
                    Unidades = 0
                    Neto = 0
                    if (Conce.empty == False):
                        Unidades = Conce["Unidades"].sum()
                        Neto = Conce["Neto"].sum()
                        SumatoriaNetoprestaciones += Neto
                    if (elemento + " / Neto" in FilaAgregar):
                        FilaAgregar[elemento + " / Unidades"] += Unidades
                        FilaAgregar[elemento + " / Neto"] += Neto 
                    else:
                        FilaAgregar[elemento + " / Unidades"] = Unidades
                        FilaAgregar[elemento + " / Neto"] = Neto
                FilaAgregar["Subtotal a pagar prestaciones"] = SumatoriaNetoprestaciones
            
            ##INDEMNIZACION
            FilaAgregar["1441 - Indemnización / Neto"] = ContratoPos.iloc[0]['Sub Total Neto indemnizacion']
            NetoIndemnizacion = 0
            NetoIndemnizacion = ContratoPos.iloc[0]['Sub Total Neto indemnizacion'].sum()
            FilaAgregar["Neto a pagar"] = SumatoriaNetoprestaciones + Subtotal + NetoIndemnizacion
            Horizontal = pd.concat([Horizontal,pd.DataFrame.from_records([FilaAgregar])],ignore_index=True)

    # Dataframe final para obtener los indices de las primeras columnas 
    Horizontal_heads_end = pd.DataFrame()
    Horizontal_heads_end = Horizontal
            
    NombreDocumento = "Horizontal_Reliquidaciones_" + Horizontal.iloc[0]['Empresa']
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
    # worksheet.write_string(1, 3,str(Horizontal_heads_end.iloc[0]['ID Periodo']),format)
    # worksheet.write_string(1, 4,str(Horizontal_heads_end.iloc[0]['Tipo de Perido']),format)
    # worksheet.write_string(1, 5,str(Horizontal_heads_end.iloc[0]['Mes']),format)
    
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
    


Empresa_ = "RANSA COLOMBIA SAS".replace(" ", "%20")    
# Estado_ = "[Enviada a Pago,Pagada,Pendiente,Enviada a Aprobación,Aprobada]".replace(" ","%20")
# Estado_ = "[Enviada a Pago,Pagada,Pendiente]".replace(" ","%20")
Estado_ = "[Enviada a Pago,Pagada,Pendiente]".replace(" ","%20")
#Traer información
# URL = "https://creatorapp.zohopublic.com/hq5colombia/compensacionhq5/xls/Generar_Re_Liquidaci_n_Report/U6jN3utwhQmfu2D7CqC6Mp20sh3J8xjHSdjtAN6GsxsqvuKYk0MEQ0vUHgyFwPdPOKj5B5mXmW5bNrK3yPDR3w3t1CPj5dDex2PR?Empresa_Usuaria="+Empresa_+"&Estado="+Estado_
URL = "https://creatorapp.zohopublic.com/hq5colombia/compensacionhq5/xls/Conceptos_De_Re_Liquidaci_n_Report/3juBT5YjxpXsDvAmfX76TkE4B4v2gwsDbZtxgrqZfDjHE7zFw5T8rHnjpFZuruae3PC7g6uww4761Xtm5h97yDj4hka5ws5xXabR?reliquidacion_lp.Empresa_Usuaria="+Empresa_+"&reliquidacion_lp.Estado="+Estado_
            
# url_ = "https://creatorapp.zohopublic.com/hq5colombia/compensacionhq5/xls/Prenomina/WWjRAOJ2MGyyNGd5BxdvwApYGzgq5A9AQ5Q6bUmpsTQvWTMJE4qE5MyKnY4KKPXneurq8RnTZ2O698AO8N2KQ7Fa7qt4hpwSet0K?Periodo=" + _idPeriodo + "&zc_FileName=PreNomina_" + _idPeriodo;
df = pd.read_excel(URL)
df1 = pd.DataFrame(df)
# RECORRER LAS PRESTACIONES SOCIALES
URL = "https://creatorapp.zohopublic.com/hq5colombia/compensacionhq5/xls/Prestaci_n_Social_Re_Liquidaci_n_Report/BC3FExKk0GgnAgbYaqrODw2gNbJe505hXt6OCHB2G0gvAuNTe7Ora79UMead2XdWFtUGVQbYb4epCSDwwZJ5SdMe98hd3YOeghhH?reliquidacion_lp.Empresa_Usuaria="+Empresa_+"&reliquidacion_lp.Estado="+Estado_
df2 = pd.read_excel(URL)
df3 = pd.DataFrame(df2)
if(df1.empty):
    print("No existe registro")
else:
    Documento_one = procesar(df1,df3)
    print(Documento_one)