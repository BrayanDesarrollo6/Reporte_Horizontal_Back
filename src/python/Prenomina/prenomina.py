import requests
import json
import time
from os import remove
import os
import pandas as pd
from xlsxwriter import Workbook
import numpy as np
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl import Workbook
from openpyxl.drawing.image import Image
from openpyxl import load_workbook
import sys
import cv2
import asyncio
#Importar funciones para obtener BD
from accessToken import funcionesGenerales

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

async def procesar(Horizontal):
    # Dataframe final para obtener los indices de las primeras columnas 
    global Empresa_
    Horizontal_heads_end = pd.DataFrame()
    #ELIMINAR REPETIDOS
    Horizontal.drop_duplicates(subset =['Numero de Contrato'], keep="last", inplace=True)
    Horizontal_heads_end = Horizontal
    #Nombre documento
    Empresa_ = Horizontal.iloc[0]['Empresa']
    NombreDocumento = "Prenomina_" + Empresa_
    # Normalizar nombre del documento
    NombreDocumento = normalize(NombreDocumento)
    NombreDocumento = replacement(NombreDocumento)
    heads = Horizontal.columns.values
    FilaAgregar = {}
    Validador = False

    wb = Workbook()
    ws = wb.active
    
    for r in dataframe_to_rows(Horizontal, index=False, header=True):
        ws.append(r)

    ws.insert_rows(1)
    ws.insert_rows(1)
    ws.insert_rows(1)
    writer = pd.ExcelWriter("./src/"+NombreDocumento+".xlsx", engine='xlsxwriter')
    Horizontal.to_excel(writer, sheet_name='Sheet1',index = False, header = False , startrow = 4)
    workbook = writer.book
    worksheet = writer.sheets["Sheet1"]
    
    format = workbook.add_format()
    format.set_pattern(1)
    format.set_bg_color('#FF2828')
    format.set_bold(True) 
    ##Informacion cliente
    worksheet.write_string(1, 0, "Temporal",format)
    worksheet.write_string(2, 0,"Cliente",format)
    worksheet.write_string(3, 0,"Periodo",format)
    worksheet.write_string(1, 1, str(Horizontal_heads_end.iloc[0]['Temporal']),format)
    worksheet.write_string(2, 1,str(Horizontal_heads_end.iloc[0]['Empresa']),format)
    worksheet.write_string(3, 1,str(Horizontal_heads_end.iloc[0]['Periodo']),format)
    
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
        # worksheet.write_string(MaxFilas-1, contador,Dato ,format)
        contador += 1
        
    writer.close()
    funcionesGenerales().updatedata("./src/"+NombreDocumento+".xlsx",IDregistro_)
  
async def procesar_prenomina(data):
    # -----------------------------
    global IDregistro_,Periodo
    json_object = json.loads(data)
    Periodo = json_object['data']['periodo']
    IDregistro_ = json_object['data']['id_registro']
    # RECORRER LAS PRESTACIONES SOCIALES
    URL = "https://creatorapp.zohopublic.com/hq5colombia/compensacionhq5/xls/Prenomina/WWjRAOJ2MGyyNGd5BxdvwApYGzgq5A9AQ5Q6bUmpsTQvWTMJE4qE5MyKnY4KKPXneurq8RnTZ2O698AO8N2KQ7Fa7qt4hpwSet0K?Temporal_op=30&Periodo=" + Periodo
    df = pd.read_excel(URL)
    df1 = pd.DataFrame(df)
    if(df1.empty):
        print("No existe registro")
    else:
        print("Existe y se esta procesando")
        await procesar(df1)
        # Documento_one = procesar(df1,df3)
        # print(Documento_one)  
        
async def main():
    info_received = sys.argv[1]
    await procesar_prenomina(info_received) 
if __name__ == "__main__":
    asyncio.run(main())