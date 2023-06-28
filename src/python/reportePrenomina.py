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
def AccesToken():
    refresh_ = '1000.33a5526ca066d526ee93e1d4eac2c002.c2707bc77a9bb1cfcb29e590da4581b8'
    url = 'https://accounts.zoho.com/oauth/v2/token?refresh_token=' + refresh_ + '&client_id=1000.1X8CFKQHNVMIQYBM2LD5D630UAMMXB&client_secret=ed77d9ad812478a75cb46e11db1bbc262b8f1d49bf&grant_type=refresh_token'
    cabeceras = {"Content-Type": "application/json", "Access-Control-Allow-Origin": "*"} 
    auth_data = {"answer": "42" }
    resp = requests.post(url, data=auth_data,headers=cabeceras)
    posts = resp.json()
    # print("Respuesta de token")
    # print(posts)
    if "access_token" in posts:
        token = posts['access_token']
    else:
        time.sleep(5)
        refresh_ = '1000.33a5526ca066d526ee93e1d4eac2c002.c2707bc77a9bb1cfcb29e590da4581b8'
        url = 'https://accounts.zoho.com/oauth/v2/token?refresh_token=' + refresh_ + '&client_id=1000.1X8CFKQHNVMIQYBM2LD5D630UAMMXB&client_secret=ed77d9ad812478a75cb46e11db1bbc262b8f1d49bf&grant_type=refresh_token'
        cabeceras = {"Content-Type": "application/json", "Access-Control-Allow-Origin": "*"} 
        auth_data = {"answer": "42" }
        resp = requests.post(url, data=auth_data,headers=cabeceras)
        posts = resp.json()
        token = posts['access_token']
    return token
def procesar(Horizontal):
    # Dataframe final para obtener los indices de las primeras columnas 
    global Empresa_
    Horizontal_heads_end = pd.DataFrame()
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
    
    # for k in heads:
    #     if(str(k).__contains__("Salario Base")):
    #         Validador = True
    #     if(Validador):
    #         Horizontal[k] = Horizontal[k].astype('float')
    #         FilaAgregar[k] = sum(Horizontal[k])
        
    # Horizontal = pd.concat([Horizontal,pd.DataFrame.from_records([FilaAgregar])],ignore_index=True)

    wb = Workbook()
    ws = wb.active
    
    for r in dataframe_to_rows(Horizontal, index=False, header=True):
        ws.append(r)

    ws.insert_rows(1)
    ws.insert_rows(1)
    ws.insert_rows(1)
    writer = pd.ExcelWriter("./src/database/"+NombreDocumento+".xlsx", engine='xlsxwriter')
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
    insertarDocumento()
def insertarDocumento():
    global IDregistro_,Empresa_
    # Llamar API de zoho para actualizar datos del consecutivo
    # Token_ = AccesToken()
    Token_ = AccesToken()
    # Realizar consulta
    NombreDocumento = "./src/database/Prenomina_" + Empresa_
    with open(NombreDocumento +'.xlsx','rb') as archivo:
        Datos_ = {'file': archivo}
        url_ = "https://creator.zoho.com/api/v2/hq5colombia/compensacionhq5/report/Generar_pre_nomina_Report/"+str(IDregistro_)+"/Adjunto1/upload"
        header = {"Authorization":"Zoho-oauthtoken "+Token_, "Access-Control-Allow-Origin": "*"} 
        r = requests.post(url_,files=Datos_,headers=header)
    if(r.status_code == 200):
        time.sleep(5)
        os.remove(NombreDocumento +'.xlsx')   
        #hacer registro para el correo
        # Realizar consulta
        Datos_ = {"data": {"Generar_pre_nomina":IDregistro_,"Proceso":"Prenomina"}}
        Datos_ = json.dumps(Datos_)
        url_ = "https://creator.zoho.com/api/v2/hq5colombia/compensacionhq5/form/templatesVarios"
        header = {"Authorization":"Zoho-oauthtoken "+Token_, "Access-Control-Allow-Origin": "*"} 
        r = requests.post(url_,data=Datos_,headers=header)
        
    
    
    
# -----------------------------
global IDregistro_,Periodo
Periodo = sys.argv[1]
IDregistro_ = sys.argv[2]
# RECORRER LAS PRESTACIONES SOCIALES
URL = "https://creatorapp.zohopublic.com/hq5colombia/compensacionhq5/xls/Prenomina/WWjRAOJ2MGyyNGd5BxdvwApYGzgq5A9AQ5Q6bUmpsTQvWTMJE4qE5MyKnY4KKPXneurq8RnTZ2O698AO8N2KQ7Fa7qt4hpwSet0K?Periodo=" + Periodo
df = pd.read_excel(URL)
df1 = pd.DataFrame(df)
if(df1.empty and df3.empty):
    print("No existe registro")
else:
    procesar(df1)
    # Documento_one = procesar(df1,df3)
    # print(Documento_one)   