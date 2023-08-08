import pandas as pd
from xlsxwriter import Workbook
import numpy as np
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl import Workbook
from openpyxl import load_workbook
import sys
import json
from Directories.Directory import DirectoryObtenerEmpresasLQ, DirectoryObtenerEmpresasReLQ

# Obtener si es para liquidacion o reliquidacion
estado = sys.argv[1]

if estado == "0":
    URL = "https://creatorapp.zohopublic.com/hq5colombia/compensacionhq5/xls/Conceptos_De_Liquidaci_n_Retiros_Report/JPdZda7vkNjCJEanQ6P4x4eBB6m8BJKR4wfNXDSyz5q2qdn8nZdjdz0nFvaqYaegJ5qSmj8pnkNqTMTYwwhtwJW1XPR2ae2Vdmbe"
if estado == "1":
    URL = "https://creatorapp.zohopublic.com/hq5colombia/compensacionhq5/xls/Conceptos_De_Re_Liquidaci_n_Report/3juBT5YjxpXsDvAmfX76TkE4B4v2gwsDbZtxgrqZfDjHE7zFw5T8rHnjpFZuruae3PC7g6uww4761Xtm5h97yDj4hka5ws5xXabR"           

df = pd.read_excel(URL)
df1 = pd.DataFrame(df)

if(df1.empty == False):
    File_Json = {}
    Empresas = df1['Empresa Usuaria'].unique().tolist()
    Empresas = [str(x) for x in Empresas]
    Empresas = [x for x in Empresas if x != 'nan']
    File_Json["Empresas"] = Empresas
    Estados = df1['Estado'].unique().tolist()
    Estados = [str(x) for x in Estados]
    Estados = [x for x in Estados if x != 'nan']
    File_Json["Estados"] = Estados
    if estado == "0":
        with open(DirectoryObtenerEmpresasLQ,"w") as file:
            json.dump(File_Json, file, indent=4)
        print("VariablesEntornoLQ.json")    
    if estado == "1":
        with open(DirectoryObtenerEmpresasReLQ,"w") as file:
            json.dump(File_Json, file, indent=4)
        print("VariablesEntornoReLQ.json")
    
else:
    print("No existe registro")