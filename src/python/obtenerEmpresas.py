from django.http import HttpResponse
from django.shortcuts import render
import pandas as pd
from xlsxwriter import Workbook
import numpy as np
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl import Workbook
from openpyxl import load_workbook
import sys

# Obtener si es para liquidiacion o reliquidacion
estado = sys.argv[1]
print(estado)
Estado_ = "[Enviada a Pago,Enviada a pago sin paz y salvo,Pagada]".replace(" ","%20")
URL = "https://creatorapp.zohopublic.com/hq5colombia/compensacionhq5/xls/Conceptos_De_Liquidaci_n_Retiros_Report/JPdZda7vkNjCJEanQ6P4x4eBB6m8BJKR4wfNXDSyz5q2qdn8nZdjdz0nFvaqYaegJ5qSmj8pnkNqTMTYwwhtwJW1XPR2ae2Vdmbe?liquidacion_lp.Estado="+Estado_
if (estado == "1"):
    URL = "https://creatorapp.zohopublic.com/hq5colombia/compensacionhq5/xls/Conceptos_De_Re_Liquidaci_n_Report/3juBT5YjxpXsDvAmfX76TkE4B4v2gwsDbZtxgrqZfDjHE7zFw5T8rHnjpFZuruae3PC7g6uww4761Xtm5h97yDj4hka5ws5xXabR?reliquidacion_lp.Estado="+Estado_
           
df = pd.read_excel(URL)
df1 = pd.DataFrame(df)
Empresas = df1['Empresa'].unique().tolist()
print(Empresas)