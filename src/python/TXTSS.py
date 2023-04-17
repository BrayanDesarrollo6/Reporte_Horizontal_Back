from django.http import HttpResponse

#improtar shortcuts, metodo para renderisar plantillas y optimizar codigo al cargar plantillas
from django.shortcuts import render
##Librerias para el proceso del reporte horizontal
import pandas as pd
from xlsxwriter import Workbook
# from io import BytesIO
# import io,csv
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl import Workbook
from openpyxl import load_workbook
import sys
# import json

# request.GET.get("id")
# Dataframe final
# return render(request, "Resultado.html")

def procesarTXTSS(NombreTemporal,Anio,Mes,Grupo):
    
    NombreTemporal_ = NombreTemporal.replace(" ","%20")
    
    Grupo = Grupo.replace(", ", ",")
    Grupo = Grupo.replace(" ", "%20")
    Grupo = '[' + Grupo + ']'
    
    TXT_Final = pd.DataFrame()
    
    # URL DEL XLS DE SS
    URL = "https://creatorapp.zohopublic.com/hq5colombia/compensacionhq5/xls/TXT_SS_DESARROLLO/3HO1RZORhePyRgar44EefyEhhD27umsJE7GeJJhCDwwx2ngQ2KEHHGTCB1mYQtFktzmgSyHG2qRWsnu3ZGbW8N97TZtX709N3DAC?NOMBRE_EMPRESA=" + NombreTemporal_ + "&PENSION_ANO=" + Anio + "&PENSION_MES=" + Mes + "&agrupacion_de_empresa_seg_soc=" + Grupo
    # URL = "https://creatorapp.zohopublic.com/hq5colombia/compensacionhq5/xls/TXT_SS_DESARROLLO/3HO1RZORhePyRgar44EefyEhhD27umsJE7GeJJhCDwwx2ngQ2KEHHGTCB1mYQtFktzmgSyHG2qRWsnu3ZGbW8N97TZtX709N3DAC?NOMBRE_EMPRESA=HQ5%20S.A.S&PENSION_ANO=2022&PENSION_MES=10"
    #CNVERTIR XLS EN DATAFRAME
    df = pd.read_excel(URL)
    df1 = pd.DataFrame(df)
    
    if(df1.empty):
        print("No existe registro")
    else:
        #Armar la primera linea dependiendo temporal con nit y demas
        #Tomar sumatoria de ibc
        IBCTotal_ = sum(df1["IBC EPS"])
        ##Contador total de lineas del txt
        MaxFilas = len(df1.axes[0])
        ##Datos fijos del txt
        E_TRegristro_ = "01"
        E_Modalidad_ = "2"
        E_Secuencia_ = "0001"
        #Funcion para reemplazar acentos del txt
        def normalize(s):
            replacements = (
                ("á", "a"),
                ("é", "e"),
                ("í", "i"),
                ("ó", "o"),
                ("ú", "u"),
                ("Á", "A"),
                ("É", "E"),
                ("Í", "I"),
                ("Ó", "O"),
                ("Ú", "U")
            )
            for a, b in replacements:
                s = s.replace(a, b).replace(a.upper(), b.upper())
            return s
        #Normalizar nombre de la temporal
        NombreTemporal = normalize(NombreTemporal)
        NombreTemporal = NombreTemporal.replace(".", "")
        NombreTemporal = NombreTemporal.replace("-", "_")
        #Completar 200 espacios y se rellenan con espacios en blanco
        Espacios_ = 200 - len(NombreTemporal)
        NombreTemporal = NombreTemporal + (Espacios_ * " ")
        E_RazonSocial_ = NombreTemporal
        #Nit tempora - se obtiene del dataframe y ademas se completan 16 espacios en blanco
        E_TDocumento_ = "NI"
        E_NDocumento_ = str(df.iloc[0]['NIT'])
        Espacios_ = 16 - len(E_NDocumento_)
        E_NDocumento_ = E_NDocumento_ + (Espacios_ * " ")
        #Numero de verificacion - se obtiene del dataframe
        E_DVerificacion_ = str(df.iloc[0]['Número de verificacación'])
        ##Informacion fija, planilla E 
        E_TPlantilla_ = "E"
        E_Blanco1_ = 20 * " "
        E_FPresentacion_ = "U"
        #Datos fijos y completar con espacios en blanco
        E_CSucursal_ = "01" + (8 * " ")
        E_NSucursal_ = "01" + (38 * " ")
        #Codigo de la ARL, se obtiene del data frame y se completan 6 espacios
        E_CodARL_ = str(df.iloc[0]['Código ARL'])
        Espacios_ = 6 - len(E_CodARL_)
        E_CodARL_ = E_CodARL_ + (Espacios_ * " ")
        ##Fechas de pension y salud
        E_FechaPension_ = str(df.iloc[0]['Año pensión']) + "-" + str(df.iloc[0]['Mes pensión'])
        E_FechaSalud_ = str(df.iloc[0]['Año pensión']) + "-" + str(int(df.iloc[0]['Mes pensión'] + 1))
        #Informaicon fija, con 10 espacios en 0
        E_NRadicacion_ = "0000000000"
        E_FechaPago = 10 * "0"

        ##Total de empleados, completadno con 0 a la izquierda
        E_TotalEmpleados_ = str(MaxFilas)
        Ceros_ = 5 - len(E_TotalEmpleados_)
        E_TotalEmpleados_ = (Ceros_ * "0") + E_TotalEmpleados_ 
        ##Total de ibc a pagar completando con 0 espacios hasta 12 
        E_TotalNomina_ = str(IBCTotal_).replace(".", "")
        Ceros_ = 12 - len(E_TotalNomina_)
        E_TotalNomina_ = (Ceros_ * "0") + E_TotalNomina_
        #Informacion Fija
        E_TAportente_ = "01"
        E_CodOperador_ = "00"
        #Se concatena la informacio anterior en una variable y luego se agrega al diccionario
        TXT_ = E_TRegristro_ + E_Modalidad_ + E_Secuencia_ + E_RazonSocial_ + E_TDocumento_ + E_NDocumento_ + E_DVerificacion_ + E_TPlantilla_ + E_Blanco1_ + E_FPresentacion_ + E_CSucursal_ + E_NSucursal_ + E_CodARL_ + E_FechaPension_ + E_FechaSalud_ + E_NRadicacion_ + E_FechaPago + E_TotalEmpleados_ + E_TotalNomina_ + E_TAportente_ + E_CodOperador_
        FilaAgregar = {}
        FilaAgregar["TXT"] = TXT_
        ##Se agrega el diccionario a un nuevo dataframe
        TXT_Final = pd.concat([TXT_Final,pd.DataFrame.from_records([FilaAgregar])],ignore_index=True)
        
        ## Dar consecutivo y completar linea de empleados
        ## Se añade al dataframe final
        for i in range(len(df)):
            FilaAgregar = {}
            Ceros_ = 5 - len(str(i +1))
            Contador = (Ceros_* "0") + str(i+1)
            FilaAgregar["TXT"] =  "02"+ Contador + str(df.iloc[i]['TXT'])
            TXT_Final = pd.concat([TXT_Final,pd.DataFrame.from_records([FilaAgregar])],ignore_index=True)

        ##Para exportar en txt
        NombreTXT_ = "TXT-" + NombreTemporal_ + "-" + Anio + "-" + Mes 
        file_name = open("./src/database/"+NombreTXT_ + ".txt", "w+", encoding="ANSI")
        Texto_ = ""
        for fila in TXT_Final["TXT"]:
            Texto_ += (str(fila)+"\n")
        
        file_name.write(Texto_)
        file_name.close()
        print(str(NombreTXT_)+".txt")
        # # to read the content of it
        # read_file = open(NombreTXT_ + ".txt", "r")
        # response = HttpResponse(read_file.read(), content_type="text/plain,charset=utf8")
        # read_file.close()

        # response['Content-Disposition'] = 'attachment; filename="{}.txt"'.format(NombreTXT_)
        # return response
#procesarTXTSS()

# TRAER GRUPOS DE SEGURIDAD SOCIAL ---------------------------------------------------------------------
def procesargrupostxt(NombreTemporal_,Anio,Mes):
    
    # URL DEL XLS DE SS
    URL = "https://creatorapp.zohopublic.com/hq5colombia/compensacionhq5/xls/TXT_SS_DESARROLLO/3HO1RZORhePyRgar44EefyEhhD27umsJE7GeJJhCDwwx2ngQ2KEHHGTCB1mYQtFktzmgSyHG2qRWsnu3ZGbW8N97TZtX709N3DAC?NOMBRE_EMPRESA=" + NombreTemporal_ + "&PENSION_ANO=" + Anio + "&PENSION_MES=" + Mes
    # URL = "https://creatorapp.zohopublic.com/hq5colombia/compensacionhq5/xls/TXT_SS_DESARROLLO/3HO1RZORhePyRgar44EefyEhhD27umsJE7GeJJhCDwwx2ngQ2KEHHGTCB1mYQtFktzmgSyHG2qRWsnu3ZGbW8N97TZtX709N3DAC?NOMBRE_EMPRESA=HQ5%20S.A.S&PENSION_ANO=2022&PENSION_MES=10"
    #CONVERTIR XLS EN DATAFRAME
    df = pd.read_excel(URL)
    df1 = pd.DataFrame(df)
    
    if(df1.empty):
        print("No existe registro")
    else:
        List_Groups = df1['Agrupación de empresa'].unique().tolist()
        print(List_Groups)
    
# Controlar las acciones del script
def controlador():
    NombreTemporal = sys.argv[1]
    NombreTemporal_ = NombreTemporal.replace(" ","%20")
    Anio = sys.argv[2]
    Mes = sys.argv[3]
    Grupo = sys.argv[4]
    
    if(Grupo == "Search_group"):
        procesargrupostxt(NombreTemporal_,Anio,Mes)
    else:
        procesarTXTSS(NombreTemporal,Anio,Mes,Grupo)
        
    
controlador()