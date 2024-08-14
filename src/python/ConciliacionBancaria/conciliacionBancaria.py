import sys
import json
import os

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from Access.Getaccess import obtener_access_token
from downloadFile import download_file_zoho_creator as download_zoho_file
from getData import procesar_excel
from solver import obtener_mejor_combinacion
from saveResults import guardar_resultados
from uploadFile import upload_file_zoho_creator as upload_zoho_file

def main():
    req_data = json.loads(sys.argv[1])
    
    app = req_data["app"]
    report = req_data["report"]
    record_id = req_data["recordId"]
    field_name = req_data["fieldName"]
    file_name = req_data["fileName"]
    value_min = req_data["valueMin"]
    
    token = obtener_access_token()
    file_path = download_zoho_file(app, report, record_id, field_name, file_name, token)
    values, objective = procesar_excel(file_path)
    
    if values is not None and objective is not None:
        resultado, diferencia = obtener_mejor_combinacion(values, objective, 100, value_min)
        guardar_resultados(file_path, resultado, diferencia, values, objective)
        access_token = obtener_access_token()
        file_path = upload_zoho_file(app, report, record_id, field_name, access_token, file_path)
        print(file_path)

    else:
        print("Error no se pudieron obtener los datos del archivo")

if __name__ == "__main__":
    main()