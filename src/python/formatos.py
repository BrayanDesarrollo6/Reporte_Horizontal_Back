import sys
import json
from Access.Getaccess import *
import Formatos.formatoOrdenIngreso as process_orden_ingreso
import Formatos.formatoOrdenDHL as process_orden_dhl
from Formatos.fileUploadOrdenIngreso import updatedata
from Formatos.fileUpload import Updatedata

def procesar_formato(data):
    
    try:
        json_object = json.loads(data)
        formato = json_object['data']['formato']['nombre']
        record_id = json_object['data']['ID_generado']

        formatos = {
            "orden_ingreso": process_orden_ingreso,
            "orden_dhl": process_orden_dhl,
        }

        if formato in formatos:
            modulo_formato = formatos[formato]
            
            # Procesamiento y creación del archivo
            file_path = modulo_formato.process_json_data(data)
            
            access_token = obtener_access_token()
        
            if access_token:
                    if formato == "orden_ingreso":
                        updatedata(access_token,file_path,record_id)
                    elif formato == "orden_dhl":
                        report = json_object['data']['formato']['reporte']
                        field = json_object['data']['formato']['campo']
                        Updatedata(access_token,file_path,record_id,report,field)
            
        else:
            print(f"Error: Módulo no encontrado para el formato {formato}")

    except json.JSONDecodeError as e:
        print(f"Error al decodificar JSON: {e}")
    except Exception as e:
        print(f"Error inesperado: {e}")

if __name__ == "__main__":
    INFO_RECEIVED = sys.argv[1]
    procesar_formato(INFO_RECEIVED)