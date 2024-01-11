import sys
import json
from Access.Getaccess import *
import Formatos.formatoOrdenIngreso as process_orden_ingreso
from Formatos.fileUploadOrdenIngreso import *

def procesar_formato(data):
    
    try:
        json_object = json.loads(data)
        formato = json_object['data']['formato']['nombre']
        record_id = json_object['data']['ID_generado']

        formatos = {
            "orden_ingreso": process_orden_ingreso,
        }

        if formato in formatos:
            modulo_formato = formatos[formato]
            
            # Procesamiento y creación del archivo
            file_path = modulo_formato.process_json_data(data)
            
            access_token = obtener_access_token()
        
            if access_token:
                if(access_token != None):
                    updatedata(access_token,file_path,record_id)
            
        else:
            print(f"Error: Módulo no encontrado para el formato {formato}")

    except json.JSONDecodeError as e:
        print(f"Error al decodificar JSON: {e}")
    except Exception as e:
        print(f"Error inesperado: {e}")

if __name__ == "__main__":
    INFO_RECEIVED = sys.argv[1]
    procesar_formato(INFO_RECEIVED)