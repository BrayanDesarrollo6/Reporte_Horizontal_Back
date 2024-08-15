import requests
from updateRecord import patch_zoho_creator

def upload_file_zoho_creator(app, report, record_id, field, token, file_path):
    url = f"https://creator.zoho.com/api/v2/hq5colombia/{app}/report/{report}/{record_id}/{field}/upload"

    headers = {
        "Authorization": f"Zoho-oauthtoken {token}"
    }
    
    try:
        with open(file_path, "rb") as file:
            files = { "file": file }
        
            response = requests.post(url, headers=headers, files=files)
            response.raise_for_status()
            
            patch_zoho_creator(app, report, record_id, { "data": { "estado_con_ban": "Completado"} }, token)
            
            return file_path
        
    except requests.exceptions.RequestException as e:
        print(f'Error al cargar archivo: {e}')
        return None

    