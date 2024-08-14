import requests
import os

def get_file_type(content_type):
    types = {
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': 'xlsx',
        'application/pdf': 'pdf',
        'image/jpeg': 'jpg',
        'image/png': 'png',
        'image/gif': 'gif',
        'image/bmp': 'bmp',
        'application/msword': 'doc',
        'application/vnd.openxmlformats-officedocument.wordprocessingml.document': 'docx',
        'application/vnd.ms-excel': 'xls',
        'application/vnd.ms-powerpoint': 'ppt',
        'application/vnd.openxmlformats-officedocument.presentationml.presentation': 'pptx',
        'text/plain': 'txt',
        'application/zip': 'zip',
        'application/x-rar-compressed': 'rar',
        'application/octet-stream': 'bin'
    }
    return types.get(content_type.split(';')[0], None)

def download_file_zoho_creator(app, report, record_id, field, name, token):
    url = f"https://creator.zoho.com/api/v2.1/hq5colombia/{app}/report/{report}/{record_id}/{field}/download"
    
    headers = {
        'Authorization': f'Zoho-oauthtoken {token}'
    }

    try:
        response = requests.get(url, headers=headers)
        response.raise_for_status()

        content_type = response.headers.get('Content-Type')
        ext = get_file_type(content_type)

        if ext is None:
            print(f'Error tipo de archivo no reconocido: {response.headers}')
            return None

        file_name = f'{name}.{ext}'
        file_path = os.path.abspath(os.path.join(os.path.dirname(__file__), f'../ConciliacionBancaria/{file_name}'))

        with open(file_path, 'wb') as file:
            file.write(response.content)

        return file_path

    except requests.exceptions.RequestException as e:
        print(f'Error al descargar archivo: {e}')
        return None