import requests

def updatedata(access_token,file_path,record_id):

    url = f"https://creator.zoho.com/api/v2/hq5colombia/hq5/report/Formato_orden_de_ingreso_Report/{record_id}/formato_file_for_ord_ing/upload"

    headers = {
        "Authorization": f"Zoho-oauthtoken {access_token}"
    }

    with open(file_path, "rb") as file:
        files = {"file": file}
        response = requests.post(url, headers=headers, files=files)

    print(file_path)