import requests

def updatedata(access_token,file_path,record_id):

    url = f"https://creator.zoho.com/api/v2/hq5colombia/compensacionhq5/report/Plantillas_Report/{record_id}/plantilla_plan/upload"

    headers = {
        "Authorization": f"Zoho-oauthtoken {access_token}"
    }

    with open(file_path, "rb") as file:
        files = {"file": file}
        response = requests.post(url, headers=headers, files=files)

    print(file_path)