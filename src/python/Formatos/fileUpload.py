import requests

def Updatedata(access_token,file_path,record_id,report,field):

    url = f"https://creator.zoho.com/api/v2/hq5colombia/hq5/report/{report}/{record_id}/{field}/upload"

    headers = {
        "Authorization": f"Zoho-oauthtoken {access_token}"
    }

    with open(file_path, "rb") as file:
        files = {"file": file}
        response = requests.post(url, headers=headers, files=files)

    print(file_path)