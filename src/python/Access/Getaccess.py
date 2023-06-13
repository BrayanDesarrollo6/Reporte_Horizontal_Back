import requests

def obtener_access_token():
    url = 'https://accounts.zoho.com/oauth/v2/token'
    data = {
        'grant_type': 'refresh_token',
        'client_id': '1000.IIM2A185O6YWU0SVCV5SU8N1WADV5O',
        'client_secret': '3fa85b34e476b4acb29ab2d8154fc16876c0e14fe7',
        'refresh_token':'1000.bf3f1b541c296f7b0000b62c5a4320cd.49dbce948d23d77dcb0bd39449f96a18'
    }
    response = requests.post(url, data=data)
    if response.status_code == 200:
        token_data = response.json()
        access_token = token_data['access_token']
        return access_token
    else:
        return None