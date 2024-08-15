import requests

def patch_zoho_creator(app, report, record_id, data, token):
    url = f"https://creator.zoho.com/api/v2.1/hq5colombia/{app}/report/{report}/{record_id}"

    headers = {
        'Content-Type': 'application/json',
        'Authorization': f'Zoho-oauthtoken {token}'
    }

    try:
        response = requests.patch(url, headers=headers, json=data)

        if response.status_code != 200:
            return None

        return response.json()

    except requests.RequestException as e:
        return None
