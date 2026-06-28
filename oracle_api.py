import requests

BASE_URL = "https://fa-euth-saasfademo1.ds-fa.oraclepdemos.com"

USERNAME = "scm01.student"

PASSWORD = "L%b3V^M9"


def get_planners():

    url = BASE_URL + "/fscmRestApi/resources/11.13.18.05/planners"

    response = requests.get(
        url,
        auth=(USERNAME, PASSWORD),
        headers={
            "Accept": "application/json"
        }
    )

    return response

def test_erp_integration():

    url = BASE_URL + "/fscmRestApi/resources/11.13.18.05/erpintegrations"

    response = requests.options(
        url,
        auth=(USERNAME, PASSWORD),
        headers={
            "Accept": "application/json"
        }
    )

    return response

import base64


def upload_fbdi_zip():

    zip_file = "output/ScpPlannersImport.zip"

    with open(zip_file, "rb") as f:
        file_data = base64.b64encode(f.read()).decode()

    url = BASE_URL + "/fscmRestApi/resources/11.13.18.05/erpintegrations"

    payload = {
        "OperationName": "uploadFileToUCM",
        "DocumentName": "ScpPlannersImport.zip",
        "Content": file_data,
        "ContentType": "zip"
    }

    response = requests.post(
        url,
        json=payload,
        auth=(USERNAME, PASSWORD),
        headers={
            "Content-Type": "application/json",
            "Accept": "application/json"
        }
    )

    return response



