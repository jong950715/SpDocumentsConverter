import json

import requests


class EcountApi:
    def __init__(self):
        pass

import base64

# Base64 디코딩
# base64_decoded_bytes = base64.b64decode('eG1gS/71VWEzF9t15vgs/K0PtTVnt3MEDLKfxSfK6YxsRZ8bonj9G6dYdJqu7R/SMItLfjm1NRrQhjxAs9GCsQ==')
# xm`KUa3\x17u,\x0f5gs\x04\x0c'lE\x1bx\x1bXt\x1f0K~95\x1aІ<@т

# # print(str(base64.b64decode(base64_decoded_bytes), encoding='utf-8'))
# base64_decoded_str = base64_decoded_bytes.decode('euc-kr', errors='ignore')
# print(base64_decoded_str)


#
# url = 'https://sboapi{0}.ecount.com/OAPI/V2/InventoryBalance/GetListInventoryBalanceStatus?SESSION_ID={1}'.format('CB','3635373435347c535036393932:CB-AR9bEtF7GVVnb')
# headers = {'Content-Type': 'application/json'}
# # data = {'COM_CODE': 657454,
# #         'USER_ID': 'sp6992',
# #         'API_CERT_KEY': '5',
# #         'LAN_TYPE': 'ko-KR',
# #         'ZONE': 'CB', }
# data = {
#     'SESSION_ID': '3635373435347c535036393932:CB-AR9bEtF7GVVnb',
#     'BASE_DATE': '20240423',
# }
#
# res = requests.post(url=url, headers=headers, json=data)
# print(res.text)
# print(json.dumps(res.json(), indent=4, ensure_ascii=False))
#
