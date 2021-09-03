from xlwings.rest.api import api
import pandas as pd
import requests, json
base_url = 'http://192.168.0.5:5000'
endpoint = '/book/TableData.xlsm/names/akhil/range'
r = requests.get(base_url+endpoint)
data = r.json()
df = pd.DataFrame(data['formula'][0:], columns=data['formula'[0]
df