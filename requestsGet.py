import requests
import os
from requests.exceptions import Timeout

baseURL = "https://gorest.co.in/public/v2/users"

for i in range(4):
    try:
        r = requests.get(baseURL,timeout=0.1)
        print(r.text)
        print(r.status_code)
        break
    except Timeout as e:
        print(e)

print(os.getenv('Token_Name'))