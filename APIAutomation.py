import pandas as pd
import requests
from dotenv import load_dotenv
import os

# --- 1. CONFIGURATION ---
# Load environment variables from the config.env file
load_dotenv(dotenv_path='config.env')

API_URL = os.getenv("API_URL")
API_KEY = os.getenv("API_KEY")
EXCEL_FILE_PATH = os.getenv("EXCEL_FILE_PATH")
SHEET_NAME = os.getenv("SHEET_NAME")


@retry(
    stop=stop_after_attempt(3), 
    wait=wait_fixed(2),
    retry=(retry_if_exception_type(requests.exceptions.ConnectionError) | 
           retry_if_exception_type(requests.exceptions.Timeout))
)
def post_data_to_api(payload):
    
    headers = {
        'Content-Type': 'application/json',
        'Authorization': f'Bearer {API_KEY}'
    }

    response = requests.post(API_URL, headers=headers, json=payload, timeout=5)

    return response


def main():
    
    if not all([API_URL, API_KEY, EXCEL_FILE_PATH, SHEET_NAME]):
        return
    
    try:
        df = pd.read_excel(EXCEL_FILE_PATH, sheet_name=SHEET_NAME)
    except FileNotFoundError:
        return
    except Exception as e:
        return

    for index, row in df.iterrows():

        payload = {
            "key1": row['column_one'],
            "key2": row['column_two'],
            "key3": int(row['column_three']) 
        }

        try:
            post_data_to_api(payload)
        except requests.exceptions.HTTPError as e:
            print(e)
        except Exception as e:
            print(e)

if __name__ == "__main__":
    main()

    