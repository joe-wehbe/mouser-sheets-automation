import gspread # Allows interaction with Google Sheets via its API
from google.oauth2.service_account import Credentials # Allows authentication using a service account
import os # Used to get environment variabled
from dotenv import load_dotenv # Used to load environment variables 
import requests # Used for making HTTP requests

load_dotenv() # Loading environment variables
SEARCH_API_KEY = os.getenv("SEARCH_API_KEY") # Geting the SEARCH_API_KEY from .env
BASE_URL = "https://api.mouser.com/api/v2/" # Storing the base URL in a variable

# ===== CONNECTING TO GOOGLE SHEETS =====
def get_google_sheet(sheet_id: str, credentials: str):
    scopes = ["https://www.googleapis.com/auth/spreadsheets"] # Access scope for the Google Sheets API
    creds = Credentials.from_service_account_file(credentials, scopes=scopes) # Storing the credentials from credentials.json in a variable
    client = gspread.authorize(creds) # Authorizing the gspread client to interact with Google Sheets
    sheet = client.open_by_key(sheet_id) # Opening the sheet using its unique ID
    return sheet.sheet1 # Returning the first sheet

# ===== FILLING THE COLUMNS IN THE SHEET =====
def fill_sheet(worksheet):
    part_numbers = worksheet.col_values(1)[1:] # Stores all the part numbers (first column) in a variable, except the header
    manufacturers, categories, descriptions = [], [], [] # Lists to store data for each processed part number before uploading them to the sheet

    for index, part_number in enumerate(part_numbers, start=2): # Iterating over part numbers (the first column)
        manufacturer, category, description = process_part_number(part_number) # Processing each part number (function below) and storing the result in variables

        manufacturers.append([manufacturer]) # Adding the obtained manufacturer to the array 
        categories.append([category]) # Adding the obtained category to the array
        descriptions.append([description]) # Adding the obtained description to the array

        print(f"({index - 1}/{len(part_numbers)}) Processed part number: {part_number}")

    if manufacturers and categories and descriptions:
        print("Updating Google Sheet...")
        worksheet.update(manufacturers, range_name=f"B2:B{len(part_numbers)+1}") # Adding manufacturers array to column B
        worksheet.update(categories, range_name=f"C2:C{len(part_numbers)+1}") # Adding categories array to column C
        worksheet.update(descriptions, range_name=f"D2:D{len(part_numbers)+1}") # Adding descriptions array to column D
        print("Update Successful")

# ===== HELPER METHOD THAT FETCHES DATA USING MOUSER SEARCH API =====
def process_part_number(part_number):
    params = {"apiKey": SEARCH_API_KEY} # Setting API key in params variable
    payload = {"SearchByPartRequest": {"mouserPartNumber": part_number}} # Using mouserPartNumber to search by part number

    response = requests.post(BASE_URL + "search/partnumber", params=params, json=payload) # Sending the request and storing the response in a variable

    if response.status_code == 200: # If status code is OK
        try:
            data = response.json() # Storing the respinse in json format in a variable to access it
            part = data["SearchResults"]["Parts"][0] # Getting the first part, which is the most relevant search result
            return [part.get("Manufacturer", "N/A"), part.get("Category", "N/A"), part.get("Description", "N/A")] # Getting values of manufacturer, category, description or N/A as a fallback value
        except (IndexError, KeyError): # Handling potential errors
            return ["N/A", "N/A", "N/A"] # Returning N/A if IndexError or KeyError occured
    return ["N/A", "N/A", "N/A"] # Returning N/A if the status code is not 200

# ===== MAIN PROGRAM =====
if __name__ == "__main__":
    sheet_id = "1HFwrHn94KpDu287DEjAbL8Y2ujM4oWFlg41H3qznz60"
    credentials = "credentials.json"

    sheet = get_google_sheet(sheet_id, credentials)
    fill_sheet(sheet)