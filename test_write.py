import gspread
from google.oauth2.service_account import Credentials

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SPREADSHEET_ID = '1ZB6ri0fzqTRk_ciHibcGXEuViNmW1Ag9kkazE8A5iKc'
SHEET_NAME = '請購單'

def get_google_sheets_client():
    creds = Credentials.from_service_account_file('service-account-key.json', scopes=SCOPES)
    return gspread.authorize(creds)

client = get_google_sheets_client()
worksheet = client.open_by_key(SPREADSHEET_ID).worksheet(SHEET_NAME)

worksheet.append_row(['測試單號', '20250720', '總務部', '測試人', 'test@mail.com'])
print('寫入成功')