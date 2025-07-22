from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

SCOPES = ['https://www.googleapis.com/auth/drive']
creds = Credentials.from_service_account_file('service-account-key.json', scopes=SCOPES)
service = build('drive', 'v3', credentials=creds)

folder_id = '14HZ0hn4dxzTrI4vM2Y8xDt0PRq1EWwSz'  # 請購單附件資料夾ID
file_metadata = {
    'name': 'apitest.txt',
    'parents': [folder_id]
}
media = None  # 不上傳內容，只建空檔
try:
    file = service.files().create(body=file_metadata, media_body=media, fields='id,webViewLink').execute()
    print('建立成功:', file)
except Exception as e:
    import traceback
    print('建立失敗:', e)
    traceback.print_exc()