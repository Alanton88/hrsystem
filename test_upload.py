#!/usr/bin/env python3
"""
測試檔案上傳功能的腳本
"""

import os
from dotenv import load_dotenv
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

# 載入環境變數
load_dotenv()

def test_file_upload():
    """測試檔案上傳功能"""
    try:
        print("開始測試檔案上傳...")
        
        # 使用 OAuth 2.0 流程
        SCOPES = ['https://www.googleapis.com/auth/drive.file']
        
        flow = InstalledAppFlow.from_client_secrets_file(
            'client_secret_530091699052-b5t9bbrevaodte3fi0u4n5urd9l6rhvs.apps.googleusercontent.com.json', 
            SCOPES
        )
        
        print("正在啟動本地伺服器進行授權...")
        creds = flow.run_local_server(port=55551)
        service = build('drive', 'v3', credentials=creds)
        
        # 創建一個測試檔案
        test_content = "這是一個測試檔案，用於測試 Google Drive 上傳功能。"
        test_filename = "test_upload.txt"
        
        with open(test_filename, 'w', encoding='utf-8') as f:
            f.write(test_content)
        
        print(f"創建測試檔案: {test_filename}")
        
        # 上傳檔案
        folder_id = os.getenv('PURCHASE_REQUEST_ATTACHMENT_FOLDER_ID')
        if not folder_id:
            print("未設定資料夾ID")
            return
        
        print(f"上傳到資料夾: {folder_id}")
        
        file_metadata = {
            'name': test_filename,
            'parents': [folder_id]
        }
        
        media = MediaFileUpload(test_filename, resumable=True)
        file = service.files().create(
            body=file_metadata, 
            media_body=media, 
            fields='id,webViewLink,webContentLink'
        ).execute()
        
        print("檔案上傳成功！")
        print(f"檔案ID: {file.get('id')}")
        print(f"webViewLink: {file.get('webViewLink')}")
        print(f"webContentLink: {file.get('webContentLink')}")
        
        # 設定檔案權限
        try:
            permission = {
                'type': 'anyone',
                'role': 'reader'
            }
            service.permissions().create(
                fileId=file.get('id'),
                body=permission,
                fields='id'
            ).execute()
            print("檔案權限設定成功")
        except Exception as perm_error:
            print(f"設定檔案權限失敗: {perm_error}")
        
        # 清理測試檔案
        os.remove(test_filename)
        
        return file.get('webViewLink')
        
    except Exception as e:
        print(f"測試失敗: {e}")
        import traceback
        traceback.print_exc()
        return None

if __name__ == '__main__':
    result = test_file_upload()
    if result:
        print(f"\n測試成功！檔案連結: {result}")
    else:
        print("\n測試失敗！") 