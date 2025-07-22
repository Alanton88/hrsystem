#!/usr/bin/env python3
"""
Google OAuth 測試腳本
用於測試 OAuth 設定是否正確
"""

from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
import os

# OAuth 範圍
SCOPES = ['https://www.googleapis.com/auth/drive.file']

def test_oauth():
    """測試 OAuth 流程"""
    try:
        print("開始 OAuth 測試...")
        
        # 使用 OAuth 客戶端設定檔案
        flow = InstalledAppFlow.from_client_secrets_file(
            'client_secret_530091699052-b5t9bbrevaodte3fi0u4n5urd9l6rhvs.apps.googleusercontent.com.json', 
            SCOPES
        )
        
        print("正在啟動本地伺服器進行授權...")
        creds = flow.run_local_server(port=55551)
        
        print("授權成功！正在測試 Google Drive API...")
        
        # 測試 Google Drive API
        service = build('drive', 'v3', credentials=creds)
        
        # 列出根目錄的檔案（測試 API 存取）
        results = service.files().list(pageSize=5, fields="nextPageToken, files(id, name)").execute()
        files = results.get('files', [])
        
        if not files:
            print('未找到檔案。')
        else:
            print('檔案列表:')
            for file in files:
                print(f"- {file['name']} ({file['id']})")
        
        print("OAuth 測試成功！")
        return True
        
    except Exception as e:
        print(f"OAuth 測試失敗: {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == '__main__':
    test_oauth() 