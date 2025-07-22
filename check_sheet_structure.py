#!/usr/bin/env python3
"""
檢查 Google Sheets 欄位結構的腳本
"""

import os
from dotenv import load_dotenv
import gspread
from google.oauth2.service_account import Credentials
import json

# 載入環境變數
load_dotenv()

def get_google_sheets_client():
    """建立Google Sheets客戶端"""
    try:
        # 從環境變數或服務帳戶金鑰檔案讀取憑證
        service_account_info = os.getenv('GOOGLE_SERVICE_ACCOUNT_INFO')
        if service_account_info:
            creds_dict = json.loads(service_account_info)
        else:
            # 如果沒有環境變數，嘗試從檔案讀取
            creds_dict = json.load(open('service-account-key.json'))
        
        creds = Credentials.from_service_account_info(creds_dict, scopes=[
            'https://www.googleapis.com/auth/spreadsheets',
            'https://www.googleapis.com/auth/drive'
        ])
        client = gspread.authorize(creds)
        return client
    except Exception as e:
        print(f"Google Sheets 客戶端建立失敗: {e}")
        return None

def check_sheet_structure():
    """檢查請購單工作表的欄位結構"""
    try:
        client = get_google_sheets_client()
        if not client:
            return
        
        spreadsheet_id = os.getenv('SPREADSHEET_ID')
        if not spreadsheet_id:
            print("未設定試算表ID")
            return
        
        spreadsheet = client.open_by_key(spreadsheet_id)
        worksheet = spreadsheet.worksheet('請購單')
        
        # 取得標題列
        headers = worksheet.row_values(1)
        print("=== 請購單工作表欄位結構 ===")
        for i, header in enumerate(headers, 1):
            print(f"{i:2d}. {header}")
        
        # 檢查最後幾筆資料
        all_values = worksheet.get_all_values()
        if len(all_values) > 1:
            print(f"\n=== 最後一筆資料 ===")
            last_row = all_values[-1]
            for i, (header, value) in enumerate(zip(headers, last_row), 1):
                print(f"{i:2d}. {header}: {value}")
        
        # 檢查附件欄位位置
        try:
            attachment_col_index = headers.index('附件') + 1
            print(f"\n附件欄位位置: 第 {attachment_col_index} 欄")
        except ValueError:
            print("\n警告: 找不到 '附件' 欄位")
            print("可用的欄位:", headers)
        
    except Exception as e:
        print(f"檢查工作表結構失敗: {e}")
        import traceback
        traceback.print_exc()

if __name__ == '__main__':
    check_sheet_structure() 