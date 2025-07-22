#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
測試使用者資訊功能
"""

import os
import json
from dotenv import load_dotenv
import gspread
from google.oauth2.service_account import Credentials

# 載入環境變數
load_dotenv()

def get_google_sheets_client():
    """建立Google Sheets客戶端"""
    try:
        service_account_info = os.getenv('GOOGLE_SERVICE_ACCOUNT_INFO')
        if service_account_info:
            creds_dict = json.loads(service_account_info)
        else:
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

def get_user_info(username):
    """依帳號取得姓名與mail"""
    try:
        client = get_google_sheets_client()
        spreadsheet_id = os.getenv('SPREADSHEET_ID')
        spreadsheet = client.open_by_key(spreadsheet_id)
        worksheet = spreadsheet.worksheet('使用者帳號')
        records = worksheet.get_all_records()
        for r in records:
            if r.get('帳號') == username:
                return {
                    'name': r.get('姓名', ''),
                    'mail': r.get('mail', '')
                }
        return {'name': '', 'mail': ''}
    except Exception as e:
        print(f"取得user info失敗: {e}")
        return {'name': '', 'mail': ''}

def test_user_info():
    """測試使用者資訊功能"""
    print("使用者資訊測試")
    print("=" * 50)
    
    # 測試不同的使用者帳號
    test_users = ['admin', 'user1', 'test']
    
    for username in test_users:
        user_info = get_user_info(username)
        print(f"帳號: {username}")
        print(f"  姓名: {user_info.get('name', 'N/A')}")
        print(f"  信箱: {user_info.get('mail', 'N/A')}")
        print()
    
    # 測試驗收人員自動填入功能
    print("驗收人員自動填入測試:")
    print("-" * 30)
    
    # 模擬登入者為 admin
    current_user = 'admin'
    user_info = get_user_info(current_user)
    current_user_name = user_info.get('name', current_user)
    
    print(f"登入者帳號: {current_user}")
    print(f"登入者姓名: {current_user_name}")
    print(f"驗收人員欄位預設值: {current_user_name}")
    
    # 如果沒有姓名，使用帳號
    if not current_user_name:
        current_user_name = current_user
        print(f"使用帳號作為驗收人員: {current_user_name}")

if __name__ == "__main__":
    test_user_info() 