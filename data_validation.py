#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
資料驗證工具
用於驗證不同頁面顯示的資料是否一致
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

def validate_data_consistency():
    """驗證資料一致性"""
    try:
        client = get_google_sheets_client()
        if not client:
            print("無法建立 Google Sheets 客戶端")
            return
        
        spreadsheet_id = os.getenv('SPREADSHEET_ID')
        if not spreadsheet_id:
            print("未設定試算表ID")
            return
        
        spreadsheet = client.open_by_key(spreadsheet_id)
        worksheet = spreadsheet.worksheet('請購單')
        
        # 取得所有資料
        all_records = worksheet.get_all_records()
        
        print("資料一致性驗證報告")
        print("=" * 80)
        
        # 檢查資料完整性
        print(f"總記錄數: {len(all_records)}")
        
        # 檢查必要欄位
        required_fields = ['請購單號', '請購日期', '請購部門', '申請人', '品名', '數量', '單位', '需求日期', '請購單簽核']
        missing_fields = []
        
        for field in required_fields:
            missing_count = sum(1 for record in all_records if not record.get(field))
            if missing_count > 0:
                missing_fields.append((field, missing_count))
        
        if missing_fields:
            print("\n發現缺失欄位:")
            for field, count in missing_fields:
                print(f"  {field}: {count} 筆記錄缺失")
        else:
            print("\n所有必要欄位都有資料")
        
        # 檢查資料格式一致性
        print("\n資料格式檢查:")
        
        # 檢查請購單號格式
        purchase_no_formats = set()
        for record in all_records:
            purchase_no = record.get('請購單號', '')
            if purchase_no:
                # 檢查格式是否為 YYYYMMDD-XXX
                if len(purchase_no) >= 8 and '-' in purchase_no:
                    format_type = "標準格式 (YYYYMMDD-XXX)"
                else:
                    format_type = "非標準格式"
                purchase_no_formats.add(format_type)
        
        print(f"  請購單號格式: {', '.join(purchase_no_formats)}")
        
        # 檢查部門名稱
        departments = set()
        for record in all_records:
            dept = record.get('請購部門', '')
            if dept:
                departments.add(dept)
        
        print(f"  請購部門: {', '.join(sorted(departments))}")
        
        # 檢查申請人
        applicants = set()
        for record in all_records:
            applicant = record.get('申請人', '')
            if applicant:
                applicants.add(applicant)
        
        print(f"  申請人: {', '.join(sorted(applicants))}")
        
        # 檢查簽核狀態
        approval_statuses = set()
        for record in all_records:
            status = record.get('請購單簽核', '')
            if status:
                approval_statuses.add(status)
        
        print(f"  簽核狀態: {', '.join(sorted(approval_statuses))}")
        
        # 檢查已核准的請購單
        approved_records = [r for r in all_records if r.get('請購單簽核') == '核准']
        print(f"\n已核准請購單: {len(approved_records)} 筆")
        
        # 檢查特定請購單號的資料
        target_purchase_no = "20250718-001"
        target_records = [r for r in all_records if r.get('請購單號') == target_purchase_no]
        
        if target_records:
            print(f"\n請購單號 {target_purchase_no} 的資料驗證:")
            print("-" * 60)
            record = target_records[0]  # 應該只有一筆
            print(f"  請購單號: {record.get('請購單號', 'N/A')}")
            print(f"  請購日期: {record.get('請購日期', 'N/A')}")
            print(f"  請購部門: {record.get('請購部門', 'N/A')}")
            print(f"  申請人: {record.get('申請人', 'N/A')}")
            print(f"  品名: {record.get('品名', 'N/A')}")
            print(f"  規格: {record.get('規格', 'N/A')}")
            print(f"  數量: {record.get('數量', 'N/A')}")
            print(f"  單位: {record.get('單位', 'N/A')}")
            print(f"  需求日期: {record.get('需求日期', 'N/A')}")
            print(f"  簽核狀態: {record.get('請購單簽核', 'N/A')}")
            print(f"  驗收單驗收狀態: {record.get('驗收狀態', 'N/A')}")
        else:
            print(f"\n找不到請購單號 {target_purchase_no}")
        
        # 資料一致性建議
        print("\n資料一致性建議:")
        print("-" * 60)
        print("1. 確保所有頁面都從同一個資料來源讀取資料")
        print("2. 避免在前端硬編碼資料")
        print("3. 使用 API 調用取得真實資料")
        print("4. 定期檢查資料完整性")
        print("5. 建立資料驗證機制")
        
    except Exception as e:
        print(f"資料驗證時發生錯誤: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    validate_data_consistency() 