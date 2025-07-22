#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
請購單資料一致性檢查工具
用於診斷為什麼驗收單作業區顯示的數量與預期不符
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

def get_safe_records(worksheet):
    """安全地從工作表取得記錄"""
    try:
        return worksheet.get_all_records()
    except Exception as e:
        if "duplicates" in str(e):
            print("檢測到重複標題，使用手動標題設定")
            headers = worksheet.row_values(1)
            cleaned_headers = []
            for i, header in enumerate(headers):
                if header:
                    cleaned_headers.append(header)
                else:
                    cleaned_headers.append(f'Column_{i+1}')
            
            all_values = worksheet.get_all_values()
            if len(all_values) > 1:
                data_rows = all_values[1:]
                all_records = []
                for row in data_rows:
                    record = {}
                    for i, value in enumerate(row):
                        if i < len(cleaned_headers):
                            record[cleaned_headers[i]] = value
                    all_records.append(record)
                return all_records
            else:
                return []
        else:
            raise e

def analyze_purchase_data():
    """分析請購單資料"""
    print("=== 請購單資料一致性檢查 ===\n")
    
    try:
        client = get_google_sheets_client()
        if not client:
            print("❌ 無法建立 Google Sheets 客戶端")
            return
        
        spreadsheet_id = os.getenv('SPREADSHEET_ID', '1ZB6ri0fzqTRk_ciHibcGXEuViNmW1Ag9kkazE8A5iKc')
        spreadsheet = client.open_by_key(spreadsheet_id)
        worksheet = spreadsheet.worksheet('請購單')
        
        # 取得所有資料
        all_records = get_safe_records(worksheet)
        print(f"📊 總記錄數: {len(all_records)}")
        
        if not all_records:
            print("❌ 沒有找到任何記錄")
            return
        
        # 顯示欄位名稱
        print(f"\n📋 欄位名稱: {list(all_records[0].keys())}")
        
        # 分析簽核狀態
        approval_fields = ['簽核', '簽核狀態', 'approval_status', '狀態']
        approval_statuses = {}
        approved_records = []
        pending_receipt_records = []
        
        print(f"\n🔍 分析簽核狀態...")
        
        for i, record in enumerate(all_records):
            purchase_no = record.get('請購單號', f'Record_{i+1}')
            
            # 檢查所有可能的簽核欄位
            approval_status = None
            used_field = None
            for field in approval_fields:
                if record.get(field):
                    approval_status = record.get(field)
                    used_field = field
                    break
            
            if not approval_status:
                approval_status = '未設定'
                used_field = '無'
            
            approval_statuses[purchase_no] = {
                'status': approval_status,
                'field': used_field,
                'record': record
            }
            
            # 檢查是否為已核准
            if approval_status in ['核准', 'approved', 'APPROVED', 'Approved', '已核准', '已批准']:
                approved_records.append(record)
                
                # 檢查驗收單驗收狀態
                receipt_status = (
                    record.get('驗收單狀態') or 
                    record.get('驗收狀態') or 
                    record.get('receipt_status') or 
                    '待驗收'
                )
                
                if receipt_status == '待驗收':
                    pending_receipt_records.append(record)
        
        # 顯示統計結果
        print(f"\n📈 統計結果:")
        print(f"   ✅ 已核准記錄: {len(approved_records)} 筆")
        print(f"   ⏳ 待驗收記錄: {len(pending_receipt_records)} 筆")
        
        # 顯示簽核狀態分布
        print(f"\n📊 簽核狀態分布:")
        status_counts = {}
        field_counts = {}
        
        for purchase_no, info in approval_statuses.items():
            status = info['status']
            field = info['field']
            
            status_counts[status] = status_counts.get(status, 0) + 1
            field_counts[field] = field_counts.get(field, 0) + 1
        
        print("   簽核狀態:")
        for status, count in sorted(status_counts.items()):
            print(f"     {status}: {count} 筆")
        
        print("   使用的欄位:")
        for field, count in sorted(field_counts.items()):
            print(f"     {field}: {count} 筆")
        
        # 顯示待驗收的詳細資料
        print(f"\n🔍 待驗收請購單詳細資料:")
        for record in pending_receipt_records:
            purchase_no = record.get('請購單號', 'N/A')
            department = record.get('請購部門', 'N/A')
            applicant = record.get('申請人', 'N/A')
            item_name = record.get('品名', 'N/A')
            approval_status = record.get('簽核', record.get('簽核狀態', 'N/A'))
            
            print(f"   📋 {purchase_no} | {department} | {applicant} | {item_name} | 簽核: {approval_status}")
        
        # 顯示已核准但非待驗收的記錄
        non_pending_approved = [r for r in approved_records if r not in pending_receipt_records]
        if non_pending_approved:
            print(f"\n⚠️  已核准但非待驗收的記錄 ({len(non_pending_approved)} 筆):")
            for record in non_pending_approved:
                purchase_no = record.get('請購單號', 'N/A')
                receipt_status = (
                    record.get('驗收單狀態') or 
                    record.get('驗收狀態') or 
                    record.get('receipt_status') or 
                    '待驗收'
                )
                print(f"   �� {purchase_no} | 驗收單驗收狀態: {receipt_status}")
        
        # 問題診斷
        print(f"\n🔧 問題診斷:")
        if len(pending_receipt_records) != 13:
            print(f"   ❌ 待驗收記錄數量不符預期 (預期: 13, 實際: {len(pending_receipt_records)})")
            
            # 檢查是否有記錄被錯誤分類
            all_pending = [r for r in all_records if (
                r.get('驗收單狀態') == '待驗收' or 
                r.get('驗收狀態') == '待驗收' or 
                r.get('receipt_status') == '待驗收'
            )]
            print(f"   📊 所有標記為待驗收的記錄: {len(all_pending)} 筆")
            
            # 檢查簽核狀態為空但驗收狀態為待驗收的記錄
            empty_approval_pending = [r for r in all_records if (
                not r.get('簽核') and 
                not r.get('簽核狀態') and 
                not r.get('approval_status') and 
                not r.get('狀態') and
                (r.get('驗收單狀態') == '待驗收' or 
                 r.get('驗收狀態') == '待驗收' or 
                 r.get('receipt_status') == '待驗收')
            )]
            if empty_approval_pending:
                print(f"   ⚠️  簽核狀態為空但驗收單驗收狀態為待驗收的記錄: {len(empty_approval_pending)} 筆")
                for record in empty_approval_pending:
                    print(f"     📋 {record.get('請購單號', 'N/A')}")
        else:
            print(f"   ✅ 待驗收記錄數量正確")
        
        print(f"\n💡 建議:")
        print(f"   1. 檢查簽核狀態欄位是否統一使用 '簽核' 或 '簽核狀態'")
        print(f"   2. 確保所有已核准的請購單都有正確的簽核狀態值")
        print(f"   3. 檢查是否有重複記錄或資料不一致的問題")
        print(f"   4. 確認驗收單驗收狀態欄位名稱是否統一")
        
    except Exception as e:
        print(f"❌ 分析過程發生錯誤: {e}")

if __name__ == "__main__":
    analyze_purchase_data() 