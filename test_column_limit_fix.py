#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
測試欄位限制修復
驗證系統是否正確處理 Google Sheets 的欄位限制問題
"""

import requests
import json
import os
from datetime import datetime

# 測試配置
BASE_URL = "http://127.0.0.1:5000"
TEST_USERNAME = "admin"
TEST_PASSWORD = "admin123"

def test_login():
    """測試登入功能"""
    print("🔐 測試登入功能...")
    
    login_data = {
        'username': TEST_USERNAME,
        'password': TEST_PASSWORD
    }
    
    try:
        response = requests.post(f"{BASE_URL}/login", data=login_data)
        if response.status_code == 200:
            print("✅ 登入成功")
            return True
        else:
            print(f"❌ 登入失敗: {response.status_code}")
            return False
    except Exception as e:
        print(f"❌ 登入錯誤: {e}")
        return False

def test_update_receipt_status():
    """測試更新驗收單驗收狀態"""
    print("\n✍️ 測試更新驗收單驗收狀態...")
    
    # 測試資料
    test_data = {
        'purchase_no': '20250718-004',  # 使用您提到的請購單號
        'receipt_status': '已驗收'
    }
    
    try:
        response = requests.post(
            f"{BASE_URL}/update-receipt-status",
            headers={'Content-Type': 'application/json'},
            json=test_data
        )
        
        if response.status_code == 200:
            result = response.json()
            if result.get('success'):
                print("✅ 驗收單驗收狀態更新成功")
                print(f"   驗收人員: {result.get('receipt_person', 'N/A')}")
                print(f"   驗收日期: {result.get('receipt_date', 'N/A')}")
                return True
            else:
                print(f"❌ 驗收單驗收狀態更新失敗: {result.get('message', '未知錯誤')}")
                return False
        else:
            print(f"❌ 請求失敗: {response.status_code}")
            return False
    except Exception as e:
        print(f"❌ 更新驗收單驗收狀態錯誤: {e}")
        return False

def test_update_receipt_approval():
    """測試更新驗收簽核狀態"""
    print("\n✍️ 測試更新驗收簽核狀態...")
    
    # 測試資料
    test_data = {
        'purchase_no': '20250718-004',  # 使用您提到的請購單號
        'approval_status': '核准',
        'approval_date': datetime.now().strftime('%Y%m%d'),
        'approver': '測試人員',
        'approval_note': '測試欄位限制修復'
    }
    
    try:
        response = requests.post(
            f"{BASE_URL}/update-receipt-approval",
            headers={'Content-Type': 'application/json'},
            json=test_data
        )
        
        if response.status_code == 200:
            result = response.json()
            if result.get('success'):
                print("✅ 驗收簽核狀態更新成功")
                print(f"   簽核人員: {result.get('approver', 'N/A')}")
                print(f"   簽核日期: {result.get('approval_date', 'N/A')}")
                print(f"   是否唯讀: {result.get('is_readonly', 'N/A')}")
                return True
            else:
                print(f"❌ 驗收簽核狀態更新失敗: {result.get('message', '未知錯誤')}")
                return False
        else:
            print(f"❌ 請求失敗: {response.status_code}")
            return False
    except Exception as e:
        print(f"❌ 更新驗收簽核狀態錯誤: {e}")
        return False

def test_get_purchase_detail():
    """測試取得請購單詳細資料"""
    print("\n🔍 測試取得請購單詳細資料...")
    
    try:
        response = requests.get(f"{BASE_URL}/purchase-detail/20250718-004")
        if response.status_code == 200:
            result = response.json()
            if result.get('success'):
                data = result.get('data', {})
                print("✅ 成功取得請購單詳細資料")
                print(f"   請購單號: {data.get('請購單號', 'N/A')}")
                print(f"   驗收單驗收狀態: {data.get('驗收單狀態', 'N/A')}")
                print(f"   驗收人員: {data.get('驗收人員', 'N/A')}")
                print(f"   驗收日期: {data.get('驗收日期', 'N/A')}")
                print(f"   驗收簽核狀態: {data.get('驗收簽核狀態', 'N/A')}")
                print(f"   編輯狀態: {data.get('編輯狀態', 'N/A')}")
                return True
            else:
                print(f"❌ 取得詳細資料失敗: {result.get('message', '未知錯誤')}")
                return False
        else:
            print(f"❌ 請求失敗: {response.status_code}")
            return False
    except Exception as e:
        print(f"❌ 取得詳細資料錯誤: {e}")
        return False

def test_debug_data_consistency():
    """測試資料一致性檢查"""
    print("\n🔍 測試資料一致性檢查...")
    
    try:
        response = requests.get(f"{BASE_URL}/debug-receipt-data")
        if response.status_code == 200:
            result = response.json()
            if result.get('success'):
                print("✅ 成功取得資料一致性檢查結果")
                print(f"   總記錄數: {result.get('total_records', 'N/A')}")
                print(f"   已核准記錄: {result.get('approved_count', 'N/A')}")
                
                # 檢查欄位資訊
                columns = result.get('columns', [])
                print(f"   欄位數量: {len(columns)}")
                print(f"   欄位列表: {columns}")
                
                # 檢查是否有唯讀記錄
                approval_analysis = result.get('approval_analysis', {})
                readonly_count = 0
                for purchase_no, info in approval_analysis.items():
                    if info.get('edit_status') == '唯讀':
                        readonly_count += 1
                        print(f"   唯讀記錄: {purchase_no}")
                
                print(f"   唯讀記錄數: {readonly_count}")
                return True
            else:
                print(f"❌ 資料一致性檢查失敗: {result.get('message', '未知錯誤')}")
                return False
        else:
            print(f"❌ 請求失敗: {response.status_code}")
            return False
    except Exception as e:
        print(f"❌ 資料一致性檢查錯誤: {e}")
        return False

def main():
    """主測試函數"""
    print("🚀 開始測試欄位限制修復")
    print("=" * 60)
    
    # 檢查服務是否運行
    try:
        response = requests.get(f"{BASE_URL}/")
        if response.status_code != 200:
            print("❌ Flask 應用程式未運行，請先啟動 app.py")
            return
    except:
        print("❌ 無法連接到 Flask 應用程式，請先啟動 app.py")
        return
    
    # 執行測試
    tests = [
        test_login,
        test_update_receipt_status,
        test_update_receipt_approval,
        test_get_purchase_detail,
        test_debug_data_consistency
    ]
    
    passed = 0
    total = len(tests)
    
    for test in tests:
        try:
            if test():
                passed += 1
        except Exception as e:
            print(f"❌ 測試執行錯誤: {e}")
    
    print("\n" + "=" * 60)
    print(f"📊 測試結果: {passed}/{total} 通過")
    
    if passed == total:
        print("🎉 所有測試通過！欄位限制修復成功")
    else:
        print("⚠️ 部分測試失敗，請檢查系統設定")

if __name__ == "__main__":
    main() 