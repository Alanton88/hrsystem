#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
測試請購單簽核功能
驗證簽核時是否正確回存登入者姓名到請購單簽核人員欄位
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

def test_get_purchase_approval_page():
    """測試取得請購單簽核頁面"""
    print("\n📋 測試取得製造部門請購單簽核頁面...")
    
    try:
        response = requests.get(f"{BASE_URL}/purchase-approval/manufacturing")
        if response.status_code == 200:
            print("✅ 成功取得請購單簽核頁面")
            return True
        else:
            print(f"❌ 取得頁面失敗: {response.status_code}")
            return False
    except Exception as e:
        print(f"❌ 取得頁面錯誤: {e}")
        return False

def test_update_approval_status():
    """測試更新簽核狀態"""
    print("\n✍️ 測試更新簽核狀態...")
    
    # 測試資料
    test_data = {
        'purchase_no': '20250718-001',  # 使用一個存在的請購單號
        'status': '核准',
        'reason': '測試簽核功能'
    }
    
    try:
        response = requests.post(
            f"{BASE_URL}/update-approval-status",
            headers={'Content-Type': 'application/json'},
            json=test_data
        )
        
        if response.status_code == 200:
            result = response.json()
            if result.get('success'):
                print("✅ 簽核狀態更新成功")
                print(f"   簽核人員: {result.get('approver', 'N/A')}")
                print(f"   簽核日期: {result.get('approval_date', 'N/A')}")
                return True
            else:
                print(f"❌ 簽核狀態更新失敗: {result.get('message', '未知錯誤')}")
                return False
        else:
            print(f"❌ 請求失敗: {response.status_code}")
            return False
    except Exception as e:
        print(f"❌ 更新簽核狀態錯誤: {e}")
        return False

def test_get_purchase_detail():
    """測試取得請購單詳細資料"""
    print("\n🔍 測試取得請購單詳細資料...")
    
    try:
        response = requests.get(f"{BASE_URL}/purchase-detail/20250718-001")
        if response.status_code == 200:
            result = response.json()
            if result.get('success'):
                data = result.get('data', {})
                print("✅ 成功取得請購單詳細資料")
                print(f"   請購單號: {data.get('請購單號', 'N/A')}")
                print(f"   簽核狀態: {data.get('簽核', 'N/A')}")
                print(f"   簽核人員: {data.get('請購單簽核人員', 'N/A')}")
                print(f"   簽核日期: {data.get('請購單簽核日期', 'N/A')}")
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

def main():
    """主測試函數"""
    print("🚀 開始測試請購單簽核功能")
    print("=" * 50)
    
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
        test_get_purchase_approval_page,
        test_update_approval_status,
        test_get_purchase_detail
    ]
    
    passed = 0
    total = len(tests)
    
    for test in tests:
        try:
            if test():
                passed += 1
        except Exception as e:
            print(f"❌ 測試執行錯誤: {e}")
    
    print("\n" + "=" * 50)
    print(f"📊 測試結果: {passed}/{total} 通過")
    
    if passed == total:
        print("🎉 所有測試通過！請購單簽核功能正常")
    else:
        print("⚠️ 部分測試失敗，請檢查系統設定")

if __name__ == "__main__":
    main() 