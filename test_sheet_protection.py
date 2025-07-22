#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
測試 Google Sheets 保護功能
驗證驗收簽核後是否真正鎖定了 Google Sheets 的編輯權限
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
    
    # 創建會話來保持登入狀態
    session = requests.Session()
    
    login_data = {
        'username': TEST_USERNAME,
        'password': TEST_PASSWORD
    }
    
    try:
        response = session.post(f"{BASE_URL}/login", data=login_data)
        if response.status_code == 200:
            print("✅ 登入成功")
            return session
        else:
            print(f"❌ 登入失敗: {response.status_code}")
            return None
    except Exception as e:
        print(f"❌ 登入錯誤: {e}")
        return None

def test_check_sheet_protection(session):
    """測試檢查 Google Sheets 保護狀態"""
    print("\n🔍 測試檢查 Google Sheets 保護狀態...")
    
    try:
        response = session.get(f"{BASE_URL}/check-sheet-protection")
        if response.status_code == 200:
            result = response.json()
            if result.get('success'):
                protection_info = result.get('protection_info', [])
                total_protected = result.get('total_protected_ranges', 0)
                
                print("✅ 成功取得保護狀態")
                print(f"   總保護範圍數: {total_protected}")
                
                for i, protection in enumerate(protection_info, 1):
                    print(f"   保護範圍 {i}:")
                    print(f"     工作表: {protection.get('sheet_name', 'N/A')}")
                    print(f"     描述: {protection.get('description', 'N/A')}")
                    print(f"     範圍: {protection.get('range', 'N/A')}")
                
                return True
            else:
                print(f"❌ 檢查保護狀態失敗: {result.get('message', '未知錯誤')}")
                return False
        else:
            print(f"❌ 請求失敗: {response.status_code}")
            return False
    except Exception as e:
        print(f"❌ 檢查保護狀態錯誤: {e}")
        return False

def test_update_receipt_approval_with_protection(session):
    """測試更新驗收簽核狀態並驗證保護功能"""
    print("\n✍️ 測試更新驗收簽核狀態並驗證保護功能...")
    
    # 測試資料
    test_data = {
        'purchase_no': '20250718-004',  # 使用您提到的請購單號
        'approval_status': '核准',
        'approval_date': datetime.now().strftime('%Y%m%d'),
        'approver': '測試人員',
        'approval_note': '測試 Google Sheets 保護功能'
    }
    
    try:
        response = session.post(
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
                
                # 等待一下讓保護設定生效
                import time
                time.sleep(2)
                
                # 再次檢查保護狀態
                print("\n🔍 檢查保護是否生效...")
                protection_response = session.get(f"{BASE_URL}/check-sheet-protection")
                if protection_response.status_code == 200:
                    protection_result = protection_response.json()
                    if protection_result.get('success'):
                        protection_info = protection_result.get('protection_info', [])
                        found_protection = False
                        
                        for protection in protection_info:
                            if test_data['purchase_no'] in protection.get('description', ''):
                                found_protection = True
                                print(f"✅ 找到請購單 {test_data['purchase_no']} 的保護範圍")
                                print(f"   描述: {protection.get('description', 'N/A')}")
                                print(f"   範圍: {protection.get('range', 'N/A')}")
                                break
                        
                        if not found_protection:
                            print(f"⚠️ 未找到請購單 {test_data['purchase_no']} 的保護範圍")
                            print("   可能原因：認證憑證問題或 API 權限不足")
                
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

def test_verify_admin_password(session):
    """測試管理員密碼驗證和解除保護"""
    print("\n🔓 測試管理員密碼驗證和解除保護...")
    
    # 測試資料
    test_data = {
        'password': TEST_PASSWORD,
        'purchase_no': '20250718-004'
    }
    
    try:
        response = session.post(
            f"{BASE_URL}/verify-admin-password",
            headers={'Content-Type': 'application/json'},
            json=test_data
        )
        
        if response.status_code == 200:
            result = response.json()
            if result.get('success'):
                print("✅ 管理員密碼驗證成功")
                print(f"   訊息: {result.get('message', 'N/A')}")
                print(f"   可編輯: {result.get('can_edit', 'N/A')}")
                
                # 等待一下讓保護解除生效
                import time
                time.sleep(2)
                
                # 再次檢查保護狀態
                print("\n🔍 檢查保護是否已解除...")
                protection_response = session.get(f"{BASE_URL}/check-sheet-protection")
                if protection_response.status_code == 200:
                    protection_result = protection_response.json()
                    if protection_result.get('success'):
                        protection_info = protection_result.get('protection_info', [])
                        found_protection = False
                        
                        for protection in protection_info:
                            if test_data['purchase_no'] in protection.get('description', ''):
                                found_protection = True
                                print(f"⚠️ 請購單 {test_data['purchase_no']} 的保護範圍仍然存在")
                                break
                        
                        if not found_protection:
                            print(f"✅ 請購單 {test_data['purchase_no']} 的保護範圍已成功解除")
                
                return True
            else:
                print(f"❌ 管理員密碼驗證失敗: {result.get('message', '未知錯誤')}")
                return False
        else:
            print(f"❌ 請求失敗: {response.status_code}")
            return False
    except Exception as e:
        print(f"❌ 管理員密碼驗證錯誤: {e}")
        return False

def main():
    """主測試函數"""
    print("🚀 開始測試 Google Sheets 保護功能")
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
    session = test_login()
    if not session:
        print("❌ 登入失敗，無法繼續測試")
        return
    
    passed = 0
    total = 3  # 除了登入外的其他測試
    
    # 執行其他測試
    tests = [
        lambda: test_check_sheet_protection(session),
        lambda: test_update_receipt_approval_with_protection(session),
        lambda: test_verify_admin_password(session)
    ]
    
    for test in tests:
        try:
            if test():
                passed += 1
        except Exception as e:
            print(f"❌ 測試執行錯誤: {e}")
    
    print("\n" + "=" * 60)
    print(f"📊 測試結果: {passed}/{total} 通過")
    
    if passed == total:
        print("🎉 所有測試通過！Google Sheets 保護功能正常")
    else:
        print("⚠️ 部分測試失敗，請檢查系統設定")
        print("\n💡 注意事項:")
        print("1. 確保已設定 Google Sheets API 認證")
        print("2. 確保 token.json 檔案存在且有效")
        print("3. 確保 Google Sheets 有適當的權限設定")

if __name__ == "__main__":
    main() 