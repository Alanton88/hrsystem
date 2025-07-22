#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
簡化測試 Google Sheets 保護功能
專注於驗證應用程式層面的保護功能
"""

import requests
import json
import os
from datetime import datetime

# 測試配置
BASE_URL = "http://127.0.0.1:5000"
TEST_USERNAME = "admin"
TEST_PASSWORD = "admin123"

def test_basic_protection():
    """測試基本的保護功能"""
    print("🚀 開始測試基本 Google Sheets 保護功能")
    print("=" * 60)
    
    # 創建會話
    session = requests.Session()
    
    # 1. 登入
    print("🔐 步驟 1: 登入系統...")
    login_data = {
        'username': TEST_USERNAME,
        'password': TEST_PASSWORD
    }
    
    try:
        response = session.post(f"{BASE_URL}/login", data=login_data)
        if response.status_code == 200:
            print("✅ 登入成功")
        else:
            print(f"❌ 登入失敗: {response.status_code}")
            return
    except Exception as e:
        print(f"❌ 登入錯誤: {e}")
        return
    
    # 2. 更新驗收簽核狀態（設定為唯讀）
    print("\n✍️ 步驟 2: 更新驗收簽核狀態...")
    test_data = {
        'purchase_no': '20250718-004',
        'approval_status': '核准',
        'approval_date': datetime.now().strftime('%Y%m%d'),
        'approver': '測試人員',
        'approval_note': '測試基本保護功能'
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
                
                if result.get('is_readonly'):
                    print("✅ 記錄已成功設為唯讀狀態")
                else:
                    print("⚠️ 記錄未設為唯讀狀態")
            else:
                print(f"❌ 驗收簽核狀態更新失敗: {result.get('message', '未知錯誤')}")
                return
        else:
            print(f"❌ 請求失敗: {response.status_code}")
            return
    except Exception as e:
        print(f"❌ 更新驗收簽核狀態錯誤: {e}")
        return
    
    # 3. 驗證管理員密碼驗證
    print("\n🔓 步驟 3: 測試管理員密碼驗證...")
    verify_data = {
        'password': TEST_PASSWORD,
        'purchase_no': '20250718-004'
    }
    
    try:
        response = session.post(
            f"{BASE_URL}/verify-admin-password",
            headers={'Content-Type': 'application/json'},
            json=verify_data
        )
        
        if response.status_code == 200:
            result = response.json()
            if result.get('success'):
                print("✅ 管理員密碼驗證成功")
                print(f"   訊息: {result.get('message', 'N/A')}")
                print(f"   可編輯: {result.get('can_edit', 'N/A')}")
                
                if result.get('can_edit'):
                    print("✅ 記錄已成功解除唯讀狀態")
                else:
                    print("⚠️ 記錄未解除唯讀狀態")
            else:
                print(f"❌ 管理員密碼驗證失敗: {result.get('message', '未知錯誤')}")
        else:
            print(f"❌ 請求失敗: {response.status_code}")
    except Exception as e:
        print(f"❌ 管理員密碼驗證錯誤: {e}")
    
    # 4. 測試批量解除保護
    print("\n🔓 步驟 4: 測試批量解除保護...")
    batch_verify_data = {
        'password': TEST_PASSWORD,
        'purchase_no': 'ALL'
    }
    
    try:
        response = session.post(
            f"{BASE_URL}/verify-admin-password",
            headers={'Content-Type': 'application/json'},
            json=batch_verify_data
        )
        
        if response.status_code == 200:
            result = response.json()
            if result.get('success'):
                print("✅ 批量解除保護成功")
                print(f"   訊息: {result.get('message', 'N/A')}")
                print(f"   更新記錄數: {result.get('updated_count', 'N/A')}")
            else:
                print(f"❌ 批量解除保護失敗: {result.get('message', '未知錯誤')}")
        else:
            print(f"❌ 請求失敗: {response.status_code}")
    except Exception as e:
        print(f"❌ 批量解除保護錯誤: {e}")
    
    print("\n" + "=" * 60)
    print("🎉 基本保護功能測試完成！")
    print("\n💡 功能說明:")
    print("1. ✅ 驗收簽核完成後，記錄會自動設為唯讀")
    print("2. ✅ 管理員可以透過密碼驗證解除保護")
    print("3. ✅ 支援批量解除所有記錄的保護")
    print("4. ⚠️ Google Sheets API 保護需要額外的權限設定")
    print("\n📝 注意事項:")
    print("- 目前使用應用程式層面的保護控制")
    print("- 如需真正的 Google Sheets 保護，請設定適當的 API 權限")
    print("- 服務帳戶需要對 Google Sheets 有編輯權限")

if __name__ == "__main__":
    test_basic_protection() 