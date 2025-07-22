"""
Python 與 Google Apps Script 整合模組
用於實現 Google Sheets 單一列唯讀保護功能
"""

import requests
import json
import os
from typing import Dict, Any, Optional

class GASProtectionManager:
    """Google Apps Script 保護管理器"""
    
    def __init__(self, web_app_url: str, admin_password: str):
        """
        初始化保護管理器
        
        Args:
            web_app_url: GAS Web App 的 URL
            admin_password: 管理員密碼
        """
        self.web_app_url = web_app_url
        self.admin_password = admin_password
        self.session = requests.Session()
    
    def set_row_protection(self, row_number: int, reason: str) -> Dict[str, Any]:
        """
        設定單一列為唯讀狀態
        
        Args:
            row_number: 要保護的列號
            reason: 保護原因
            
        Returns:
            操作結果字典
        """
        payload = {
            'action': 'setProtection',
            'rowNumber': row_number,
            'reason': reason,
            'password': self.admin_password
        }
        
        try:
            response = self.session.post(
                self.web_app_url,
                json=payload,
                headers={'Content-Type': 'application/json'}
            )
            response.raise_for_status()
            return response.json()
        except requests.exceptions.RequestException as e:
            return {
                'success': False,
                'message': f'網路請求失敗: {str(e)}'
            }
        except json.JSONDecodeError as e:
            return {
                'success': False,
                'message': f'回應解析失敗: {str(e)}'
            }
    
    def remove_row_protection(self, row_number: int) -> Dict[str, Any]:
        """
        解除單一列的唯讀保護
        
        Args:
            row_number: 要解除保護的列號
            
        Returns:
            操作結果字典
        """
        payload = {
            'action': 'removeProtection',
            'rowNumber': row_number,
            'password': self.admin_password
        }
        
        try:
            response = self.session.post(
                self.web_app_url,
                json=payload,
                headers={'Content-Type': 'application/json'}
            )
            response.raise_for_status()
            return response.json()
        except requests.exceptions.RequestException as e:
            return {
                'success': False,
                'message': f'網路請求失敗: {str(e)}'
            }
        except json.JSONDecodeError as e:
            return {
                'success': False,
                'message': f'回應解析失敗: {str(e)}'
            }
    
    def check_row_protection(self, row_number: int) -> Dict[str, Any]:
        """
        檢查列是否為唯讀狀態
        
        Args:
            row_number: 要檢查的列號
            
        Returns:
            檢查結果字典
        """
        payload = {
            'action': 'checkProtection',
            'rowNumber': row_number
        }
        
        try:
            response = self.session.post(
                self.web_app_url,
                json=payload,
                headers={'Content-Type': 'application/json'}
            )
            response.raise_for_status()
            return response.json()
        except requests.exceptions.RequestException as e:
            return {
                'success': False,
                'message': f'網路請求失敗: {str(e)}'
            }
        except json.JSONDecodeError as e:
            return {
                'success': False,
                'message': f'回應解析失敗: {str(e)}'
            }
    
    def protect_by_purchase_no(self, purchase_no: str, reason: str) -> Dict[str, Any]:
        """
        根據請購單號設定保護
        
        Args:
            purchase_no: 請購單號
            reason: 保護原因
            
        Returns:
            操作結果字典
        """
        payload = {
            'action': 'protectByPurchaseNo',
            'purchaseNo': purchase_no,
            'reason': reason,
            'password': self.admin_password
        }
        
        try:
            response = self.session.post(
                self.web_app_url,
                json=payload,
                headers={'Content-Type': 'application/json'}
            )
            response.raise_for_status()
            return response.json()
        except requests.exceptions.RequestException as e:
            return {
                'success': False,
                'message': f'網路請求失敗: {str(e)}'
            }
        except json.JSONDecodeError as e:
            return {
                'success': False,
                'message': f'回應解析失敗: {str(e)}'
            }

# 全域保護管理器實例
_gas_protection_manager = None

def get_gas_protection_manager() -> Optional[GASProtectionManager]:
    """
    取得 GAS 保護管理器實例
    
    Returns:
        GASProtectionManager 實例或 None
    """
    global _gas_protection_manager
    
    if _gas_protection_manager is None:
        web_app_url = os.getenv('GAS_WEB_APP_URL')
        admin_password = os.getenv('GAS_ADMIN_PASSWORD')
        
        if web_app_url and admin_password:
            _gas_protection_manager = GASProtectionManager(web_app_url, admin_password)
    
    return _gas_protection_manager

def set_purchase_protection(purchase_no: str, reason: str = "驗收簽核完成") -> Dict[str, Any]:
    """
    設定請購單保護（整合到現有系統）
    
    Args:
        purchase_no: 請購單號
        reason: 保護原因
        
    Returns:
        操作結果字典
    """
    manager = get_gas_protection_manager()
    
    if not manager:
        return {
            'success': False,
            'message': 'GAS 保護管理器未初始化，請檢查環境變數設定'
        }
    
    return manager.protect_by_purchase_no(purchase_no, reason)

def remove_purchase_protection(purchase_no: str) -> Dict[str, Any]:
    """
    解除請購單保護
    
    Args:
        purchase_no: 請購單號
        
    Returns:
        操作結果字典
    """
    # 需要先找到請購單對應的列號
    # 這裡需要整合現有的 Google Sheets 查詢邏輯
    from app import get_google_sheets_client, get_safe_records
    import os
    
    try:
        client = get_google_sheets_client()
        if not client:
            return {
                'success': False,
                'message': '無法連接到 Google Sheets'
            }
        
        spreadsheet_id = os.getenv('SPREADSHEET_ID')
        spreadsheet = client.open_by_key(spreadsheet_id)
        worksheet = spreadsheet.worksheet('請購單')
        
        # 找到請購單對應的列號
        all_records = get_safe_records(worksheet)
        target_row = None
        
        for i, record in enumerate(all_records, start=2):
            if record.get('請購單號') == purchase_no:
                target_row = i
                break
        
        if not target_row:
            return {
                'success': False,
                'message': f'找不到請購單號: {purchase_no}'
            }
        
        manager = get_gas_protection_manager()
        if not manager:
            return {
                'success': False,
                'message': 'GAS 保護管理器未初始化'
            }
        
        return manager.remove_row_protection(target_row)
        
    except Exception as e:
        return {
            'success': False,
            'message': f'解除保護失敗: {str(e)}'
        }

def check_purchase_protection(purchase_no: str) -> Dict[str, Any]:
    """
    檢查請購單保護狀態
    
    Args:
        purchase_no: 請購單號
        
    Returns:
        檢查結果字典
    """
    # 需要先找到請購單對應的列號
    from app import get_google_sheets_client, get_safe_records
    import os
    
    try:
        client = get_google_sheets_client()
        if not client:
            return {
                'success': False,
                'message': '無法連接到 Google Sheets'
            }
        
        spreadsheet_id = os.getenv('SPREADSHEET_ID')
        spreadsheet = client.open_by_key(spreadsheet_id)
        worksheet = spreadsheet.worksheet('請購單')
        
        # 找到請購單對應的列號
        all_records = get_safe_records(worksheet)
        target_row = None
        
        for i, record in enumerate(all_records, start=2):
            if record.get('請購單號') == purchase_no:
                target_row = i
                break
        
        if not target_row:
            return {
                'success': False,
                'message': f'找不到請購單號: {purchase_no}'
            }
        
        manager = get_gas_protection_manager()
        if not manager:
            return {
                'success': False,
                'message': 'GAS 保護管理器未初始化'
            }
        
        return manager.check_row_protection(target_row)
        
    except Exception as e:
        return {
            'success': False,
            'message': f'檢查保護狀態失敗: {str(e)}'
        }

# 測試函數
def test_gas_integration():
    """測試 GAS 整合功能"""
    print("開始測試 GAS 整合功能...")
    
    # 測試保護管理器初始化
    manager = get_gas_protection_manager()
    if manager:
        print("✓ GAS 保護管理器初始化成功")
    else:
        print("✗ GAS 保護管理器初始化失敗")
        return
    
    # 測試設定保護
    result = set_purchase_protection("20241219-001", "測試保護")
    print(f"設定保護結果: {result}")
    
    # 測試檢查保護狀態
    result = check_purchase_protection("20241219-001")
    print(f"檢查保護狀態結果: {result}")
    
    # 測試解除保護
    result = remove_purchase_protection("20241219-001")
    print(f"解除保護結果: {result}")
    
    print("GAS 整合功能測試完成")

if __name__ == "__main__":
    test_gas_integration() 