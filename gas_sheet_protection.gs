/**
 * Google Apps Script - 單一列唯讀保護系統
 * 功能：設定特定列為唯讀，需要密碼才能解除保護
 */

// ====== 參數設定 ======
const SPREADSHEET_ID = '1ZB6ri0fzqTRk_ciHibcGXEuViNmW1Ag9kkazE8A5iKc'; // 您的 Google Sheets ID
const ADMIN_PASSWORD = 'admin123'; // 管理員密碼
const PROTECTION_STATUS_COL = 'Z'; // 保護狀態欄位

// ====== 設定單一列為唯讀 ======
function setRowReadOnly(rowNumber, reason, password) {
  try {
    if (password !== ADMIN_PASSWORD) {
      return { success: false, message: '密碼錯誤，無法設定保護' };
    }
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName('請購單');
    if (!sheet) return { success: false, message: '找不到指定的工作表' };
    if (rowNumber < 2) return { success: false, message: '列號必須大於1（標題列）' };
    const range = sheet.getRange(rowNumber, 1, 1, sheet.getLastColumn());
    const protection = range.protect();
    protection.setDescription(`列 ${rowNumber} 保護 - ${reason}`);
    protection.setWarningOnly(false);
    const me = Session.getEffectiveUser();
    protection.removeEditors(protection.getEditors());
    protection.addEditor(me);
    // 標記保護狀態
    const statusCol = sheet.getRange(`${PROTECTION_STATUS_COL}1`);
    if (statusCol.getValue() !== '保護狀態') statusCol.setValue('保護狀態');
    const statusCell = sheet.getRange(`${PROTECTION_STATUS_COL}${rowNumber}`);
    statusCell.setValue(`已保護 - ${reason} - ${new Date().toLocaleString('zh-TW')}`);
    return {
      success: true,
      message: `列 ${rowNumber} 已成功設定為唯讀狀態`
    };
  } catch (error) {
    return { success: false, message: `設定保護失敗: ${error.toString()}` };
  }
}

// ====== 解除單一列唯讀保護 ======
function removeRowProtection(rowNumber, password) {
  try {
    if (password !== ADMIN_PASSWORD) {
      return { success: false, message: '密碼錯誤，無法解除保護' };
    }
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName('請購單');
    if (!sheet) return { success: false, message: '找不到指定的工作表' };
    const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
    let removed = false;
    for (let protection of protections) {
      const range = protection.getRange();
      if (range.getRow() === rowNumber) {
        protection.remove();
        removed = true;
        break;
      }
    }
    if (!removed) return { success: false, message: `列 ${rowNumber} 沒有找到保護設定` };
    const statusCell = sheet.getRange(`${PROTECTION_STATUS_COL}${rowNumber}`);
    statusCell.clearContent();
    return { success: true, message: `列 ${rowNumber} 的保護已成功解除` };
  } catch (error) {
    return { success: false, message: `解除保護失敗: ${error.toString()}` };
  }
}

// ====== 檢查單列保護狀態 ======
function isRowProtected(rowNumber) {
  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName('請購單');
    if (!sheet) return { success: false, message: '找不到指定的工作表' };
    const statusCell = sheet.getRange(`${PROTECTION_STATUS_COL}${rowNumber}`);
    const status = statusCell.getValue();
    const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
    let hasProtection = false;
    for (let protection of protections) {
      const range = protection.getRange();
      if (range.getRow() === rowNumber) {
        hasProtection = true;
        break;
      }
    }
    return {
      success: true,
      isProtected: hasProtection,
      status: status,
      message: hasProtection ? '該列已受保護' : '該列未受保護'
    };
  } catch (error) {
    return { success: false, message: `檢查失敗: ${error.toString()}` };
  }
}

// ====== 根據請購單號設定/解除保護 ======
function protectByPurchaseNo(purchaseNo, reason, password) {
  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName('請購單');
    if (!sheet) return { success: false, message: '找不到指定的工作表' };
    const data = sheet.getDataRange().getValues();
    let targetRow = null;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === purchaseNo) {
        targetRow = i + 1;
        break;
      }
    }
    if (!targetRow) return { success: false, message: `找不到請購單號: ${purchaseNo}` };
    return setRowReadOnly(targetRow, reason, password);
  } catch (error) {
    return { success: false, message: `設定保護失敗: ${error.toString()}` };
  }
}

function removeProtectionByPurchaseNo(purchaseNo, password) {
  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName('請購單');
    if (!sheet) return { success: false, message: '找不到指定的工作表' };
    const data = sheet.getDataRange().getValues();
    let targetRow = null;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === purchaseNo) {
        targetRow = i + 1;
        break;
      }
    }
    if (!targetRow) return { success: false, message: `找不到請購單號: ${purchaseNo}` };
    return removeRowProtection(targetRow, password);
  } catch (error) {
    return { success: false, message: `解除保護失敗: ${error.toString()}` };
  }
}

// ====== Web App 入口 ======
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;
    switch (action) {
      case 'setProtection':
        return ContentService.createTextOutput(JSON.stringify(
          setRowReadOnly(data.rowNumber, data.reason, data.password)
        )).setMimeType(ContentService.MimeType.JSON);
      case 'removeProtection':
        return ContentService.createTextOutput(JSON.stringify(
          removeRowProtection(data.rowNumber, data.password)
        )).setMimeType(ContentService.MimeType.JSON);
      case 'checkProtection':
        return ContentService.createTextOutput(JSON.stringify(
          isRowProtected(data.rowNumber)
        )).setMimeType(ContentService.MimeType.JSON);
      case 'protectByPurchaseNo':
        return ContentService.createTextOutput(JSON.stringify(
          protectByPurchaseNo(data.purchaseNo, data.reason, data.password)
        )).setMimeType(ContentService.MimeType.JSON);
      case 'removeProtectionByPurchaseNo':
        return ContentService.createTextOutput(JSON.stringify(
          removeProtectionByPurchaseNo(data.purchaseNo, data.password)
        )).setMimeType(ContentService.MimeType.JSON);
      default:
        return ContentService.createTextOutput(JSON.stringify({
          success: false,
          message: '未知的操作類型'
        })).setMimeType(ContentService.MimeType.JSON);
    }
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: `處理失敗: ${error.toString()}`
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// ====== 測試函數（可在 GAS 編輯器直接執行）======
function testProtection() {
  console.log('開始測試保護功能...');
  const result1 = setRowReadOnly(5, '測試保護', ADMIN_PASSWORD);
  console.log('設定保護結果:', result1);
  const result2 = isRowProtected(5);
  console.log('檢查保護狀態結果:', result2);
  const result3 = removeRowProtection(5, ADMIN_PASSWORD);
  console.log('解除保護結果:', result3);
  const result4 = protectByPurchaseNo('20250719-001', '依單號測試', ADMIN_PASSWORD);
  console.log('依單號設定保護結果:', result4);
  const result5 = removeProtectionByPurchaseNo('20250719-001', ADMIN_PASSWORD);
  console.log('依單號解除保護結果:', result5);
  console.log('測試完成');
} 