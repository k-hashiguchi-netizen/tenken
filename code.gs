/**
 * 両替依頼アプリ Backend Logic
 * ファイル名: コード.gs
 */

function doGet(e) {
  // URLパラメータから店舗名を取得 (?store=中山店 など)
  const storeParam = (e && e.parameter && e.parameter.store) ? e.parameter.store : "";
  
  const template = HtmlService.createTemplateFromFile('index');
  template.store = storeParam; // HTML側の <?= store ?> に値を渡す
  
  return template.evaluate()
    .setTitle('両替依頼システム')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * 店舗リストの取得
 */
function getStoreList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('店舗マスタ');
  if (!sheet) return [];
  
  const values = sheet.getRange(1, 1, sheet.getLastRow(), 1).getValues();
  return values.flat().filter(v => v !== "" && v !== "店舗名" && v !== "店名");
}

/**
 * 特定店舗の担当者マスタリストを取得
 */
function getStaffList(storeName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('担当者マスタ');
    if (!sheet) {
      sheet = ss.insertSheet('担当者マスタ');
      sheet.appendRow(['店舗名', '担当者名']);
      return [];
    }
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    
    const values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    // 店舗名が一致する担当者のみをフィルタリングして返却
    return values
      .filter(r => String(r[0]) === String(storeName) && r[1] !== "")
      .map(r => r[1]);
  } catch(e) {
    return [];
  }
}

/**
 * 店舗別担当者マスタに新規登録
 */
function addStaff(storeName, name) {
  if (!storeName || !name) return { success: false, error: '店舗名または担当者名が指定されていません。' };
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('担当者マスタ');
    if (!sheet) {
      sheet = ss.insertSheet('担当者マスタ');
      sheet.appendRow(['店舗名', '担当者名']);
    }
    const values = sheet.getDataRange().getValues();
    const exists = values.some(r => String(r[0]) === String(storeName) && String(r[1]) === String(name));
    if (exists) {
      return { success: false, error: 'この担当者は既に登録されています。' };
    }
    sheet.appendRow([storeName, name]);
    return { success: true };
  } catch(e) {
    return { success: false, error: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

/**
 * 店舗別担当者マスタから削除
 */
function deleteStaff(storeName, name) {
  if (!storeName || !name) return { success: false, error: '店舗名または担当者名が指定されていません。' };
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('担当者マスタ');
    if (!sheet) return { success: false, error: 'マスタシートが見つかりません。' };
    
    const values = sheet.getRange(1, 1, sheet.getLastRow(), 2).getValues();
    let deletedCount = 0;
    // 重複を避けるため下からスキャンして削除
    for (let i = values.length - 1; i >= 0; i--) {
      if (String(values[i][0]) === String(storeName) && String(values[i][1]) === String(name)) {
        sheet.deleteRow(i + 1);
        deletedCount++;
      }
    }
    return { success: deletedCount > 0, error: deletedCount === 0 ? '指定された担当者が見つかりません。' : null };
  } catch(e) {
    return { success: false, error: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

/**
 * 依頼データの保存（同時送信ロック対応）
 */
function submitExchangeRequest(formData) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('依頼データ');
    
    if (!sheet) {
      sheet = ss.insertSheet('依頼データ');
      sheet.appendRow(['日時', '店舗名', '担当者', '5000円', '1000円', '500円', '100円', '50円', '10円', '5円', '1円', '合計']);
    }
    
    const now = new Date();
    const todayStr = Utilities.formatDate(now, "JST", "yyyy/MM/dd");
    
    const rowData = [
      now,
      formData.storeName,
      formData.staffName,
      Number(formData.bill5000) || 0,
      Number(formData.bill1000) || 0,
      Number(formData.coin500) || 0,
      Number(formData.coin100) || 0,
      Number(formData.coin50) || 0,
      Number(formData.coin10) || 0,
      Number(formData.coin5) || 0,
      Number(formData.coin1) || 0,
      Number(formData.totalAmount) || 0
    ];
    
    const lastRow = sheet.getLastRow();
    let targetRow = -1;
    
    if (lastRow > 1) {
      const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues(); 
      for (let i = 0; i < data.length; i++) {
        const d = data[i][0];
        if (!(d instanceof Date)) continue;
        const rowDate = Utilities.formatDate(d, "JST", "yyyy/MM/dd");
        if (rowDate === todayStr && String(data[i][1]) === String(formData.storeName)) {
          targetRow = i + 2;
          break;
        }
      }
    }
    
    if (targetRow !== -1) {
      sheet.getRange(targetRow, 1, 1, rowData.length).setValues([rowData]);
    } else {
      sheet.appendRow(rowData);
    }
    
    return { success: true, updated: targetRow !== -1 };
  } catch (e) {
    return { success: false, error: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

/**
 * 事務所用：全データの取得
 */
function getAdminData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('依頼データ');
    if (!sheet || sheet.getLastRow() < 2) return { headers: [], rows: [] };
    
    const values = sheet.getDataRange().getValues();
    const headers = values[0];
    const rows = values.slice(1).reverse().map(row => {
      const formattedRow = [...row];
      if (formattedRow[0] instanceof Date) {
        formattedRow[0] = formattedRow[0].toISOString();
      }
      return formattedRow;
    });
    
    return { headers: headers, rows: rows };
  } catch (e) {
    return { error: e.toString() };
  }
}
