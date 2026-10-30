/**
 * GOOGLE APPS SCRIPT - ĐỒNG BỘ TRỰC TIẾP VỚI GOOGLE FORM & GOOGLE SHEET
 * Sheet ID: 1aYdZGk3-l-RuVWwrwQjJLoY1Jw38TbTbzuLYq6g3gOY
 */

const SHEET_ID = '1aYdZGk3-l-RuVWwrwQjJLoY1Jw38TbTbzuLYq6g3gOY';
const DEFAULT_COMPANY_COL = 2; // Mặc định Cột B (ngay sau cột Dấu thời gian)

function detectCompanyColIndex(headers, fallbackCol) {
  if (!headers || !headers.length) return fallbackCol - 1;
  const keywords = ['tên doanh nghiệp', 'doanh nghiệp', 'tên công ty', 'công ty', 'tên đơn vị', 'đơn vị', 'company'];
  for (let kw of keywords) {
    for (let i = 1; i < headers.length; i++) {
      const h = (headers[i] || '').toString().toLowerCase().trim();
      if (h.indexOf(kw) !== -1) {
        return i;
      }
    }
  }
  return Math.max(0, fallbackCol - 1);
}

function doGet(e) {
  try {
    let ss;
    try {
      ss = SpreadsheetApp.getActiveSpreadsheet();
    } catch (err) {
      ss = null;
    }
    if (!ss) {
      ss = SpreadsheetApp.openById(SHEET_ID);
    }

    const sheet = ss.getSheets()[0];
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    const items = [];
    let headers = [];

    if (lastRow >= 1 && lastCol >= 1) {
      const data = sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
      headers = data[0] || [];

      let colIdx;
      if (e && e.parameter && e.parameter.col && e.parameter.col !== 'auto') {
        colIdx = Math.max(0, (parseInt(e.parameter.col, 10) || DEFAULT_COMPANY_COL) - 1);
      } else {
        colIdx = detectCompanyColIndex(headers, DEFAULT_COMPANY_COL);
      }

      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const timestamp = (row[0] || '').toString().trim();
        let companyName = (row[colIdx] || '').toString().trim();

        // Nếu cột mặc định trống nhưng các cột kế tiếp có dữ liệu, lấy cột có dữ liệu đầu tiên
        if (!companyName) {
          for (let c = 1; c < row.length; c++) {
            if (row[c] && row[c].toString().trim() !== '') {
              companyName = row[c].toString().trim();
              break;
            }
          }
        }

        if (!companyName) continue;

        items.push({
          rowNumber: i + 1,
          id: 'row_' + (i + 1) + '_' + encodeURIComponent(timestamp) + '_' + encodeURIComponent(companyName),
          timestamp: timestamp,
          companyName: companyName
        });
      }
    }

    const jsonString = JSON.stringify({
      ok: true,
      total: items.length,
      items: items
    });

    if (e && e.parameter && e.parameter.callback) {
      const cb = e.parameter.callback.replace(/[^a-zA-Z0-9_$.]/g, '');
      return ContentService.createTextOutput(cb + '(' + jsonString + ');')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }

    return ContentService.createTextOutput(jsonString)
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    const errStr = JSON.stringify({ ok: false, error: error.toString() });
    if (e && e.parameter && e.parameter.callback) {
      const cb = e.parameter.callback.replace(/[^a-zA-Z0-9_$.]/g, '');
      return ContentService.createTextOutput(cb + '(' + errStr + ');')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService.createTextOutput(errStr)
      .setMimeType(ContentService.MimeType.JSON);
  }
}
