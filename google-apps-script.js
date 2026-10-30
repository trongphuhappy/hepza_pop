/**
 * GOOGLE APPS SCRIPT - KẾT NỐI DỮ LIỆU GOOGLE FORM / SHEET VỚI MÀN HÌNH SỰ KIỆN
 * Sheet ID mặc định: 1aYdZGk3-l-RuVWwrwQjJLoY1Jw38TbTbzuLYq6g3gOY
 *
 * HƯỚNG DẪN CÀI ĐẶT:
 * 1. Mở file Google Sheet của bạn:
 *    https://docs.google.com/spreadsheets/d/1aYdZGk3-l-RuVWwrwQjJLoY1Jw38TbTbzuLYq6g3gOY/edit
 * 2. Trên thanh menu, chọn: Tiện ích mở rộng (Extensions) -> Apps Script
 * 3. Xóa toàn bộ mã cũ (nếu có) và dán toàn bộ đoạn mã này vào.
 * 4. Nhấn biểu tượng Đĩa mềm (Lưu / Ctrl+S).
 * 5. Nhấn nút "Triển khai" (Deploy) ở góc trên bên phải -> Chọn "Tùy chọn triển khai mới" (New deployment).
 * 6. Ở mục "Chọn loại" (biểu tượng bánh răng), chọn "Ứng dụng web" (Web app).
 * 7. Cấu hình:
 *    - Thực thi dưới dạng (Execute as): "Tôi" (Me)
 *    - Ai có quyền truy cập (Who has access): "Bất kỳ ai" (Anyone)  <-- QUAN TRỌNG!
 * 8. Nhấn "Triển khai" (Deploy), cấp quyền (Authorize access) nếu Google hỏi.
 * 9. Copy đường link "URL ứng dụng web" (có đuôi /exec) và dán vào phần Cài đặt trên trang web!
 */

const SHEET_ID = '1aYdZGk3-l-RuVWwrwQjJLoY1Jw38TbTbzuLYq6g3gOY';
// Cột chứa Tên doanh nghiệp (1 = Cột A, 2 = Cột B, 3 = Cột C...)
const COMPANY_COLUMN_INDEX = 2; // Mặc định Cột B (ngay sau cột Dấu thời gian)

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

    // Lấy trang tính đầu tiên (thường là Câu trả lời biểu mẫu 1 / Form Responses 1)
    const sheet = ss.getSheets()[0];
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();

    const items = [];
    let headers = [];

    if (lastRow >= 1 && lastCol >= 1) {
      const data = sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
      headers = data[0] || [];

      // Cho phép truyền tham số ?col=2 từ web nếu muốn đổi cột linh hoạt
      const colParam = e && e.parameter && e.parameter.col ? parseInt(e.parameter.col, 10) : COMPANY_COLUMN_INDEX;
      const colIdx = Math.max(0, (isNaN(colParam) ? COMPANY_COLUMN_INDEX : colParam) - 1);

      // Duyệt từ dòng 2 (bỏ qua dòng tiêu đề đầu tiên)
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const timestamp = (row[0] || '').toString().trim();
        const companyName = (row[colIdx] || '').toString().trim();

        // Bỏ qua dòng trống hoàn toàn
        if (!companyName && !timestamp) continue;

        // Gom các cột phụ (nếu có thêm thông tin người đại diện, chức vụ...)
        const extraFields = [];
        for (let c = 1; c < row.length; c++) {
          if (c !== colIdx && row[c] && row[c].toString().trim() !== '') {
            extraFields.push({
              label: headers[c] || `Cột ${c + 1}`,
              value: row[c].toString().trim()
            });
          }
        }

        items.push({
          rowNumber: i + 1,
          id: `row_${i + 1}_${encodeURIComponent(timestamp)}_${encodeURIComponent(companyName)}`,
          timestamp: timestamp,
          companyName: companyName,
          extra: extraFields,
          rawRow: row
        });
      }
    }

    const payload = {
      ok: true,
      sheetName: sheet.getName(),
      updatedAt: new Date().toISOString(),
      headers: headers,
      total: items.length,
      items: items
    };

    const jsonString = JSON.stringify(payload);

    // Hỗ trợ JSONP nếu gọi qua ?callback=... (giúp chạy trực tiếp từ file:// không bị lỗi CORS)
    if (e && e.parameter && e.parameter.callback) {
      const callbackName = e.parameter.callback.replace(/[^a-zA-Z0-9_$.]/g, '');
      return ContentService
        .createTextOutput(`${callbackName}(${jsonString});`)
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }

    return ContentService
      .createTextOutput(jsonString)
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    const errPayload = JSON.stringify({
      ok: false,
      error: error.toString()
    });

    if (e && e.parameter && e.parameter.callback) {
      const callbackName = e.parameter.callback.replace(/[^a-zA-Z0-9_$.]/g, '');
      return ContentService
        .createTextOutput(`${callbackName}(${errPayload});`)
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }

    return ContentService
      .createTextOutput(errPayload)
      .setMimeType(ContentService.MimeType.JSON);
  }
}

