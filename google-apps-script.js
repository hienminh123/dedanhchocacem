// Dán toàn bộ file này vào Apps Script của Google Sheet (Tiện ích mở rộng → Apps Script),
// rồi Deploy → New deployment → Web app, "Who has access" = Anyone.
// Mỗi bạn nữ chơi xong phần nhập tên sẽ thêm một dòng vào trang tính "Câu trả lời".

const SHEET_NAME = "Câu trả lời";
const HEADERS = ["Thời gian", "Tên", "Ai xinh nhất 66D", "Ai đẹp trai nhất 66D", "Kỉ niệm với 66D", "Số câu đúng ngay"];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
    if (sh.getLastRow() === 0) {
      sh.appendRow(HEADERS);
      sh.setFrozenRows(1);
      sh.getRange(1, 1, 1, HEADERS.length).setFontWeight("bold");
    }
    // cắt bớt chữ quá dài và chặn chữ bắt đầu bằng = + - @ để không bị hiểu là công thức
    const clean = v => String(v == null ? "" : v).slice(0, 1000).replace(/^[=+\-@]/, "'$&");
    sh.appendRow([
      new Date(),
      clean(d.ten),
      clean(d.xinh_nhat),
      clean(d.dep_trai_nhat),
      clean(d.ki_niem),
      clean(d.so_cau_dung)
    ]);
    return ContentService.createTextOutput("ok");
  } finally {
    lock.releaseLock();
  }
}

// Mở link /exec trên trình duyệt sẽ thấy dòng này, để biết web app đã chạy.
function doGet() {
  return ContentService.createTextOutput("Web app nhận câu trả lời 20/10 của 66D đang chạy.");
}
