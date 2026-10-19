// Dán toàn bộ file này vào Apps Script của Google Sheet (Tiện ích mở rộng → Apps Script).
// Sau khi sửa code: Triển khai → Quản lý các bản triển khai → ✎ sửa → Phiên bản: "Phiên bản mới" → Triển khai.
// (Làm vậy thì link /exec giữ nguyên.)
//
// Mỗi bạn nữ là một dòng trong trang tính "Câu trả lời":
//   - nhập tên xong: thêm dòng với các câu trả lời
//   - quay xong vòng quay: điền "Hôm nào nhận quà" vào đúng dòng đó (tìm theo "Mã lượt chơi")

const SHEET_NAME = "Câu trả lời";
const COLUMNS = [
  ["thoi_gian", "Thời gian"],
  ["ten", "Tên"],
  ["xinh_nhat", "Ai xinh nhất 66D"],
  ["dep_trai_nhat", "Ai đẹp trai nhất 66D"],
  ["ki_niem", "Kỉ niệm với 66D"],
  ["so_cau_dung", "Số câu đúng ngay"],
  ["ngay_nhan_qua", "Hôm nào nhận quà"],
  ["id", "Mã lượt chơi"]
];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (d.type === "hit") { recordHit(ss, d.boy); return ContentService.createTextOutput("ok"); }
    const sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);

    // luôn giữ dòng tiêu đề đúng thứ tự cột (tự thêm cột mới nếu Sheet cũ còn thiếu)
    const header = sh.getRange(1, 1, 1, COLUMNS.length);
    header.setValues([COLUMNS.map(c => c[1])]).setFontWeight("bold");
    sh.setFrozenRows(1);

    // cắt bớt chữ quá dài và chặn chữ bắt đầu bằng = + - @ để không bị hiểu là công thức
    const clean = v => String(v == null ? "" : v).slice(0, 1000).replace(/^[=+\-@]/, "'$&");
    const idCol = COLUMNS.findIndex(c => c[0] === "id") + 1;

    // đã có dòng của lượt chơi này thì chỉ điền thêm các ô được gửi lên
    let row = 0;
    if (d.id && sh.getLastRow() > 1) {
      const ids = sh.getRange(2, idCol, sh.getLastRow() - 1, 1).getValues();
      const i = ids.findIndex(r => String(r[0]) === String(d.id));
      if (i >= 0) row = i + 2;
    }
    if (row) {
      COLUMNS.forEach(([key], k) => {
        if (key !== "thoi_gian" && key !== "id" && d[key] != null && d[key] !== "") {
          sh.getRange(row, k + 1).setValue(clean(d[key]));
        }
      });
    } else {
      sh.appendRow(COLUMNS.map(([key]) => key === "thoi_gian" ? new Date() : clean(d[key])));
    }
    return ContentService.createTextOutput("ok");
  } finally {
    lock.releaseLock();
  }
}

// Trò đập chuột: trang tính "Đập chuột" đếm mỗi bạn nam đã bị đập bao nhiêu lần (tính chung mọi bạn nữ).
const HIT_SHEET = "Đập chuột";
function hitSheet(ss) {
  const sh = ss.getSheetByName(HIT_SHEET) || ss.insertSheet(HIT_SHEET);
  if (sh.getLastRow() === 0) {
    sh.appendRow(["Bạn nam", "Số lần bị đập"]);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, 2).setFontWeight("bold");
  }
  return sh;
}
function recordHit(ss, boy) {
  const name = String(boy || "").slice(0, 60).replace(/^[=+\-@]/, "'$&");
  if (!name) return;
  const sh = hitSheet(ss);
  const n = sh.getLastRow() - 1;
  const names = n > 0 ? sh.getRange(2, 1, n, 1).getValues() : [];
  const i = names.findIndex(r => String(r[0]) === name);
  if (i >= 0) {
    const cell = sh.getRange(i + 2, 2);
    cell.setValue((Number(cell.getValue()) || 0) + 1);
  } else {
    sh.appendRow([name, 1]);
  }
}

// GET ?action=counts → {"Tên bạn nam": số lần bị đập, ...} để trang biết ai đã đủ lượt.
// Mở link /exec bình thường trên trình duyệt sẽ thấy dòng chữ, để biết web app đã chạy.
function doGet(e) {
  if (e && e.parameter && e.parameter.action === "counts") {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HIT_SHEET);
    const out = {};
    if (sh && sh.getLastRow() > 1) {
      sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(([name, count]) => {
        if (name) out[String(name)] = Number(count) || 0;
      });
    }
    return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
  }
  return ContentService.createTextOutput("Web app nhận câu trả lời 20/10 của 66D đang chạy (bản có ngày nhận quà + đập chuột).");
}
