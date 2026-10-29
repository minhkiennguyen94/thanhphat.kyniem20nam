// Dán toàn bộ file này vào Google Sheet: Tiện ích mở rộng > Apps Script (thay hết code cũ).
const SHEET_NAME = "Đăng ký";
const HEADERS = [
  "Thời gian", "Họ và tên", "Chức vụ", "Số điện thoại", "Vị trí làm việc", "Địa chỉ",
  "Chi nhánh", "Size áo", "Số người thân",
  "NT1 - Họ tên", "NT1 - Tuổi", "NT1 - Quan hệ",
  "NT2 - Họ tên", "NT2 - Tuổi", "NT2 - Quan hệ",
  "NT3 - Họ tên", "NT3 - Tuổi", "NT3 - Quan hệ"
];

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight("bold");
    sh.setFrozenRows(1);
  }
  return sh;
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const d = JSON.parse(e.postData.contents);
    const n = Math.min(parseInt(d.so_nguoi, 10) || 0, 3);
    const row = [
      Utilities.formatDate(new Date(), "Asia/Ho_Chi_Minh", "dd/MM/yyyy HH:mm:ss"),
      d.ho_ten, d.chuc_vu, "'" + d.sdt, d.vi_tri, d.dia_chi,
      d.chi_nhanh, d.size_ao, n
    ];
    for (let i = 1; i <= 3; i++) {
      row.push(i <= n ? d["nt" + i + "_ten"] : "", i <= n ? d["nt" + i + "_tuoi"] : "", i <= n ? d["nt" + i + "_quan_he"] : "");
    }
    getSheet_().appendRow(row);
    return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return ContentService.createTextOutput("OK");
}
