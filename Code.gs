const TAB = 'Rà soát đối thủ';
const DT_PREFIX = 'DT-';

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(10000);
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB);
    if (!sheet) return json({ ok: false, error: 'Không tìm thấy tab' });

    let data;
    try { data = JSON.parse(e.postData.contents); }
    catch (err) { return json({ ok: false, error: 'Dữ liệu không hợp lệ' }); }

    const competitors = (data && data.competitors) || [];
    if (!competitors.length) return json({ ok: false, error: 'Không có đối thủ' });

    const ts = data.timestamp ? new Date(data.timestamp) : new Date();

    // Đọc Mã ĐT hiện tại để sinh mã mới
    const lastRow = sheet.getLastRow();
    let maxDT = 0;
    const index = {};
    if (lastRow > 1) {
      const vals = sheet.getRange(2, 1, lastRow - 1, 22).getValues();
      for (let i = 0; i < vals.length; i++) {
        const ma = String(vals[i][5] || '').trim();
        const ten = String(vals[i][10] || '').trim();
        if (ma && ten) index[ma + '||' + ten] = i + 2;
        // Tìm Mã ĐT lớn nhất
        const dt = String(vals[i][21] || '').trim();
        const m = dt.match(/^DT-(\d+)$/);
        if (m) { const n = parseInt(m[1], 10); if (n > maxDT) maxDT = n; }
      }
    }

    let nextDT = maxDT + 1;
    let added = 0, updated = 0;

    competitors.forEach(function (c) {
      const row = [
        isoWeek(ts),                 // A: Tuần
        fmtDT(ts),                   // B: Ngày sửa
        data.mien || '',             // C: Miền
        data.gdv || '',              // D: GĐV
        data.qlkv || '',             // E: QLKV
        data.ma_ch || '',            // F: Mã CH
        data.ten_ch || '',           // G: Tên CH
        data.tinh || '',             // H: Tỉnh/TP
        data.huyen || '',            // I: Quận/Huyện
        c.nhom || '',                // J: Nhóm đối thủ
        c.ten || '',                 // K: Tên đối thủ
        c.diaChi || '',              // L: Địa chỉ
        (c.lat != null && c.lat !== '') ? c.lat : '',   // M: Vĩ độ
        (c.lng != null && c.lng !== '') ? c.lng : '',   // N: Kinh độ
        (c.kc_km != null && c.kc_km !== '') ? c.kc_km : '', // O: KC
        c.ngay || '',                // P: Ngày KT
        c.doanhSo || '',             // Q: Doanh số
        c.luongKhach || '',          // R: Khách
        c.mt || '',                  // S: MT
        c.giaThue || '',             // T: Giá thuê
        c.dienTich || '',            // U: Diện tích
      ];

      const key = (data.ma_ch || '') + '||' + (c.ten || '');
      const rn = index[key];
      
      if (rn) {
        // UPDATE: giữ nguyên Mã ĐT cũ, chỉ cập nhật 21 cột đầu
        sheet.getRange(rn, 1, 1, 21).setValues([row]);
        updated++;
      } else {
        // NEW: sinh Mã ĐT mới
        const maDT = DT_PREFIX + String(nextDT).padStart(4, '0');
        row.push(maDT);
        sheet.appendRow(row);
        nextDT++;
        added++;
      }
    });

    return json({ ok: true, added: added, updated: updated });
  } finally { lock.releaseLock(); }
}

function isoWeek(d) {
  const date = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
  const dayNum = date.getUTCDay() || 7;
  date.setUTCDate(date.getUTCDate() + 4 - dayNum);
  const yearStart = new Date(Date.UTC(date.getUTCFullYear(), 0, 1));
  return Math.ceil((((date - yearStart) / 86400000) + 1) / 7);
}

function fmtDT(d) {
  const p = function (n) { return String(n).padStart(2, '0'); };
  return p(d.getDate()) + '/' + p(d.getMonth() + 1) + '/' + d.getFullYear() + ' ' + p(d.getHours()) + ':' + p(d.getMinutes());
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}