var SHEET_ID = '1UMtB9MOxGOaZvxN5VAUGAS1bPaAL2tVVFEViGANd4Tk';
var MAIN_SHEET = 'Rà soát đối thủ';
var LOG_SHEET = 'Tracklogs';
var DT_PREFIX = 'DT-';

var MAIN_HEADER = [
  'Mã ĐT', 'Mã CH', 'Tên CH', 'Miền', 'GĐV', 'QLKV',
  'Tỉnh/TP', 'Quận/Huyện', 'Nhóm đối thủ', 'Tên đối thủ', 'Địa chỉ',
  'Vĩ độ', 'Kinh độ', 'KC (km)', 'Ngày khai trương',
  'Doanh số TB/ngày (trđ)', 'Khách TB/ngày', 'MT (m)',
  'Giá thuê (tr/th)', 'Diện tích (m²)', 'Ngày khai báo', 'Ngày update'
];

var LOG_HEADER = [
  'Tuần', 'Thời gian', 'Miền', 'GĐV', 'QLKV', 'Mã CH', 'Tên CH',
  'Hành động', 'Số đối thủ'
];

function doPost(e) {
  var data;
  try {
    data = JSON.parse(e && e.postData && e.postData.contents || '{}');
  } catch (err) {
    return out({ ok: false, error: 'Dữ liệu không hợp lệ' });
  }
  if (!data || typeof data !== 'object' || Array.isArray(data)) data = {};

  var ma_ch = String(data.ma_ch || '').trim();
  if (!ma_ch) return out({ ok: false, error: 'Thiếu ma_ch' });

  var action = data.action || 'save';
  var competitors = Array.isArray(data.competitors) ? data.competitors : [];
  var lock = LockService.getScriptLock();
  var lockAcquired = false;

  try {
    // Do not continue without the lock: otherwise parallel saves can allocate
    // the same DT number or overwrite one another.
    lockAcquired = lock.tryLock(10000);
    if (!lockAcquired) {
      return out({ ok: false, error: 'Hệ thống đang xử lý nhiều yêu cầu. Vui lòng thử lại.' });
    }

    var ss = SpreadsheetApp.openById(SHEET_ID);
    var main = ss.getSheetByName(MAIN_SHEET);
    if (!main) return out({ ok: false, error: 'Không tìm thấy sheet hiện trạng: ' + MAIN_SHEET });

    var log = ss.getSheetByName(LOG_SHEET);
    if (!log) log = ss.insertSheet(LOG_SHEET);
    if (log.getLastRow() === 0) {
      log.getRange(1, 1, 1, LOG_HEADER.length).setValues([LOG_HEADER]);
    }
    if (main.getLastRow() === 0) {
      main.getRange(1, 1, 1, MAIN_HEADER.length).setValues([MAIN_HEADER]);
    }

    var now = new Date();
    var tuan = getISOWeek(now);
    var ts = Utilities.formatDate(now, 'Asia/Ho_Chi_Minh', 'dd/MM/yyyy HH:mm');
    var mien = data.mien || '', gdv = data.gdv || '', qlkv = data.qlkv || '';
    var ten_ch = data.ten_ch || '';
    var actionLabel = action === 'nochange' ? 'Giữ nguyên' : action === 'none' ? 'Không có' : 'Thay đổi';
    var soDoiThu = (data.so_doi_thu != null) ? data.so_doi_thu : competitors.length;

    // One log row write. This remains inside the lock to serialize concurrent requests.
    var logRow = log.getLastRow() + 1;
    log.getRange(logRow, 1, 1, LOG_HEADER.length).setValues([[
      tuan, ts, mien, gdv, qlkv, ma_ch, ten_ch, actionLabel, soDoiThu
    ]]);

    var added = 0, updated = 0;

    if (action === 'save') {
      var lastMain = main.getLastRow();
      var existingRows = lastMain > 1
        ? main.getRange(2, 1, lastMain - 1, MAIN_HEADER.length).getValues()
        : [];

      // Build ID index and next ID from one read rather than reading column A twice.
      var mapDT = Object.create(null);
      var maxDT = 0;
      for (var i = 0; i < existingRows.length; i++) {
        var existingId = String(existingRows[i][0] || '').trim();
        if (existingId) {
          mapDT[existingId] = i + 2;
          var match = existingId.match(/^DT-(\d+)$/);
          if (match) maxDT = Math.max(maxDT, parseInt(match[1], 10));
        }
      }

      var nextDT = maxDT + 1;
      var updatesByRow = Object.create(null);
      var newRows = [];

      for (var j = 0; j < competitors.length; j++) {
        var c = competitors[j] || {};
        var maDT = String(c.ma_dt || '').trim();
        var validId = maDT.match(/^DT-(\d+)$/);
        var rowNumber = validId ? mapDT[maDT] : null;

        // A valid-looking ID is not enough to update a row: it must already
        // exist and belong to this store. Otherwise issue a fresh unique ID.
        if (rowNumber && rowNumber <= lastMain) {
          var storedRow = existingRows[rowNumber - 2];
          if (!storedRow || String(storedRow[1] || '').trim() !== ma_ch) rowNumber = null;
        } else if (rowNumber) {
          var stagedRow = newRows[rowNumber - lastMain - 1];
          if (!stagedRow || String(stagedRow[1] || '').trim() !== ma_ch) rowNumber = null;
        }

        if (!rowNumber) {
          while (mapDT[DT_PREFIX + pad4(nextDT)]) nextDT++;
          maDT = DT_PREFIX + pad4(nextDT);
          nextDT++;
        } else {
          nextDT = Math.max(nextDT, parseInt(validId[1], 10) + 1);
        }

        var row = [
          maDT, ma_ch, ten_ch, mien, gdv, qlkv, data.tinh || '', data.huyen || '',
          c.nhom || '', c.ten || '', c.diaChi || '',
          (c.lat != null && c.lat !== '') ? String(c.lat) : '',
          (c.lng != null && c.lng !== '') ? String(c.lng) : '',
          (c.kc != null && c.kc !== '') ? String(c.kc) : '',
          parseVNDate(c.ngay) || '',
          (c.doanhSo != null && c.doanhSo !== '') ? String(c.doanhSo) : '',
          (c.luongKhach != null && c.luongKhach !== '') ? String(c.luongKhach) : '',
          (c.mt != null && c.mt !== '') ? String(c.mt) : '',
          (c.giaThue != null && c.giaThue !== '') ? String(c.giaThue) : '',
          (c.dienTich != null && c.dienTich !== '') ? String(c.dienTich) : '',
          '', ts
        ];

        if (rowNumber) {
          var oldRow = rowNumber <= lastMain
            ? existingRows[rowNumber - 2]
            : newRows[rowNumber - lastMain - 1];
          row[20] = (oldRow && oldRow[20]) || ts;
          if (rowNumber <= lastMain) updatesByRow[rowNumber] = row;
          else newRows[rowNumber - lastMain - 1] = row;
          updated++;
        } else {
          row[20] = ts;
          newRows.push(row);
          rowNumber = lastMain + newRows.length;
          mapDT[maDT] = rowNumber;
          added++;
        }
      }

      // Batch contiguous updates to existing rows; no per-row setValues calls.
      var updateNumbers = Object.keys(updatesByRow).map(Number).sort(function(a, b) { return a - b; });
      var pos = 0;
      while (pos < updateNumbers.length) {
        var startRow = updateNumbers[pos];
        var groupRows = [updatesByRow[startRow]];
        var endRow = startRow;
        pos++;
        while (pos < updateNumbers.length && updateNumbers[pos] === endRow + 1) {
          endRow = updateNumbers[pos];
          groupRows.push(updatesByRow[endRow]);
          pos++;
        }
        main.getRange(startRow, 1, groupRows.length, MAIN_HEADER.length).setValues(groupRows);
      }

      // Append all new rows in a single write.
      if (newRows.length) {
        main.getRange(lastMain + 1, 1, newRows.length, MAIN_HEADER.length).setValues(newRows);
      }
    }

    return out({
      ok: true,
      action: actionLabel,
      rows: competitors.length,
      added: added,
      updated: updated
    });
  } catch (err) {
    return out({ ok: false, error: err && err.message ? err.message : String(err) });
  } finally {
    if (lockAcquired) lock.releaseLock();
  }
}

function doGet(e) {
  var params = (e && e.parameter) || {};
  var action = params.action || 'data';
  if (action === 'data') {
    try {
      var ss = SpreadsheetApp.openById(SHEET_ID);
      var dataset = params.dataset || 'all';
      var result = { ok: true };
      if (dataset === 'all' || dataset === 'main') {
        var main = ss.getSheetByName(MAIN_SHEET);
        result.main = main ? main.getDataRange().getValues() : [];
      }
      if (dataset === 'all' || dataset === 'log') {
        var log = ss.getSheetByName(LOG_SHEET);
        result.log = log ? log.getDataRange().getValues() : [];
      }
      return out(result);
    } catch (err) {
      return out({ ok: false, error: err && err.message ? err.message : String(err) });
    }
  }
  return out({ ok: true, status: 'RSDT alive' });
}

function getISOWeek(d) {
  d = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
  d.setUTCDate(d.getUTCDate() + 4 - (d.getUTCDay() || 7));
  var y = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  return Math.ceil(((d - y) / 86400000 + 1) / 7);
}

function pad4(n) {
  return String(n).padStart(4, '0');
}

function parseVNDate(raw) {
  if (!raw) return '';
  var s = String(raw).trim();
  var m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) return new Date(parseInt(m[3], 10), parseInt(m[2], 10) - 1, parseInt(m[1], 10));
  m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m) return new Date(parseInt(m[1], 10), parseInt(m[2], 10) - 1, parseInt(m[3], 10));
  m = s.match(/^(\d{4})\/(\d{2})\/(\d{2})$/);
  if (m) return new Date(parseInt(m[1], 10), parseInt(m[2], 10) - 1, parseInt(m[3], 10));
  return s;
}

function out(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
