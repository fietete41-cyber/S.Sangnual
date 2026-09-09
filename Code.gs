/**
 * ส.แสงนวล ผลไม้ — ระบบสต็อก + บันทึกการขาย (Backend / Google Apps Script)
 * ---------------------------------------------------------------------------
 * ตั้งค่าครั้งเดียว:
 *   1) สร้าง Google Sheet ใหม่ 1 ไฟล์
 *   2) Extensions > Apps Script  แล้ววางโค้ดนี้ทั้งหมด กด 💾
 *   3) (ถ้าต้องการรหัส) ⚙ Project Settings > Script properties > เพิ่ม  PIN = 1234
 *   4) เมนู Run > เลือกฟังก์ชัน  setup  แล้วกด Run (ครั้งแรกต้องกด Authorize)
 *   5) Deploy > New deployment > Web app
 *        Execute as: Me   |   Who has access: Anyone
 *      คัดลอก URL ที่ลงท้าย /exec ไปวางในหน้า "ตั้งค่า" ของแอป
 *   ** แก้โค้ดภายหลัง: Deploy > Manage deployments > ✏ > Version: New version > Deploy **
 */

var SH = { stock: 'สต็อก', sales: 'การขาย', cost: 'ทุนรายวัน', adjust: 'ปรับสต็อก' };

var STOCK_H  = ['id','รายการ','คงเหลือลัง','คงเหลือโล','โลต่อลัง','ราคาต่อโล','ทุนต่อโล','ขั้นต่ำลัง','ขั้นต่ำโล','อัปเดตล่าสุด'];
var SALES_H  = ['billId','วันที่เวลา','รอบ','ปลายทาง','ลูกค้า','ช่องทาง','สถานะ','รายการ','ลัง','โลต่อลัง','แถมโล','โลรวม','ราคาต่อโล','ทุนต่อโล','ยอดเงิน','หมายเหตุ'];
var COST_H   = ['วันที่','ทุน','หมายเหตุ'];
var ADJUST_H = ['id','วันที่เวลา','รายการ','ลัง','โล','เหตุผล'];

/* ---------------- entry points ---------------- */
function doGet(e)  { return handle((e && e.parameter) || {}, ((e && e.parameter && e.parameter.action) || '')); }
function doPost(e) {
  var b = {};
  try { b = JSON.parse(e.postData.contents); } catch (err) {}
  return handle(b, b.action || '');
}

function handle(p, action) {
  try {
    checkPin(p.pin);
    var out;
    switch (action) {
      case '':
      case 'ping':       out = { pong: true }; break;
      case 'list':       out = listAll(); break;
      case 'addSale':    out = addSale(p); break;
      case 'deleteSale': out = deleteSale(p); break;
      case 'setStatus':  out = setStatus(p); break;
      case 'saveItem':   out = saveItem(p); break;
      case 'deleteItem': out = deleteItem(p); break;
      case 'adjust':     out = adjustStock(p); break;
      case 'saveCost':   out = saveCost(p); break;
      default: throw new Error('ไม่รู้จักคำสั่ง: ' + action);
    }
    return json(mix({ ok: true }, out));
  } catch (err) {
    return json({ ok: false, error: String((err && err.message) || err) });
  }
}

/* ---------------- helpers ---------------- */
function json(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function mix(a, b) { for (var k in b) a[k] = b[k]; return a; }
function num(v) { var n = parseFloat(v); return isNaN(n) ? 0 : n; }
var _uidN = 0;
function uid(pre) { return pre + Date.now() + '_' + (_uidN++) + '_' + Math.floor(Math.random() * 1e6); }
function checkPin(pin) {
  var want = PropertiesService.getScriptProperties().getProperty('PIN');
  if (want && String(pin == null ? '' : pin) !== String(want)) throw new Error('PIN ไม่ถูกต้อง');
}
function ss()  { return SpreadsheetApp.getActiveSpreadsheet(); }
function tz()  { return ss().getSpreadsheetTimeZone() || 'Asia/Bangkok'; }
function now() { return Utilities.formatDate(new Date(), tz(), "yyyy-MM-dd'T'HH:mm:ss"); }
function today() { return Utilities.formatDate(new Date(), tz(), 'yyyy-MM-dd'); }
function toIso(v) {
  if (Object.prototype.toString.call(v) === '[object Date]')
    return Utilities.formatDate(v, tz(), "yyyy-MM-dd'T'HH:mm:ss");
  return String(v == null ? '' : v);
}
function sheet(name, headers) {
  var s = ss().getSheetByName(name);
  if (!s) s = ss().insertSheet(name);
  if (headers && s.getLastRow() === 0) {
    s.getRange(1, 1, 1, headers.length).setValues([headers]);
    s.setFrozenRows(1);
  }
  return s;
}
function rows(name, headers) {
  var s = sheet(name, headers), lr = s.getLastRow();
  if (lr < 2) return [];
  return s.getRange(2, 1, lr - 1, headers.length).getValues().map(function (r) {
    var o = {}; headers.forEach(function (h, i) { o[h] = r[i]; }); return o;
  });
}
function writeRow(s, r1, headers, obj) {
  s.getRange(r1, 1, 1, headers.length).setValues([headers.map(function (h) { return obj[h]; })]);
}
function idxById(a, id)     { for (var i = 0; i < a.length; i++) if (String(a[i].id) === String(id)) return i; return -1; }
function idxByName(a, name) { for (var i = 0; i < a.length; i++) if (String(a[i]['รายการ']).trim() === String(name).trim()) return i; return -1; }

/* ---------------- actions ---------------- */
function listAll() {
  return {
    stock: rows(SH.stock, STOCK_H).filter(function (r) { return String(r.id) !== ''; }),
    sales: rows(SH.sales, SALES_H).filter(function (r) { return String(r.billId) !== ''; })
                                  .map(function (r) { r['วันที่เวลา'] = toIso(r['วันที่เวลา']); return r; }),
    costs: rows(SH.cost, COST_H).filter(function (r) { return String(r['วันที่']) !== ''; })
                                .map(function (r) { r['วันที่'] = toIso(r['วันที่']).slice(0, 10); return r; })
  };
}

function saveItem(p) {
  var s = sheet(SH.stock, STOCK_H), data = rows(SH.stock, STOCK_H);
  var id = p.id || uid('I');
  var o = {
    id: id,
    'รายการ': String(p['รายการ'] || '').trim(),
    'คงเหลือลัง': num(p['คงเหลือลัง']), 'คงเหลือโล': num(p['คงเหลือโล']),
    'โลต่อลัง': num(p['โลต่อลัง']) || 25,
    'ราคาต่อโล': num(p['ราคาต่อโล']), 'ทุนต่อโล': num(p['ทุนต่อโล']),
    'ขั้นต่ำลัง': num(p['ขั้นต่ำลัง']), 'ขั้นต่ำโล': num(p['ขั้นต่ำโล']),
    'อัปเดตล่าสุด': now()
  };
  if (!o['รายการ']) throw new Error('ต้องมีชื่อรายการ');
  var i = idxById(data, id);
  if (i === -1) s.appendRow(STOCK_H.map(function (h) { return o[h]; }));
  else          writeRow(s, i + 2, STOCK_H, o);
  return { item: o };
}

function deleteItem(p) {
  var s = sheet(SH.stock, STOCK_H), i = idxById(rows(SH.stock, STOCK_H), p.id);
  if (i === -1) throw new Error('ไม่พบสินค้า');
  s.deleteRow(i + 2);
  return { deleted: 1 };
}

function addSale(p) {
  var bill = p.bill || {};
  var items = (p.items || []).filter(function (it) {
    return String(it['รายการ'] || '').trim() && (num(it['ลัง']) || num(it['โลรวม']) || num(it['ยอดเงิน']));
  });
  if (!items.length) throw new Error('ไม่มีรายการสินค้า');

  var billId = uid('B'), ts = now();
  var s = sheet(SH.sales, SALES_H);
  var lines = items.map(function (it) {
    return [
      billId, ts, bill['รอบ'] || '', bill['ปลายทาง'] || '', bill['ลูกค้า'] || '',
      bill['ช่องทาง'] || '', bill['สถานะ'] || 'จ่ายแล้ว',
      String(it['รายการ']).trim(), num(it['ลัง']), num(it['โลต่อลัง']) || 25,
      num(it['แถมโล']), num(it['โลรวม']), num(it['ราคาต่อโล']), num(it['ทุนต่อโล']),
      num(it['ยอดเงิน']), bill['หมายเหตุ'] || ''
    ];
  });
  s.getRange(s.getLastRow() + 1, 1, lines.length, SALES_H.length).setValues(lines);

  var warn = [], st = sheet(SH.stock, STOCK_H), sd = rows(SH.stock, STOCK_H);
  items.forEach(function (it) {
    var i = idxByName(sd, it['รายการ']);
    if (i === -1) { warn.push('ไม่พบในสต็อก: ' + it['รายการ']); return; }
    sd[i]['คงเหลือลัง'] = num(sd[i]['คงเหลือลัง']) - num(it['ลัง']);
    sd[i]['คงเหลือโล']  = num(sd[i]['คงเหลือโล'])  - num(it['โลรวม']);
    sd[i]['อัปเดตล่าสุด'] = ts;
    writeRow(st, i + 2, STOCK_H, sd[i]);
  });
  return { billId: billId, warnings: warn };
}

function deleteSale(p) {
  if (!p.billId) throw new Error('ไม่ระบุบิล');
  var restock = p.restock !== false;
  var s = sheet(SH.sales, SALES_H), data = rows(SH.sales, SALES_H);
  var t = [];
  data.forEach(function (r, i) { if (String(r.billId) === String(p.billId)) t.push(i); });
  if (!t.length) throw new Error('ไม่พบบิล');

  if (restock) {
    var st = sheet(SH.stock, STOCK_H), sd = rows(SH.stock, STOCK_H);
    t.forEach(function (i) {
      var r = data[i], k = idxByName(sd, r['รายการ']);
      if (k === -1) return;
      sd[k]['คงเหลือลัง'] = num(sd[k]['คงเหลือลัง']) + num(r['ลัง']);
      sd[k]['คงเหลือโล']  = num(sd[k]['คงเหลือโล'])  + num(r['โลรวม']);
      sd[k]['อัปเดตล่าสุด'] = now();
      writeRow(st, k + 2, STOCK_H, sd[k]);
    });
  }
  t.sort(function (a, b) { return b - a; }).forEach(function (i) { s.deleteRow(i + 2); });
  return { deleted: t.length };
}

function setStatus(p) {
  if (!p.billId) throw new Error('ไม่ระบุบิล');
  var s = sheet(SH.sales, SALES_H), data = rows(SH.sales, SALES_H), n = 0;
  var cS = SALES_H.indexOf('สถานะ') + 1, cC = SALES_H.indexOf('ช่องทาง') + 1;
  data.forEach(function (r, i) {
    if (String(r.billId) !== String(p.billId)) return;
    if (p['สถานะ'] != null) s.getRange(i + 2, cS).setValue(p['สถานะ']);
    if (p['ช่องทาง'] != null) s.getRange(i + 2, cC).setValue(p['ช่องทาง']);
    n++;
  });
  if (!n) throw new Error('ไม่พบบิล');
  return { updated: n };
}

function adjustStock(p) {
  var st = sheet(SH.stock, STOCK_H), sd = rows(SH.stock, STOCK_H), i = idxById(sd, p.id);
  if (i === -1) throw new Error('ไม่พบสินค้า');
  var dL = num(p['ลัง']), dK = num(p['โล']);
  sd[i]['คงเหลือลัง'] = num(sd[i]['คงเหลือลัง']) + dL;
  sd[i]['คงเหลือโล']  = num(sd[i]['คงเหลือโล'])  + dK;
  sd[i]['อัปเดตล่าสุด'] = now();
  writeRow(st, i + 2, STOCK_H, sd[i]);
  sheet(SH.adjust, ADJUST_H).appendRow([uid('A'), now(), sd[i]['รายการ'], dL, dK, p['เหตุผล'] || '']);
  return { item: sd[i] };
}

function saveCost(p) {
  var d = String(p['วันที่'] || today()).slice(0, 10);
  var s = sheet(SH.cost, COST_H), data = rows(SH.cost, COST_H);
  var i = -1;
  data.forEach(function (r, k) { if (String(toIso(r['วันที่']).slice(0, 10)) === d) i = k; });
  var o = { 'วันที่': d, 'ทุน': num(p['ทุน']), 'หมายเหตุ': p['หมายเหตุ'] || '' };
  if (i === -1) s.appendRow([o['วันที่'], o['ทุน'], o['หมายเหตุ']]);
  else          writeRow(s, i + 2, COST_H, o);
  return { cost: o };
}

/* ---------------- one-time setup ---------------- */
function setup() {
  sheet(SH.stock, STOCK_H);
  sheet(SH.sales, SALES_H);
  sheet(SH.cost, COST_H);
  sheet(SH.adjust, ADJUST_H);
  if (!rows(SH.stock, STOCK_H).length) {
    saveItem({ 'รายการ': 'กิมจูยอด1', 'คงเหลือลัง': 5, 'คงเหลือโล': 125, 'โลต่อลัง': 25, 'ราคาต่อโล': 40, 'ทุนต่อโล': 30, 'ขั้นต่ำลัง': 2 });
    saveItem({ 'รายการ': 'ต้าลิน',     'คงเหลือลัง': 0, 'คงเหลือโล': 0,   'โลต่อลัง': 25, 'ราคาต่อโล': 37, 'ทุนต่อโล': 28 });
  }
}
