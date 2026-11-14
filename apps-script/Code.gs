/**
 * TMJH 訂餐站 — Google Apps Script 後端
 *
 * 資料存在這個 Apps Script 所綁定的 Google 試算表裡：
 *   「菜單」工作表：每列一份菜單（最後一欄是完整 JSON）
 *   「訂單」工作表：每列一筆訂單（可以直接在試算表裡看、排序、列印）
 *
 * 需要在「專案設定 → 指令碼屬性」設定：
 *   ADMIN_PASSWORD     管理菜單用的密碼（必填）
 *
 * 菜單照片辨識使用 Google 雲端硬碟內建的文字辨識（免費），
 * 需要在左側「服務 ＋」加入 Drive API。
 */

var MENU_SHEET = '菜單';
var ORDER_SHEET = '訂單';
var MENU_HEAD = ['id', '名稱', '類型', '開放', '建立時間', '資料JSON'];
var ORDER_HEAD = ['時間', '菜單', '訂購人', '品項', '規格', '選項', '備註', '數量', '單價', '金額', 'id', 'menuId', 'key', 'optsJSON'];

// ---------- 網頁 ----------
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('TMJH 訂餐站')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ---------- 試算表工具 ----------
function sheet_(name, head) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(head);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, head.length).setFontWeight('bold');
  }
  return sh;
}
function rows_(sh) {
  var last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
}
function findRow_(sh, col, value) {
  var r = rows_(sh);
  for (var i = 0; i < r.length; i++) if (String(r[i][col]) === String(value)) return i + 2;
  return -1;
}
function prop_(k) { return PropertiesService.getScriptProperties().getProperty(k) || ''; }
function checkAdmin_(pw) {
  var real = prop_('ADMIN_PASSWORD');
  if (!real) throw new Error('尚未設定管理密碼（ADMIN_PASSWORD），請到 Apps Script 的專案設定新增。');
  if (String(pw) !== real) throw new Error('管理密碼錯誤');
}
function withLock_(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}
function id_() { return Utilities.getUuid().replace(/-/g, '').slice(0, 12); }
function str_(v, max) { return String(v == null ? '' : v).trim().slice(0, max || 200); }

function readMenus_() {
  return rows_(sheet_(MENU_SHEET, MENU_HEAD)).map(function (r) {
    var m = {};
    try { m = JSON.parse(r[5] || '{}'); } catch (e) {}
    m.id = String(r[0]);
    m.open = r[3] === true || r[3] === 'TRUE' || r[3] === '開放';
    m.createdAt = Number(r[4]) || 0;
    return m;
  });
}
function readOrders_() {
  return rows_(sheet_(ORDER_SHEET, ORDER_HEAD)).map(function (r) {
    var opts = [];
    try { opts = JSON.parse(r[13] || '[]'); } catch (e) {}
    return {
      at: r[0] instanceof Date ? r[0].getTime() : Number(new Date(r[0])) || 0,
      menuName: String(r[1]), name: String(r[2]), item: String(r[3]), variant: String(r[4]),
      note: String(r[6]), qty: Number(r[7]) || 1, unit: Number(r[8]) || 0, total: Number(r[9]) || 0,
      id: String(r[10]), menuId: String(r[11]), opts: opts
    };
  });
}

// ---------- 公開 API（任何人都能呼叫） ----------
function getData() {
  return { menus: readMenus_(), orders: readOrders_(), aiReady: ocrReady_() };
}

function submitOrder(o) {
  return withLock_(function () {
    var menu = readMenus_().filter(function (m) { return m.id === String(o.menuId); })[0];
    if (!menu) throw new Error('找不到這份菜單，請重新整理頁面。');
    if (!menu.open) throw new Error('這份菜單已經截止點餐了。');
    var name = str_(o.name, 40);
    if (!name) throw new Error('請填寫訂購人');
    var item = (menu.items || []).filter(function (x) { return x.id === o.itemId; })[0];
    if (!item) throw new Error('找不到這個品項，請重新整理頁面。');

    // 在伺服器端重新計算價格，避免被竄改
    var unit = Number(item.price) || 0, variant = '';
    if (item.variants && item.variants.length) {
      var v = item.variants.filter(function (x) { return x.name === o.variant; })[0];
      if (!v) throw new Error('請選擇尺寸／規格');
      unit = Number(v.price) || 0; variant = v.name;
    }
    var opts = [];
    (menu.options || []).forEach(function (g) {
      var picked = (o.opts && o.opts[g.id]) || [];
      if (!g.multi) picked = picked.slice(0, 1);
      if (g.required && !picked.length) throw new Error('請選擇「' + g.name + '」');
      picked.forEach(function (cn) {
        var c = g.choices.filter(function (x) { return x.name === cn; })[0];
        if (c) { unit += Number(c.price) || 0; opts.push({ g: g.name, c: c.name, p: Number(c.price) || 0 }); }
      });
    });
    var qty = Math.max(1, Math.min(99, parseInt(o.qty, 10) || 1));
    var id = id_(), key = id_();
    var optText = opts.map(function (x) { return x.c + (x.p ? '(+' + x.p + ')' : ''); }).join('、');
    sheet_(ORDER_SHEET, ORDER_HEAD).appendRow([new Date(), menu.name, name, item.name, variant, optText, str_(o.note, 100), qty, unit, unit * qty, id, menu.id, key, JSON.stringify(opts)]);
    return { id: id, key: key };
  });
}

function deleteOrder(id, key, pw) {
  return withLock_(function () {
    var sh = sheet_(ORDER_SHEET, ORDER_HEAD);
    var row = findRow_(sh, 10, id);
    if (row < 0) return true;
    var realKey = String(sh.getRange(row, 13).getValue());
    if (!(key && String(key) === realKey)) checkAdmin_(pw);
    sh.deleteRow(row);
    return true;
  });
}

// ---------- 管理 API（需要密碼） ----------
function adminLogin(pw) { checkAdmin_(pw); return true; }

function saveMenu(pw, m) {
  checkAdmin_(pw);
  return withLock_(function () {
    var sh = sheet_(MENU_SHEET, MENU_HEAD);
    var clean = {
      name: str_(m.name, 60) || '未命名菜單', type: m.type === 'drink' ? 'drink' : 'bento',
      deadline: str_(m.deadline, 60), note: str_(m.note, 200),
      items: (m.items || []).slice(0, 300).map(function (it) {
        return { id: str_(it.id, 20) || id_(), category: str_(it.category, 40), name: str_(it.name, 60), price: Number(it.price) || 0,
          variants: (it.variants || []).slice(0, 10).map(function (v) { return { name: str_(v.name, 20), price: Number(v.price) || 0 }; }) };
      }),
      options: (m.options || []).slice(0, 30).map(function (g) {
        return { id: str_(g.id, 20) || id_(), name: str_(g.name, 40), multi: !!g.multi, required: !!g.required,
          choices: (g.choices || []).slice(0, 40).map(function (c) { return { name: str_(c.name, 40), price: Number(c.price) || 0 }; }) };
      })
    };
    var json = JSON.stringify(clean);
    if (m.id) {
      var row = findRow_(sh, 0, m.id);
      if (row > 0) {
        sh.getRange(row, 2, 1, 5).setValues([[clean.name, clean.type === 'drink' ? '飲料' : '便當', !!m.open, sh.getRange(row, 5).getValue(), json]]);
        return m.id;
      }
    }
    var id = id_();
    sh.appendRow([id, clean.name, clean.type === 'drink' ? '飲料' : '便當', !!m.open, Date.now(), json]);
    return id;
  });
}

function setMenuOpen(pw, id, open) {
  checkAdmin_(pw);
  return withLock_(function () {
    var sh = sheet_(MENU_SHEET, MENU_HEAD);
    var row = findRow_(sh, 0, id);
    if (row > 0) sh.getRange(row, 4).setValue(!!open);
    return true;
  });
}

function deleteMenu(pw, id) {
  checkAdmin_(pw);
  return withLock_(function () {
    var osh = sheet_(ORDER_SHEET, ORDER_HEAD);
    var r = rows_(osh);
    for (var i = r.length - 1; i >= 0; i--) if (String(r[i][11]) === String(id)) osh.deleteRow(i + 2);
    var sh = sheet_(MENU_SHEET, MENU_HEAD);
    var row = findRow_(sh, 0, id);
    if (row > 0) sh.deleteRow(row);
    return true;
  });
}

// ---------- 菜單辨識（Google 雲端硬碟內建文字辨識，免費、不需金鑰） ----------
function ocrReady_() { return typeof Drive !== 'undefined' && !!Drive.Files; }

function ocrImage_(img, i) {
  var blob = Utilities.newBlob(Utilities.base64Decode(img.data), img.mime || 'image/jpeg', 'menu' + i + '.jpg');
  var file;
  if (Drive.Files.create) { // Drive API v3
    file = Drive.Files.create({ name: 'TMJH 菜單辨識暫存', mimeType: 'application/vnd.google-apps.document' }, blob, { ocrLanguage: 'zh-TW' });
  } else { // Drive API v2
    file = Drive.Files.insert({ title: 'TMJH 菜單辨識暫存', mimeType: blob.getContentType() }, blob, { ocr: true, ocrLanguage: 'zh-TW' });
  }
  try {
    return DocumentApp.openById(file.id).getBody().getText();
  } finally {
    try { DriveApp.getFileById(file.id).setTrashed(true); } catch (e) {}
  }
}

function recognizeMenu(pw, images, forcedType) {
  checkAdmin_(pw);
  if (!ocrReady_()) throw new Error('尚未開啟「Drive API」服務。請到 Apps Script 左側「服務 ＋」加入 Drive API 後再試。');
  if (!images || !images.length) throw new Error('請先選擇菜單照片');
  var text = images.slice(0, 6).map(ocrImage_).join('\n');
  var data = parseMenu_(text, forcedType);
  if (!data.items.length) throw new Error('照片裡沒有找到「品名＋價格」。請換一張更清楚、正面拍攝的照片，或改用「貼上文字」。');
  data.rawText = text;
  return data;
}

function parseMenuText(pw, text, forcedType) {
  checkAdmin_(pw);
  text = String(text || '').trim();
  if (!text) throw new Error('請先貼上菜單文字');
  // 允許直接貼上 JSON（例如請 Claude 聊天整理好的結果）
  var s = text.indexOf('{'), e = text.lastIndexOf('}');
  if (s >= 0 && e > s) {
    try {
      var j = JSON.parse(text.slice(s, e + 1));
      if (j && j.items && j.items.length) return j;
    } catch (err) {}
  }
  var data = parseMenu_(text, forcedType);
  if (!data.items.length) throw new Error('沒有找到「品名＋價格」。每行請寫成像「排骨便當 90」或「紅茶 M30 L40」的格式。');
  return data;
}

/** 把一段菜單文字整理成 {store, type, items, addons}。盡量猜，不準的部分讓管理員在畫面上修改。 */
function parseMenu_(text, forcedType) {
  var SIZE_WORDS = /^(M|L|XL|S|中|大|小|特大|中杯|大杯|小杯|瓶|冰|熱)$/i;
  var lines = String(text).replace(/\r/g, '').split('\n')
    .map(function (l) { return l.replace(/[　\t]+/g, ' ').replace(/[｜|•●○◎★☆※]/g, ' ').trim(); })
    .filter(function (l) { return l; });
  var items = [], addons = [], category = '', prevCategory = '', sizeLabels = null, pendingName = '', store = '';

  function toHalf(str) { return str.replace(/＄/g, '$').replace(/[０-９]/g, function (c) { return String.fromCharCode(c.charCodeAt(0) - 0xFEE0); }); }
  function prices(line) {
    var out = [], re = /(?:([A-Za-z一-鿿]{1,2})\s*[:：]?\s*)?\$?\s*(\d{2,4})\s*(?:元)?/g, m;
    while ((m = re.exec(line))) {
      var n = +m[2];
      if (n >= 5 && n <= 3000) out.push({ label: m[1] && SIZE_WORDS.test(m[1]) ? m[1] : '', price: n, index: m.index });
    }
    return out;
  }
  function cleanName(str) {
    return str.replace(/\$?\s*\d{2,4}\s*元?/g, ' ').replace(/(^|\s)(M|L|XL|中杯|大杯|小杯)\s*[:：]?(?=\s|$)/gi, ' ')
      .replace(/[.。…·\-–—_~～:：、,，()（）\[\]【】]+/g, ' ').replace(/\s+/g, ' ').trim();
  }
  function addItem(name, ps) {
    var variants = [];
    if (ps.length > 1) {
      variants = ps.map(function (p, i) {
        return { name: p.label || (sizeLabels && sizeLabels[i]) || (ps.length === 2 ? ['M', 'L'][i] : '規格' + (i + 1)), price: p.price };
      });
    }
    var item = { category: category || '', name: name, price: Math.min.apply(null, ps.map(function (p) { return p.price; })), variants: variants };
    if (/^加/.test(name) && ps.length === 1) addons.push({ group: '加購', name: name.replace(/^加\s*/, '加'), price: ps[0].price });
    else items.push(item);
  }

  lines.forEach(function (raw) {
    var line = toHalf(raw);
    // 表頭：「M L」、「中杯 大杯」之類，記下尺寸名稱
    var words = line.split(/\s+/);
    if (words.length >= 2 && words.length <= 4 && words.every(function (w) { return SIZE_WORDS.test(w); })) { sizeLabels = words; return; }
    // 電話、地址、營業時間等不是品項
    if (/電話|tel|地址|營業|外送|line|訂購專線|\d{2,4}-\d{3,4}-?\d{3,4}|\d{1,2}:\d{2}/i.test(line)) return;
    var ps = prices(line);
    var name = cleanName(line);
    if (!ps.length) {
      // 沒有價格的短句：可能是分類標題，也可能是品名（價格在下一行）
      if (name && name.length <= 12 && !/\d/.test(name)) {
        if (!store && !items.length && !category) { store = name; return; }
        prevCategory = category; category = name; pendingName = name;
      }
      return;
    }
    if (!name) {
      // 只有價格的一行：接到上一行的品名
      if (pendingName) { category = prevCategory; addItem(pendingName, ps); }
      pendingName = '';
      return;
    }
    pendingName = '';
    addItem(name, ps);
  });

  var all = items.map(function (i) { return i.name + i.category; }).join('');
  var drinkScore = (all.match(/茶|奶|拿鐵|咖啡|果汁|冰沙|多多|飲|鮮乳|養樂多|檸檬|冬瓜|綠|烏龍/g) || []).length;
  var bentoScore = (all.match(/飯|便當|麵|雞|豬|排骨|魚|牛|餐|丼|咖哩|滷|炒/g) || []).length;
  var type = forcedType === 'bento' || forcedType === 'drink' ? forcedType : (drinkScore > bentoScore ? 'drink' : 'bento');
  items.forEach(function (i) { if (!i.category) i.category = type === 'drink' ? '飲品' : '主餐'; });
  return { store: store, type: type, items: items, addons: addons };
}
