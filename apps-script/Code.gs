/**
 * TMJH 訂餐站 — Google Apps Script 後端
 *
 * 資料存在這個 Apps Script 所綁定的 Google 試算表裡：
 *   「菜單」工作表：每列一份菜單（最後一欄是完整 JSON）
 *   「訂單」工作表：每列一筆訂單（可以直接在試算表裡看、排序、列印）
 *
 * 需要在「專案設定 → 指令碼屬性」設定：
 *   ADMIN_PASSWORD     管理菜單用的密碼（必填）
 *   ANTHROPIC_API_KEY  Claude API 金鑰（選填；沒填就只能手動建立菜單）
 */

var MENU_SHEET = '菜單';
var ORDER_SHEET = '訂單';
var MENU_HEAD = ['id', '名稱', '類型', '開放', '建立時間', '資料JSON'];
var ORDER_HEAD = ['時間', '菜單', '訂購人', '品項', '規格', '選項', '備註', '數量', '單價', '金額', 'id', 'menuId', 'key', 'optsJSON'];
var CLAUDE_MODEL = 'claude-opus-5-5';

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
  return { menus: readMenus_(), orders: readOrders_(), aiReady: !!prop_('ANTHROPIC_API_KEY') };
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

// ---------- 菜單辨識（呼叫 Claude API） ----------
function recognizeMenu(pw, images, forcedType) {
  checkAdmin_(pw);
  var key = prop_('ANTHROPIC_API_KEY');
  if (!key) throw new Error('尚未設定 ANTHROPIC_API_KEY，無法自動辨識。可以先用手動建立。');
  if (!images || !images.length) throw new Error('請先選擇菜單照片');

  var prompt = '你是台灣餐飲菜單辨識助手。請閱讀附上的菜單照片（可能多張，屬於同一家店），把所有可點的品項整理成 JSON。\n' +
    '規則：\n' +
    '- 只輸出 JSON，不要任何說明文字或 markdown。\n' +
    '- 價格用新台幣整數。看不清楚的價格填 0。\n' +
    '- 同一品項有多種尺寸／價格（例如中杯 M、大杯 L，或小/大份）時，放在 variants 陣列，price 填最低價；只有一種價格時 variants 給空陣列。\n' +
    '- category 用菜單上的分類（例如：招牌便當、炸物、茶類、奶茶、鮮奶），沒有分類就填「主餐」或「飲品」。\n' +
    '- addons：菜單上明確列出的加購／加料／升級項目（例如：加珍珠 +10、加蛋 +10、升級套餐 +30），沒有就給空陣列。\n' +
    '- type：便當／餐點類填 "bento"，飲料店填 "drink"。' + (forcedType === 'bento' || forcedType === 'drink' ? '本菜單已確定是 "' + forcedType + '"。' : '') + '\n' +
    '- store：店名（看得到的話），否則空字串。\n' +
    '格式：{"store":"","type":"bento","items":[{"category":"","name":"","price":0,"variants":[{"name":"M","price":0}]}],"addons":[{"group":"加料","name":"珍珠","price":10}]}';

  var content = images.slice(0, 6).map(function (img) {
    return { type: 'image', source: { type: 'base64', media_type: img.mime || 'image/jpeg', data: img.data } };
  });
  content.push({ type: 'text', text: prompt });

  var res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post',
    contentType: 'application/json',
    muteHttpExceptions: true,
    headers: {
      'x-api-key': key,
      'anthropic-version': '2026-10-21',
      'anthropic-beta': 'server-side-fallback-2026-07-01'
    },
    payload: JSON.stringify({
      model: CLAUDE_MODEL,
      max_tokens: 16000,
      output_config: { effort: 'low' },
      fallbacks: 'default',
      messages: [{ role: 'user', content: content }]
    })
  });

  var code = res.getResponseCode();
  var body = {};
  try { body = JSON.parse(res.getContentText()); } catch (e) {}
  if (code === 401) throw new Error('Claude API 金鑰無效，請檢查 ANTHROPIC_API_KEY。');
  if (code === 429) throw new Error('辨識次數太頻繁，請稍等一分鐘再試。');
  if (code !== 200) throw new Error('辨識失敗（' + code + '）：' + ((body.error && body.error.message) || '請再試一次'));
  if (body.stop_reason === 'refusal') throw new Error('這張圖片無法辨識，請換一張菜單照片。');

  var text = (body.content || []).filter(function (b) { return b.type === 'text'; }).map(function (b) { return b.text; }).join('');
  var s = text.indexOf('{'), e = text.lastIndexOf('}');
  if (s < 0 || e < s) throw new Error('沒有辨識到品項，請換一張更清楚、正面拍攝的照片。');
  var data;
  try { data = JSON.parse(text.slice(s, e + 1)); } catch (err) { throw new Error('辨識結果格式不正確，請再試一次。'); }
  if (!data || !data.items || !data.items.length) throw new Error('沒有辨識到品項，請換一張更清楚、正面拍攝的照片。');
  return data;
}
