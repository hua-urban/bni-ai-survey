/**
 * BNI 華都更分會｜AI 課程課前問卷
 * 前後端都在這支 GAS，帳號＝hua.urban.renewal（依《分會系統建置與交接原則》原則 1）
 *
 * 問卷： .../exec
 * 統計： .../exec?p=stats
 * 資料： 試算表「AI課程課前問卷_回覆」
 *
 * 改版流程：改完 code → 部署 → 管理部署作業 → 編輯（鉛筆）→ 版本＝新版本 → 部署。
 * ⚠️ 絕不用 clasp deploy，會把「任何人」存取打回「需登入」。
 */

var SHEET_ID = '16eiKAtX-o_-GLF-2gwgWmiMwoF5Xa9n-NuEhLPG1ihE';
var SHEET_NAME = '回覆';

var FIELDS = [
  ['ts',  '填答時間'],
  ['q1',  'Q1 姓名與編號'],
  ['q2',  'Q2 使用程度'],
  ['q3',  'Q3 在用哪些工具'],
  ['q4',  'Q4 最近一週做過'],
  ['q5',  'Q5 卡在哪（最多3）'],
  ['q6',  'Q6 最想解決的一件事'],
  ['q7',  'Q7 已經試過什麼'],
  ['q8',  'Q8 想上哪些單元（最多3）'],
  ['q9',  'Q9 方式與時段'],
  ['q10', 'Q10 合理費用'],
  ['q11', 'Q11 只能先開一個'],
  ['q12', 'Q12 會不會第一批報名'],
  ['q13', 'Q13 願不願意分享'],
  ['q14', 'Q14 聯絡方式'],
  ['q15', 'Q15 想說的話'],
  ['ua',  '裝置']
];

var REQUIRED = ['q1','q2','q3','q4','q5','q6','q8','q9','q10','q11','q12','q13'];

function sheet_() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  return sh;
}

function setup() {
  var sh = sheet_();
  var head = FIELDS.map(function (f) { return f[1]; });
  sh.getRange(1, 1, 1, head.length).setValues([head]);
  sh.getRange(1, 1, 1, head.length).setFontWeight('bold').setBackground('#F5F3EF');
  sh.setFrozenRows(1);
  sh.setColumnWidth(1, 150);
  return '標題列已建立，共 ' + head.length + ' 欄';
}

/** 刪掉測試列（Q1 以「000」開頭），正式發布前跑一次 */
function clearTestRows() {
  var sh = sheet_();
  var last = sh.getLastRow();
  if (last < 2) return '沒有資料';
  var col = sh.getRange(2, 2, last - 1, 1).getValues();
  var removed = 0;
  for (var i = col.length - 1; i >= 0; i--) {
    if (String(col[i][0]).indexOf('000') === 0) { sh.deleteRow(i + 2); removed++; }
  }
  return '刪除 ' + removed + ' 列測試資料';
}

// ── 寫入 ──────────────────────────────────────────────

function save_(data) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    for (var i = 0; i < REQUIRED.length; i++) {
      if (!data[REQUIRED[i]] || String(data[REQUIRED[i]]).trim() === '') {
        return { ok: false, error: '缺少必填：' + REQUIRED[i] };
      }
    }
    var sh = sheet_();
    if (sh.getLastRow() === 0) setup();
    var row = FIELDS.map(function (f) {
      var k = f[0];
      if (k === 'ts') return new Date();
      return data[k] == null ? '' : String(data[k]).slice(0, 2000);
    });
    sh.appendRow(row);
    return { ok: true, n: sh.getLastRow() - 1 };
  } catch (err) {
    return { ok: false, error: String(err) };
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}

/** 給前端 google.script.run 呼叫 */
function submitSurvey(data) {
  return save_(data);
}

/** 保留給外部 POST（測試、其他前端） */
function doPost(e) {
  var out;
  try {
    out = save_(JSON.parse(e.postData.contents));
  } catch (err) {
    out = { ok: false, error: String(err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── 統計 ──────────────────────────────────────────────

function getStats() {
  try {
    var sh = sheet_();
    var last = sh.getLastRow();
    if (last < 2) return { ok: true, total: 0 };

    var head = FIELDS.map(function (f) { return f[0]; });
    var values = sh.getRange(2, 1, last - 1, head.length).getValues();
    function idx(k) { return head.indexOf(k); }

    function tally(key, split) {
      var c = {}, col = idx(key);
      values.forEach(function (r) {
        var raw = String(r[col] || '').trim();
        if (!raw) return;
        (split ? raw.split(' / ') : [raw]).forEach(function (it) {
          it = it.trim();
          if (it) c[it] = (c[it] || 0) + 1;
        });
      });
      return c;
    }

    return {
      ok: true,
      total: values.length,
      level: tally('q2', false),
      tools: tally('q3', true),
      lastWeek: tally('q4', true),
      pains: tally('q5', true),
      units: tally('q8', true),
      firstUnit: tally('q11', false),
      mode: tally('q9', true),
      price: tally('q10', false),
      intent: tally('q12', false),
      share: tally('q13', false),
      wishes: values.map(function (r) { return String(r[idx('q6')] || '').trim(); })
                    .filter(function (s) { return s; }),
      contacts: values.filter(function (r) { return String(r[idx('q14')] || '').trim(); }).length
    };
  } catch (err) {
    return { ok: false, error: String(err) };
  }
}

// ── 頁面 ──────────────────────────────────────────────

function doGet(e) {
  var p = (e && e.parameter && e.parameter.p) || '';

  if (p === 'json') {
    return ContentService.createTextOutput(JSON.stringify(getStats()))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var file = (p === 'stats') ? 'stats' : 'form';
  var title = (p === 'stats') ? '分會 AI 程度分布｜BNI 華都更' : 'BNI 華都更｜AI 課程課前問卷';

  return HtmlService.createHtmlOutputFromFile(file)
    .setTitle(title)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
