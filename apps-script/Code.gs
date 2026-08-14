/**
 * The Peach Lab 🍑 — приймач даних трекера тренувань.
 *
 * Прив'язується до таблиці (Розширення → Apps Script) і при кожній синхронізації
 * перезаписує вкладки: Журнал, Прогрес вправ, Заміри, Тиждень 1…12 і BACKUP.
 * Трекер надсилає вже готові рядки, тому програма тренувань живе лише в трекері.
 *
 * doPost — прийняти дані з трекера; doGet — віддати копію для відновлення.
 */

var BACKUP_TAB = 'BACKUP (не редагувати)';
var CHUNK = 40000;           // ліміт клітинки Google Sheets — 50 000 символів
var HEADER_BG = '#F4D9C6';
var HEADER_FG = '#5A2E17';

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (err) {
    return out({ ok: false, error: 'Таблиця зайнята іншим записом, спробуй ще раз' });
  }
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return out({ ok: false, error: 'Порожній запит' });
    }
    var p = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActive();
    var names = [];

    if (p.journal)  names.push(writeTable(ss, 'Журнал', p.journal.header, p.journal.rows));
    if (p.progress) names.push(writeTable(ss, 'Прогрес вправ', p.progress.header, p.progress.rows));
    if (p.meas)     names.push(writeTable(ss, 'Заміри', p.meas.header, p.meas.rows));
    if (p.weeks && p.weeks.length) {
      for (var i = 0; i < p.weeks.length; i++) {
        var w = p.weeks[i];
        names.push(writeTable(ss, w.name, w.header, w.rows));
      }
    }
    if (p.backup) writeBackup(ss, p.backup);
    reorderTabs(ss, names);

    return out({ ok: true, at: nowStr(), tabs: names });
  } catch (err) {
    return out({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (ignored) {}
  }
}

function doGet() {
  try {
    var sh = SpreadsheetApp.getActive().getSheetByName(BACKUP_TAB);
    if (!sh || sh.getLastRow() < 2) {
      return out({ ok: false, error: 'У таблиці ще немає збереженої копії' });
    }
    var vals = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
    var s = '';
    for (var i = 0; i < vals.length; i++) s += (vals[i][0] || '');
    if (!s) return out({ ok: false, error: 'Копія порожня' });
    return out({ ok: true, data: s });
  } catch (err) {
    return out({ ok: false, error: String(err) });
  }
}

/** Перезаписує вкладку таблицею значень і форматує шапку. Повертає назву вкладки. */
function writeTable(ss, name, header, rows) {
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  sh.clear();
  header = header || [];
  rows = rows || [];
  var width = header.length;
  for (var i = 0; i < rows.length; i++) width = Math.max(width, rows[i].length);
  if (width === 0) { width = 1; header = ['—']; }

  var data = [pad(header, width)];
  for (var j = 0; j < rows.length; j++) data.push(pad(rows[j], width));

  sh.getRange(1, 1, data.length, width).setValues(data);
  var head = sh.getRange(1, 1, 1, width);
  head.setFontWeight('bold').setBackground(HEADER_BG).setFontColor(HEADER_FG);
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, width);
  if (rows.length) {
    sh.getRange(2, 1, rows.length, width).setVerticalAlignment('middle');
  }
  return name;
}

/** Кладе повну JSON-копію в одну вкладку, розбиту на шматки по рядках. */
function writeBackup(ss, json) {
  var sh = ss.getSheetByName(BACKUP_TAB) || ss.insertSheet(BACKUP_TAB);
  sh.clear();
  var parts = [];
  for (var i = 0; i < json.length; i += CHUNK) parts.push([json.substr(i, CHUNK)]);
  if (!parts.length) parts = [['']];
  sh.getRange(1, 1, 1, 1).setValues([['Копія даних трекера від ' + nowStr() + ' — не редагувати вручну']]);
  sh.getRange(1, 1).setFontWeight('bold').setBackground(HEADER_BG).setFontColor(HEADER_FG);
  sh.getRange(2, 1, parts.length, 1).setValues(parts);
  sh.setColumnWidth(1, 420);
  sh.hideSheet();
}

/** Ставить вкладки в зрозумілому порядку: зведення → тижні → копія. */
function reorderTabs(ss, names) {
  for (var i = 0; i < names.length; i++) {
    var sh = ss.getSheetByName(names[i]);
    if (!sh) continue;
    ss.setActiveSheet(sh);
    ss.moveActiveSheet(i + 1);
  }
  var first = ss.getSheetByName(names[0]);
  if (first) ss.setActiveSheet(first);
}

function pad(row, width) {
  var r = (row || []).slice();
  while (r.length < width) r.push('');
  return r;
}

function nowStr() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd.MM.yyyy HH:mm');
}

function out(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
