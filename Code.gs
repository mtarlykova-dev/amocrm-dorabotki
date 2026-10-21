/**
 * Хранилище отметок и комментариев для лендинга «Доработка amoCRM».
 * Лист «status»:   id | status | updated | title
 * Лист «comments»: id | author | text | created | title
 * Лист «deadlines»: id | deadline | updated | title
 *
 * Установка:
 * 1. Создайте Google-таблицу → Расширения → Apps Script.
 * 2. Вставьте этот код вместо содержимого Code.gs, сохраните.
 * 3. Развернуть → Новое развёртывание → тип «Веб-приложение».
 *    Выполнять от имени: «Я». Доступ: «Все» (Anyone).
 * 4. Скопируйте URL вида https://script.google.com/macros/s/.../exec
 *    и вставьте его в STATUS_URL в коде лендинга.
 *
 * Если старая версия скрипта уже развёрнута: вставьте новый код, затем
 * Развернуть → Управление развёртываниями → карандаш → Версия: «Новая версия» → Развернуть.
 * Так URL останется прежним, менять его на лендинге не нужно.
 *
 * Если отметки «слетали» на старой версии скрипта: удалите листы «status» и «comments»
 * целиком — скрипт создаст их заново с правильным форматом.
 *
 * Удалить комментарий: удалите его строку на листе «comments».
 */

var ALLOWED = ['todo', 'wip', 'done'];

function sheet_(name, header) {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  if (sh.getLastRow() === 0) sh.appendRow(header);
  /* Колонка id — строго текст. Иначе таблица превращает «1-1» в дату 1 января,
     и отметка перестаёт находиться — после обновления страница её «теряет». */
  sh.getRange('A:A').setNumberFormat('@');
  /* На листе дедлайнов дату тоже храним текстом «ГГГГ-ММ-ДД», чтобы не сбивался часовой пояс. */
  if (name === 'deadlines') sh.getRange('B:B').setNumberFormat('@');
  return sh;
}
function statusSheet_() { return sheet_('status', ['id', 'status', 'updated', 'title']); }
function deadlineSheet_() { return sheet_('deadlines', ['id', 'deadline', 'updated', 'title']); }
function commentSheet_() { return sheet_('comments', ['id', 'author', 'text', 'created', 'title']); }

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
function iso_(v) { return v instanceof Date ? v.toISOString() : String(v); }
/* Текст, начинающийся с = + - @, таблица считает формулой — экранируем. */
function safe_(v, max) {
  v = String(v || '').slice(0, max);
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}

/* Обновить строку с этим id или добавить новую. */
function upsert_(sh, id, row) {
  var ids = sh.getRange(1, 1, sh.getLastRow(), 1).getDisplayValues().map(function (r) { return r[0]; });
  var idx = ids.indexOf(id);
  if (idx > -1) sh.getRange(idx + 1, 1, 1, row.length).setValues([row]);
  else sh.appendRow(row);
}

function doGet() {
  var status = {}, comments = {};
  var sh = statusSheet_();
  var sv = sh.getDataRange().getValues();
  var sid = sh.getDataRange().getDisplayValues();
  for (var i = 1; i < sv.length; i++) {
    if (!sid[i][0]) continue;
    status[sid[i][0]] = { s: sv[i][1], ts: iso_(sv[i][2]) };
  }
  var ch = commentSheet_();
  var cv = ch.getDataRange().getValues();
  var cid = ch.getDataRange().getDisplayValues();
  for (var j = 1; j < cv.length; j++) {
    var id = cid[j][0];
    if (!id || !cv[j][2]) continue;
    (comments[id] = comments[id] || []).push({ a: String(cv[j][1]), x: String(cv[j][2]), ts: iso_(cv[j][3]) });
  }
  var deadlines = {};
  var dh = deadlineSheet_();
  var dv = dh.getDataRange().getValues();
  var dd = dh.getDataRange().getDisplayValues();
  for (var k = 1; k < dd.length; k++) {
    if (!dd[k][0] || !dd[k][1]) continue;
    deadlines[dd[k][0]] = { d: dd[k][1], ts: iso_(dv[k][2]) };
  }
  return json_({ v: 3, status: status, comments: comments, deadlines: deadlines });
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var d = JSON.parse(e.postData.contents);
    var id = String(d.id || '').slice(0, 40);
    if (!id) return json_({ ok: false, error: 'no id' });
    var now = new Date();

    if (d.type === 'comment') {
      var text = String(d.x || '').trim();
      if (!text) return json_({ ok: false, error: 'empty' });
      commentSheet_().appendRow([id, safe_(d.a, 60), safe_(text, 2000), now, safe_(d.title, 200)]);
      return json_({ ok: true, ts: now.toISOString() });
    }

    if (d.type === 'deadline') {
      var dl = String(d.d || '');
      if (dl && !/^\d{4}-\d{2}-\d{2}$/.test(dl)) return json_({ ok: false, error: 'bad date' });
      upsert_(deadlineSheet_(), id, [id, dl, now, safe_(d.title, 200)]);
      return json_({ ok: true, ts: now.toISOString() });
    }

    var s = String(d.s || '');
    if (ALLOWED.indexOf(s) < 0) return json_({ ok: false, error: 'bad status' });
    upsert_(statusSheet_(), id, [id, s, now, safe_(d.title, 200)]);
    return json_({ ok: true, ts: now.toISOString() });
  } finally {
    lock.releaseLock();
  }
}
