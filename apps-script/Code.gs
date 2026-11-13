/**
 * Прийом карток вивчення педагогічного досвіду (атестація).
 * Скрипт прив'язаний до Google Таблиці: кожна надіслана картка — рядок на аркуші «Відповіді»,
 * повна картка — окремий Google Документ у папці на Диску власника, власнику — лист-сповіщення.
 * Повторне надсилання тієї самої картки (той самий ID) оновлює рядок і документ, а не дублює їх.
 *
 * Розгортання: Розгорнути → Нове розгортання → Вебзастосунок;
 * «Виконувати від імені»: Я; «Хто має доступ»: Усі.
 */
const SHEET_NAME = 'Відповіді';
const FOLDER_NAME = 'Картки атестації — відповіді вчителів';
const NOTIFY = true;
const MAX_BYTES = 400000;

const COMP_CODES = ['А1', 'А2', 'А3', 'Б1', 'Б2', 'Б3', 'В1', 'В2', 'Г1', 'Г2', 'Г3', 'Д1'];
const HEADERS = ['Надіслано', 'ID', 'ПІБ', 'Посада', 'Предмет(и)', 'Email', 'Телефон', 'Освітній рівень', 'Пед. стаж, років',
  'Наявна категорія', 'Наявне звання', 'Вид атестації', 'Мета: категорія', 'Мета: звання',
  'Підвищення кваліфікації, год', 'з них ООП, год', 'Тема досвіду', 'Самооцінювання, середнє']
  .concat(COMP_CODES)
  .concat(['Досягнення учнів, записів', 'Методична діяльність, записів', 'Портфоліо', 'Картка (Google Документ)', 'Надсилань', 'Дані (JSON)']);

function doGet() {
  return json_({ ok: true, service: 'attestation-card' });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const raw = e && e.postData && e.postData.contents;
    if (!raw || raw.length > MAX_BYTES) return json_({ ok: false, error: 'empty_or_too_large' });
    const p = JSON.parse(raw);
    if (!p || !p.id || !p.data || !String(p.data.fio || '').trim()) return json_({ ok: false, error: 'bad_payload' });
    const id = String(p.id).slice(0, 64);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sh = ss.getSheetByName(SHEET_NAME);
    if (!sh) sh = ss.insertSheet(SHEET_NAME, 0);
    if (sh.getLastRow() === 0) {
      sh.appendRow(HEADERS);
      sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#e3eaf5').setWrap(true);
      sh.setFrozenRows(1);
    }

    const idCol = 2, docCol = HEADERS.indexOf('Картка (Google Документ)') + 1, cntCol = HEADERS.indexOf('Надсилань') + 1;
    let rowIdx = -1, docId = '', count = 0;
    const last = sh.getLastRow();
    if (last > 1) {
      const ids = sh.getRange(2, idCol, last - 1, 1).getValues();
      for (let i = 0; i < ids.length; i++) { if (String(ids[i][0]) === id) { rowIdx = i + 2; break; } }
      if (rowIdx > 0) {
        const link = String(sh.getRange(rowIdx, docCol).getValue() || '');
        const m = link.match(/\/d\/([a-zA-Z0-9_-]+)/);
        docId = m ? m[1] : '';
        count = Number(sh.getRange(rowIdx, cntCol).getValue()) || 0;
      }
    }

    const fio = String(p.data.fio).trim();
    const docUrl = writeDoc_(docId, 'Картка атестації — ' + fio, p.blocks || []);
    const s = p.summary || {};
    const comps = s.comps || {};
    const row = [new Date(), id, fio, p.data.position || '', p.data.subjects || '', p.data.email || '', p.data.phone || '',
      s.eduLevel || '', p.data.stazh || '', s.curCat || '', s.curTitle || '', s.attType || '', s.goalCat || '', s.goalTitle || '',
      s.pkHours || 0, s.pkOop || 0, p.data.expTopic || '', s.compAvg || '']
      .concat(COMP_CODES.map(function (c) { return comps[c] || ''; }))
      .concat([s.achCount || 0, s.methCount || 0, p.data.portfolio || '', docUrl, count + 1, JSON.stringify(p.data).slice(0, 49000)]);

    if (rowIdx > 0) sh.getRange(rowIdx, 1, 1, row.length).setValues([row]);
    else sh.appendRow(row);

    if (NOTIFY) {
      const me = Session.getEffectiveUser().getEmail();
      if (me) {
        MailApp.sendEmail({
          to: me,
          subject: (rowIdx > 0 ? 'Оновлено картку атестації: ' : 'Нова картка атестації: ') + fio,
          htmlBody: '<p><b>' + esc_(fio) + '</b>' + (p.data.subjects ? ', ' + esc_(p.data.subjects) : '') + '</p>' +
            '<p>Мета атестації: ' + esc_(s.goalCat || '—') + (s.goalTitle ? '; ' + esc_(s.goalTitle) : '') + '</p>' +
            '<p><a href="' + docUrl + '">Відкрити картку</a> · <a href="' + ss.getUrl() + '">Таблиця відповідей</a></p>'
        });
      }
    }
    return json_({ ok: true, updated: rowIdx > 0 });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  } finally {
    try { lock.releaseLock(); } catch (e2) { }
  }
}

function writeDoc_(docId, title, blocks) {
  let doc = null;
  if (docId) { try { doc = DocumentApp.openById(docId); } catch (e) { doc = null; } }
  if (!doc) {
    doc = DocumentApp.create(title);
    const file = DriveApp.getFileById(doc.getId());
    file.moveTo(folder_());
  } else {
    doc.setName(title);
  }
  const body = doc.getBody();
  body.clear();
  body.setMarginLeft(56).setMarginRight(42).setMarginTop(42).setMarginBottom(42);
  blocks.forEach(function (b) {
    if (!b || !b.t) return;
    if (b.t === 'title') {
      body.appendParagraph(String(b.text || '')).setAlignment(DocumentApp.HorizontalAlignment.CENTER).editAsText().setBold(true).setFontSize(13);
    } else if (b.t === 'note') {
      body.appendParagraph(String(b.text || '')).setAlignment(DocumentApp.HorizontalAlignment.CENTER).editAsText().setItalic(true).setFontSize(10);
    } else if (b.t === 'h') {
      body.appendParagraph(String(b.text || '')).setHeading(DocumentApp.ParagraphHeading.HEADING2);
    } else if (b.t === 'kv') {
      const rows = (b.rows || []).map(function (r) { return [String(r[0] == null ? '' : r[0]), String(r[1] == null || r[1] === '' ? '—' : r[1])]; });
      if (!rows.length) return;
      const t = body.appendTable(rows);
      for (let i = 0; i < t.getNumRows(); i++) { t.getRow(i).getCell(0).editAsText().setBold(true); t.getRow(i).getCell(0).setBackgroundColor('#eef2f7'); }
    } else if (b.t === 'table') {
      const rows = (b.rows || []).map(function (r) { return r.map(function (c) { return String(c == null ? '' : c); }); });
      if (!rows.length) { body.appendParagraph('—'); return; }
      const t = body.appendTable([b.head.map(String)].concat(rows));
      const hr = t.getRow(0);
      for (let j = 0; j < hr.getNumCells(); j++) { hr.getCell(j).editAsText().setBold(true); hr.getCell(j).setBackgroundColor('#e3eaf5'); }
    } else if (b.t === 'p') {
      const p = body.appendParagraph('');
      if (b.label) p.appendText(String(b.label) + ': ').setBold(true);
      p.appendText(String(b.text == null || b.text === '' ? '—' : b.text)).setBold(false);
    }
  });
  const url = doc.getUrl();
  doc.saveAndClose();
  return url;
}

function folder_() {
  const it = DriveApp.getFoldersByName(FOLDER_NAME);
  return it.hasNext() ? it.next() : DriveApp.createFolder(FOLDER_NAME);
}

function esc_(s) {
  return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; });
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/** Запустіть один раз вручну, щоб надати дозволи до розгортання. */
function setup() {
  folder_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName(SHEET_NAME)) ss.insertSheet(SHEET_NAME, 0);
  const me = Session.getEffectiveUser().getEmail();
  Logger.log('Готово. Сповіщення надходитимуть на: ' + me);
  if (false) MailApp.sendEmail(me, 'test', 'test');
}
