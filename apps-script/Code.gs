/**
 * ימי הולדת במשפחה – צד השרת (Google Apps Script)
 *
 * הגיליון הוא מסד הנתונים. הסקריפט הזה:
 *   1. משמש API לדף (doPost) – קריאה, הוספה, עריכה, מחיקה.
 *   2. שולח לך מייל ב-1 לכל חודש עם ימי ההולדת וכפתור לוואטסאפ.
 *
 * הגדרות נשמרות ב-Script Properties (לא בקוד):
 *   VIEW_KEY   – המפתח הסודי שבקישור המשפחתי (נוצר אוטומטית ב-setup).
 *   EDIT_CODE  – קוד העריכה הקצר (נוצר אוטומטית ב-setup, אפשר לשנות).
 *   SITE_URL   – כתובת הדף שלך (לא חובה; משמש לקישור במייל).
 *   OWNER_EMAIL – לאן לשלוח את התזכורת (לא חובה; ברירת מחדל: אתה).
 */

const SHEET_NAME = 'בני משפחה';
const HEADERS = ['id', 'שם פרטי', 'שם משפחה', 'יום', 'חודש', 'שנה', 'נמחק', 'עודכן'];
const MONTHS = ['ינואר', 'פברואר', 'מרץ', 'אפריל', 'מאי', 'יוני',
  'יולי', 'אוגוסט', 'ספטמבר', 'אוקטובר', 'נובמבר', 'דצמבר'];
const MAX_EDIT_FAILS = 10;      // ניסיונות קוד שגוי לפני חסימה זמנית
const LOCK_SECONDS = 10 * 60;   // משך החסימה

/* ───────────── הפעלה ראשונה ───────────── */

/** להריץ פעם אחת מהעורך: יוצר גיליון, מפתחות וטריגר חודשי. */
function setup() {
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('VIEW_KEY')) {
    props.setProperty('VIEW_KEY', Utilities.getUuid().replace(/-/g, '').slice(0, 24));
  }
  if (!props.getProperty('EDIT_CODE')) {
    props.setProperty('EDIT_CODE', String(Math.floor(1000 + Math.random() * 9000)));
  }
  sheet_();

  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'monthlyReminder')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('monthlyReminder').timeBased().onMonthDay(1).atHour(8).create();

  showSettings();
}

/** מדפיס ללוג את המפתח, קוד העריכה והקישור המשפחתי. */
function showSettings() {
  const props = PropertiesService.getScriptProperties();
  const key = props.getProperty('VIEW_KEY');
  const site = props.getProperty('SITE_URL');
  Logger.log('מפתח צפייה (VIEW_KEY): ' + key);
  Logger.log('קוד עריכה (EDIT_CODE): ' + props.getProperty('EDIT_CODE'));
  Logger.log(site
    ? 'הקישור לשלוח למשפחה: ' + familyLink_()
    : 'הקישור למשפחה: <כתובת הדף שלך>#k=' + key + '  (הגדר SITE_URL כדי לקבל אותו מוכן)');
}

/** בדיקה: שולח עכשיו את מייל התזכורת של החודש הנוכחי. */
function testMonthlyEmail() {
  monthlyReminder();
  Logger.log('נשלח מייל אל ' + ownerEmail_());
}

/* ───────────── API ───────────── */

function doGet() {
  return ContentService.createTextOutput('ימי הולדת במשפחה – השרת פעיל ✓');
}

function doPost(e) {
  let req;
  try {
    req = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, error: 'bad_request' });
  }
  try {
    return json_(handle_(req));
  } catch (err) {
    if (err.code) return json_({ ok: false, error: err.code, message: err.message });
    console.error(err);
    return json_({ ok: false, error: 'server' });
  }
}

function handle_(req) {
  const props = PropertiesService.getScriptProperties();
  if (!req.key || req.key !== props.getProperty('VIEW_KEY')) throw fail_('no_access');

  switch (req.action) {
    case 'list':
      return { ok: true, people: getPeople_() };
    case 'checkCode':
      checkEditCode_(req.code);
      return { ok: true };
    case 'add':
    case 'update':
    case 'delete':
      checkEditCode_(req.code);
      return withLock_(() => {
        if (req.action === 'add') addPerson_(clean_(req.person));
        if (req.action === 'update') updatePerson_(req.person && req.person.id, clean_(req.person));
        if (req.action === 'delete') deletePerson_(req.id);
        return { ok: true, people: getPeople_() };
      });
    default:
      throw fail_('bad_request');
  }
}

function checkEditCode_(code) {
  const cache = CacheService.getScriptCache();
  const fails = Number(cache.get('editFails') || 0);
  if (fails >= MAX_EDIT_FAILS) throw fail_('locked');
  const real = PropertiesService.getScriptProperties().getProperty('EDIT_CODE');
  if (String(code || '').trim() !== String(real)) {
    cache.put('editFails', String(fails + 1), LOCK_SECONDS);
    throw fail_('bad_code');
  }
}

/* ───────────── הגיליון ───────────── */

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.setRightToLeft(true);
    sh.getRange('B:C').setNumberFormat('@');
  }
  return sh;
}

function getPeople_() {
  const sh = sheet_();
  const last = sh.getLastRow();
  if (last < 2) return [];
  const rows = sh.getRange(2, 1, last - 1, HEADERS.length).getValues();
  const people = [];
  rows.forEach((r, i) => {
    const first = String(r[1]).trim();
    const day = Number(r[3]);
    const month = Number(r[4]);
    if (!first || !day || !month || r[6] === true) return;
    let id = String(r[0]).trim();
    if (!id) { // שורה שהוקלדה ידנית בגיליון – נותנים לה מזהה
      id = Utilities.getUuid();
      sh.getRange(i + 2, 1).setValue(id);
    }
    people.push({
      id: id,
      firstName: first,
      lastName: String(r[2]).trim(),
      day: day,
      month: month,
      year: Number(r[5]) || null,
    });
  });
  return people;
}

function addPerson_(p) {
  sheet_().appendRow([Utilities.getUuid(), p.firstName, p.lastName, p.day, p.month, p.year || '', false, new Date()]);
}

function updatePerson_(id, p) {
  const row = findRow_(id);
  sheet_().getRange(row, 2, 1, 7)
    .setValues([[p.firstName, p.lastName, p.day, p.month, p.year || '', false, new Date()]]);
}

/** מחיקה "רכה": מסמן נמחק=TRUE. לשחזור – לשנות ל-FALSE בגיליון. */
function deletePerson_(id) {
  const row = findRow_(id);
  sheet_().getRange(row, 7, 1, 2).setValues([[true, new Date()]]);
}

function findRow_(id) {
  const sh = sheet_();
  const last = sh.getLastRow();
  if (id && last >= 2) {
    const ids = sh.getRange(2, 1, last - 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === String(id)) return i + 2;
    }
  }
  throw fail_('not_found');
}

/** בדיקת תקינות וניקוי של רשומה שמגיעה מהדף. */
function clean_(p) {
  p = p || {};
  const name = s => String(s || '').trim().replace(/^[=+\-@]+/, '').slice(0, 40);
  const firstName = name(p.firstName);
  const lastName = name(p.lastName);
  const day = parseInt(p.day, 10);
  const month = parseInt(p.month, 10);
  const year = (p.year === '' || p.year == null) ? null : parseInt(p.year, 10);
  const thisYear = Number(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy'));

  if (!firstName) throw fail_('invalid', 'חסר שם פרטי');
  if (!(month >= 1 && month <= 12)) throw fail_('invalid', 'חודש לא תקין');
  if (!(day >= 1 && day <= daysInMonth_(2000, month))) throw fail_('invalid', 'יום לא תקין');
  if (year !== null) {
    if (isNaN(year) || year < 1900 || year > thisYear) throw fail_('invalid', 'שנת לידה לא תקינה');
    if (day > daysInMonth_(year, month)) throw fail_('invalid', 'התאריך הזה לא קיים בשנה ' + year);
  }
  return { firstName, lastName, day, month, year };
}

/* ───────────── תזכורת חודשית ───────────── */

function monthlyReminder() {
  const tz = Session.getScriptTimeZone();
  const now = new Date();
  const year = Number(Utilities.formatDate(now, tz, 'yyyy'));
  const month = Number(Utilities.formatDate(now, tz, 'M'));
  const list = birthdaysInMonth_(getPeople_(), year, month);
  const monthName = MONTHS[month - 1];
  const site = familyLink_();

  let html, text;
  if (list.length === 0) {
    text = 'אין ימי הולדת ב' + monthName + '.';
    html = '<div dir="rtl" style="font-family:Arial,sans-serif;font-size:16px">' + text +
      (site ? '<p><a href="' + site + '">לדף ימי ההולדת</a></p>' : '') + '</div>';
  } else {
    const message = buildMessage_(list, month);
    const wa = 'https://wa.me/?text=' + encodeURIComponent(message);
    const rows = list.map(x =>
      '<li style="margin:4px 0"><b>' + x.day + '.' + month + '</b> – ' + esc_(fullName_(x.p)) +
      (x.age ? ' (גיל ' + x.age + ')' : '') + '</li>').join('');
    html =
      '<div dir="rtl" style="font-family:Arial,sans-serif;font-size:16px;line-height:1.5;color:#222">' +
      '<h2 style="margin:0 0 8px">🎂 ימי הולדת ב' + monthName + '</h2>' +
      '<ul style="padding-right:20px;margin:0 0 18px">' + rows + '</ul>' +
      '<a href="' + wa + '" style="display:inline-block;background:#25D366;color:#fff;text-decoration:none;' +
      'padding:12px 22px;border-radius:999px;font-weight:bold">שליחה לקבוצה בוואטסאפ</a>' +
      '<p style="color:#666;font-size:13px">הכפתור פותח את וואטסאפ עם ההודעה מוכנה. נשאר רק לבחור את הקבוצה וללחוץ שליחה.</p>' +
      (site ? '<p><a href="' + site + '">לדף ימי ההולדת</a></p>' : '') +
      '</div>';
    text = message + '\n\nלשליחה בוואטסאפ: ' + wa;
  }

  MailApp.sendEmail({
    to: ownerEmail_(),
    subject: '🎂 ימי הולדת ב' + monthName + (list.length ? ' (' + list.length + ')' : ''),
    body: text,
    htmlBody: html,
  });
}

/* ───────────── עזרים (זהים ללוגיקה בדף) ───────────── */

function birthdaysInMonth_(people, year, month) {
  return people
    .filter(p => p.month === month)
    .map(p => ({
      p: p,
      day: Math.min(p.day, daysInMonth_(year, month)), // 29.2 בשנה רגילה → 28.2
      age: p.year ? year - p.year : null,
    }))
    .filter(x => x.age === null || x.age > 0)
    .sort((a, b) => a.day - b.day || fullName_(a.p).localeCompare(fullName_(b.p), 'he'));
}

function buildMessage_(list, month) {
  const RLM = '‏';
  const lines = list.map(x =>
    RLM + x.day + '.' + month + ' – ' + fullName_(x.p) + (x.age ? ' (גיל ' + x.age + ')' : ''));
  return '🎂 ימי הולדת ב' + MONTHS[month - 1] + '\n\n' + lines.join('\n') + '\n\nמזל טוב לכולם! 🎉';
}

function fullName_(p) { return (p.firstName + ' ' + (p.lastName || '')).trim(); }
function daysInMonth_(y, m) { return new Date(y, m, 0).getDate(); }

function familyLink_() {
  const props = PropertiesService.getScriptProperties();
  const site = props.getProperty('SITE_URL');
  return site ? site.replace(/#.*$/, '') + '#k=' + props.getProperty('VIEW_KEY') : '';
}

function ownerEmail_() {
  return PropertiesService.getScriptProperties().getProperty('OWNER_EMAIL') ||
    Session.getEffectiveUser().getEmail();
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) throw fail_('busy');
  try { return fn(); } finally { lock.releaseLock(); }
}

function fail_(code, message) {
  const e = new Error(message || code);
  e.code = code;
  return e;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function esc_(s) {
  return String(s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}
