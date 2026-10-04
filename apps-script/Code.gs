/**
 * ימי הולדת במשפחה – צד השרת (Google Apps Script)
 *
 * הגיליון הוא מסד הנתונים, והוא משרת כמה משפחות במקביל:
 *   • לשונית "משפחות" – שורה לכל משפחה: שם, מפתח סודי (בקישור), קוד עריכה, מייל לתזכורת.
 *   • לשונית נפרדת לכל משפחה – בני המשפחה שלה. משפחה אחת לא רואה את השנייה.
 *
 * הסקריפט:
 *   1. משמש API לדף (doPost) – המפתח שבקישור קובע לאיזו משפחה הבקשה שייכת.
 *   2. שולח ב-1 לכל חודש מייל לכל משפחה עם ימי ההולדת וכפתור לוואטסאפ.
 *
 * הוספת משפחה – בלי לשנות קוד ובלי פריסה מחדש:
 *   בגיליון: תפריט 🎂 ימי הולדת ← הוספת משפחה חדשה.
 *   (או: לכתוב שם בשורה חדשה בלשונית "משפחות" – השאר יתמלא לבד.)
 *
 * Script Properties:
 *   SITE_URL    – כתובת הדף (לקישורים).
 *   OWNER_EMAIL – לא חובה; לאן לשלוח תזכורת למשפחה בלי מייל משלה (ברירת מחדל: אתה).
 */

const FAMILIES_SHEET = 'משפחות';
const FAMILY_HEADERS = ['שם המשפחה', 'מפתח (בקישור)', 'קוד עריכה', 'מייל לתזכורת', 'לשונית', 'פעיל', 'נוצר'];
const FIRST_FAMILY_TAB = 'בני משפחה'; // הלשונית של המשפחה הראשונה (מלפני שהיו כמה משפחות)
const HEADERS = ['id', 'שם פרטי', 'שם משפחה', 'יום', 'חודש', 'שנה', 'נמחק', 'עודכן'];
const MONTHS = ['ינואר', 'פברואר', 'מרץ', 'אפריל', 'מאי', 'יוני',
  'יולי', 'אוגוסט', 'ספטמבר', 'אוקטובר', 'נובמבר', 'דצמבר'];
const MAX_EDIT_FAILS = 10;      // ניסיונות קוד שגוי לפני חסימה זמנית
const LOCK_SECONDS = 10 * 60;   // משך החסימה

/* ───────────── הפעלה ראשונה ───────────── */

/** להריץ פעם אחת מהעורך: יוצר את לשונית המשפחות, משפחה ראשונה וטריגר חודשי. */
function setup() {
  const fams = families_();
  if (!fams.length) addFamily('המשפחה שלי', '');

  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'monthlyReminder')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('monthlyReminder').timeBased().onMonthDay(1).atHour(8).create();

  showSettings();
}

/** מדפיס ללוג את הקישור וקוד העריכה של כל משפחה. */
function showSettings() {
  const site = PropertiesService.getScriptProperties().getProperty('SITE_URL');
  if (!site) Logger.log('⚠️ חסר SITE_URL ב-Script Properties – בלעדיו הקישורים לא שלמים.');
  families_().forEach(f => {
    Logger.log(f.name + (f.active ? '' : ' (מושבתת)') + '\n  קישור: ' + familyLink_(f) + '\n  קוד עריכה: ' + f.code);
  });
}

/** בדיקה: שולח עכשיו את מיילי התזכורת של החודש הנוכחי (לכל המשפחות). */
function testMonthlyEmail() {
  monthlyReminder();
}

/* ───────────── משפחות ───────────── */

/** יוצר משפחה חדשה ומחזיר אותה (עם הקישור והקוד). */
function addFamily(name, email) {
  name = String(name || '').trim().slice(0, 40);
  if (!name) throw new Error('חסר שם משפחה');
  return withLock_(() => {
    const reg = registry_();
    const f = {
      name: name,
      key: newKey_(),
      code: newCode_(),
      email: String(email || '').trim(),
      tab: uniqueTab_(name),
      active: true,
    };
    reg.appendRow([f.name, f.key, f.code, f.email, f.tab, true, new Date()]);
    sheet_(f.tab);
    return f;
  });
}

/** כל המשפחות. משלים לבד שורות שמישהו כתב בהן רק שם. */
function families_() {
  const reg = registry_();
  const last = reg.getLastRow();
  if (last < 2) return [];
  const rows = reg.getRange(2, 1, last - 1, FAMILY_HEADERS.length).getValues();
  const out = [];
  rows.forEach((r, i) => {
    const name = String(r[0]).trim();
    if (!name) return;
    if (!String(r[1]).trim() || !String(r[2]).trim() || !String(r[4]).trim()) { // שורה חלקית – משלימים
      r[1] = String(r[1]).trim() || newKey_();
      r[2] = String(r[2]).trim() || newCode_();
      r[4] = String(r[4]).trim() || uniqueTab_(name);
      if (r[5] === '') r[5] = true;
      if (!r[6]) r[6] = new Date();
      reg.getRange(i + 2, 1, 1, FAMILY_HEADERS.length).setValues([r]);
      sheet_(r[4]);
    }
    out.push({
      name: name,
      key: String(r[1]).trim(),
      code: String(r[2]).trim(),
      email: String(r[3]).trim(),
      tab: String(r[4]).trim(),
      active: r[5] !== false && String(r[5]).toUpperCase() !== 'FALSE',
    });
  });
  return out;
}

function familyByKey_(key) {
  if (!key) return null;
  return families_().find(f => f.active && f.key === String(key)) || null;
}

/**
 * לשונית המשפחות. בפעם הראשונה – יוצרת אותה, ואם המערכת כבר עבדה
 * עם משפחה אחת (VIEW_KEY / EDIT_CODE ישנים) – מעבירה אותה לשם, כך שהקישור הקיים ממשיך לעבוד.
 */
function registry_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let reg = ss.getSheetByName(FAMILIES_SHEET);
  if (reg) return reg;
  reg = ss.insertSheet(FAMILIES_SHEET, 0);
  reg.getRange(1, 1, 1, FAMILY_HEADERS.length).setValues([FAMILY_HEADERS]).setFontWeight('bold');
  reg.setFrozenRows(1);
  reg.setRightToLeft(true);
  reg.getRange('B:C').setNumberFormat('@');
  const props = PropertiesService.getScriptProperties();
  const oldKey = props.getProperty('VIEW_KEY');
  if (oldKey) {
    reg.appendRow(['המשפחה שלי', oldKey, props.getProperty('EDIT_CODE') || newCode_(),
      props.getProperty('OWNER_EMAIL') || '', FIRST_FAMILY_TAB, true, new Date()]);
  }
  return reg;
}

function uniqueTab_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const base = name.replace(/[\[\]*?:\/\\']/g, '').trim().slice(0, 60) || 'משפחה';
  let tab = base, n = 2;
  while (ss.getSheetByName(tab) || tab === FAMILIES_SHEET) tab = base + ' ' + n++;
  return tab;
}

function newKey_() { return Utilities.getUuid().replace(/-/g, '').slice(0, 24); }
function newCode_() { return String(Math.floor(1000 + Math.random() * 9000)); }

/* ───────────── תפריט בגיליון ───────────── */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('🎂 ימי הולדת')
    .addItem('הוספת משפחה חדשה', 'menuAddFamily')
    .addItem('הצגת הקישורים והקודים', 'menuShowLinks')
    .addItem('ניקוי כפילויות', 'removeDuplicates')
    .addToUi();
}

function menuAddFamily() {
  const ui = SpreadsheetApp.getUi();
  const n = ui.prompt('משפחה חדשה', 'שם המשפחה (למשל: משפחת לוי)', ui.ButtonSet.OK_CANCEL);
  if (n.getSelectedButton() !== ui.Button.OK || !n.getResponseText().trim()) return;
  const m = ui.prompt('מייל לתזכורת החודשית',
    'למי לשלוח ב-1 לחודש את רשימת ימי ההולדת? (אפשר להשאיר ריק – יישלח אליך)', ui.ButtonSet.OK_CANCEL);
  if (m.getSelectedButton() !== ui.Button.OK) return;
  const f = addFamily(n.getResponseText(), m.getResponseText());
  ui.alert('✓ ' + f.name + ' נוספה',
    'קישור לשלוח לקבוצה של המשפחה:\n' + familyLink_(f) + '\n\nקוד עריכה: ' + f.code, ui.ButtonSet.OK);
}

function menuShowLinks() {
  const text = families_().map(f =>
    f.name + (f.active ? '' : ' (מושבתת)') + '\n' + familyLink_(f) + '\nקוד עריכה: ' + f.code).join('\n\n');
  SpreadsheetApp.getUi().alert('קישורים וקודים', text || 'אין עדיין משפחות', SpreadsheetApp.getUi().ButtonSet.OK);
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
  const fam = familyByKey_(req.key);
  if (!fam) throw fail_('no_access');

  switch (req.action) {
    case 'list':
      return { ok: true, family: fam.name, people: getPeople_(fam) };
    case 'checkCode':
      checkEditCode_(fam, req.code);
      return { ok: true };
    case 'add':
    case 'update':
    case 'delete':
      checkEditCode_(fam, req.code);
      return withLock_(() => {
        if (req.action === 'add') {
          const p = clean_(req.person);
          assertNotDuplicate_(fam, p, null);
          addPerson_(fam, p);
        }
        if (req.action === 'update') {
          const p = clean_(req.person);
          assertNotDuplicate_(fam, p, req.person.id);
          updatePerson_(fam, req.person.id, p);
        }
        if (req.action === 'delete') deletePerson_(fam, req.id);
        return { ok: true, family: fam.name, people: getPeople_(fam) };
      });
    default:
      throw fail_('bad_request');
  }
}

function checkEditCode_(fam, code) {
  const cache = CacheService.getScriptCache();
  const failsKey = 'editFails:' + fam.key;
  const fails = Number(cache.get(failsKey) || 0);
  if (fails >= MAX_EDIT_FAILS) throw fail_('locked');
  if (String(code || '').trim() !== fam.code) {
    cache.put(failsKey, String(fails + 1), LOCK_SECONDS);
    throw fail_('bad_code');
  }
}

/* ───────────── הגיליון ───────────── */

/** הלשונית של משפחה (נוצרת אם חסרה). */
function sheet_(tab) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(tab);
  if (!sh) {
    sh = ss.insertSheet(tab);
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.setRightToLeft(true);
    sh.getRange('B:C').setNumberFormat('@');
  }
  return sh;
}

function getPeople_(fam) {
  const sh = sheet_(fam.tab);
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

function addPerson_(fam, p) {
  sheet_(fam.tab).appendRow([Utilities.getUuid(), p.firstName, p.lastName, p.day, p.month, p.year || '', false, new Date()]);
}

function updatePerson_(fam, id, p) {
  const row = findRow_(fam, id);
  sheet_(fam.tab).getRange(row, 2, 1, 7)
    .setValues([[p.firstName, p.lastName, p.day, p.month, p.year || '', false, new Date()]]);
}

/** חוסם רשומה עם אותו שם ואותו תאריך (למשל שני בני משפחה שהוסיפו את אותו אדם). */
function assertNotDuplicate_(fam, p, exceptId) {
  const twin = getPeople_(fam).find(x => x.id !== exceptId && sameKey_(x) === sameKey_(p));
  if (twin) throw fail_('duplicate', fullName_(twin) + ' כבר ברשימה עם אותו תאריך.');
}

function sameKey_(p) {
  const n = s => String(s || '').trim().replace(/\s+/g, ' ').toLowerCase();
  return [n(p.firstName), n(p.lastName), p.day, p.month].join('|');
}

/**
 * ניקוי חד-פעמי: מסמן כנמחקות כפילויות (אותו שם + אותו יום וחודש).
 * נשארת הרשומה הראשונה; אם רק לאחת יש שנת לידה – היא זו שנשארת.
 * להריץ מהעורך. הכול הפיך: בגיליון אפשר להחזיר "נמחק" ל-FALSE.
 */
function removeDuplicates() {
  families_().forEach(removeFamilyDuplicates_);
}

function removeFamilyDuplicates_(fam) {
  withLock_(() => {
    const sh = sheet_(fam.tab);
    const people = getPeople_(fam);
    const keep = {};
    people.forEach(p => {
      const k = sameKey_(p);
      if (!keep[k] || (!keep[k].year && p.year)) keep[k] = p;
    });
    const keepIds = new Set(Object.keys(keep).map(k => keep[k].id));
    const ids = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 1).getValues().map(r => String(r[0]));
    let removed = 0;
    people.forEach(p => {
      if (keepIds.has(p.id)) return;
      sh.getRange(ids.indexOf(p.id) + 2, 7, 1, 2).setValues([[true, new Date()]]);
      removed++;
      Logger.log(fam.name + ' – הוסרה כפילות: ' + fullName_(p) + ' ' + p.day + '.' + p.month);
    });
    Logger.log(fam.name + ' – סה"כ הוסרו ' + removed + ' כפילויות.');
  });
}

/** מחיקה "רכה": מסמן נמחק=TRUE. לשחזור – לשנות ל-FALSE בגיליון. */
function deletePerson_(fam, id) {
  const row = findRow_(fam, id);
  sheet_(fam.tab).getRange(row, 7, 1, 2).setValues([[true, new Date()]]);
}

function findRow_(fam, id) {
  const sh = sheet_(fam.tab);
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
  families_().filter(f => f.active).forEach(f => {
    try {
      sendMonthlyEmail_(f);
    } catch (err) {
      console.error('שליחה נכשלה עבור ' + f.name + ': ' + err); // משפחה אחת לא עוצרת את השאר
    }
  });
}

function sendMonthlyEmail_(fam) {
  const tz = Session.getScriptTimeZone();
  const now = new Date();
  const year = Number(Utilities.formatDate(now, tz, 'yyyy'));
  const month = Number(Utilities.formatDate(now, tz, 'M'));
  const list = birthdaysInMonth_(getPeople_(fam), year, month);
  const monthName = MONTHS[month - 1];
  const site = familyLink_(fam);

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
      '<h2 style="margin:0 0 8px">🎂 ימי הולדת ב' + monthName + ' – ' + esc_(fam.name) + '</h2>' +
      '<ul style="padding-right:20px;margin:0 0 18px">' + rows + '</ul>' +
      '<a href="' + wa + '" style="display:inline-block;background:#25D366;color:#fff;text-decoration:none;' +
      'padding:12px 22px;border-radius:999px;font-weight:bold">שליחה לקבוצה בוואטסאפ</a>' +
      '<p style="color:#666;font-size:13px">הכפתור פותח את וואטסאפ עם ההודעה מוכנה. נשאר רק לבחור את הקבוצה וללחוץ שליחה.</p>' +
      (site ? '<p><a href="' + site + '">לדף ימי ההולדת</a></p>' : '') +
      '</div>';
    text = message + '\n\nלשליחה בוואטסאפ: ' + wa;
  }

  const to = fam.email || ownerEmail_();
  MailApp.sendEmail({
    to: to,
    subject: '🎂 ימי הולדת ב' + monthName + ' – ' + fam.name + (list.length ? ' (' + list.length + ')' : ''),
    body: text,
    htmlBody: html,
  });
  Logger.log('נשלח מייל ל' + fam.name + ' אל ' + to);
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

function familyLink_(fam) {
  const site = PropertiesService.getScriptProperties().getProperty('SITE_URL') || '<כתובת הדף>';
  return site.replace(/#.*$/, '') + '#k=' + fam.key;
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
