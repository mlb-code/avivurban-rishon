/**
 * דיירי תמורה — ראשל״צ (מוהליבר + ויצברד + החפץ חיים) / אביב אורבן
 * Apps Script: מקבל הגשת טופס, יוצר תיקייה לדייר ב-Drive ושומר שורה בגיליון.
 */

const SHEET_ID = '1GRFXZ4kDY-MsYKiWh4uFYQfhMmbnNak3KWF3VA1wIOQ';
const SHEET_TAB_NAME = 'דיירים';
const MAOZ_AVIV_FOLDER_ID = '1fLblc97zvxEwVPqh5jfc0ZEYCL6bBZIa';

const FILE_FIELDS = {
  C29_rental_contract: 'חוזה שכירות',
  C35_arnona: 'חשבון ארנונה',
  C36_water: 'חשבון מים',
  C37_electricity: 'חשבון חשמל',
  C38_gas: 'חשבון גז',
  C39_id_copy: 'צילום ת.ז',
  C43_bank_confirmation: 'אישור בנק',
  deed_file: 'נסח טאבו',
};

const COLUMN_ORDER = [
  '__timestamp__',
  'C7_first_name', 'C8_last_name', 'C10_id', 'C11_mobile', 'C12_email',
  'C3_address', 'C4_entrance', 'C5_floor', 'C6_apt_num', 'C13_current_address',
  'has_contact', 'C15_contact_name', 'C16_contact_phone', 'C17_contact_email',
  'C18_senior', 'C19_senior_details', 'C20_age',
  'C21_status', 'C23_tenant_names', 'C24_tenant_phone', 'C25_eviction_notice',
  'C40_bank_name', 'C41_branch', 'C42_account',
  'C2_tat_helka', 'deed_address', 'deed_floor', 'deed_apt_num',
  'C44_signed', 'final_confirm',
  '__folder_link__',
  'C29_rental_contract', 'C35_arnona', 'C36_water', 'C37_electricity',
  'C38_gas', 'C39_id_copy', 'C43_bank_confirmation', 'deed_file',
  '__submission_status__',
  '__legal_doc_version__', '__legal_ip__', '__legal_user_agent__',
];

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    // ללא חובות — אם חסרים שדות, נשתמש ב-fallback כדי לא לחסום את הדייר.
    const parent = DriveApp.getFolderById(MAOZ_AVIV_FOLDER_ID);
    const tsPretty = Utilities.formatDate(new Date(), 'Asia/Jerusalem', 'yyyy-MM-dd HHmm');

    const address = (data.C3_address || '').trim();
    let buildingFolder;
    if (address) {
      buildingFolder = findFolderByName(parent, address);
      if (!buildingFolder) {
        // כתובת שהוזנה לא תואמת אף תיקייה — fallback לתיקיית "ללא כתובת"
        buildingFolder = getOrCreateFolder(parent, 'ללא כתובת');
      }
    } else {
      buildingFolder = getOrCreateFolder(parent, 'ללא כתובת');
    }

    const firstName = (data.C7_first_name || '').trim();
    const lastName = (data.C8_last_name || '').trim();
    const residentName = `${firstName} ${lastName}`.trim();
    const aptNum = (String(data.C6_apt_num || '')).trim();

    let residentFolderName;
    if (aptNum && residentName) {
      residentFolderName = `דירה ${aptNum} — ${residentName}`;
    } else if (residentName) {
      residentFolderName = residentName;
    } else if (aptNum) {
      residentFolderName = `דירה ${aptNum}`;
    } else {
      residentFolderName = `הגשה ללא פרטים — ${tsPretty}`;
    }

    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_TAB_NAME);
    if (!sheet) return jsonResponse({ success: false, error: 'גיליון לא נמצא: ' + SHEET_TAB_NAME });

    const isResubmission = idAlreadySubmitted(sheet, data.C10_id);

    let residentFolder = findFolderByName(buildingFolder, residentFolderName);
    let uploadTarget;

    if (residentFolder) {
      let completions = findFolderByName(residentFolder, 'השלמות');
      if (!completions) completions = residentFolder.createFolder('השלמות');
      uploadTarget = completions;
    } else {
      residentFolder = buildingFolder.createFolder(residentFolderName);
      uploadTarget = residentFolder;
    }

    const isCompletion = uploadTarget.getName() === 'השלמות';
    const tsPrefix = Utilities.formatDate(new Date(), 'Asia/Jerusalem', 'yyyy-MM-dd_HHmm');

    const fileLinks = {};
    for (const field of Object.keys(FILE_FIELDS)) {
      const f = data[field];
      if (!f || !f.base64) continue;
      const ext = (f.name && f.name.indexOf('.') >= 0) ? f.name.split('.').pop() : 'bin';
      const baseName = FILE_FIELDS[field];
      const filename = isCompletion
        ? `${tsPrefix}_${baseName}.${ext}`
        : `${baseName}.${ext}`;
      const blob = Utilities.newBlob(Utilities.base64Decode(f.base64), f.mimeType || 'application/octet-stream', filename);
      const driveFile = uploadTarget.createFile(blob);
      fileLinks[field] = driveFile.getUrl();
    }

    const row = COLUMN_ORDER.map(col => {
      if (col === '__timestamp__') return new Date();
      if (col === '__folder_link__') return residentFolder.getUrl();
      if (col === '__submission_status__') {
        return isResubmission ? (isCompletion ? 'השלמה' : 'הגשה ראשונה') : 'הגשה ראשונה';
      }
      if (col === '__legal_doc_version__') return data.__legal_doc_version__ || '';
      if (col === '__legal_ip__') return data.__legal_ip__ || '';
      if (col === '__legal_user_agent__') return data.__legal_user_agent__ || '';
      if (FILE_FIELDS[col]) return fileLinks[col] || '';
      const v = data[col];
      return (v === undefined || v === null) ? '' : v;
    });

    sheet.appendRow(row);

    return jsonResponse({ success: true, folder: residentFolder.getUrl() });

  } catch (err) {
    return jsonResponse({ success: false, error: String(err && err.message || err) });
  }
}

function doGet() {
  return ContentService
    .createTextOutput('דיירי תמורה — Web App פעיל')
    .setMimeType(ContentService.MimeType.TEXT);
}

function findFolderByName(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : null;
}

function getOrCreateFolder(parent, name) {
  return findFolderByName(parent, name) || parent.createFolder(name);
}

function idAlreadySubmitted(sheet, id) {
  if (!id) return false;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;
  const values = sheet.getRange(2, 4, lastRow - 1, 1).getValues(); // column D = ת.ז.
  const idStr = String(id).trim();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim() === idStr) return true;
  }
  return false;
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * הרץ בדיקות אבחון — בודק שהכל מחובר נכון ושאין דבר חסר.
 * תוצאות מודפסות ב-Logger (View → Logs / Cmd+Enter).
 */
function runTests() {
  const expectedBuildings = [
    'מוהליבר 28',
    'ויצברד 17',
    'החפץ חיים 5',
    'החפץ חיים 7',
    'החפץ חיים 9',
  ];

  const results = [];
  const check = (name, fn) => {
    try {
      const info = fn();
      results.push({ ok: true, name, info: info || '' });
    } catch (err) {
      results.push({ ok: false, name, info: String(err && err.message || err) });
    }
  };

  check('פתיחת הגיליון לפי ID', () => {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    return 'שם הגיליון: ' + ss.getName();
  });

  check('קיום Tab "' + SHEET_TAB_NAME + '"', () => {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_TAB_NAME);
    if (!sheet) throw new Error('לא נמצא tab בשם זה');
    return 'OK';
  });

  check('כותרות בשורה הראשונה', () => {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_TAB_NAME);
    const lastCol = sheet.getLastColumn();
    const expected = COLUMN_ORDER.length;
    if (lastCol < expected) {
      throw new Error('נמצאו ' + lastCol + ' עמודות, צפויות ' + expected);
    }
    const first = sheet.getRange(1, 1).getValue();
    if (!first) throw new Error('שורה ראשונה ריקה');
    return lastCol + ' עמודות, מתחיל ב-"' + first + '"';
  });

  check('פתיחת תיקיית "פרויקט מעוז אביב" ב-Drive', () => {
    const folder = DriveApp.getFolderById(MAOZ_AVIV_FOLDER_ID);
    return 'שם: ' + folder.getName();
  });

  for (const name of expectedBuildings) {
    check('תיקיית בניין: ' + name, () => {
      const parent = DriveApp.getFolderById(MAOZ_AVIV_FOLDER_ID);
      const f = findFolderByName(parent, name);
      if (!f) throw new Error('לא נמצאה תת-תיקייה בשם זה');
      return 'OK';
    });
  }

  check('הרשאת כתיבה ל-Drive (יצירת + מחיקת תיקיית בדיקה)', () => {
    const parent = DriveApp.getFolderById(MAOZ_AVIV_FOLDER_ID);
    const test = parent.createFolder('__test_delete_me__');
    test.setTrashed(true);
    return 'OK';
  });

  check('הרשאת כתיבה לגיליון (כתיבה + מחיקה של תא בדיקה)', () => {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_TAB_NAME);
    const lastCol = sheet.getLastColumn();
    const testCell = sheet.getRange(1, lastCol + 1);
    testCell.setValue('__test__');
    testCell.clearContent();
    return 'OK';
  });

  const ok = results.filter(r => r.ok).length;
  const fail = results.length - ok;
  const lines = results.map(r => (r.ok ? '✅ ' : '❌ ') + r.name + (r.info ? ' — ' + r.info : ''));
  const summary = '\n────────────────\n' +
    'סיכום: ' + ok + ' עברו · ' + fail + ' נכשלו';

  const report = lines.join('\n') + summary;
  Logger.log(report);
  return report;
}

/**
 * הרצה חד-פעמית — מוסיפה 3 עמודות לוג משפטי בסוף הגיליון:
 * גרסת מסמך משפטי, IP של המגיש, User Agent.
 * בטוח להרצה חוזרת — לא ידרוס נתונים אם הכותרות כבר קיימות.
 */
function addLegalLogColumns() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_TAB_NAME);
  if (!sheet) throw new Error('גיליון לא נמצא: ' + SHEET_TAB_NAME);

  const newHeaders = ['גרסת מסמך משפטי', 'IP של המגיש', 'User Agent'];
  const lastCol = sheet.getLastColumn();
  const existing = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(String);

  if (newHeaders.every(h => existing.indexOf(h) !== -1)) {
    Logger.log('הכותרות כבר קיימות — לא נדרשה פעולה.');
    return;
  }

  const range = sheet.getRange(1, lastCol + 1, 1, newHeaders.length);
  range.setValues([newHeaders]);
  range.setFontWeight('bold');
  range.setBackground('#1a1a1a');
  range.setFontColor('#ffffff');
  range.setHorizontalAlignment('center');
  Logger.log('נוספו ' + newHeaders.length + ' כותרות חדשות בהצלחה.');
}

/**
 * הרצה ידנית פעם אחת — מגדיר RTL, מקפיא שורה ראשונה, מעצב כותרת.
 */
function setupSheet() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_TAB_NAME);
  if (!sheet) throw new Error('גיליון לא נמצא: ' + SHEET_TAB_NAME);
  sheet.setRightToLeft(true);
  sheet.setFrozenRows(1);
  const lastCol = sheet.getLastColumn();
  const header = sheet.getRange(1, 1, 1, lastCol);
  header.setFontWeight('bold');
  header.setBackground('#1a1a1a');
  header.setFontColor('#ffffff');
  header.setHorizontalAlignment('center');
  sheet.autoResizeColumns(1, lastCol);
}
