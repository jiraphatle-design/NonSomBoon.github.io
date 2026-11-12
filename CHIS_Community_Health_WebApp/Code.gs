/**
 * Community Health Intelligence System (CHIS)
 * Google Apps Script backend
 *
 * วิธีใช้:
 * 1) เปิดโปรเจกต์ Apps Script ที่ผูกกับ Google Sheet
 * 2) เพิ่มไฟล์ HTML ชื่อ index แล้ววางเนื้อหาจาก index.html
 * 3) ตรวจสอบ SPREADSHEET_ID และสิทธิ์ของบัญชีที่ใช้ Deploy
 * 4) Deploy เวอร์ชันใหม่เป็น Web app แล้วเปิด URL ของ Web App
 *
 * หมายเหตุ: ข้อมูลสุขภาพเป็นข้อมูลส่วนบุคคลที่มีความอ่อนไหว
 * ต้องกำหนดสิทธิ์และมาตรการคุ้มครองข้อมูลให้เหมาะสมก่อนใช้งานจริง
 */
const SPREADSHEET_ID = '14-AgcLs-TamNKANfQCifsX9LyGvItv3YW9Xt49XQ_so';

const SHEETS = {
  HOUSEHOLDS: 'Households',
  HEPATITIS: 'Hepatitis',
  PETS: 'Pets',
  COMMUNITY_SURVEYS: 'CommunitySurveys'
};

function doGet(e) {
  const action = String((e && e.parameter && e.parameter.action) || '').trim();
  if (!action) {
    return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('CHIS | ระบบสารสนเทศสุขภาพชุมชน');
  }
  try {
    if (action === 'getDashboard') return json(getDashboard());
    if (action === 'getHouseholds') return json(getHouseholds());
    if (action === 'getCommunitySurveys') return json(getCommunitySurveys());
    return json({ok:false,error:'ไม่พบคำสั่งที่ร้องขอ'});
  } catch (err) {
    return json({ok:false,error:errorMessage(err)});
  }
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) throw new Error('ไม่พบข้อมูลที่ส่งเข้ามา');
    const body = JSON.parse(e.postData.contents);
    const payload = body.payload || {};
    if (body.action === 'saveHousehold') return json(saveHousehold(payload));
    if (body.action === 'updateHouseholdCoordinates') return json(updateHouseholdCoordinates(payload));
    if (body.action === 'saveHepatitis') return json(saveHepatitis(payload));
    if (body.action === 'savePet') return json(savePet(payload));
    if (body.action === 'saveCommunitySurvey') return json(saveCommunitySurvey(payload));
    return json({ok:false,error:'ไม่พบคำสั่งที่ร้องขอ'});
  } catch (err) {
    return json({ok:false,error:errorMessage(err)});
  }
}

function json(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function ss() {
  if (!SPREADSHEET_ID || SPREADSHEET_ID.indexOf('PUT_') === 0) {
    throw new Error('กรุณากำหนดรหัส Spreadsheet ให้ถูกต้อง');
  }
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function ensureSheets() {
  const book = ss();
  ensureSheet(book, SHEETS.HOUSEHOLDS, [
    'id','householdId','headName','members','surveyStatus','lat','lng',
    'elderly','ncd','disabled','pregnant','note','createdAt'
  ]);
  ensureSheet(book, SHEETS.HEPATITIS, [
    'id','person','virus','screening','confirm','treatment','date','createdAt'
  ]);
  ensureSheet(book, SHEETS.PETS, [
    'id','householdId','type','count','vaccine','vaccineDate','status','createdAt'
  ]);
  ensureSheet(book, SHEETS.COMMUNITY_SURVEYS, [
    'id','householdId','headName','collector','surveyDate','data','createdAt'
  ]);
}

function ensureSheet(book, name, headers) {
  let sh = book.getSheetByName(name);
  if (!sh) sh = book.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.appendRow(headers);
  } else {
    const currentHeaders = sh.getRange(1, 1, 1, headers.length).getDisplayValues()[0];
    if (headers.some((header, index) => currentHeaders[index] !== header)) {
      throw new Error('หัวตารางในชีต ' + name + ' ไม่ตรงกับรูปแบบที่ระบบกำหนด');
    }
  }
  return sh;
}

function saveHousehold(d) {
  ensureSheets();
  const householdId = requiredText(d.householdId, 'รหัสครัวเรือน');
  const headName = requiredText(d.headName, 'ชื่อหัวหน้าครัวเรือน');
  const status = enumValue(d.surveyStatus, ['complete','partial','pending'], 'pending', 'สถานะการสำรวจ');
  const sh = ss().getSheetByName(SHEETS.HOUSEHOLDS);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const existing = sh.getLastRow() > 1 ? sh.getRange(2, 2, sh.getLastRow() - 1, 1).getDisplayValues().flat() : [];
    if (existing.indexOf(householdId) !== -1) throw new Error('มีรหัสครัวเรือนนี้ในระบบแล้ว');
    const id = Utilities.getUuid();
    sh.appendRow([id,householdId,headName,n(d.members, 'จำนวนสมาชิก', 1),status,
      coordinate(d.lat, 'ละติจูด'),coordinate(d.lng, 'ลองจิจูด'),n(d.elderly, 'ผู้สูงอายุ'),
      n(d.ncd, 'ผู้ป่วยโรคไม่ติดต่อเรื้อรัง (NCDs)'),n(d.disabled, 'คนพิการ'),n(d.pregnant, 'หญิงตั้งครรภ์'),
      safeText(d.note),new Date()]);
    return {ok:true,id:id};
  } finally {
    lock.releaseLock();
  }
}

function updateHouseholdCoordinates(d) {
  ensureSheets();
  const householdId = requiredText(d.householdId, 'รหัสครัวเรือน');
  const lat = coordinate(d.lat, 'ละติจูด');
  const lng = coordinate(d.lng, 'ลองจิจูด');
  if (lat === '' || lng === '') throw new Error('กรุณาเลือกตำแหน่งที่มีทั้งละติจูดและลองจิจูด');

  const sh = ss().getSheetByName(SHEETS.HOUSEHOLDS);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ids = sh.getLastRow() > 1 ? sh.getRange(2, 2, sh.getLastRow() - 1, 1).getDisplayValues().flat() : [];
    const rowIndex = ids.indexOf(householdId);
    if (rowIndex < 0) throw new Error('ไม่พบรหัสครัวเรือนนี้ในระบบ');
    sh.getRange(rowIndex + 2, 6, 1, 2).setValues([[lat, lng]]);
    return {ok:true,householdId:householdId,lat:lat,lng:lng};
  } finally {
    lock.releaseLock();
  }
}

function saveHepatitis(d) {
  ensureSheets();
  const sh = ss().getSheetByName(SHEETS.HEPATITIS);
  const id = Utilities.getUuid();
  appendRow(sh, [id,requiredText(d.person, 'ชื่อ–สกุลหรือรหัสผู้รับบริการ'),
    enumValue(d.virus, ['B','C','B และ C'], 'B', 'ชนิดไวรัส'),
    enumValue(d.screening, ['negative','positive','pending'], 'pending', 'ผลคัดกรอง'),
    enumValue(d.confirm, ['pending','confirmed','notconfirmed'], 'pending', 'ผลตรวจยืนยัน'),
    enumValue(d.treatment, ['none','followup','treating','completed'], 'none', 'สถานะการรักษา'),
    dateText(d.date, 'วันที่บันทึก'),new Date()]);
  return {ok:true,id:id};
}

function savePet(d) {
  ensureSheets();
  const sh = ss().getSheetByName(SHEETS.PETS);
  const id = Utilities.getUuid();
  appendRow(sh, [id,requiredText(d.householdId, 'รหัสครัวเรือน'),
    enumValue(d.type, ['สุนัข','แมว','สัตว์ปีก','อื่น ๆ'], 'อื่น ๆ', 'ชนิดสัตว์'),
    n(d.count, 'จำนวนสัตว์', 1),safeText(d.vaccine),dateText(d.vaccineDate, 'วันที่ได้รับวัคซีน'),
    enumValue(d.status, ['ได้รับวัคซีนครบตามข้อมูลที่บันทึกไว้','ยังไม่ได้รับวัคซีนหรือได้รับไม่ครบ','ไม่ทราบ'], 'ไม่ทราบ', 'สถานะการได้รับวัคซีน'),new Date()]);
  return {ok:true,id:id};
}

function saveCommunitySurvey(d) {
  ensureSheets();
  const householdId = requiredText(d.householdId, 'รหัสครัวเรือน');
  const headName = requiredText(d.headName, 'ชื่อหัวหน้าครัวเรือน');
  const collector = requiredText(d.collector, 'ชื่อผู้เก็บข้อมูล');
  const surveyDate = dateText(d.surveyDate, 'วันที่สำรวจ');
  if (!surveyDate) throw new Error('กรุณาระบุวันที่สำรวจ');
  const surveyStatus = enumValue(d.surveyStatus, ['complete','partial'], 'partial', 'สถานะการสำรวจ');
  if (!d.data || typeof d.data !== 'object' || Array.isArray(d.data)) {
    throw new Error('ไม่พบคำตอบในแบบสำรวจ');
  }

  const formData = cleanSurveyData(d.data);
  const serialized = JSON.stringify(formData);
  if (serialized.length > 45000) throw new Error('ข้อมูลแบบสำรวจยาวเกินไป กรุณาลดข้อความในช่องหมายเหตุ');
  const household = formData.household || {};
  const members = Array.isArray(formData.members) ? formData.members : [];
  const memberCount = n(household.totalMembers, 'จำนวนสมาชิกทั้งหมด', 1);
  if (members.length > 40) throw new Error('บันทึกสมาชิกได้ไม่เกิน 40 คนต่อแบบสำรวจ');

  const book = ss();
  const surveySheet = book.getSheetByName(SHEETS.COMMUNITY_SURVEYS);
  const householdSheet = book.getSheetByName(SHEETS.HOUSEHOLDS);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const rows = householdSheet.getLastRow() > 1
      ? householdSheet.getRange(2, 1, householdSheet.getLastRow() - 1, 13).getValues()
      : [];
    const rowIndex = rows.findIndex(row => String(row[1]) === householdId);
    const existing = rowIndex < 0 ? null : rows[rowIndex];
    let elderly = existing ? existing[7] : 0;
    let ncd = existing ? existing[8] : 0;
    let disabled = existing ? existing[9] : 0;
    let pregnant = existing ? existing[10] : 0;
    if (members.some(person => person.age !== '')) elderly = members.filter(person => Number(person.age) >= 60).length;
    if (members.some(person => Array.isArray(person.chronicDiseases) && person.chronicDiseases.length)) {
      ncd = members.filter(person => Array.isArray(person.chronicDiseases) &&
        person.chronicDiseases.some(value => value !== 'ไม่มีโรคประจำตัว')).length;
    }
    if (members.some(person => String(person.disabilityType || '').trim())) {
      disabled = members.filter(person => String(person.disabilityType || '').trim() &&
        person.disabilityType !== 'ไม่มี').length;
    }
    if (formData.maternalChild && formData.maternalChild.pregnant === 'มี') pregnant = 1;
    if (formData.maternalChild && formData.maternalChild.pregnant === 'ไม่มี') pregnant = 0;
    if (rowIndex < 0) {
      householdSheet.appendRow([Utilities.getUuid(),householdId,headName,memberCount,surveyStatus,
        '', '',elderly,ncd,disabled,pregnant,'',new Date()]);
    } else {
      const sheetRow = rowIndex + 2;
      householdSheet.getRange(sheetRow, 3, 1, 3).setValues([[headName,memberCount,surveyStatus]]);
      householdSheet.getRange(sheetRow, 8, 1, 4).setValues([[elderly,ncd,disabled,pregnant]]);
    }
    const id = Utilities.getUuid();
    surveySheet.appendRow([id,householdId,headName,collector,surveyDate,serialized,new Date()]);
    return {ok:true,id:id,householdId:householdId};
  } finally {
    lock.releaseLock();
  }
}

function cleanSurveyData(value) {
  if (Array.isArray(value)) return value.map(cleanSurveyData);
  if (value && typeof value === 'object') {
    const result = {};
    Object.keys(value).forEach(key => { result[safeText(key)] = cleanSurveyData(value[key]); });
    return result;
  }
  if (typeof value === 'string') return value.replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F]/g, '').trim();
  return value == null || typeof value === 'number' || typeof value === 'boolean' ? value : String(value);
}

function getHouseholds() {
  ensureSheets();
  return {ok:true,households:readObjects(SHEETS.HOUSEHOLDS)};
}

function getDashboard() {
  ensureSheets();
  return {
    ok:true,
    households:readObjects(SHEETS.HOUSEHOLDS),
    hepatitis:readObjects(SHEETS.HEPATITIS),
    pets:readObjects(SHEETS.PETS),
    communitySurveys:readCommunitySurveySummaries()
  };
}

function getCommunitySurveys() {
  ensureSheets();
  return {ok:true,communitySurveys:readObjects(SHEETS.COMMUNITY_SURVEYS)};
}

function readCommunitySurveySummaries() {
  return readObjects(SHEETS.COMMUNITY_SURVEYS).map(row => {
    let memberCount = 0;
    try {
      const data = JSON.parse(row.data || '{}');
      memberCount = data.household && Number(data.household.totalMembers)
        ? Number(data.household.totalMembers)
        : (Array.isArray(data.members) ? data.members.length : 0);
    } catch (err) {}
    return {
      id:row.id,householdId:row.householdId,headName:row.headName,collector:row.collector,
      surveyDate:row.surveyDate,createdAt:row.createdAt,memberCount:memberCount
    };
  });
}

function readObjects(sheetName) {
  const sh = ss().getSheetByName(sheetName);
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values.shift().map(String);
  return values.filter(row => row.some(value => value !== '')).map(row => {
    const obj = {};
    headers.forEach((h,i) => {
      let v = row[i];
      if (v instanceof Date) v = Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      obj[h] = v;
    });
    return obj;
  });
}

function appendRow(sheet, values) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    sheet.appendRow(values);
  } finally {
    lock.releaseLock();
  }
}

function requiredText(value, label) {
  const text = safeText(value).trim();
  if (!text) throw new Error('กรุณาระบุ' + label);
  return text;
}

function safeText(value) {
  let text = value == null ? '' : String(value).trim();
  text = text.replace(/[\u0000-\u001F\u007F]/g, '');
  if (/^[=+@]/.test(text) || /^-[^0-9]/.test(text)) text = "'" + text;
  return text;
}

function n(value, label, minimum) {
  minimum = minimum == null ? 0 : minimum;
  if (value === '' || value == null) {
    if (minimum > 0) throw new Error('กรุณาระบุ' + label);
    return 0;
  }
  const number = Number(value);
  if (!Number.isInteger(number) || number < minimum) {
    throw new Error(label + 'ต้องเป็นจำนวนเต็มตั้งแต่ ' + minimum + ' ขึ้นไป');
  }
  return number;
}

function dateText(value, label) {
  const text = safeText(value);
  if (!text) return '';
  if (!/^\d{4}-\d{2}-\d{2}$/.test(text)) throw new Error('รูปแบบ' + label + 'ไม่ถูกต้อง');
  const parts = text.split('-').map(Number);
  const date = new Date(Date.UTC(parts[0], parts[1] - 1, parts[2]));
  if (date.getUTCFullYear() !== parts[0] || date.getUTCMonth() !== parts[1] - 1 || date.getUTCDate() !== parts[2]) {
    throw new Error('รูปแบบ' + label + 'ไม่ถูกต้อง');
  }
  return text;
}

function coordinate(value, label) {
  if (value === '' || value == null) return '';
  const number = Number(value);
  if (!Number.isFinite(number)) throw new Error(label + 'ต้องเป็นตัวเลข');
  if (label === 'ละติจูด' && (number < -90 || number > 90)) throw new Error('ละติจูดต้องอยู่ระหว่าง -90 ถึง 90 องศา');
  if (label === 'ลองจิจูด' && (number < -180 || number > 180)) throw new Error('ลองจิจูดต้องอยู่ระหว่าง -180 ถึง 180 องศา');
  return number;
}

function enumValue(value, allowed, fallback, label) {
  const text = value == null || value === '' ? fallback : String(value);
  if (allowed.indexOf(text) === -1) throw new Error('ค่า' + label + 'ไม่ถูกต้อง');
  return text;
}

function errorMessage(err) {
  return err && err.message ? err.message : String(err);
}
