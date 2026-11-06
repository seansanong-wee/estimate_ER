// ==========================================
// ระบบประเมินประสิทธิภาพ: ฝั่งเซิร์ฟเวอร์ (Google Apps Script)
//
// ตั้งค่าใน Project Settings > Script Properties (ห้ามใส่ไว้ในโค้ด):
//   MANAGER_USERNAME  ชื่อผู้ใช้ของหัวหน้า
//   MANAGER_PASSWORD  รหัสผ่านของหัวหน้า
//
// ครั้งแรกให้เรียกใช้ setupOnce หนึ่งครั้ง เพื่อเติม id ให้แถวเก่าและเพิ่มคอลัมน์รหัสแก้ไข
// ==========================================

var COL = {
  NAME: 0, POSITION: 1, DEPARTMENT: 2, PERIOD: 3, KPI_FINAL: 4, BEHAVIOR_FINAL: 5, TOTAL: 6, GRADE: 7,
  EVALUATOR: 8, COMMENTS: 9, DATE: 10, KPI_DATA: 11, BEHAVIOR_DATA: 12, ID: 13, EDIT_CODE: 14
};
var ROW_LENGTH = 15;

var PERIODS = ['ครั้งที่ 1 (1 ต.ค. - 31 มี.ค.)', 'ครั้งที่ 2 (1 เม.ย. - 30 ก.ย.)'];
var FISCAL_YEARS = ['2568', '2569', '2570', '2571', '2572'];

var DEFAULT_KPI = [
  { id: 1, name: 'อัตราการเสียชีวิตของผู้ป่วยวิกฤตฉุกเฉิน (triage level 1) ภายใน 24 ชม.', target: 5, weight: 20 },
  { id: 2, name: 'หน้าที่รับผิดชอบหลัก', target: 5, weight: 30 },
  { id: 3, name: 'พฤติกรรมบริการ', target: 5, weight: 20 },
  { id: 4, name: 'หน้าที่ที่ได้รับมอบหมาย', target: 5, weight: 20 },
  { id: 5, name: 'กิจกรรม ๕ส.', target: 5, weight: 10 }
];
var DEFAULT_BEHAVIOR = [
  { id: 1, name: 'การมุ่งผลสัมฤทธิ์', target: 5, weight: 30 },
  { id: 2, name: 'การบริการที่ดี', target: 5, weight: 20 },
  { id: 3, name: 'การสั่งสมความเชี่ยวชาญในงานอาชีพ', target: 5, weight: 20 },
  { id: 4, name: 'การยึดมั่นในความถูกต้องคุณธรรมและจริยธรรม', target: 5, weight: 10 },
  { id: 5, name: 'การทำงานเป็นทีม', target: 5, weight: 20 }
];

var MIN_EVALUATED_FOR_PUBLIC_AVG = 5; // ต่ำกว่านี้ค่าเฉลี่ยอาจเผยคะแนนรายคน
var MIN_EDIT_CODE_LENGTH = 4;
var MAX_FAILED_ATTEMPTS = 5;
var FAILED_ATTEMPT_WINDOW_SECONDS = 900;
var SESSION_SECONDS = 21600; // 6 ชั่วโมง ซึ่งเป็นค่าสูงสุดของ CacheService

// ==========================================
// จุดรับคำขอ
// ==========================================

// ข้อมูลสาธารณะเท่านั้น ห้ามส่งคะแนน ความเห็น หรือรหัสแก้ไขออกจากฟังก์ชันนี้
function doGet() {
  var records = readRecords_(getSheet_());
  return json_({ ok: true, rows: records.map(toPublicRow_), stats: buildPublicStats_(records) });
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000); // ให้ทำงานทีละคำขอ ป้องกันแถวซ้ำเมื่อบันทึกพร้อมกัน
  } catch (err) {
    return fail_('BUSY');
  }
  try {
    var data = JSON.parse(e.postData.contents);
    var handler = ACTIONS_[data.action];
    if (!handler) return fail_('INVALID');
    if (handler.manager && !isValidSession_(data.token)) return fail_('AUTH');
    return handler.run(data);
  } catch (err) {
    console.error(err);
    return fail_(err.code || 'SERVER');
  } finally {
    lock.releaseLock();
  }
}

var ACTIONS_ = {
  login: { run: login_ },
  logout: { run: logout_ },
  selfCreate: { run: selfCreate_ },
  openForEdit: { run: openForEdit_ },
  selfSave: { run: selfSave_ },
  list: { manager: true, run: managerList_ },
  save: { manager: true, run: managerSave_ },
  delete: { manager: true, run: managerDelete_ },
  setUnlocked: { manager: true, run: managerSetUnlocked_ }
};

// เรียกใช้เองหนึ่งครั้งจากหน้า Apps Script: เติม id ให้แถวเก่าและตั้งชื่อคอลัมน์รหัสแก้ไข
function setupOnce() {
  var sheet = getSheet_();
  var values = sheet.getDataRange().getValues();
  var base = new Date().getTime();
  var filled = 0;
  for (var i = 1; i < values.length; i++) {
    if (values[i][COL.NAME] && !values[i][COL.ID]) {
      sheet.getRange(i + 1, COL.ID + 1).setValue(base + i);
      filled++;
    }
  }
  sheet.getRange(1, COL.EDIT_CODE + 1).setValue('รหัสแก้ไข (เข้ารหัสแล้ว)');
  console.log('เติม id ให้ ' + filled + ' แถว');
}

// ==========================================
// หัวหน้า
// ==========================================

function setManagerAccount() {
  PropertiesService.getScriptProperties().setProperties({
    MANAGER_USERNAME: 'wee',
    MANAGER_PASSWORD: '2112'
  });
  CacheService.getScriptCache().remove('login_fail');
}

  var key = 'login_fail';
  if (isRateLimited_(key)) return fail_('RATE_LIMITED');
  if (data.username !== username || data.password !== password) {
    recordFailure_(key);
    return fail_('BAD_LOGIN');
  }
  clearFailures_(key);

  var token = Utilities.getUuid() + Utilities.getUuid();
  CacheService.getScriptCache().put('session_' + token, '1', SESSION_SECONDS);
  return json_({ ok: true, token: token });
}

function logout_(data) {
  if (data.token) CacheService.getScriptCache().remove('session_' + data.token);
  return json_({ ok: true });
}

function isValidSession_(token) {
  return !!token && CacheService.getScriptCache().get('session_' + token) === '1';
}

function managerList_() {
  return json_({ ok: true, rows: readRecords_(getSheet_()).map(toManagerRow_) });
}

function managerSave_(data) {
  var sheet = getSheet_();
  var values = sheet.getDataRange().getValues();
  var rowNumbers = data.id ? findRowNumbers_(values, data.id) : [];
  var existing = rowNumbers.length ? values[rowNumbers[0] - 1] : null;

  var editCodeCell = existing ? existing[COL.EDIT_CODE] || '' : '';
  if (data.newEditCode) {
    if (!isValidEditCode_(data.newEditCode)) return fail_('INVALID_CODE');
    editCodeCell = makeEditCodeCell_(data.newEditCode);
  }

  // สถานะปลดล็อกเปลี่ยนได้ทางปุ่มล็อกเท่านั้น ฟอร์มที่เปิดค้างไว้จึงเขียนทับไม่ได้
  var kpiData = Array.isArray(data.kpiData) ? data.kpiData : [];
  if (kpiData[0]) {
    kpiData[0].selfEditUnlocked = !!existing && (parseJson_(existing[COL.KPI_DATA]) || [{}])[0].selfEditUnlocked === true;
  }

  var id = data.id || newId_(values);
  var row = [
    safeText_(data.employeeName || data.name), safeText_(data.position), safeText_(data.department), safeText_(data.period),
    Number(data.kpiScoreFinal) || 0, Number(data.behaviorScoreFinal) || 0, Number(data.totalScore) || 0,
    safeText_(data.grade), safeText_(data.evaluatorName), safeText_(data.comments),
    data.date || today_(),
    JSON.stringify(kpiData), JSON.stringify(data.behaviorData || []),
    id, editCodeCell
  ];
  writeRow_(sheet, rowNumbers, row);
  return json_({ ok: true, id: id });
}

function managerDelete_(data) {
  if (!data.id) return fail_('INVALID');
  var sheet = getSheet_();
  var rowNumbers = findRowNumbers_(sheet.getDataRange().getValues(), data.id);
  for (var i = rowNumbers.length - 1; i >= 0; i--) {
    sheet.deleteRow(rowNumbers[i]);
  }
  return json_({ ok: true });
}

function managerSetUnlocked_(data) {
  var sheet = getSheet_();
  var values = sheet.getDataRange().getValues();
  var rowNumbers = findRowNumbers_(values, data.id);
  if (!rowNumbers.length) return fail_('NOT_FOUND');

  var row = values[rowNumbers[0] - 1];
  var kpiData = parseJson_(row[COL.KPI_DATA]) || [];
  if (!kpiData[0]) return fail_('INVALID');
  kpiData[0].selfEditUnlocked = data.unlocked === true;
  // เก็บปีงบประมาณไว้ เพราะชีตไม่มีคอลัมน์นี้
  kpiData[0].fiscalYear = fiscalYearOf_(kpiData, row[COL.DATE], row[COL.PERIOD]);
  row[COL.KPI_DATA] = JSON.stringify(kpiData);
  writeRow_(sheet, rowNumbers, padRow_(row));
  return json_({ ok: true });
}

// ==========================================
// ผู้รับการประเมิน
// ==========================================

function selfCreate_(data) {
  var name = normalizeName_(data.employeeName);
  var fiscalYear = String(data.fiscalYear);
  if (!name || PERIODS.indexOf(data.period) < 0 || FISCAL_YEARS.indexOf(fiscalYear) < 0) return fail_('INVALID');
  if (!isValidEditCode_(data.editCode)) return fail_('INVALID_CODE');

  var sheet = getSheet_();
  var values = sheet.getDataRange().getValues();
  if (hasRecordInSlot_(values, name, fiscalYear, data.period, null)) return fail_('DUPLICATE');

  var kpiData = mergeSelfScores_(cloneDefaults_(DEFAULT_KPI), data.kpiData, true);
  var behaviorData = mergeSelfScores_(cloneDefaults_(DEFAULT_BEHAVIOR), data.behaviorData, false);
  if (!kpiData || !behaviorData) return fail_('INVALID');
  kpiData[0].fiscalYear = fiscalYear;

  var totals = computeTotals_(kpiData, behaviorData);
  var id = newId_(values);
  sheet.appendRow([
    safeText_(name), safeText_(data.position), safeText_(data.department), data.period,
    totals.kpi, totals.behavior, totals.total, totals.grade,
    '', '', today_(),
    JSON.stringify(kpiData), JSON.stringify(behaviorData),
    id, makeEditCodeCell_(data.editCode)
  ]);
  return json_({ ok: true, id: id });
}

function openForEdit_(data) {
  var found = findEditableRow_(getSheet_().getDataRange().getValues(), data);
  if (found.error) return fail_(found.error);

  var row = found.row;
  var kpiData = parseJson_(row[COL.KPI_DATA]) || [];
  return json_({
    ok: true,
    record: {
      id: row[COL.ID],
      name: row[COL.NAME],
      employeeName: row[COL.NAME],
      position: row[COL.POSITION],
      department: row[COL.DEPARTMENT],
      period: row[COL.PERIOD],
      date: row[COL.DATE],
      fiscalYear: fiscalYearOf_(kpiData, row[COL.DATE], row[COL.PERIOD]),
      kpiData: kpiData.map(toSelfItem_),
      behaviorData: (parseJson_(row[COL.BEHAVIOR_DATA]) || []).map(toSelfItem_)
    }
  });
}

function selfSave_(data) {
  var sheet = getSheet_();
  var values = sheet.getDataRange().getValues();
  var found = findEditableRow_(values, data);
  if (found.error) return fail_(found.error);

  var row = found.row;
  var fiscalYear = String(data.fiscalYear);
  if (PERIODS.indexOf(data.period) < 0 || FISCAL_YEARS.indexOf(fiscalYear) < 0) return fail_('INVALID');
  if (hasRecordInSlot_(values, row[COL.NAME], fiscalYear, data.period, row[COL.ID])) return fail_('DUPLICATE');

  // คะแนนหัวหน้า น้ำหนัก และความเห็น มาจากชีตเสมอ รับจากผู้ส่งเฉพาะคะแนนตนเอง
  var wasEvaluated = isEvaluated_(parseJson_(row[COL.KPI_DATA]), parseJson_(row[COL.BEHAVIOR_DATA]));
  var kpiData = mergeSelfScores_(parseJson_(row[COL.KPI_DATA]) || [], data.kpiData, true);
  var behaviorData = mergeSelfScores_(parseJson_(row[COL.BEHAVIOR_DATA]) || [], data.behaviorData, false);
  if (!kpiData || !behaviorData || !kpiData[0]) return fail_('INVALID');
  kpiData[0].fiscalYear = fiscalYear;
  // ปลดล็อกใช้ได้ครั้งเดียว บันทึกแล้วกลับไปล็อกทันที เพื่อให้หัวหน้าตรวจผลล่าสุดได้
  if (wasEvaluated) kpiData[0].selfEditUnlocked = false;

  var totals = computeTotals_(kpiData, behaviorData);
  row[COL.POSITION] = safeText_(data.position);
  row[COL.DEPARTMENT] = safeText_(data.department);
  row[COL.PERIOD] = data.period;
  row[COL.KPI_FINAL] = totals.kpi;
  row[COL.BEHAVIOR_FINAL] = totals.behavior;
  row[COL.TOTAL] = totals.total;
  row[COL.GRADE] = totals.grade;
  row[COL.KPI_DATA] = JSON.stringify(kpiData);
  row[COL.BEHAVIOR_DATA] = JSON.stringify(behaviorData);
  writeRow_(sheet, found.rowNumbers, padRow_(row));
  return json_({ ok: true });
}

// ตรวจตามลำดับ: มีแถวนี้ไหม ล็อกอยู่ไหม มีรหัสไหม รหัสถูกไหม
function findEditableRow_(values, data) {
  var rowNumbers = data.id ? findRowNumbers_(values, data.id) : [];
  if (!rowNumbers.length) return { error: 'NOT_FOUND' };

  var row = values[rowNumbers[0] - 1];
  if (!canSelfEdit_(parseJson_(row[COL.KPI_DATA]), parseJson_(row[COL.BEHAVIOR_DATA]))) return { error: 'LOCKED' };
  if (!row[COL.EDIT_CODE]) return { error: 'NO_CODE' };

  var key = 'code_fail_' + data.id;
  if (isRateLimited_(key)) return { error: 'RATE_LIMITED' };
  if (!matchesEditCode_(row[COL.EDIT_CODE], data.code)) {
    recordFailure_(key);
    return { error: 'BAD_CODE' };
  }
  clearFailures_(key);
  return { row: row, rowNumbers: rowNumbers };
}

// ==========================================
// การคำนวณคะแนน: ต้องตรงกับฟอร์มในหน้าเว็บ
// ==========================================

function computeTotals_(kpiData, behaviorData) {
  var kpi = sectionPercent_(kpiData);
  var behavior = sectionPercent_(behaviorData);
  var total = kpi * 0.7 + behavior * 0.3;
  return { kpi: kpi, behavior: behavior, total: total, grade: gradeOf_(total) };
}

function sectionPercent_(items) {
  var sum = items.reduce(function (acc, item) {
    return acc + rowAverage_(item.selfScore, item.managerScore) * (Number(item.weight) / 100);
  }, 0);
  return sum * 100 / 5;
}

function rowAverage_(self, manager) {
  var sum = 0, count = 0;
  if (Number(self) > 0) { sum += Number(self); count++; }
  if (Number(manager) > 0) { sum += Number(manager); count++; }
  return count === 0 ? 0 : sum / count;
}

function gradeOf_(score) {
  if (score === 0) return 'รอการประเมิน';
  if (score >= 90) return 'ดีเด่น';
  if (score >= 80) return 'ดีมาก';
  if (score >= 70) return 'ดี';
  if (score >= 60) return 'พอใช้';
  return 'ต้องปรับปรุง';
}

// คืน null ถ้าคะแนนตนเองไม่ครบหรือไม่อยู่ในช่วง 1-5
function mergeSelfScores_(stored, submitted, allowFirstNameEdit) {
  if (!Array.isArray(submitted)) return null;
  var byId = {};
  submitted.forEach(function (item) { if (item) byId[item.id] = item; });

  var merged = [];
  for (var i = 0; i < stored.length; i++) {
    var incoming = byId[stored[i].id];
    var score = incoming ? Number(incoming.selfScore) : NaN;
    if (!(score >= 1 && score <= 5) || Math.floor(score) !== score) return null;
    var item = stored[i];
    item.selfScore = score;
    if (allowFirstNameEdit && item.id === 1 && typeof incoming.name === 'string') {
      item.name = incoming.name.slice(0, 500);
    }
    merged.push(item);
  }
  return merged;
}

function cloneDefaults_(items) {
  return items.map(function (item) {
    return { id: item.id, name: item.name, target: item.target, weight: item.weight, selfScore: 0, managerScore: 0 };
  });
}

// ==========================================
// แปลงแถวเป็นข้อมูลที่ส่งออก
// ==========================================

function readRecords_(sheet) {
  var values = sheet.getDataRange().getValues();
  var byId = {};
  var order = [];
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    if (!row[COL.NAME]) continue;
    var key = String(row[COL.ID] || 'missing-' + (i + 1));
    if (!(key in byId)) order.push(key);
    byId[key] = toRecord_(row, key); // ถ้ามี id ซ้ำ ใช้แถวล่างสุด
  }
  return order.map(function (key) { return byId[key]; });
}

function toRecord_(row, key) {
  var kpiData = parseJson_(row[COL.KPI_DATA]);
  var behaviorData = parseJson_(row[COL.BEHAVIOR_DATA]);
  return {
    id: row[COL.ID] || key,
    employeeName: row[COL.NAME],
    name: row[COL.NAME],
    position: row[COL.POSITION] || '',
    department: row[COL.DEPARTMENT] || '',
    period: row[COL.PERIOD] || PERIODS[0],
    kpiScoreFinal: parseFloat(row[COL.KPI_FINAL]) || 0,
    behaviorScoreFinal: parseFloat(row[COL.BEHAVIOR_FINAL]) || 0,
    totalScore: parseFloat(row[COL.TOTAL]) || 0,
    grade: row[COL.GRADE] || '',
    evaluatorName: row[COL.EVALUATOR] || '',
    comments: row[COL.COMMENTS] || '',
    date: row[COL.DATE] || '',
    kpiData: kpiData,
    behaviorData: behaviorData,
    fiscalYear: fiscalYearOf_(kpiData, row[COL.DATE], row[COL.PERIOD]),
    hasEditCode: !!row[COL.EDIT_CODE]
  };
}

function toManagerRow_(record) {
  return record;
}

function toPublicRow_(record) {
  return {
    id: record.id,
    employeeName: record.employeeName,
    name: record.name,
    position: record.position,
    department: record.department,
    period: record.period,
    date: record.date,
    fiscalYear: record.fiscalYear,
    managerEvaluated: isEvaluated_(record.kpiData, record.behaviorData),
    selfEditUnlocked: isUnlocked_(record.kpiData)
  };
}

function toSelfItem_(item) {
  return { id: item.id, name: item.name, target: item.target, weight: item.weight, selfScore: item.selfScore, managerScore: 0 };
}

// ค่าเฉลี่ยของแต่ละปีและรอบ ส่งเฉพาะกลุ่มที่ประเมินแล้วถึงเกณฑ์
function buildPublicStats_(records) {
  var groups = {};
  records.forEach(function (record) {
    if (!isEvaluated_(record.kpiData, record.behaviorData)) return;
    [
      'all|all',
      record.fiscalYear + '|all',
      'all|' + record.period,
      record.fiscalYear + '|' + record.period
    ].forEach(function (key) {
      groups[key] = groups[key] || { sum: 0, count: 0 };
      groups[key].sum += record.totalScore;
      groups[key].count++;
    });
  });

  var stats = {};
  Object.keys(groups).forEach(function (key) {
    if (groups[key].count >= MIN_EVALUATED_FOR_PUBLIC_AVG) {
      stats[key] = { avg: Math.round(groups[key].sum / groups[key].count * 100) / 100 };
    }
  });
  return stats;
}

// ==========================================
// เครื่องมือช่วย
// ==========================================

function isEvaluated_(kpiData, behaviorData) {
  var hasScore = function (items) {
    return Array.isArray(items) && items.some(function (item) { return Number(item.managerScore) > 0; });
  };
  return hasScore(kpiData) || hasScore(behaviorData);
}

function isUnlocked_(kpiData) {
  return Array.isArray(kpiData) && !!kpiData[0] && kpiData[0].selfEditUnlocked === true;
}

function canSelfEdit_(kpiData, behaviorData) {
  return !isEvaluated_(kpiData, behaviorData) || isUnlocked_(kpiData);
}

function currentFiscalYear_() {
  var now = new Date();
  var year = String(now.getFullYear() + 543 + (now.getMonth() >= 9 ? 1 : 0));
  return FISCAL_YEARS.indexOf(year) >= 0 ? year : FISCAL_YEARS[1];
}

function fiscalYearOf_(kpiData, dateValue, period) {
  var stored = Array.isArray(kpiData) && kpiData[0] ? String(kpiData[0].fiscalYear || '') : '';
  if (FISCAL_YEARS.indexOf(stored) >= 0) return stored;

  var date = dateValue ? new Date(dateValue) : new Date();
  if (isNaN(date.getTime())) return currentFiscalYear_();
  var month = date.getMonth();
  // ครั้งที่ 2 มักประเมินหลังจบรอบ เช่น ต.ค. จึงยังเป็นปีงบประมาณที่เพิ่งสิ้นสุด
  var offset = period === PERIODS[1] ? (month < 3 ? -1 : 0) : (month >= 9 ? 1 : 0);
  var year = String(date.getFullYear() + 543 + offset);
  return FISCAL_YEARS.indexOf(year) >= 0 ? year : currentFiscalYear_();
}

function hasRecordInSlot_(values, name, fiscalYear, period, exceptId) {
  var target = normalizeName_(name);
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    if (!row[COL.NAME] || (exceptId && String(row[COL.ID]) === String(exceptId))) continue;
    if (normalizeName_(row[COL.NAME]) === target
        && (row[COL.PERIOD] || PERIODS[0]) === period
        && fiscalYearOf_(parseJson_(row[COL.KPI_DATA]), row[COL.DATE], row[COL.PERIOD]) === fiscalYear) {
      return true;
    }
  }
  return false;
}

function findRowNumbers_(values, id) {
  var rows = [];
  var target = String(id || '');
  if (!target) return rows;
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][COL.ID] || '') === target) rows.push(i + 1);
  }
  return rows;
}

// ถ้ามีแถว id ซ้ำ ให้แก้ทุกแถวเพื่อให้ข้อมูลตรงกัน
function writeRow_(sheet, rowNumbers, row) {
  if (!rowNumbers.length) {
    sheet.appendRow(row);
    return;
  }
  rowNumbers.forEach(function (rowNumber) {
    sheet.getRange(rowNumber, 1, 1, row.length).setValues([row]);
  });
}

function padRow_(row) {
  var padded = row.slice(0, ROW_LENGTH);
  while (padded.length < ROW_LENGTH) padded.push('');
  return padded;
}

function newId_(values) {
  var used = {};
  for (var i = 1; i < values.length; i++) used[String(values[i][COL.ID])] = true;
  var id = new Date().getTime();
  while (used[String(id)]) id++;
  return id;
}

function isValidEditCode_(code) {
  return typeof code === 'string' && code.trim().length >= MIN_EDIT_CODE_LENGTH && code.length <= 100;
}

function hashEditCode_(salt, code) {
  var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + code.trim(), Utilities.Charset.UTF_8);
  return Utilities.base64Encode(digest);
}

function makeEditCodeCell_(code) {
  var salt = Utilities.getUuid();
  return salt + ':' + hashEditCode_(salt, code);
}

function matchesEditCode_(cell, code) {
  if (typeof code !== 'string') return false;
  var parts = String(cell).split(':');
  return parts.length === 2 && hashEditCode_(parts[0], code) === parts[1];
}

function isRateLimited_(key) {
  return Number(CacheService.getScriptCache().get(key)) >= MAX_FAILED_ATTEMPTS;
}

function recordFailure_(key) {
  var cache = CacheService.getScriptCache();
  cache.put(key, String(Number(cache.get(key) || 0) + 1), FAILED_ATTEMPT_WINDOW_SECONDS);
}

function clearFailures_(key) {
  CacheService.getScriptCache().remove(key);
}

function normalizeName_(name) {
  return String(name === undefined || name === null ? '' : name).trim().replace(/\s+/g, ' ');
}

// กันข้อความที่ขึ้นต้นด้วยอักขระสูตรไม่ให้ชีตรันเป็นสูตร
function safeText_(value) {
  var text = String(value === undefined || value === null ? '' : value).slice(0, 2000);
  return /^[=+\-@]/.test(text) ? "'" + text : text;
}

function parseJson_(text) {
  if (!text) return null;
  try {
    return JSON.parse(text);
  } catch (err) {
    return null;
  }
}

function today_() {
  return new Date().toISOString().split('T')[0];
}

function getSheet_() {
  return SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function fail_(code) {
  return json_({ ok: false, code: code });
}
