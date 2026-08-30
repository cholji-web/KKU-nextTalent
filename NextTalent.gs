/**
 * ==================================================================
 * NextTalent Prototype — Part 1 + Part 2 + Avatar + Skill Gap Analysis (Google Apps Script)
 * Entry point เดียวของ Web App — single-page app ตั้งแต่แรก
 * เสิร์ฟ App.html ไฟล์เดียวเสมอ การสลับ view (หน้าหลัก/Admin/ค้นหา/แดชบอร์ด)
 * ทำด้วย JS ฝั่ง client ทั้งหมด — ไม่มี doPost()
 * ==================================================================
 */
function doGet() {
  return HtmlService.createTemplateFromFile('App')
    .evaluate()
    .setTitle('NextTalent')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/* ==================================================================
 * SETUP — รันครั้งเดียวก่อนใช้งานจริง
 * ================================================================== */

/**
 * รันครั้งเดียวก่อนใช้งาน — สร้าง Sheet Users / Submissions / Config พร้อม header
 * ไม่เขียนทับถ้ามี Sheet นั้นอยู่แล้ว (เช็คแค่ชื่อ Sheet ไม่เช็คข้อมูลข้างใน)
 * Users มีคอลัมน์ photoUrl (D) ไว้ใส่ลิงก์รูปโปรไฟล์ — ปล่อยว่างได้ ฝั่ง client จะ
 * generate เป็น avatar ตัวอักษรย่อสีอัตโนมัติแทน
 */
function setupSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var usersSheet = ss.getSheetByName('Users');
  if (!usersSheet) {
    usersSheet = ss.insertSheet('Users');
    usersSheet.appendRow(['userId', 'name', 'department', 'photoUrl']);
    usersSheet.setFrozenRows(1);
  }

  var subSheet = ss.getSheetByName('Submissions');
  if (!subSheet) {
    subSheet = ss.insertSheet('Submissions');
    subSheet.appendRow([
      'id', 'timestamp', 'userId', 'fileUrl', 'ocrText',
      'skill', 'level', 'rarity', 'xp', 'status'
    ]);
    subSheet.setFrozenRows(1);
  }

  var configSheet = ss.getSheetByName('Config');
  if (!configSheet) {
    configSheet = ss.insertSheet('Config');
    configSheet.appendRow(['category', 'key', 'value']);
    configSheet.setFrozenRows(1);
  }

  Logger.log('setupSheet() เสร็จสิ้น: Users / Submissions / Config พร้อมใช้งาน');
}

/** ข้อความ prompt เริ่มต้นสำหรับวิเคราะห์ทักษะ (ใช้ตอน seed ครั้งแรก) */
var DEFAULT_SKILL_PROMPT_ =
  'คุณคือผู้ช่วยวิเคราะห์ทักษะจากข้อความที่ได้ โดยวิเคราะห์หาว่าเป็นทักษะด้านใดใน 3 ทักษะนี้ ' +
  '1.IT 2.บริหาร 3.ทั่วไป และวิเคราะห์ว่าเป็นทักษะระดับใด 1.มหาวิทยาลัย 2.ระดับประเทศ 3.นานาชาติ ' +
  'ตอบกลับเป็น JSON ที่ถูกต้องเท่านั้น รูปแบบ {"skill":"...","level":"..."} ห้ามมีข้อความอื่นนอกเหนือจาก JSON';

/**
 * รันครั้งเดียวหลัง setupSheet() — ใส่ค่า default ลง Config และ user ตัวอย่างลง Users
 * เช็คว่า Sheet มีข้อมูลอยู่แล้วหรือยัง (getLastRow() <= 1 คือมีแค่ header)
 * ถ้ามีข้อมูลแล้วจะไม่เขียนทับ (ยกเว้น prompt/benchmark ซึ่งเช็คแยกเพื่อรองรับ migration)
 */
function setupSeedData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- seed Users ---
  var usersSheet = ss.getSheetByName('Users');
  if (!usersSheet) {
    throw new Error('ไม่พบ Sheet "Users" — กรุณารัน setupSheet() ก่อน');
  }
  if (usersSheet.getLastRow() <= 1) {
    usersSheet.getRange(2, 1, 3, 3).setValues([
      ['u001', 'มานะ ใจดี', 'ฝ่าย IT'],
      ['u002', 'สมหญิง รักเรียน', 'ฝ่ายบริหาร'],
      ['u003', 'วิชัย ตั้งใจทำงาน', 'ฝ่ายทั่วไป']
    ]);
  }

  // --- seed Config (level / rarity / curve) เฉพาะตอนที่ยังไม่มีข้อมูลเลย ---
  var configSheet = ss.getSheetByName('Config');
  if (!configSheet) {
    throw new Error('ไม่พบ Sheet "Config" — กรุณารัน setupSheet() ก่อน');
  }
  if (configSheet.getLastRow() <= 1) {
    var rows = [
      ['level', 'มหาวิทยาลัย', 10],
      ['level', 'ประเทศ', 30],
      ['level', 'นานาชาติ', 50],
      ['rarity', 'IT', 1.0],
      ['rarity', 'บริหาร', 1.5],
      ['rarity', 'ทั่วไป', 1.0],
      ['curve', 'base', 100],
      ['curve', 'exponent', 1.5],
      ['specialCurve', 'base', 50],
      ['specialCurve', 'exponent', 1.5]
    ];
    configSheet.getRange(2, 1, rows.length, 3).setValues(rows);
  }

  // เติมแถว prompt ให้เสมอ ถ้ายังไม่มี (ไม่สนใจว่า Config มีข้อมูลอื่นอยู่แล้วหรือไม่)
  // เพื่อรองรับ migration กรณีที่รัน setupSeedData() ไปแล้วรอบก่อนหน้า (ก่อนมี feature นี้)
  ensurePromptSeed_();

  // เติมแถว benchmark ให้เสมอ ถ้ายังไม่มี (ใช้กับ Skill Gap Analysis บนหน้า Dashboard)
  ensureBenchmarkSeed_();

  Logger.log('setupSeedData() เสร็จสิ้น: Config และ Users มีข้อมูลตัวอย่างพร้อมใช้งาน');
}

/** เติมแถว category="prompt" ลง Config ถ้ายังไม่มี (เรียกซ้ำได้ ไม่ทำให้เกิดแถวซ้ำ) */
function ensurePromptSeed_() {
  var configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
  if (!configSheet) return;

  var data = configSheet.getDataRange().getValues();
  var hasPrompt = data.some(function (row) {
    return row[0] === 'prompt' && row[1] === 'skillAnalysis';
  });

  if (!hasPrompt) {
    configSheet.appendRow(['prompt', 'skillAnalysis', DEFAULT_SKILL_PROMPT_]);
    Logger.log('เพิ่มแถว prompt เริ่มต้นลง Config เรียบร้อย');
  }
}

/**
 * เติมแถว category="benchmark" ลง Config ถ้ายังไม่มี (เรียกซ้ำได้ ไม่ทำให้เกิดแถวซ้ำ)
 * ใช้เป็นเกณฑ์เป้าหมาย XP เฉลี่ยของแต่ละหน่วยงาน/ทักษะ สำหรับ Skill Gap Analysis บน Dashboard
 * key รูปแบบ "หน่วยงาน::ทักษะ" — มี "default::ทักษะ" เป็นเกณฑ์กลาง (fallback เมื่อไม่มี override เฉพาะหน่วยงาน)
 */
function ensureBenchmarkSeed_() {
  var configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
  if (!configSheet) return;

  var data = configSheet.getDataRange().getValues();
  var hasBenchmark = data.some(function (row) { return row[0] === 'benchmark'; });

  if (!hasBenchmark) {
    var rows = [
      ['benchmark', getBenchmarkKey_('default', 'IT'), 300],
      ['benchmark', getBenchmarkKey_('default', 'บริหาร'), 300],
      ['benchmark', getBenchmarkKey_('default', 'ทั่วไป'), 200],
      ['benchmark', getBenchmarkKey_('ฝ่าย IT', 'IT'), 500],
      ['benchmark', getBenchmarkKey_('ฝ่ายบริหาร', 'บริหาร'), 500],
      ['benchmark', getBenchmarkKey_('ฝ่ายทั่วไป', 'ทั่วไป'), 300]
    ];
    configSheet.getRange(configSheet.getLastRow() + 1, 1, rows.length, 3).setValues(rows);
    Logger.log('เพิ่มแถว benchmark เริ่มต้นลง Config เรียบร้อย');
  }
}

/**
 * MIGRATION: ใช้ตอน Sheet Users มีอยู่แล้วก่อนหน้า (deploy รอบก่อนยังไม่มี photoUrl)
 * รันครั้งเดียวเพื่อเพิ่มคอลัมน์ photoUrl ที่ท้ายตาราง Users แบบไม่กระทบข้อมูลเดิม
 * ถ้ามีคอลัมน์ photoUrl อยู่แล้วจะไม่ทำอะไร (เช็คจาก header row)
 */
function migrateAddPhotoUrlColumn() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  if (!sheet) {
    throw new Error('ไม่พบ Sheet "Users" — กรุณารัน setupSheet() ก่อน');
  }
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (headers.indexOf('photoUrl') !== -1) {
    Logger.log('คอลัมน์ photoUrl มีอยู่แล้ว ไม่ต้อง migrate');
    return;
  }
  var newCol = sheet.getLastColumn() + 1;
  sheet.getRange(1, newCol).setValue('photoUrl');
  Logger.log('เพิ่มคอลัมน์ photoUrl ที่ column ' + newCol + ' เรียบร้อย');
}

/* ==================================================================
 * PART 1 — PUBLIC API: หน้าหลัก (user)
 * ================================================================== */

/** คืนรายชื่อ user ทั้งหมดจาก Sheet "Users" สำหรับ dropdown เลือก user (รวม photoUrl) */
function getUsersList() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  if (!sheet) {
    throw new Error('ไม่พบ Sheet "Users" — กรุณารัน setupSheet() ก่อน');
  }

  var data = sheet.getDataRange().getValues();
  var users = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    users.push({
      userId: String(data[i][0]),
      name: String(data[i][1] || ''),
      department: String(data[i][2] || ''),
      photoUrl: String(data[i][3] || '') // คอลัมน์ D — ว่างได้ ฝั่ง client จะ fallback เป็น avatar ตัวอักษร
    });
  }
  return users;
}

/**
 * คืนข้อมูลสรุปของ user คนหนึ่ง: ชื่อ, หน่วยงาน, photoUrl, growupLevel (approvedXp/pendingXp/level/badge
 * ภาพรวมทุก skill), specialList (approvedXp/pendingXp/level แยกตาม skill คนละก้อน),
 * และรายการที่ส่งผลทั้งหมด — level/badge นับจาก approved เท่านั้น, submitted แสดงเป็น pendingXp
 * (ไม่กระทบ level), rejected ไม่นับทั้งคู่แต่เก็บประวัติไว้
 */
function getUserState(userId) {
  if (!userId) {
    throw new Error('ไม่พบ userId');
  }

  try {
    var user = getUsersList().filter(function (u) { return u.userId === userId; })[0];
    if (!user) {
      throw new Error('ไม่พบ user: ' + userId);
    }

    var submissions = getSubmissionsByUser_(userId);
    var config = loadConfig_();

    if (!config.curve || config.curve.base === undefined || config.curve.exponent === undefined) {
      throw new Error('Config sheet ไม่มีค่า curve.base / curve.exponent ครบ — ตรวจสอบ Sheet "Config"');
    }
    if (!config.specialCurve || config.specialCurve.base === undefined || config.specialCurve.exponent === undefined) {
      throw new Error('Config sheet ไม่มีค่า specialCurve.base / specialCurve.exponent ครบ — ตรวจสอบ Sheet "Config"');
    }
    if (!config.rarity || Object.keys(config.rarity).length === 0) {
      throw new Error('Config sheet ไม่มีค่า rarity เลย — ตรวจสอบ Sheet "Config"');
    }

    // --- growupLevel: รวม XP ทุก skill เข้าด้วยกันเป็นก้อนเดียว ---
    var overallApprovedXp = submissions
      .filter(function (s) { return s.status === 'approved'; })
      .reduce(function (sum, s) { return sum + Number(s.xp || 0); }, 0);

    var overallPendingXp = submissions
      .filter(function (s) { return s.status === 'submitted'; })
      .reduce(function (sum, s) { return sum + Number(s.xp || 0); }, 0);

    var growupLevelInfo = calcLevelFromXp_(overallApprovedXp, config.curve.base, config.curve.exponent);

    var growupLevel = {
      approvedXp: overallApprovedXp,
      pendingXp: overallPendingXp,
      level: growupLevelInfo.level,
      xpIntoLevel: growupLevelInfo.xpIntoLevel,
      xpForNextLevel: growupLevelInfo.xpForNextLevel,
      badge: getBadge_(growupLevelInfo.level)
    };

    // --- specialList: แยก XP/level เป็นก้อนอิสระต่อ skill (IT / บริหาร / ทั่วไป) ---
    var skillNames = Object.keys(config.rarity);
    var specialList = skillNames.map(function (skill) {
      var approvedXp = submissions
        .filter(function (s) { return s.status === 'approved' && s.skill === skill; })
        .reduce(function (sum, s) { return sum + Number(s.xp || 0); }, 0);

      var pendingXp = submissions
        .filter(function (s) { return s.status === 'submitted' && s.skill === skill; })
        .reduce(function (sum, s) { return sum + Number(s.xp || 0); }, 0);

      var levelInfo = calcLevelFromXp_(approvedXp, config.specialCurve.base, config.specialCurve.exponent);

      return {
        skill: skill,
        approvedXp: approvedXp,
        pendingXp: pendingXp,
        level: levelInfo.level,
        xpIntoLevel: levelInfo.xpIntoLevel,
        xpForNextLevel: levelInfo.xpForNextLevel,
        badge: getBadge_(levelInfo.level)
      };
    });

    var result = {
      userId: user.userId,
      name: user.name,
      department: user.department,
      photoUrl: user.photoUrl,
      growupLevel: growupLevel,
      specialList: specialList,
      submissions: submissions
    };

    JSON.stringify(result); // บังคับให้ error ถ้า serialize ไม่ได้ (ป้องกัน silent-null)

    return result;

  } catch (err) {
    Logger.log('getUserState() error: ' + err.message + '\n' + (err.stack || ''));
    throw new Error('getUserState ล้มเหลว: ' + err.message);
  }
}

/**
 * ขั้นตอนหลัก: เซฟไฟล์ลง Drive -> OCR ผ่าน n8n -> วิเคราะห์ทักษะผ่าน AI Gateway -> คำนวณ XP
 * คืนค่าผลลัพธ์ให้ client แสดงเป็นคะแนนชั่วคราว (ยังไม่ insert ลง Sheet Submissions)
 */
function analyzeCertificate(fileBase64, fileName, mimeType) {
  if (!fileBase64) {
    throw new Error('ไม่พบข้อมูลไฟล์ (fileBase64)');
  }

  var blob = Utilities.newBlob(
    Utilities.base64Decode(fileBase64),
    mimeType || 'application/pdf',
    fileName || 'certificate.pdf'
  );

  var fileUrl = saveToDrive_(blob);
  var extractedText = callOcrWebhook_(blob);
  var aiResult = callAiGateway_(extractedText);
  var config = loadConfig_();

  var rarity = config.rarity[aiResult.skill];
  var baseXp = config.level[aiResult.level];

  if (rarity === undefined || baseXp === undefined) {
    throw new Error(
      'ไม่พบค่า rarity/baseXP ของ skill="' + aiResult.skill + '" level="' + aiResult.level +
      '" ใน Config sheet (skill ที่มี: ' + Object.keys(config.rarity).join(', ') +
      ' / level ที่มี: ' + Object.keys(config.level).join(', ') + ')'
    );
  }

  return {
    fileUrl: fileUrl,
    ocrText: extractedText,
    skill: aiResult.skill,
    level: aiResult.level,
    rarity: Number(rarity),
    xp: Number(baseXp) * Number(rarity)
  };
}

/**
 * บันทึกผลที่ user กด "ส่งผล" ลง Sheet "Submissions" (status = submitted)
 * คืนค่า getUserState() ล่าสุด เพื่อให้ client รีเฟรช XP/level/badge/รายการได้ทันที
 */
function submitResult(userId, result) {
  if (!userId) {
    throw new Error('ไม่พบ userId');
  }
  if (!result || !result.skill || !result.level || result.xp === undefined) {
    throw new Error('ข้อมูลผลลัพธ์ไม่ครบ ไม่สามารถส่งผลได้');
  }

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Submissions');
  if (!sheet) {
    throw new Error('ไม่พบ Sheet "Submissions" — กรุณารัน setupSheet() ก่อน');
  }

  sheet.appendRow([
    generateId_(),
    new Date(),
    userId,
    result.fileUrl || '',
    result.ocrText || '',
    result.skill,
    result.level,
    Number(result.rarity) || 1,
    Number(result.xp) || 0,
    'submitted'
  ]);

  return getUserState(userId);
}

/* ==================================================================
 * PART 2 — PUBLIC API: หน้า Admin
 * ================================================================== */

/** คืนตัวเลือก skill/level สำหรับ dropdown ในหน้า Admin (แก้ไขก่อนอนุมัติ) — ดึงจาก Config sheet จริง */
function getConfigOptions() {
  var config = loadConfig_();
  return {
    skills: Object.keys(config.rarity),
    levels: Object.keys(config.level)
  };
}

/**
 * คืนรายการ status=submitted ของทุก user (join กับ Sheet Users เพื่อได้ชื่อ/หน่วยงาน/รูปโปรไฟล์)
 * เรียงจากเก่าไปใหม่ (คิวประมวลผล — รายการที่รอนานที่สุดอยู่บนสุด)
 * เรียกใหม่ทุกครั้งที่ client สลับเข้า view Admin เพื่อให้เห็นข้อมูลล่าสุด
 */
function getPendingSubmissions() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Submissions');
  if (!sheet) {
    throw new Error('ไม่พบ Sheet "Submissions" — กรุณารัน setupSheet() ก่อน');
  }

  var data = sheet.getDataRange().getValues();
  var headers = data[0];

  var userMap = {};
  getUsersList().forEach(function (u) { userMap[u.userId] = u; });

  var list = [];
  for (var i = 1; i < data.length; i++) {
    var obj = {};
    headers.forEach(function (h, idx) { obj[h] = data[i][idx]; });

    if (String(obj.status) !== 'submitted') continue;

    obj.id = String(obj.id || '');
    obj.timestamp = (obj.timestamp instanceof Date) ? obj.timestamp.toISOString() : String(obj.timestamp || '');
    obj.userId = String(obj.userId || '');
    obj.fileUrl = String(obj.fileUrl || '');
    obj.ocrText = String(obj.ocrText || '');
    obj.skill = String(obj.skill || '');
    obj.level = String(obj.level || '');
    obj.rarity = Number(obj.rarity) || 1;
    obj.xp = Number(obj.xp) || 0;
    obj.status = String(obj.status || '');

    var user = userMap[obj.userId];
    obj.userName = user ? user.name : '(ไม่พบ user: ' + obj.userId + ')';
    obj.department = user ? user.department : '-';
    obj.userPhotoUrl = user ? user.photoUrl : '';
    obj.previewUrl = getPreviewUrlFromDriveUrl_(obj.fileUrl);

    list.push(obj);
  }

  list.sort(function (a, b) { return new Date(a.timestamp) - new Date(b.timestamp); });
  return list;
}

/**
 * แก้ไข skill/level ของรายการที่ยังรออนุมัติ (ก่อนกดอนุมัติจริง) — เปิดจาก modal ในหน้า Admin
 * คำนวณ rarity/xp ใหม่โดยอัตโนมัติจาก Config sheet ตาม skill/level ที่แก้ (ไม่ให้กรอก xp เอง)
 * status ไม่เปลี่ยน (ยังเป็น submitted) — ต้องกด "อนุมัติ" แยกทีหลังจึงมีผลกับ level จริง
 */
function updateSubmission(id, skill, level) {
  if (!id) throw new Error('ไม่พบ id ของรายการ');
  if (!skill || !level) throw new Error('กรุณาระบุ skill และ level ให้ครบ');

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Submissions');
  if (!sheet) throw new Error('ไม่พบ Sheet "Submissions" — กรุณารัน setupSheet() ก่อน');

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var rowIndex = findSubmissionRowIndex_(sheet, id);
  if (rowIndex === -1) throw new Error('ไม่พบรายการ id=' + id);

  var config = loadConfig_();
  var rarity = config.rarity[skill];
  var baseXp = config.level[level];
  if (rarity === undefined || baseXp === undefined) {
    throw new Error(
      'ไม่พบค่า rarity/baseXP ของ skill="' + skill + '" level="' + level + '" ใน Config sheet'
    );
  }
  var xp = Number(baseXp) * Number(rarity);

  var skillCol = headers.indexOf('skill') + 1;
  var levelCol = headers.indexOf('level') + 1;
  var rarityCol = headers.indexOf('rarity') + 1;
  var xpCol = headers.indexOf('xp') + 1;

  sheet.getRange(rowIndex, skillCol).setValue(skill);
  sheet.getRange(rowIndex, levelCol).setValue(level);
  sheet.getRange(rowIndex, rarityCol).setValue(rarity);
  sheet.getRange(rowIndex, xpCol).setValue(xp);

  return getPendingSubmissions();
}

/** อนุมัติรายการ — เปลี่ยน status เป็น approved (มีผลกับ level/badge ทันทีที่ getUserState ครั้งถัดไป) */
function approveSubmission(id) {
  return setSubmissionStatus_(id, 'approved');
}

/** ปฏิเสธรายการ — เปลี่ยน status เป็น rejected (ไม่ลบแถว เก็บประวัติไว้) */
function rejectSubmission(id) {
  return setSubmissionStatus_(id, 'rejected');
}

/* ==================================================================
 * PART 2 — PUBLIC API: หน้าค้นหา
 * ================================================================== */

/**
 * ค้นหา user จากชื่อ หรือชื่อทักษะใน specialList (IT / บริหาร / ทั่วไป)
 * skill ที่ใช้ค้นหาและ tag ที่แสดงนับจากรายการ approved เท่านั้น (approvedXp > 0)
 * query ว่าง = คืน user ทั้งหมด (ใช้ตอนเปิดหน้าค้นหาครั้งแรกก่อนพิมพ์อะไร)
 */
function searchUsers(query) {
  query = String(query || '').trim();
  var normalizedQuery = query.toLowerCase();

  var users = getUsersList();
  var config = loadConfig_();
  var skillNames = Object.keys(config.rarity);

  var results = users.map(function (user) {
    var submissions = getSubmissionsByUser_(user.userId);

    var approvedXp = submissions
      .filter(function (s) { return s.status === 'approved'; })
      .reduce(function (sum, s) { return sum + Number(s.xp || 0); }, 0);

    var growupInfo = calcLevelFromXp_(approvedXp, config.curve.base, config.curve.exponent);

    var skillTags = skillNames.filter(function (skill) {
      var xp = submissions
        .filter(function (s) { return s.status === 'approved' && s.skill === skill; })
        .reduce(function (sum, s) { return sum + Number(s.xp || 0); }, 0);
      return xp > 0;
    });

    return {
      userId: user.userId,
      name: user.name,
      department: user.department,
      photoUrl: user.photoUrl,
      growupLevel: {
        level: growupInfo.level,
        badge: getBadge_(growupInfo.level)
      },
      skillTags: skillTags
    };
  });

  if (!normalizedQuery) return results;

  return results.filter(function (r) {
    var nameMatch = r.name.toLowerCase().indexOf(normalizedQuery) !== -1;
    var skillMatch = r.skillTags.some(function (skill) {
      return skill.toLowerCase().indexOf(normalizedQuery) !== -1;
    });
    return nameMatch || skillMatch;
  });
}

/* ==================================================================
 * PART 2 — PUBLIC API: Skill Gap Analysis (สำหรับหน้า Dashboard)
 * ================================================================== */

/**
 * วิเคราะห์ช่องว่างทักษะ (Skill Gap) แยกตามหน่วยงาน x ทักษะ เทียบกับเกณฑ์เป้าหมายใน Config (category=benchmark)
 * คืนค่า:
 *  - departments: รายละเอียดต่อหน่วยงาน (avgXp, targetXp, gapXp, fulfillmentPct ต่อ skill + overallFulfillmentPct)
 *  - organizationAverage: ค่าเฉลี่ยรวมทั้งองค์กรต่อ skill (ไม่แยกหน่วยงาน) เทียบกับเกณฑ์กลาง (default::skill)
 *  - criticalGaps: top 5 คู่ (หน่วยงาน, ทักษะ) ที่ fulfillmentPct ต่ำสุด (เร่งพัฒนาก่อน)
 * ใช้แสดงผลบนหน้า Dashboard เท่านั้น ไม่กระทบการคำนวณ XP/level จริงของ user
 */
function getSkillGapAnalysis() {
  try {
    var users = getUsersList();
    if (!users.length) {
      throw new Error('ไม่มี user ในระบบ — ไม่สามารถวิเคราะห์ Skill Gap ได้');
    }

    var config = loadConfig_();
    var skillNames = Object.keys(config.rarity);
    if (!skillNames.length) {
      throw new Error('Config sheet ไม่มีค่า rarity เลย — ตรวจสอบ Sheet "Config"');
    }

    var allSubmissions = getAllSubmissions_();

    // จัดกลุ่ม userId ตาม department
    var deptMap = {};
    users.forEach(function (u) {
      var dept = u.department || '(ไม่ระบุหน่วยงาน)';
      if (!deptMap[dept]) deptMap[dept] = [];
      deptMap[dept].push(u.userId);
    });

    function sumXp(userIds, skill, status) {
      return allSubmissions
        .filter(function (s) {
          return userIds.indexOf(s.userId) !== -1 && s.skill === skill && s.status === status;
        })
        .reduce(function (sum, s) { return sum + s.xp; }, 0);
    }

    function round1_(n) { return Math.round(n * 10) / 10; }

    // --- คำนวณต่อหน่วยงาน ---
    var departments = Object.keys(deptMap).map(function (dept) {
      var userIds = deptMap[dept];
      var userCount = userIds.length;

      var skills = skillNames.map(function (skill) {
        var approvedTotal = sumXp(userIds, skill, 'approved');
        var pendingTotal = sumXp(userIds, skill, 'submitted');

        var avgApprovedXp = approvedTotal / userCount;
        var avgPendingXp = pendingTotal / userCount;
        var targetXp = resolveBenchmarkXp_(config, dept, skill);

        var gapXp = Math.max(0, targetXp - avgApprovedXp);
        var fulfillmentPct = targetXp > 0 ? Math.min(100, (avgApprovedXp / targetXp) * 100) : 100;
        var projectedFulfillmentPct = targetXp > 0
          ? Math.min(100, ((avgApprovedXp + avgPendingXp) / targetXp) * 100)
          : 100;

        return {
          skill: skill,
          avgApprovedXp: round1_(avgApprovedXp),
          avgPendingXp: round1_(avgPendingXp),
          targetXp: targetXp,
          gapXp: round1_(gapXp),
          fulfillmentPct: round1_(fulfillmentPct),
          projectedFulfillmentPct: round1_(projectedFulfillmentPct),
          hasTarget: targetXp > 0
        };
      });

      var skillsWithTarget = skills.filter(function (s) { return s.hasTarget; });
      var overallFulfillmentPct = skillsWithTarget.length
        ? round1_(skillsWithTarget.reduce(function (sum, s) { return sum + s.fulfillmentPct; }, 0) / skillsWithTarget.length)
        : 100;

      return {
        department: dept,
        userCount: userCount,
        skills: skills,
        overallFulfillmentPct: overallFulfillmentPct
      };
    });

    // --- ค่าเฉลี่ยรวมทั้งองค์กร (ไม่แยกหน่วยงาน) เทียบกับเกณฑ์กลาง (default) ---
    var allUserIds = users.map(function (u) { return u.userId; });
    var organizationAverage = skillNames.map(function (skill) {
      var approvedTotal = sumXp(allUserIds, skill, 'approved');
      var pendingTotal = sumXp(allUserIds, skill, 'submitted');
      var avgApprovedXp = approvedTotal / users.length;
      var avgPendingXp = pendingTotal / users.length;
      var targetXp = resolveBenchmarkXp_(config, 'default', skill);

      var gapXp = Math.max(0, targetXp - avgApprovedXp);
      var fulfillmentPct = targetXp > 0 ? Math.min(100, (avgApprovedXp / targetXp) * 100) : 100;
      var projectedFulfillmentPct = targetXp > 0
        ? Math.min(100, ((avgApprovedXp + avgPendingXp) / targetXp) * 100)
        : 100;

      return {
        skill: skill,
        avgApprovedXp: round1_(avgApprovedXp),
        avgPendingXp: round1_(avgPendingXp),
        targetXp: targetXp,
        gapXp: round1_(gapXp),
        fulfillmentPct: round1_(fulfillmentPct),
        projectedFulfillmentPct: round1_(projectedFulfillmentPct),
        hasTarget: targetXp > 0
      };
    });

    // --- Critical Gaps: เรียงจาก fulfillmentPct ต่ำสุด (ห่างเป้าหมายมากที่สุด) ---
    var criticalGaps = [];
    departments.forEach(function (d) {
      d.skills.forEach(function (s) {
        if (s.hasTarget && s.gapXp > 0) {
          criticalGaps.push({
            department: d.department,
            skill: s.skill,
            gapXp: s.gapXp,
            fulfillmentPct: s.fulfillmentPct
          });
        }
      });
    });
    criticalGaps.sort(function (a, b) { return a.fulfillmentPct - b.fulfillmentPct; });
    criticalGaps = criticalGaps.slice(0, 5);

    return {
      departments: departments,
      organizationAverage: organizationAverage,
      criticalGaps: criticalGaps
    };

  } catch (err) {
    Logger.log('getSkillGapAnalysis() error: ' + err.message + '\n' + (err.stack || ''));
    throw new Error('getSkillGapAnalysis ล้มเหลว: ' + err.message);
  }
}

/* ==================================================================
 * INTERNAL HELPERS — ฟังก์ชันภายใน (ไม่เรียกจาก client)
 * ================================================================== */

function getSubmissionsByUser_(userId) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Submissions');
  if (!sheet) {
    throw new Error('ไม่พบ Sheet "Submissions" — กรุณารัน setupSheet() ก่อน');
  }

  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var userIdCol = headers.indexOf('userId');

  if (userIdCol === -1) {
    throw new Error('Sheet "Submissions" ไม่มีคอลัมน์ "userId" — ตรวจสอบ header row');
  }

  return data
    .slice(1)
    .filter(function (row) { return String(row[userIdCol]) === String(userId); })
    .map(function (row) {
      var obj = {};
      headers.forEach(function (h, i) { obj[h] = row[i]; });

      obj.id = String(obj.id || '');
      obj.timestamp = (obj.timestamp instanceof Date)
        ? obj.timestamp.toISOString()
        : (obj.timestamp ? String(obj.timestamp) : new Date().toISOString());
      obj.userId = String(obj.userId || '');
      obj.fileUrl = String(obj.fileUrl || '');
      obj.ocrText = String(obj.ocrText || '');
      obj.skill = String(obj.skill || '');
      obj.level = String(obj.level || '');

      obj.rarity = Number(obj.rarity);
      if (isNaN(obj.rarity)) obj.rarity = 1;

      obj.xp = Number(obj.xp);
      if (isNaN(obj.xp)) obj.xp = 0;

      obj.status = String(obj.status || 'submitted');

      return obj;
    })
    .sort(function (a, b) { return new Date(b.timestamp) - new Date(a.timestamp); });
}

/**
 * ดึง submissions ทั้งหมดในระบบ (ทุก user) แบบ normalized
 * ใช้เฉพาะกรณีที่ต้องสรุปข้อมูลข้ามหลาย user เช่น Skill Gap Analysis / Dashboard
 * (ต่างจาก getSubmissionsByUser_ ที่กรองเฉพาะ user คนเดียว)
 */
function getAllSubmissions_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Submissions');
  if (!sheet) {
    throw new Error('ไม่พบ Sheet "Submissions" — กรุณารัน setupSheet() ก่อน');
  }

  var data = sheet.getDataRange().getValues();
  var headers = data[0];

  return data.slice(1).map(function (row) {
    var obj = {};
    headers.forEach(function (h, i) { obj[h] = row[i]; });

    obj.userId = String(obj.userId || '');
    obj.skill = String(obj.skill || '');
    obj.status = String(obj.status || 'submitted');
    obj.xp = Number(obj.xp);
    if (isNaN(obj.xp)) obj.xp = 0;

    return obj;
  });
}

/** หา row number (1-based, ตรงกับ getRange จริง) ของ submission ตาม id — คืน -1 ถ้าไม่พบ */
function findSubmissionRowIndex_(sheet, id) {
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var idCol = headers.indexOf('id');
  if (idCol === -1) return -1;

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][idCol]) === String(id)) return i + 1;
  }
  return -1;
}

/** ใช้ร่วมโดย approveSubmission/rejectSubmission เพื่อตั้งค่า status ของแถวใดแถวหนึ่ง */
function setSubmissionStatus_(id, status) {
  if (!id) throw new Error('ไม่พบ id ของรายการ');

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Submissions');
  if (!sheet) throw new Error('ไม่พบ Sheet "Submissions" — กรุณารัน setupSheet() ก่อน');

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var rowIndex = findSubmissionRowIndex_(sheet, id);
  if (rowIndex === -1) throw new Error('ไม่พบรายการ id=' + id);

  var statusCol = headers.indexOf('status') + 1;
  sheet.getRange(rowIndex, statusCol).setValue(status);

  return getPendingSubmissions();
}

function generateId_() {
  return 'sub_' + Utilities.getUuid();
}

/** เซฟไฟล์ PDF ลง Google Drive คืนค่าเป็นลิงก์ไฟล์ — ลองตั้ง sharing ให้ดู preview ได้ในองค์กร */
function saveToDrive_(blob) {
  var folder = getOrCreateUploadFolder_();
  var file = folder.createFile(blob);

  try {
    file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (e) {
    Logger.log('ไม่สามารถตั้งค่า sharing อัตโนมัติได้: ' + e.message);
  }

  return file.getUrl();
}

function getOrCreateUploadFolder_() {
  var folderName = 'NextTalent Uploads';
  var folders = DriveApp.getFoldersByName(folderName);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
}

/** แปลงลิงก์ Drive แบบ view (.../file/d/ID/view) เป็นลิงก์ preview สำหรับฝัง iframe */
function getPreviewUrlFromDriveUrl_(fileUrl) {
  if (!fileUrl) return '';
  var match = fileUrl.match(/\/d\/([a-zA-Z0-9_-]+)/);
  if (!match) return '';
  return 'https://drive.google.com/file/d/' + match[1] + '/preview';
}

/**
 * ส่งไฟล์ PDF เข้า n8n webhook เพื่อถอดข้อความ (OCR)
 * Request: multipart/form-data, field name "file"
 * Response ที่คาดไว้: JSON { "status": "...", "extracted_text": "..." }
 */
function callOcrWebhook_(fileBlob) {
  var url = PropertiesService.getScriptProperties().getProperty('OCR_WEBHOOK_URL');
  if (!url) {
    throw new Error('ไม่พบ Script Property "OCR_WEBHOOK_URL"');
  }

  var options = {
    method: 'post',
    payload: { file: fileBlob },
    muteHttpExceptions: true
  };

  var response = UrlFetchApp.fetch(url, options);
  var status = response.getResponseCode();
  var rawText = response.getContentText();

  var parsed = null;
  try { parsed = JSON.parse(rawText); } catch (e) { /* ไม่ใช่ JSON — fallback แสดง raw text */ }

  if (status < 200 || status >= 300) {
    throw new Error('OCR webhook ตอบกลับ HTTP ' + status + ': ' + rawText);
  }
  if (!parsed || typeof parsed.extracted_text !== 'string') {
    throw new Error('OCR webhook ตอบกลับไม่มี extracted_text ที่ใช้งานได้: ' + rawText);
  }

  var text = parsed.extracted_text.trim();
  if (!text) {
    throw new Error('OCR ไม่พบข้อความในไฟล์ (อาจเป็น PDF แบบสแกน/รูปภาพ)');
  }
  return text;
}

/**
 * ส่งข้อความที่ OCR ได้ไปวิเคราะห์ทักษะที่ KKU AI Gateway
 * systemPrompt ดึงมาจาก Sheet "Config" (category=prompt, key=skillAnalysis)
 * เพื่อให้ Admin แก้ไขข้อความ prompt ได้เองโดยไม่ต้องแก้โค้ด/deploy ใหม่
 * Response ที่คาดไว้จาก AI (message.content): valid JSON { "skill": "...", "level": "..." }
 */
function callAiGateway_(ocrText) {
  var props = PropertiesService.getScriptProperties();
  var baseUrl = props.getProperty('AI_GATEWAY_BASE_URL');
  var apiKey = props.getProperty('AI_GATEWAY_API_KEY');

  if (!baseUrl || !apiKey) {
    throw new Error('ไม่พบ Script Property "AI_GATEWAY_BASE_URL" หรือ "AI_GATEWAY_API_KEY"');
  }

  var systemPrompt = loadPrompt_('skillAnalysis');

  var payload = {
    model: 'deepseek-v4-pro',
    messages: [
      { role: 'system', content: systemPrompt },
      { role: 'user', content: ocrText }
    ],
    stream: false
  };

  var options = {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + apiKey },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  var response = UrlFetchApp.fetch(baseUrl + '/chat/completions', options);
  var status = response.getResponseCode();
  var rawText = response.getContentText();

  if (status < 200 || status >= 300) {
    throw new Error('AI Gateway ตอบกลับ HTTP ' + status + ': ' + rawText);
  }

  var body;
  try { body = JSON.parse(rawText); } catch (e) {
    throw new Error('AI Gateway ตอบกลับไม่ใช่ JSON: ' + rawText);
  }

  var content = body.choices && body.choices[0] && body.choices[0].message && body.choices[0].message.content;
  if (!content) {
    throw new Error('AI Gateway ไม่มีข้อความตอบกลับที่ใช้งานได้: ' + rawText);
  }

  var result;
  try { result = JSON.parse(content); } catch (e) {
    throw new Error('AI วิเคราะห์ทักษะตอบกลับไม่ใช่ JSON ที่ parse ได้: ' + content);
  }

  if (!result.skill || !result.level) {
    throw new Error('AI วิเคราะห์ทักษะตอบกลับไม่ครบ skill/level: ' + content);
  }
  return result;
}

/**
 * โหลดค่า baseXP (ตาม level), rarity (ตาม skill), curve (growupLevel), specialCurve (specialList)
 * และ benchmark (เป้าหมาย XP เฉลี่ยต่อหน่วยงาน/ทักษะ สำหรับ Skill Gap Analysis) จาก Sheet "Config"
 */
function loadConfig_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
  if (!sheet) {
    throw new Error('ไม่พบ Sheet "Config" — กรุณารัน setupSheet() และ setupSeedData() ก่อน');
  }

  var data = sheet.getDataRange().getValues();
  var config = { level: {}, rarity: {}, curve: {}, specialCurve: {}, benchmark: {} };

  for (var i = 1; i < data.length; i++) {
    var category = String(data[i][0] || '').trim();
    var key = String(data[i][1] || '').trim();
    var value = data[i][2];
    if (!category || !key) continue;
    if (config[category] !== undefined) {
      config[category][key] = Number(value);
    }
  }
  return config;
}

/**
 * โหลด prompt จาก Sheet "Config" (category = "prompt") ตาม key ที่กำหนด
 * ให้ Admin แก้ไขข้อความ prompt ได้ผ่านชีตโดยตรง ไม่ต้องแก้โค้ด/deploy ใหม่
 */
function loadPrompt_(key) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
  if (!sheet) {
    throw new Error('ไม่พบ Sheet "Config" — กรุณารัน setupSheet() ก่อน');
  }

  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var category = data[i][0];
    var rowKey = data[i][1];
    var value = data[i][2];
    if (category === 'prompt' && rowKey === key) {
      if (!value || String(value).trim() === '') {
        throw new Error('พบ prompt key="' + key + '" ใน Config แต่ค่า value เป็นค่าว่าง');
      }
      return String(value);
    }
  }

  throw new Error(
    'ไม่พบ prompt key="' + key + '" ใน Sheet "Config" (category=prompt) — ' +
    'กรุณาเพิ่มแถว category="prompt", key="' + key + '" พร้อมข้อความ prompt ใน column value'
  );
}

/**
 * สร้าง key มาตรฐานสำหรับ benchmark (department::skill) — trim ป้องกันเว้นวรรคเกินจากชีต
 * ใช้ร่วมกันทั้งตอน seed และตอน resolve ค่า
 */
function getBenchmarkKey_(department, skill) {
  return String(department).trim() + '::' + String(skill).trim();
}

/**
 * หาค่าเป้าหมาย (target XP) ของ department+skill คู่หนึ่ง
 * ลำดับการค้นหา: 1) เกณฑ์เฉพาะของ department นั้น 2) เกณฑ์กลาง (default::skill) 3) ไม่มีเกณฑ์ -> คืน 0
 * คืน 0 หมายถึง "ยังไม่ได้ตั้งเป้าหมายไว้" (ฝั่ง client จะแสดงว่ายังไม่มี benchmark ไม่ใช่ error)
 */
function resolveBenchmarkXp_(config, department, skill) {
  var overrideKey = getBenchmarkKey_(department, skill);
  if (config.benchmark[overrideKey] !== undefined && !isNaN(config.benchmark[overrideKey])) {
    return config.benchmark[overrideKey];
  }
  var defaultKey = getBenchmarkKey_('default', skill);
  if (config.benchmark[defaultKey] !== undefined && !isNaN(config.benchmark[defaultKey])) {
    return config.benchmark[defaultKey];
  }
  return 0;
}

/** คำนวณ level และ XP ที่เข้า/ต้องการของ level ปัจจุบัน จาก XP รวม ตามสูตร base × N^exponent */
function calcLevelFromXp_(totalXp, base, exponent) {
  base = Number(base) || 100;
  exponent = Number(exponent) || 1.5;
  totalXp = Number(totalXp) || 0;

  var level = 1;
  while (xpRequiredForLevel_(level + 1, base, exponent) <= totalXp) {
    level++;
    if (level > 1000) break;
  }

  var xpForCurrentLevel = xpRequiredForLevel_(level, base, exponent);
  var xpForNextLevel = xpRequiredForLevel_(level + 1, base, exponent);

  return {
    level: level,
    xpIntoLevel: totalXp - xpForCurrentLevel,
    xpForNextLevel: xpForNextLevel - xpForCurrentLevel
  };
}

function xpRequiredForLevel_(level, base, exponent) {
  if (level <= 1) return 0;
  return Math.round(base * Math.pow(level, exponent));
}

/** คืนชื่อ badge ตาม level: 1-4 Rookie, 5-9 Rising Star, 10-19 Expert, 20+ Master */
function getBadge_(level) {
  if (level >= 20) return { name: 'Master' };
  if (level >= 10) return { name: 'Expert' };
  if (level >= 5) return { name: 'Rising Star' };
  return { name: 'Rookie' };
}

/* ==================================================================
 * DEBUG HELPERS — ใช้ทดสอบระหว่างพัฒนา รันจาก Apps Script Editor เท่านั้น
 * ================================================================== */

function debugLoadConfig_() {
  Logger.log(JSON.stringify(loadConfig_(), null, 2));
}

function debugGetSubmissions_() {
  var userId = 'u001';
  Logger.log(JSON.stringify(getSubmissionsByUser_(userId), null, 2));
}

function debugGetPending_() {
  Logger.log(JSON.stringify(getPendingSubmissions(), null, 2));
}

function debugGetSkillGap_() {
  Logger.log(JSON.stringify(getSkillGapAnalysis(), null, 2));
}

/**
 * เพิ่ม user ใหม่ลง Sheet "Users" — สร้าง userId อัตโนมัติต่อจากรายการล่าสุด
 * (เช่นมี u001, u002, u003 อยู่แล้ว -> user ใหม่จะได้ u004)
 * คืนค่า getUsersList() ล่าสุดเพื่อให้ client รีเฟรช dropdown ได้ทันที
 */
function addUser(name, department, photoUrl) {
  name = String(name || '').trim();
  department = String(department || '').trim();
  photoUrl = String(photoUrl || '').trim();

  if (!name) {
    throw new Error('กรุณาระบุชื่อ-นามสกุลของ user');
  }

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  if (!sheet) {
    throw new Error('ไม่พบ Sheet "Users" — กรุณารัน setupSheet() ก่อน');
  }

  var newUserId = generateNextUserId_(sheet);
  sheet.appendRow([newUserId, name, department, photoUrl]);

  return getUsersList();
}

/** สร้าง userId ใหม่ต่อจากรูปแบบ "u001", "u002", ... โดยหาเลขสูงสุดที่มีอยู่แล้ว +1 */
function generateNextUserId_(sheet) {
  var data = sheet.getDataRange().getValues();
  var maxNum = 0;

  for (var i = 1; i < data.length; i++) {
    var id = String(data[i][0] || '');
    var match = id.match(/^u(\d+)$/i);
    if (match) {
      var num = parseInt(match[1], 10);
      if (num > maxNum) maxNum = num;
    }
  }

  var nextNum = maxNum + 1;
  var padded = String(nextNum);
  while (padded.length < 3) padded = '0' + padded;
  return 'u' + padded;
}

/**
 * รวมสถิติทั้งระบบสำหรับหน้า Dashboard:
 * จำนวน user ทั้งหมด, จำนวน submissions แยกตามสถานะ, XP รวมที่อนุมัติแล้ว,
 * การกระจายตัวของ skill (เฉพาะ approved), และ leaderboard top 5 (เรียงตาม XP รวม approved)
 */
function getDashboardStats() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var usersSheet = ss.getSheetByName('Users');
  var subSheet = ss.getSheetByName('Submissions');
  if (!usersSheet || !subSheet) {
    throw new Error('ไม่พบ Sheet "Users" หรือ "Submissions" — กรุณารัน setupSheet() ก่อน');
  }

  var users = getUsersList();

  var subData = subSheet.getDataRange().getValues();
  var headers = subData[0];
  var rows = subData.slice(1).map(function (row) {
    var obj = {};
    headers.forEach(function (h, i) { obj[h] = row[i]; });
    return obj;
  });

  var approved = rows.filter(function (r) { return r.status === 'approved'; });
  var pending = rows.filter(function (r) { return r.status === 'submitted'; });
  var rejected = rows.filter(function (r) { return r.status === 'rejected'; });

  var totalApprovedXp = approved.reduce(function (sum, r) { return sum + Number(r.xp || 0); }, 0);

  // การกระจายตัวของ skill (นับจำนวนรายการ approved ต่อ skill)
  var skillCounts = {};
  approved.forEach(function (r) {
    skillCounts[r.skill] = (skillCounts[r.skill] || 0) + 1;
  });

  // Leaderboard: รวม XP approved ต่อ user แล้วคำนวณ growupLevel ด้วย curve เดียวกับ getUserState()
  var config = loadConfig_();
  var leaderboard = users.map(function (u) {
    var userApprovedXp = approved
      .filter(function (r) { return r.userId === u.userId; })
      .reduce(function (sum, r) { return sum + Number(r.xp || 0); }, 0);

    var levelInfo = calcLevelFromXp_(userApprovedXp, config.curve.base, config.curve.exponent);

    return {
      userId: u.userId,
      name: u.name,
      department: u.department,
      photoUrl: u.photoUrl || '',
      xp: userApprovedXp,
      level: levelInfo.level,
      badge: getBadge_(levelInfo.level)
    };
  })
  .sort(function (a, b) { return b.xp - a.xp; })
  .slice(0, 5);

  return {
    totalUsers: users.length,
    totalSubmissions: rows.length,
    approvedCount: approved.length,
    pendingCount: pending.length,
    rejectedCount: rejected.length,
    totalApprovedXp: totalApprovedXp,
    skillCounts: skillCounts,
    leaderboard: leaderboard
  };
}