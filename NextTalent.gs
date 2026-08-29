/**
 * ==================================================================
 * NextTalent Prototype — Part 1 + Part 2 + Avatar (Google Apps Script)
 * Entry point เดียวของ Web App — single-page app ตั้งแต่แรก
 * เสิร์ฟ App.html ไฟล์เดียวเสมอ การสลับ view (หน้าหลัก/Admin/ค้นหา)
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

/**
 * รันครั้งเดียวหลัง setupSheet() — ใส่ค่า default ลง Config และ user ตัวอย่างลง Users
 * เช็คว่า Sheet มีข้อมูลอยู่แล้วหรือยัง (getLastRow() <= 1 คือมีแค่ header)
 * ถ้ามีข้อมูลแล้วจะไม่เขียนทับ
 */
function setupSeedData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- seed Users (photoUrl เว้นว่างไว้ — ให้ระบบ generate avatar ตัวอักษรย่อเอง) ---
  var usersSheet = ss.getSheetByName('Users');
  if (!usersSheet) {
    throw new Error('ไม่พบ Sheet "Users" — กรุณารัน setupSheet() ก่อน');
  }
  if (usersSheet.getLastRow() <= 1) {
    usersSheet.getRange(2, 1, 3, 4).setValues([
      ['u001', 'มานะ ใจดี', 'ฝ่าย IT', ''],
      ['u002', 'สมหญิง รักเรียน', 'ฝ่ายบริหาร', ''],
      ['u003', 'วิชัย ตั้งใจทำงาน', 'ฝ่ายทั่วไป', '']
    ]);
  }

  // --- seed Config ---
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

  Logger.log('setupSeedData() เสร็จสิ้น: Config และ Users มีข้อมูลตัวอย่างพร้อมใช้งาน');
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
 * Response ที่คาดไว้จาก AI (message.content): valid JSON { "skill": "...", "level": "..." }
 */
function callAiGateway_(ocrText) {
  var props = PropertiesService.getScriptProperties();
  var baseUrl = props.getProperty('AI_GATEWAY_BASE_URL');
  var apiKey = props.getProperty('AI_GATEWAY_API_KEY');

  if (!baseUrl || !apiKey) {
    throw new Error('ไม่พบ Script Property "AI_GATEWAY_BASE_URL" หรือ "AI_GATEWAY_API_KEY"');
  }

  var systemPrompt =
    'คุณคือผู้ช่วยวิเคราะห์ทักษะจากข้อความที่ได้ โดยวิเคราะห์หาว่าเป็นทักษะด้านใดใน 3 ทักษะนี้ ' +
    '1.IT 2.บริหาร 3.ทั่วไป และวิเคราะห์ว่าเป็นทักษะระดับใด 1.มหาวิทยาลัย 2.ประเทศ 3.นานาชาติ ' +
    'ตอบกลับเป็น JSON ที่ถูกต้องเท่านั้น รูปแบบ {"skill":"...","level":"..."} ' +
    'โดย skill ต้องเป็นหนึ่งใน "IT", "บริหาร", "ทั่วไป", "วิทยากร" เท่านั้น ' +
    'และ level ต้องเป็นหนึ่งใน "มหาวิทยาลัย", "ประเทศ", "นานาชาติ" เท่านั้น ' +
    'ห้ามมีข้อความอื่นนอกเหนือจาก JSON ห้ามเติมคำอธิบายหรือ markdown code block';

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

  var cleanContent = content.trim().replace(/^```json\s*/i, '').replace(/^```\s*/i, '').replace(/```\s*$/i, '').trim();

  var result;
  try { result = JSON.parse(cleanContent); } catch (e) {
    throw new Error('AI วิเคราะห์ทักษะตอบกลับไม่ใช่ JSON ที่ parse ได้: ' + content);
  }

  if (!result.skill || !result.level) {
    throw new Error('AI วิเคราะห์ทักษะตอบกลับไม่ครบ skill/level: ' + content);
  }

  return {
    skill: String(result.skill).trim(),
    level: String(result.level).trim()
  };
}

/** โหลดค่า baseXP (ตาม level), rarity (ตาม skill), curve (growupLevel) และ specialCurve (specialList) จาก Sheet "Config" */
function loadConfig_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
  if (!sheet) {
    throw new Error('ไม่พบ Sheet "Config" — กรุณารัน setupSheet() และ setupSeedData() ก่อน');
  }

  var data = sheet.getDataRange().getValues();
  var config = { level: {}, rarity: {}, curve: {}, specialCurve: {} };

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