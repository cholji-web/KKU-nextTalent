/**
 * ==================================================================
 * NextTalent Prototype — Part 1 (Google Apps Script)
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
 */
function setupSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var usersSheet = ss.getSheetByName('Users');
  if (!usersSheet) {
    usersSheet = ss.insertSheet('Users');
    usersSheet.appendRow(['userId', 'name', 'department']);
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

/* ==================================================================
 * PUBLIC API — ฟังก์ชันที่ client เรียกผ่าน google.script.run
 * ================================================================== */

/** คืนรายชื่อ user ทั้งหมดจาก Sheet "Users" สำหรับ dropdown เลือก user */
function getUsersList() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  if (!sheet) {
    throw new Error('ไม่พบ Sheet "Users" — กรุณารัน setupSheet() ก่อน');
  }

  var data = sheet.getDataRange().getValues();
  var users = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    users.push({ userId: data[i][0], name: data[i][1], department: data[i][2] });
  }
  return users;
}

/**
 * คืนข้อมูลสรุปของ user คนหนึ่ง: ชื่อ, หน่วยงาน, growupLevel (XP/level/badge ภาพรวมทุก skill),
 * specialList (XP/level แยกตาม skill/rarity คนละก้อน), และรายการที่ส่งผลทั้งหมด
 * ใช้แสดงการ์ดข้อมูล user และรายการรออนุมัติในหน้า UI
 */
/**
 * คืนข้อมูลสรุปของ user คนหนึ่ง — ครอบด้วย try/catch เพื่อให้ error จริงโผล่ไปที่
 * failureHandler ฝั่ง client แทนที่จะ silent-fail กลายเป็น null
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
      growupLevel: growupLevel,
      specialList: specialList,
      submissions: submissions
    };

    // ทดสอบ serialize ก่อนส่งกลับจริง — ถ้าพังจะ throw ตรงนี้เลย เห็น error ชัดเจน
    JSON.stringify(result);

    return result;

  } catch (err) {
    Logger.log('getUserState() error: ' + err.message + '\n' + err.stack);
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
    throw new Error('ไม่พบค่า rarity/baseXP ของ skill="' + aiResult.skill + '" level="' + aiResult.level + '" ใน Config sheet');
  }

  return {
    fileUrl: fileUrl,
    ocrText: extractedText,
    skill: aiResult.skill,
    level: aiResult.level,
    rarity: rarity,
    xp: baseXp * rarity
  };
}

/**
 * บันทึกผลที่ user กด "ส่งผล" ลง Sheet "Submissions" (status = submitted)
 * รับผลลัพธ์ที่ได้จาก analyzeCertificate() กลับมาจาก client (ยังไม่เคยเซฟมาก่อน)
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
    result.rarity,
    result.xp,
    'submitted'
  ]);

  return getUserState(userId);
}

/* ==================================================================
 * INTERNAL HELPERS — ฟังก์ชันภายใน (ไม่เรียกจาก client)
 * ================================================================== */

/** ดึงรายการ submissions ของ user คนหนึ่งจาก Sheet "Submissions" เรียงล่าสุดก่อน */
/** ดึงรายการ submissions ของ user คนหนึ่งจาก Sheet "Submissions" เรียงล่าสุดก่อน
 *  sanitize ค่าทุกฟิลด์ให้ปลอดภัยต่อการ serialize กลับไป client (กัน NaN/undefined/Date แปลกๆ)
 */
function getSubmissionsByUser_(userId) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Submissions');
  if (!sheet) {
    throw new Error('ไม่พบ Sheet "Submissions" — กรุณารัน setupSheet() ก่อน');
  }

  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var userIdCol = headers.indexOf('userId');

  return data
    .slice(1)
    .filter(function (row) { return row[userIdCol] === userId; })
    .map(function (row) {
      var obj = {};
      headers.forEach(function (h, i) { obj[h] = row[i]; });

      // --- sanitize: กันค่าที่ทำให้ JSON serialize พังหรือคำนวณเพี้ยน ---
      obj.timestamp = obj.timestamp instanceof Date
        ? obj.timestamp.toISOString()
        : String(obj.timestamp || '');

      obj.xp = Number(obj.xp);
      if (isNaN(obj.xp)) obj.xp = 0;

      obj.rarity = Number(obj.rarity);
      if (isNaN(obj.rarity)) obj.rarity = 1;

      obj.skill = String(obj.skill || '');
      obj.level = String(obj.level || '');
      obj.status = String(obj.status || 'submitted');
      obj.fileUrl = String(obj.fileUrl || '');
      obj.ocrText = String(obj.ocrText || '');
      obj.id = String(obj.id || '');

      return obj;
    })
    .sort(function (a, b) { return new Date(b.timestamp) - new Date(a.timestamp); });
}

function generateId_() {
  return 'sub_' + Utilities.getUuid();
}

/** เซฟไฟล์ PDF ลง Google Drive คืนค่าเป็นลิงก์ไฟล์ */
function saveToDrive_(blob) {
  var folder = getOrCreateUploadFolder_();
  return folder.createFile(blob).getUrl();
}

function getOrCreateUploadFolder_() {
  var folderName = 'NextTalent Uploads';
  var folders = DriveApp.getFoldersByName(folderName);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
}

/**
 * ส่งไฟล์ PDF เข้า n8n webhook เพื่อถอดข้อความ (OCR)
 * Request: multipart/form-data, field name "file"
 * Response ที่คาดไว้: JSON { "status": "...", "extracted_text": "..." }
 * Response อาจไม่ใช่ JSON เสมอไป (error จาก n8n) — parse แบบ try/catch แล้ว fallback แสดง raw text
 */
function callOcrWebhook_(fileBlob) {
  var url = PropertiesService.getScriptProperties().getProperty('OCR_WEBHOOK_URL');
  if (!url) {
    throw new Error('ไม่พบ Script Property "OCR_WEBHOOK_URL"');
  }

  var options = {
    method: 'post',
    payload: { file: fileBlob }, // UrlFetchApp ส่งเป็น multipart/form-data อัตโนมัติเมื่อ payload มี Blob
    muteHttpExceptions: true
  };

  var response = UrlFetchApp.fetch(url, options);
  var status = response.getResponseCode();
  var rawText = response.getContentText();

  var parsed = null;
  try { parsed = JSON.parse(rawText); } catch (e) { /* ไม่ใช่ JSON — ปล่อย parsed = null เพื่อ fallback แสดง raw text */ }

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
    '1.IT 2.บริหาร 3.ทั่วไป และวิเคราะห์ว่าเป็นทักษะระดับใด 1.มหาวิทยาลัย 2.ระดับประเทศ 3.นานาชาติ ' +
    'ตอบกลับเป็น JSON ที่ถูกต้องเท่านั้น รูปแบบ {"skill":"...","level":"..."} ห้ามมีข้อความอื่นนอกเหนือจาก JSON';

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

/** โหลดค่า baseXP (ตาม level), rarity (ตาม skill), curve (growupLevel) และ specialCurve (specialList) จาก Sheet "Config" */
function loadConfig_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
  if (!sheet) {
    throw new Error('ไม่พบ Sheet "Config" — กรุณารัน setupSheet() และ setupSeedData() ก่อน');
  }

  var data = sheet.getDataRange().getValues();
  var config = { level: {}, rarity: {}, curve: {}, specialCurve: {} };

  for (var i = 1; i < data.length; i++) {
    var category = data[i][0];
    var key = data[i][1];
    var value = data[i][2];
    if (config[category] !== undefined) {
      config[category][key] = value;
    }
  }
  return config;
}

/** คำนวณ level และ XP ที่เข้า/ต้องการของ level ปัจจุบัน จาก XP รวม ตามสูตร base × N^exponent */
function calcLevelFromXp_(totalXp, base, exponent) {
  base = base || 100;
  exponent = exponent || 1.5;

  var level = 1;
  while (xpRequiredForLevel_(level + 1, base, exponent) <= totalXp) {
    level++;
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

/** คืนชื่อ badge ตาม level: 1-4 Rookie, 5-9 Rising Star, 10-19 Expert, 20+ Master (ฝั่ง client วาดเป็น SVG เอง ดู BADGE_TIERS) */
function getBadge_(level) {
  if (level >= 20) return { name: 'Master' };
  if (level >= 10) return { name: 'Expert' };
  if (level >= 5) return { name: 'Rising Star' };
  return { name: 'Rookie' };
}