// --- Configuration ---
const SHEET_USERS = "Users";
const SHEET_RECORDS = "TimeRecords";
const SHEET_DUTY = "DutyRoster";

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('FineStamp - ระบบบันทึกเวลาปฏิบัติงาน')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// --- Helper Functions ---
function _formatDateForHtml(val, timezone) {
  if (!val) return "";
  try {
    if (val instanceof Date) return Utilities.formatDate(val, timezone, "yyyy-MM-dd");
    const d = new Date(val);
    if (!isNaN(d.getTime())) return Utilities.formatDate(d, timezone, "yyyy-MM-dd");
  } catch(e) {}
  return String(val);
}

function _formatTimeForHtml(val, timezone) {
  if (!val) return "";
  try {
    if (val instanceof Date) return Utilities.formatDate(val, timezone, "HH:mm");
    let str = String(val).trim();
    const parts = str.split(':');
    if (parts.length >= 2) {
      const h = parts[0].length === 1 ? '0' + parts[0] : parts[0];
      const m = parts[1].length === 1 ? '0' + parts[1] : parts[1];
      return `${h}:${m.substring(0, 2)}`;
    }
  } catch(e) {}
  return String(val);
}

// --- User & Auth ---
function loginUser(userId, password) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_USERS);
  if (!sheet) return JSON.stringify({ isOk: false, message: "System Error: Sheet not found." });

  const data = sheet.getDataRange().getDisplayValues();
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[1] == userId && row[5] == password) {
      return JSON.stringify({
        isOk: true,
        user: {
          id: row[0],
          user_id: row[1],
          first_name: row[2],
          last_name: row[3],
          work_groups: row[4],
          role: row[8] ? row[8].trim() : "User",
          profile_image: row[9] || "" // Column J
        }
      });
    }
  }
  return JSON.stringify({ isOk: false, message: "ID หรือรหัสผ่านไม่ถูกต้อง" });
}

function registerUser(payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_USERS);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_USERS);
      sheet.appendRow(["id", "user_id", "first_name", "last_name", "work_groups", "password", "created_date", "created_time", "role", "profile_image"]);
    }
    const data = JSON.parse(payload);
    const users = sheet.getDataRange().getDisplayValues();
    if (users.some(row => row[1] == data.user_id)) return JSON.stringify({ isOk: false, message: "User ID นี้มีอยู่ในระบบแล้ว" });

    const now = new Date();
    sheet.appendRow([
      data.id, data.user_id, data.first_name, data.last_name, data.work_groups, data.password,
      Utilities.formatDate(now, "GMT+7", "yyyy-MM-dd"),
      Utilities.formatDate(now, "GMT+7", "HH:mm:ss"),
      "User", ""
    ]);
    return JSON.stringify({ isOk: true });
  } catch (e) {
    return JSON.stringify({ isOk: false, message: e.toString() });
  } finally {
    lock.releaseLock();
  }
}

// --- Profile Image Handling ---
function saveProfileImage(userId, base64Data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_USERS);
    const data = sheet.getDataRange().getDisplayValues();
   
    // 1. ค้นหา User และเตรียมลบรูปเก่า (ถ้ามี)
    let rowIndex = -1;
    let oldFileUrl = "";
   
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] == userId) {
        rowIndex = i + 1; // แถวใน Sheet (เริ่มที่ 1)
        oldFileUrl = data[i][9]; // Column J (Index 9)
        break;
      }
    }

    if (rowIndex === -1) return JSON.stringify({ isOk: false, message: "User not found" });

    // ลบไฟล์เก่าทิ้ง (ถ้ามี URL และเป็นไฟล์ใน Drive)
    if (oldFileUrl && oldFileUrl.includes("drive.google.com")) {
      try {
        // ดึง ID จาก URL: https://drive.google.com/thumbnail?id=xxxxx&sz=s400
        const idMatch = oldFileUrl.match(/id=([^&]+)/);
        if (idMatch && idMatch[1]) {
          DriveApp.getFileById(idMatch[1]).setTrashed(true); // ย้ายลงถังขยะ
        }
      } catch (err) {
        // ถ้าหาไฟล์ไม่เจอ หรือลบไม่ได้ ให้ข้ามไป (ไม่ต้อง Error)
        console.log("Delete old file error: " + err);
      }
    }

    // 2. สร้างไฟล์ใหม่
    const folderName = "FineStamp_Profiles";
    const folders = DriveApp.getFoldersByName(folderName);
    let folder;
    if (folders.hasNext()) folder = folders.next();
    else folder = DriveApp.createFolder(folderName);

    const contentType = base64Data.substring(5, base64Data.indexOf(';'));
    const bytes = Utilities.base64Decode(base64Data.substr(base64Data.indexOf('base64,')+7));
    // ตั้งชื่อไฟล์ให้ไม่ซ้ำด้วย Timestamp
    const blob = Utilities.newBlob(bytes, contentType, `profile_${userId}_${Date.now()}.jpg`);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
   
    // 3. อัปเดต URL ใหม่ลง Sheet
    const fileUrl = `https://drive.google.com/thumbnail?id=${file.getId()}&sz=s400`;
    sheet.getRange(rowIndex, 10).setValue(fileUrl); // Column J

    return JSON.stringify({ isOk: true, url: fileUrl });

  } catch(e) {
    return JSON.stringify({ isOk: false, message: e.toString() });
  }
}

// --- Records Management ---
function getUserRecords(userId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_RECORDS);
  if (!sheet) return JSON.stringify([]);
  const data = sheet.getDataRange().getValues();
  const tz = Session.getScriptTimeZone();
  const records = [];
  for (let i = data.length - 1; i >= 1; i--) {
    const row = data[i];
    if (String(row[1]) === String(userId)) {
      records.push({
        id: row[0],
        clock_in_date: _formatDateForHtml(row[5], tz),
        clock_in_time: _formatTimeForHtml(row[6], tz),
        clock_out_time: _formatTimeForHtml(row[8], tz),
        work_type: row[9]
      });
    }
  }
  return JSON.stringify(records);
}

function saveTimeRecord(payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_RECORDS);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_RECORDS);
      sheet.appendRow(["id", "user_id", "first_name", "last_name", "work_groups", "clock_in_date", "clock_in_time", "clock_out_date", "clock_out_time", "work_type", "log_date", "log_time"]);
      sheet.setFrozenRows(1);
    }

    const rec = JSON.parse(payload);
    const allData = sheet.getDataRange().getValues();
    const tz = Session.getScriptTimeZone();
    let rowIndex = -1;
    let existingRow = null;

    for (let i = 1; i < allData.length; i++) {
      if (String(allData[i][0]) === String(rec.id)) {
        rowIndex = i + 1;
        existingRow = allData[i];
        break;
      }
    }

    let finalClockInDate = rec.clock_in_date;
    let finalClockInTime = rec.clock_in_time;
    let finalWorkType = rec.work_type;

    if (rowIndex !== -1 && existingRow) {
        if (!finalClockInDate) finalClockInDate = _formatDateForHtml(existingRow[5], tz);
        if (!finalClockInTime) finalClockInTime = _formatTimeForHtml(existingRow[6], tz);
        if (!finalWorkType) finalWorkType = existingRow[9];
    }

    const now = new Date();
    const rowData = [
      rec.id, rec.user_id, rec.first_name, rec.last_name, rec.work_groups,
      finalClockInDate, finalClockInTime,
      rec.clock_out_date || "", rec.clock_out_time || "",
      finalWorkType || "",
      Utilities.formatDate(now, "GMT+7", "yyyy-MM-dd"),
      Utilities.formatDate(now, "GMT+7", "HH:mm:ss")
    ];
    const stringRowData = rowData.map(d => String(d));

    if (rowIndex !== -1) {
      sheet.getRange(rowIndex, 1, 1, stringRowData.length).setValues([stringRowData]);
    } else {
      sheet.appendRow(stringRowData);
    }

// --- TELEGRAM NOTIFY LOGIC (NEW) ---
    // ตั้งค่าเริ่มต้น
    let icon = "🔵"; // สีน้ำเงิน (เข้างาน)
    let title = "ลงเวลาเข้างาน";
    let nameDisplay = `${rec.first_name} ${rec.last_name}`;
    let workTypeDisplay = rec.work_type || "-";
    let timeInDisplay = rec.clock_in_time ? `${rec.clock_in_time} น.` : "-";
    let timeOutDisplay = "";

    // --- กำหนดเงื่อนไขการแสดงผล ---
    
    // กรณีที่ 1: งาน "เจาะเลือดเช้า" (กรณีพิเศษ: เข้างานแต่มีเวลาออกเลย)
    if (rec.work_type === "เจาะเลือดเช้า") {
        icon = "🟢"; // สีเขียว (ครบทั้งเข้างานและออกงาน)
        title = "ลงเวลาเข้า-ออกงาน";
        timeOutDisplay = rec.clock_out_time ? `${rec.clock_out_time} น.` : "08:00 น.";
    } 
    // กรณีที่ 2: กดออกงานทั่วไป (มีเวลาออก และไม่ใช่งานเจาะเลือดที่เพิ่งกดเข้า)
    else if (rec.clock_out_time && rec.clock_out_time.trim() !== "") {
        icon = "🔴"; // สีแดง (ออกงาน)
        title = "ลงเวลาออกงาน";
        timeOutDisplay = `${rec.clock_out_time} น.`;
    } 
    // กรณีที่ 3: กดเข้างานทั่วไป (เวรแล็บ, HPV - ยังไม่มีเวลาออก)
    else {
        icon = "🔵"; // สีน้ำเงิน (เข้างาน)
        title = "ลงเวลาเข้างาน";
        timeOutDisplay = "<i>..............</i>"; // ข้อความแทนเวลาออก
    }

    // สร้างข้อความ HTML
    const msg = `<b>${icon} ${title}</b>\n` +
                `➖➖➖➖➖➖➖➖➖➖\n` +
                `👤 <b>ชื่อ-สกุล:</b>  ${nameDisplay}\n` +
                `💼 <b>งาน:</b>           ${workTypeDisplay}\n` +
                `🕐 <b>เวลามา:</b>      ${timeInDisplay}\n` +
                `🏁 <b>เวลากลับ:</b>   ${timeOutDisplay}`;
    
    // ส่งเข้า Telegram
    sendTelegramMsg(msg);
    // ------------------------------------

    return JSON.stringify({ isOk: true, record: rec });
  } catch (e) {
    return JSON.stringify({ isOk: false, message: e.toString() });
  } finally {
    lock.releaseLock();
  }
}

// --- DUTY ROSTER FUNCTIONS ---
function getMonthDuty(userId, month, year) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_DUTY);
  if (!sheet) return JSON.stringify([]);

  const data = sheet.getDataRange().getValues();
  const results = [];
 
  // Convert month/year to check
  const targetM = parseInt(month, 10);
  const targetY = parseInt(year, 10);

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (String(row[1]) !== String(userId)) continue; // Check User ID

    // Parse Date
    let dObj = null;
    if (row[2] instanceof Date) dObj = row[2];
    else if (typeof row[2] === 'string') dObj = new Date(row[2]);

    if (dObj && !isNaN(dObj.getTime())) {
      if ((dObj.getMonth() + 1) === targetM && dObj.getFullYear() === targetY) {
        results.push({
          date: Utilities.formatDate(dObj, "GMT+7", "yyyy-MM-dd"),
          shifts: row[3] ? row[3].split(',') : [] // shift_type stored as comma separated
        });
      }
    }
  }
  return JSON.stringify(results);
}

function saveDutyRecord(payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_DUTY);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_DUTY);
      sheet.appendRow(["id", "user_id", "shift_date", "shift_type", "created_at"]);
    }

    const data = JSON.parse(payload); // { user_id, date, shifts: [] }
    const allData = sheet.getDataRange().getValues();
    const targetDateStr = data.date;
   
    // Find existing row for this user & date
    let rowIndex = -1;
    for (let i = 1; i < allData.length; i++) {
      let dStr = "";
      if (allData[i][2] instanceof Date) dStr = Utilities.formatDate(allData[i][2], "GMT+7", "yyyy-MM-dd");
      else dStr = String(allData[i][2]);

      if (String(allData[i][1]) === String(data.user_id) && dStr === targetDateStr) {
        rowIndex = i + 1;
        break;
      }
    }

    const shiftStr = data.shifts.join(',');
    const now = new Date();

    if (rowIndex !== -1) {
      // Update or Delete (if empty)
      if (data.shifts.length === 0) {
        sheet.deleteRow(rowIndex);
      } else {
        sheet.getRange(rowIndex, 4).setValue(shiftStr); // Update shift_type
        sheet.getRange(rowIndex, 5).setValue(now); // Update timestamp
      }
    } else {
      // Create new
      if (data.shifts.length > 0) {
        sheet.appendRow([
          'D_' + Date.now(),
          data.user_id,
          targetDateStr,
          shiftStr,
          now
        ]);
      }
    }
    return JSON.stringify({ isOk: true });
  } catch (e) {
    return JSON.stringify({ isOk: false, message: e.toString() });
  } finally {
    lock.releaseLock();
  }
}

// --- REPORT FUNCTIONS ---
function getUsersByGroup(groupName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_USERS);
  if (!sheet) return JSON.stringify([]);
 
  const data = sheet.getDataRange().getDisplayValues();
  const users = [];
 
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const groups = row[4];
    if (groups && groups.includes(groupName)) {
      users.push({
        user_id: row[1],
        name: `${row[2]} ${row[3]}`
      });
    }
  }
  users.sort((a, b) => a.name.localeCompare(b.name, 'th'));
  return JSON.stringify(users);
}

function getMonthlyReport(monthStr, yearStr, targetGroup, targetUserId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_RECORDS);
    if (!sheet) return JSON.stringify({ isOk: false, message: "ไม่พบ Sheet TimeRecords" });

    const data = sheet.getDataRange().getValues();
    const tz = Session.getScriptTimeZone();
    const targetMonth = parseInt(monthStr, 10);
    const targetYear = parseInt(yearStr, 10);

    const result = [];
    const customOrderAP = ["รุ่งตะวัน", "ปวรวรรชน์", "พิสิฏฐ์", "ธนภรณ์", "บุษบา", "รัชนี"];
    const thaiMonthsShort = ["ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.", "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."];

    const GROUP_AP = "พยาธิวิทยากายวิภาค";
    const GROUP_CP = "พยาธิวิทยาคลินิกและเทคนิคการแพทย์";

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      let dObj = null;
      if (row[5] instanceof Date) { dObj = row[5]; }
      else if (typeof row[5] === 'string' && row[5].trim() !== "") { dObj = new Date(row[5]); }
     
      if (!dObj || isNaN(dObj.getTime())) continue;
     
      const m = dObj.getMonth() + 1;
      const y = dObj.getFullYear();
      if (m !== targetMonth || y !== targetYear) continue;

      const recordUserId = String(row[1]);
      if (targetUserId && targetUserId !== "ALL") {
        if (recordUserId !== targetUserId) continue;
      }

      const workType = row[9] ? String(row[9]).trim() : "";
      const recordGroup = row[4] ? String(row[4]).trim() : "";
     
      let includeRecord = false;

      if (targetGroup === GROUP_CP) {
        // CP: เอา "เจาะเลือดเช้า", "เวรแล็บ" หรือ "ค่าว่าง" (งานปกติ)
        if (workType === "เจาะเลือดเช้า" || workType === "เวรแล็บ") {
          includeRecord = true;
        } else if (workType === "" && recordGroup.includes(GROUP_CP)) {
          includeRecord = true;
        }
      } else if (targetGroup === GROUP_AP) {
        // AP: เอา "เวร HPV" หรือ "ค่าว่าง" (งานปกติ)
        if (workType === "เวร HPV") {
          includeRecord = true;
        } else if (workType === "" && recordGroup.includes(GROUP_AP)) {
          includeRecord = true;
        }
      }

      if (includeRecord) {
        const day = dObj.getDate();
        const monthIndex = dObj.getMonth();
        const yearBE = dObj.getFullYear() + 543;
        const dateThaiStr = `${day} ${thaiMonthsShort[monthIndex]} ${yearBE}`;
       
        result.push({
          dateObj: dObj,
          dateDisplay: dateThaiStr,
          fullName: `${row[2]} ${row[3]}`,
          firstName: String(row[2]).trim(),
          timeIn: _formatTimeForHtml(row[6], tz),
          timeOut: _formatTimeForHtml(row[8], tz),
          workType: workType
        });
      }
    }

    result.sort((a, b) => {
      if (a.dateObj.getTime() !== b.dateObj.getTime()) {
        return a.dateObj.getTime() - b.dateObj.getTime();
      }
      if (targetGroup === GROUP_AP) {
        let idxA = customOrderAP.indexOf(a.firstName);
        let idxB = customOrderAP.indexOf(b.firstName);
        if (idxA === -1) idxA = 999;
        if (idxB === -1) idxB = 999;
        return idxA - idxB;
      } else {
        if (a.timeIn < b.timeIn) return -1;
        if (a.timeIn > b.timeIn) return 1;
        return 0;
      }
    });

    const finalOutput = result.map(r => ({
      date: r.dateDisplay,
      name: r.fullName,
      in: r.timeIn,
      out: r.timeOut
    }));

    return JSON.stringify({ isOk: true, data: finalOutput });

  } catch (e) {
    return JSON.stringify({ isOk: false, message: e.toString() });
  }
}

// --- User Profile Management (เพิ่มส่วนนี้ต่อท้ายไฟล์) ---
function changePassword(userId, newPassword) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_USERS);
    if (!sheet) return JSON.stringify({ isOk: false, message: "Sheet User ไม่พบ" });

    const data = sheet.getDataRange().getValues();
    let rowIndex = -1;

    // ค้นหา User (Column B คือ user_id)
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]) === String(userId)) {
        rowIndex = i + 1; // Row ใน Sheet เริ่มที่ 1
        break;
      }
    }

    if (rowIndex !== -1) {
      // อัปเดตรหัสผ่าน (Column F คือ index 6)
      sheet.getRange(rowIndex, 6).setValue(newPassword);
      return JSON.stringify({ isOk: true });
    } else {
      return JSON.stringify({ isOk: false, message: "ไม่พบผู้ใช้งานนี้" });
    }

  } catch (e) {
    return JSON.stringify({ isOk: false, message: e.toString() });
  } finally {
    lock.releaseLock();
  }
}

// --- TELEGRAM FUNCTION (วางไว้ล่างสุดของไฟล์ Code.gs) ---
function sendTelegramMsg(message) {
  // ************************************************
  // ใส่ค่าของคุณตรงนี้ (อย่าลืมเครื่องหมายลบที่ Chat ID)
  const token = "XXXXX"; // <--- แก้ตรงนี้
  const chatId = "XXXXX"; // <--- แก้ตรงนี้
  // ************************************************

  const url = "https://api.telegram.org/bot" + token + "/sendMessage";
  const options = {
    "method": "post",
    "payload": {
      "chat_id": chatId,
      "text": message,
      "parse_mode": "HTML"
    },
    "muteHttpExceptions": true
  };
  try { UrlFetchApp.fetch(url, options); } catch (e) { console.log(e); }
}
