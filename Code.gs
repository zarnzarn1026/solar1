/**
 * =========================================================================
 * SOLAR1ENERGY - O&M MAINTENANCE SYSTEM BACKEND (Code.gs) - FULL VERSION
 * รองรับ:
 * 1. จัดการใบงาน MaintenanceTracker (ครบ 15 คอลัมน์ A-O)
 * 2. ส่งแจ้งเตือนจริงทั้ง Gmail (HTML Email) และ LINE Messaging API (Push Message)
 * 3. Webhook ดักเก็บ LINE User ID / Group ID ลงแท็บ LINE_Users อัตโนมัติ
 * =========================================================================
 */
const CONFIG = {
  SHEET_NAME: "MaintenanceTracker",
  USER_SHEET: "LINE_Users",
  ADMIN_EMAIL: "admin@solar1energy.com",
  // 👇 นำ Channel access token (long-lived) จากแท็บ Messaging API มาวางในเครื่องหมายคำพูดด้านล่างนี้
  LINE_CHANNEL_ACCESS_TOKEN: "17x155M/m0ryhRisgJA7nvR1mf+YgAyiXyF3wqYCBUHyVswple1aRrHuAay6R4fIVrdPbBy4tOxpCZLe/6DECdrsJbcTu/1yd5K6zMrBukXMomApF9IBdohT9mY83TVxpEojLGLCe8O1uM/k48UFYwdB04t89/1O/w1cDnyilFU="
};

/**
 * ฟังก์ชันช่วยล้างช่องว่าง คำว่า Bearer หรือการขึ้นบรรทัดใหม่ที่อาจติดมากับ Token
 */
function getCleanLineToken_() {
  return String(CONFIG.LINE_CHANNEL_ACCESS_TOKEN || "17x155M/m0ryhRisgJA7nvR1mf+YgAyiXyF3wqYCBUHyVswple1aRrHuAay6R4fIVrdPbBy4tOxpCZLe/6DECdrsJbcTu/1yd5K6zMrBukXMomApF9IBdohT9mY83TVxpEojLGLCe8O1uM/k48UFYwdB04t89/1O/w1cDnyilFU=")
    .replace(/^Bearer\s+/i, "")
    .replace(/[\r\n\s]+/g, "")
    .trim();
}

function getTargetSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAME);
  }

  const headers = [
    "Job ID",
    "Report Date",
    "Customer Name",
    "Customer Email",
    "Phone",
    "Issue",
    "Status",
    "Technician Name",
    "Technician Email",
    "LINE Tech ID",
    "LINE Customer ID",
    "Repair Result",
    "Before Image",
    "After Image",
    "System Log"
  ];

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  } else {
    const firstRowValues = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), headers.length)).getValues()[0];
    const hasHeader = firstRowValues.some(v => String(v || "").trim() !== "");
    if (!hasHeader) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    } else if (sheet.getLastColumn() < headers.length) {
      sheet.insertColumnsAfter(sheet.getLastColumn(), headers.length - sheet.getLastColumn());
    }
  }

  return sheet;
}

function normalizeJobRow_(row) {
  const normalized = Array.isArray(row) ? row.slice(0, 15) : [];
  while (normalized.length < 15) {
    normalized.push("");
  }
  return normalized;
}

/**
 * [GET] ดึงข้อมูลใบงานทั้งหมดไปแสดงบนหน้าเว็บ
 */
function doGet(e) {
  try {
    const sheet = getTargetSheet_();
    const lastRow = sheet.getLastRow();
    const jobs = lastRow > 1
      ? sheet.getRange(2, 1, lastRow - 1, 15).getValues()
          .filter(r => String(r[0]).trim() !== "")
          .map((r) => {
            let dateStr = String(r[1] || '');
            if (dateStr.includes('GMT')) {
              const d = new Date(dateStr);
              if (!isNaN(d.getTime())) {
                dateStr = `${String(d.getDate()).padStart(2, '0')}/${String(d.getMonth() + 1).padStart(2, '0')}/${d.getFullYear()}`;
              }
            }
            return {
              jobId: String(r[0] || ''),
              reportDate: dateStr,
              customerName: String(r[2] || ''),
              customerEmail: String(r[3] || ''),
              phone: String(r[4] || '').replace(/^'/, ''),
              issue: String(r[5] || ''),
              status: String(r[6] || 'รอดำเนินการ'),
              techName: String(r[7] || ''),
              techEmail: String(r[8] || ''),
              lineTech: String(r[9] || ''),
              lineCustomer: String(r[10] || ''),
              repairResult: String(r[11] || ''),
              beforeImg: String(r[12] || ''),
              afterImg: String(r[13] || ''),
              systemLog: String(r[14] || '')
            };
          })
      : [];

    const responseObj = { status: "success", count: jobs.length, data: jobs };
    const callback = e && e.parameter && e.parameter.callback;
    if (callback) {
      return ContentService.createTextOutput(callback + "(" + JSON.stringify(responseObj) + ")")
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService.createTextOutput(JSON.stringify(responseObj))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * [POST] รับคำสั่งจากทั้ง LINE Webhook และจากหน้าเว็บ Solar1Energy O&M
 */
function doPost(e) {
  try {
    const postData = e.postData.contents;
    const data = JSON.parse(postData);

    // ==========================================================
    // ส่วนที่ 1: ตรวจจับว่าเป็น Webhook จาก LINE หรือไม่ (ดักเก็บ User/Group ID)
    // ==========================================================
    if (data.events && Array.isArray(data.events)) {
      const event = data.events[0];
      if (event && event.source) {
        let targetId = "";
        let sourceType = event.source.type; // 'user', 'group', 'room'

        if (sourceType === 'user') targetId = event.source.userId;
        else if (sourceType === 'group') targetId = event.source.groupId;
        else if (sourceType === 'room') targetId = event.source.roomId;

        if (targetId) {
          const displayName = getLineProfileName_(targetId, sourceType);
          saveLineUserToSheet_(targetId, displayName, sourceType);

          if (event.type === 'message' && event.message && event.message.type === 'text') {
            replyLineMessage_(event.replyToken, `สวัสดีคุณ ${displayName}\nระบบ Solar1Energy O&M บันทึกข้อมูลของคุณแล้ว\nLINE ID ของคุณคือ:\n${targetId}`);
          }
        }
      }
      return ContentService.createTextOutput(JSON.stringify({ status: "ok" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // ==========================================================
    // ส่วนที่ 2: คำสั่งส่งแจ้งเตือน Email & LINE จากหน้าเว็บ
    // ==========================================================
    if (data.action === 'sendNotification') {
      const resNotify = handleSendNotification_(data);
      return ContentService.createTextOutput(JSON.stringify(resNotify))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // ==========================================================
    // ส่วนที่ 3: คำสั่งสร้างและอัปเดตใบงานลงแท็บ MaintenanceTracker
    // ==========================================================
    if (data.action === 'createJob' || data.action === 'updateJob') {
      const sheet = getTargetSheet_();
      let row = normalizeJobRow_(data.rowValues || []);
      if (row[1] && String(row[1]).includes('GMT')) {
        const d = new Date(row[1]);
        if (!isNaN(d.getTime())) {
          row[1] = `${String(d.getDate()).padStart(2, '0')}/${String(d.getMonth() + 1).padStart(2, '0')}/${d.getFullYear()}`;
        }
      }

      const lastRow = sheet.getLastRow();
      let targetRow = -1;
      if (lastRow > 1) {
        const jobIds = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
        for (let i = 0; i < jobIds.length; i++) {
          if (String(jobIds[i][0]).trim() === String(row[0]).trim()) {
            targetRow = i + 2;
            break;
          }
        }
      }

      if (targetRow > 0) {
        sheet.getRange(targetRow, 1, 1, row.length).setValues([row]);
      } else {
        sheet.appendRow(row);
      }

      return ContentService.createTextOutput(JSON.stringify({ status: "success", message: "Saved successfully" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "Invalid action: " + data.action }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * ฟังก์ชันจัดการการส่ง Email และ LINE พร้อมตรวจสอบสถานะจริง
 */
function handleSendNotification_(data) {
  const emailTo = (data.customEmailTo || "").trim();
  const emailSubject = (data.customEmailSubject || "แจ้งสถานะงานซ่อมบำรุง Solar1Energy").trim();
  const emailHtmlBody = data.customEmailBody || "";
  const lineTo = (data.customLineTo || "").trim();
  const lineMessage = (data.customLineMessage || "").trim();

  let sentChannels = [];
  let errors = [];

  // 1. ส่งอีเมลผ่าน GmailApp
  if (emailTo && emailTo.includes('@')) {
    try {
      GmailApp.sendEmail(emailTo, emailSubject, "", {
        htmlBody: `<div style="font-family: Sarabun, sans-serif; font-size: 14px; color: #0F172A; line-height: 1.6;">${emailHtmlBody}</div>`
      });
      sentChannels.push("Email (" + emailTo + ")");
    } catch (e) {
      errors.push("Email Error: " + e.message);
    }
  }

  // 2. ส่งข้อความผ่าน LINE Messaging API
  if (lineTo) {
    const lineRes = sendLineMessagingApi_(lineTo, lineMessage);
    if (lineRes.success) {
      sentChannels.push("LINE (" + lineTo + ")");
    } else {
      errors.push("LINE Error: " + lineRes.error);
    }
  }

  if (errors.length > 0) {
    return { status: "error", message: errors.join(" | ") };
  }

  if (sentChannels.length === 0) {
    return { status: "error", message: "กรุณาระบุอีเมลผู้รับ หรือระบุ LINE Target" };
  }

  return { status: "success", detail: sentChannels.join(", ") };
}

/**
 * ฟังก์ชันยิง Push Message เข้า LINE Messaging API
 */
function sendLineMessagingApi_(toTarget, message) {
  const token = getCleanLineToken_();
  if (!token || token.includes("วาง_")) {
    return { success: false, error: "กรุณานำ Channel Access Token (แบบยาว) มาใส่ในบรรทัดที่ 11 ของ Code.gs ก่อนครับ" };
  }

  try {
    const resLine = UrlFetchApp.fetch("https://api.line.me/v2/bot/message/push", {
      method: "post",
      headers: {
        "Content-Type": "application/json",
        "Authorization": "Bearer " + token
      },
      payload: JSON.stringify({
        to: toTarget.trim(),
        messages: [{ type: "text", text: message }]
      }),
      muteHttpExceptions: true
    });

    const code = resLine.getResponseCode();
    const body = resLine.getContentText();
    Logger.log("LINE HTTP Code: " + code + " | Response: " + body);

    if (code === 200) {
      return { success: true };
    } else {
      return { success: false, error: `HTTP ${code}: ${body}` };
    }
  } catch (err) {
    return { success: false, error: err.message };
  }
}

/**
 * บันทึกรายชื่อและ LINE ID ลงในแท็บ LINE_Users อัตโนมัติ
 */
function saveLineUserToSheet_(targetId, displayName, type) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.USER_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.USER_SHEET);
    sheet.appendRow(["Timestamp", "User/Group ID", "Display Name (ชื่อไลน์)", "Type (user / group)"]);
  }

  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const existingIds = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
    for (let i = 0; i < existingIds.length; i++) {
      if (String(existingIds[i][0]).trim() === String(targetId).trim()) return;
    }
  }

  const now = new Date();
  const timestamp = `${String(now.getDate()).padStart(2, '0')}/${String(now.getMonth() + 1).padStart(2, '0')}/${now.getFullYear()} ${String(now.getHours()).padStart(2, '0')}:${String(now.getMinutes()).padStart(2, '0')}`;
  sheet.appendRow([timestamp, targetId, displayName, type]);
}

function getLineProfileName_(targetId, type) {
  const token = getCleanLineToken_();
  if (!token || type !== 'user') return type === 'group' ? "LINE Group" : "LINE User";

  try {
    const res = UrlFetchApp.fetch(`https://api.line.me/v2/bot/profile/${targetId}`, {
      method: "get",
      headers: { "Authorization": "Bearer " + token },
      muteHttpExceptions: true
    });
    if (res.getResponseCode() === 200) {
      const profile = JSON.parse(res.getContentText());
      return profile.displayName || "LINE User";
    }
  } catch (e) {}
  return "LINE User";
}

function replyLineMessage_(replyToken, message) {
  const token = getCleanLineToken_();
  if (!token || !replyToken) return;

  try {
    UrlFetchApp.fetch("https://api.line.me/v2/bot/message/reply", {
      method: "post",
      headers: {
        "Content-Type": "application/json",
        "Authorization": "Bearer " + token
      },
      payload: JSON.stringify({
        replyToken: replyToken,
        messages: [{ type: "text", text: message }]
      }),
      muteHttpExceptions: true
    });
  } catch (e) {}
}

/**
 * 🧪 ฟังก์ชันที่ 1: กด Run เพื่อทดสอบส่งข้อความเข้า LINE ของคุณโดยตรง
 */
function testSendLineNow() {
  const myUserId = "U9ba020cd34e45a4bbe05776a9c42d82e";
  const result = sendLineMessagingApi_(myUserId, "🔔 ทดสอบการส่งแจ้งเตือนจากระบบ Solar1Energy O&M สำเร็จแล้ว!");
  if (result.success) {
    Logger.log("✅ ส่งเข้า LINE สำเร็จ (HTTP 200)! ข้อความเด้งเข้ามือถือของคุณแล้ว");
  } else {
    Logger.log("❌ ส่งไม่ผ่าน สาเหตุ: " + result.error);
  }
}

/**
 * 🧪 ฟังก์ชันที่ 2: กด Run เพื่อเปิดสิทธิ์และทดสอบส่ง Gmail
 */
function testAuthorizeEmail() {
  const me = Session.getActiveUser().getEmail();
  GmailApp.sendEmail(me, "[Solar1Energy] ทดสอบระบบส่งอีเมลสำเร็จ", "ระบบ O&M ของคุณพร้อมส่งอีเมลแจ้งเตือนแล้ว");
  Logger.log("✅ ส่งเมลทดสอบไปที่ " + me + " สำเร็จ!");
}
