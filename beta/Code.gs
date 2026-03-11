// PlugnGO Beta 序號發放系統
// 部署為 Google Apps Script Web App
// 試算表結構：
//   工作表1「序號庫」：A欄=序號, B欄=狀態(空白=未使用, "USED"=已使用)
//   工作表2「申請紀錄」：A=時間戳記, B=姓名, C=Email, D=拍攝類型, E=序號

const SHEET_ID = 'YOUR_GOOGLE_SHEET_ID'; // 替換成你的 Google Sheets ID
const CODES_SHEET = '序號庫';
const RECORDS_SHEET = '申請紀錄';
const SENDER_NAME = 'PlugnGO';
const FACEBOOK_URL = 'https://www.facebook.com/skiseiju';

function doPost(e) {
  const headers = {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'POST',
    'Access-Control-Allow-Headers': 'Content-Type',
    'Content-Type': 'application/json'
  };

  try {
    const data = JSON.parse(e.postData.contents);
    const { name, email, type } = data;

    // 基本驗證
    if (!name || !email || !type) {
      return ContentService
        .createTextOutput(JSON.stringify({ success: false, error: 'MISSING_FIELDS' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // Email 格式驗證
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (!emailRegex.test(email)) {
      return ContentService
        .createTextOutput(JSON.stringify({ success: false, error: 'INVALID_EMAIL' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const ss = SpreadsheetApp.openById(SHEET_ID);
    const recordsSheet = ss.getSheetByName(RECORDS_SHEET);
    const codesSheet = ss.getSheetByName(CODES_SHEET);

    // 檢查是否已申請過
    const records = recordsSheet.getDataRange().getValues();
    for (let i = 1; i < records.length; i++) {
      if (records[i][2] && records[i][2].toString().toLowerCase() === email.toLowerCase()) {
        return ContentService
          .createTextOutput(JSON.stringify({ success: false, error: 'ALREADY_REGISTERED' }))
          .setMimeType(ContentService.MimeType.JSON);
      }
    }

    // 取得可用序號
    const codes = codesSheet.getDataRange().getValues();
    let availableRow = -1;
    let availableCode = '';

    for (let i = 1; i < codes.length; i++) {
      if (codes[i][0] && !codes[i][1]) {
        availableRow = i + 1; // Sheets 從 1 開始
        availableCode = codes[i][0].toString();
        break;
      }
    }

    if (availableRow === -1) {
      return ContentService
        .createTextOutput(JSON.stringify({ success: false, error: 'NO_CODES_LEFT' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 標記序號為已使用
    codesSheet.getRange(availableRow, 2).setValue('USED');

    // 寫入申請紀錄
    recordsSheet.appendRow([
      new Date(),
      name,
      email,
      type,
      availableCode
    ]);

    // 寄送 Email
    sendBetaEmail(name, email, availableCode);

    return ContentService
      .createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: 'SERVER_ERROR', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  // 處理 CORS preflight
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'PlugnGO Beta API is running' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function sendBetaEmail(name, email, code) {
  const subject = '你的 PlugnGO Beta 序號來了 🎉';
  const body = `嗨 ${name}，

感謝你申請 PlugnGO Beta 試用！

你的專屬序號：${code}

有效期限：三個月

使用過程有任何問題，直接回覆這封信就可以找到我，或到我的 Facebook 隨意留言也可以。
${FACEBOOK_URL}

海哥
PlugnGO`;

  const htmlBody = `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <style>
    body { font-family: -apple-system, sans-serif; background: #f5f5f5; margin: 0; padding: 40px 20px; }
    .container { max-width: 520px; margin: 0 auto; background: #fff; border-radius: 12px; overflow: hidden; }
    .header { background: #0a0a0a; padding: 32px; text-align: center; }
    .header h1 { color: #fff; margin: 0; font-size: 22px; letter-spacing: 0.05em; }
    .header span { color: #888; font-size: 13px; }
    .body { padding: 36px 32px; }
    .greeting { font-size: 16px; color: #333; margin-bottom: 24px; }
    .code-box { background: #f0f0f0; border: 2px dashed #ccc; border-radius: 8px; padding: 20px; text-align: center; margin: 28px 0; }
    .code-label { font-size: 12px; color: #888; letter-spacing: 0.1em; text-transform: uppercase; margin-bottom: 8px; }
    .code { font-size: 24px; font-weight: 700; color: #0a0a0a; letter-spacing: 0.08em; font-family: monospace; }
    .validity { font-size: 13px; color: #888; margin-top: 8px; }
    .message { font-size: 15px; color: #555; line-height: 1.7; margin-bottom: 24px; }
    .fb-link { color: #0a0a0a; font-weight: 600; }
    .footer { border-top: 1px solid #eee; padding: 20px 32px; font-size: 13px; color: #aaa; }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <h1>PlugnGO</h1>
      <span>Beta 試用邀請</span>
    </div>
    <div class="body">
      <p class="greeting">嗨 ${name}，</p>
      <p class="message">感謝你申請 PlugnGO Beta 試用！你的專屬序號在這裡：</p>
      <div class="code-box">
        <div class="code-label">專屬序號</div>
        <div class="code">${code}</div>
        <div class="validity">有效期限：三個月</div>
      </div>
      <p class="message">
        使用過程有任何問題，直接回覆這封信就可以找到我，或到我的 Facebook 隨意留言也可以。<br><br>
        <a href="${FACEBOOK_URL}" class="fb-link">facebook.com/skiseiju</a>
      </p>
      <p class="message">海哥<br>PlugnGO</p>
    </div>
    <div class="footer">這是自動寄出的信件，有問題請直接回覆。</div>
  </div>
</body>
</html>`;

  GmailApp.sendEmail(email, subject, body, {
    htmlBody: htmlBody,
    name: SENDER_NAME,
    replyTo: 'plugngo@gmail.com'
  });
}

// ===== 工具函式：一次性執行，產生 1000 組序號寫入試算表 =====
function generateCodes() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(CODES_SHEET);

  if (!sheet) {
    sheet = ss.insertSheet(CODES_SHEET);
  }

  // 清空並加標題
  sheet.clearContents();
  sheet.getRange(1, 1).setValue('序號');
  sheet.getRange(1, 2).setValue('狀態');

  const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789'; // 移除易混淆字元 O, I, 0, 1
  const codes = new Set();

  while (codes.size < 300) {
    let code = 'PNG-BETA-';
    for (let i = 0; i < 6; i++) {
      code += chars[Math.floor(Math.random() * chars.length)];
    }
    codes.add(code);
  }

  const rows = Array.from(codes).map(c => [c, '']);
  sheet.getRange(2, 1, rows.length, 2).setValues(rows);

  Logger.log('✅ 已產生 1000 組序號');
}

// ===== 工具函式：初始化申請紀錄工作表 =====
function setupRecordsSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(RECORDS_SHEET);

  if (!sheet) {
    sheet = ss.insertSheet(RECORDS_SHEET);
  }

  sheet.getRange(1, 1, 1, 5).setValues([['時間戳記', '姓名', 'Email', '拍攝類型', '序號']]);
  Logger.log('✅ 申請紀錄工作表已建立');
}
