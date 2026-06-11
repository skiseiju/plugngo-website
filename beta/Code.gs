/**
 * PlugnGO License Management Server (Google Apps Script)
 */

const SS_ID = '1AZzLIgU8EJGmHG7k0_DAygrahooHmE94XNGWo9G9Rqg';

const SHEET_LICENSES = "Licenses";
const SHEET_TRIALS   = "Trials";

const BETA_CODES_SHEET   = "Beta_Codes";
const BETA_SIGNUPS_SHEET = "Beta_Signups";
const BETA_SENDER_NAME   = "PlugnGO";
const BETA_REPLY_TO      = "plugngo@gmail.com";
const BETA_FACEBOOK_URL  = "https://www.facebook.com/skiseiju";
const ADMIN_SECRET_PROPERTY = "PLUGNGO_ADMIN_SECRET";
const LICENSE_PRIVATE_KEY_PROPERTY = "PLUGNGO_LICENSE_PRIVATE_KEY";
const LICENSE_CERT_TTL_DAYS = 30;
const VALID_LICENSE_TIERS = ["photo", "video", "hub"];
const VALID_LICENSE_PLUGINS = ["video_proxy", "ai_vision_p1", "audio_ai"];

const COL_LICENSE_KEY = 0;
const COL_ORDER_EMAIL = 1;
const COL_BOUND_HWID  = 2;
const COL_STATUS      = 3;
const COL_TIER        = 4;
const COL_PLUGINS     = 5;
const COL_SUB_EXPIRY  = 6;
const COL_MAX_DEVICES = 7;
const COL_ALIAS       = 8;
const COL_ORDER_DATE  = 9;
const COL_PAYMENT_PROVIDER = 10;
const COL_PAYMENT_ID       = 11;
const COL_PAYMENT_AMOUNT   = 12;
const COL_PAYMENT_CURRENCY = 13;
const COL_SOURCE           = 14;
const COL_EMAIL_SENT_AT    = 15;

// ─────────────────────────────────────────
// onEdit
// ─────────────────────────────────────────
function onEdit(e) {
  const sheet = e.source.getActiveSheet();
  if (sheet.getName() !== SHEET_LICENSES) return;
  const range = e.range;
  const row = range.getRow();
  const col = range.getColumn();
  if (col !== COL_ORDER_EMAIL + 1 || row === 1) return;
  const email = range.getValue().toString().trim();
  if (!email) return;
  const existingKey = sheet.getRange(row, COL_LICENSE_KEY + 1).getValue();
  if (existingKey) return;
  const newKey = generateLicenseKey();
  sheet.getRange(row, COL_LICENSE_KEY + 1).setValue(newKey);
  const existingStatus = sheet.getRange(row, COL_STATUS + 1).getValue();
  if (!existingStatus) sheet.getRange(row, COL_STATUS + 1).setValue("Valid");
}

// ─────────────────────────────────────────
// doGet：直接處理 GET 請求（避免 redirect 問題）
// ─────────────────────────────────────────
function doGet(e) {
  const action = (e.parameter && e.parameter.action) || "";
  if (action === "beta_signup") {
    return handleBetaSignup(e.parameter);
  }
  if (action === "ping") {
    return createJsonResponse({ "success": true, "server_time_utc": new Date().toISOString() }, 200);
  }
  if (action === "trial") {
    const hwid = e.parameter && e.parameter.hwid;
    if (!hwid) return createJsonResponse({ "success": false, "error_code": "MISSING_HWID", "message": "Missing hwid." }, 400);
    return handleTrial(hwid);
  }
  if (action === "activate") {
    return handleActivate(e.parameter || {});
  }
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'PlugnGO Beta API is running' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─────────────────────────────────────────
// doPost 主路由
// ─────────────────────────────────────────
function doPost(e) {
  try {
    let postData = {};
    if (e.postData && e.postData.contents) {
      postData = JSON.parse(e.postData.contents);
    } else if (e.parameter) {
      postData = e.parameter;
    }

    const action = postData.action || "activate";

    if (action === "beta_signup") return handleBetaSignup(postData);
    if (action === "create_license") return handleCreateLicense(postData);

    const hwid = postData.hwid;
    if (!hwid) return createJsonResponse({ "success": false, "message": "Missing hwid." }, 400);

    if (action === "trial") return handleTrial(hwid);

    if (action === "ping") {
      return createJsonResponse({ "success": true, "server_time_utc": new Date().toISOString() }, 200);
    }

    if (action === "activate") return handleActivate(postData);

    return createJsonResponse({ "success": false, "message": "Invalid action." }, 400);

  } catch (err) {
    return createJsonResponse({ "success": false, "message": "Internal Server Error: " + err.toString() }, 500);
  }
}

function handleActivate(postData) {
  const hwid = postData.hwid;
  const licenseKey = normalizeLicenseKey(postData.license_key);
  const ts = postData.ts;
  const sig = postData.sig;

  if (!hwid) return createJsonResponse({ "success": false, "error_code": "MISSING_HWID", "message": "Missing hwid." }, 400);
  if (!licenseKey) return createJsonResponse({ "success": false, "error_code": "MISSING_LICENSE_KEY", "message": "Missing license_key." }, 400);
  if (!verifyZeroTrustSig(ts, hwid, licenseKey, sig)) {
    return createJsonResponse({ "success": false, "error_code": "INVALID_SIGNATURE", "message": "安全性驗證失敗" }, 403);
  }
  return verifyLicense(licenseKey, hwid);
}

// ─────────────────────────────────────────
// Beta 申請處理
// ─────────────────────────────────────────
function handleBetaSignup(postData) {
  const name  = (postData.name  || "").toString().trim();
  const email = (postData.email || "").toString().trim();
  const type  = (postData.type  || "").toString().trim();

  if (!name || !email || !type)
    return createJsonResponse({ success: false, error: "MISSING_FIELDS" }, 400);

  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  if (!emailRegex.test(email))
    return createJsonResponse({ success: false, error: "INVALID_EMAIL" }, 400);

  const ss          = SpreadsheetApp.openById(SS_ID);
  const codesSheet  = ss.getSheetByName(BETA_CODES_SHEET);
  const signupSheet = ss.getSheetByName(BETA_SIGNUPS_SHEET);
  const licSheet    = ss.getSheetByName(SHEET_LICENSES);

  if (!codesSheet || !signupSheet)
    return createJsonResponse({ success: false, error: "SERVER_ERROR", message: "Beta sheets not found." }, 500);

  // 重複申請檢查
  const signups = signupSheet.getDataRange().getValues();
  for (let i = 1; i < signups.length; i++) {
    if (signups[i][2] && signups[i][2].toString().toLowerCase() === email.toLowerCase())
      return createJsonResponse({ success: false, error: "ALREADY_REGISTERED" }, 200);
  }

  // 取得可用序號
  const codes = codesSheet.getDataRange().getValues();
  let availableRow = -1, availableCode = "";
  for (let i = 1; i < codes.length; i++) {
    if (codes[i][0] && !codes[i][1]) {
      availableRow = i + 1;
      availableCode = codes[i][0].toString();
      break;
    }
  }
  if (availableRow === -1)
    return createJsonResponse({ success: false, error: "NO_CODES_LEFT" }, 200);

  // 標記序號為已使用
  codesSheet.getRange(availableRow, 2).setValue("USED");

  // 寫入申請紀錄
  signupSheet.appendRow([new Date(), name, email, type, availableCode]);

  // ★ 同步寫入 Licenses（讓 verifyLicense 驗證，含 90 天有效期）
  const expiryDate = new Date();
  expiryDate.setDate(expiryDate.getDate() + 90);
  const expiryStr = expiryDate.toISOString().split('T')[0]; // YYYY-MM-DD

  if (licSheet) {
    licSheet.appendRow([
      availableCode,              // A: LicenseKey
      email,                      // B: OrderEmail
      "",                         // C: BoundHWID
      "Valid",                    // D: Status
      "photo",                    // E: Tier
      "",                         // F: Plugins
      expiryStr,                  // G: SubscriptionExpiry
      1,                          // H: MaxDevices
      name,                       // I: Alias
      new Date().toISOString()    // J: OrderDate
    ]);
  }

  sendBetaEmail(name, email, availableCode);
  return createJsonResponse({ success: true }, 200);
}

// ─────────────────────────────────────────
// 正式購買序號建立
// ─────────────────────────────────────────
function handleCreateLicense(postData) {
  if (!verifyAdminSecret(postData)) {
    return createJsonResponse({ "success": false, "error": "UNAUTHORIZED" }, 403);
  }

  const email      = (postData.email       || "").toString().trim();
  const tier       = normalizeTier(postData.tier || "photo");
  const plugins    = normalizePlugins(postData.plugins || "");
  const maxDevices = normalizeMaxDevices(postData.max_devices || 1);
  const alias      = (postData.alias       || "").toString().trim();
  const subExpiry  = (postData.sub_expiry  || "never").toString().trim();
  const paymentProvider = normalizePaymentValue(postData.payment_provider || "");
  const paymentId       = normalizePaymentValue(postData.payment_id || "");
  const paymentAmount   = normalizePaymentValue(postData.payment_amount || "");
  const paymentCurrency = normalizePaymentValue(postData.payment_currency || "TWD");
  const source          = normalizePaymentValue(postData.source || "");
  const shouldSendEmail = String(postData.send_email || "true").toLowerCase() !== "false";

  if (!email) return createJsonResponse({ "success": false, "error": "MISSING_EMAIL" }, 400);
  if (!tier) return createJsonResponse({ "success": false, "error": "INVALID_TIER" }, 400);
  if (plugins === null) return createJsonResponse({ "success": false, "error": "INVALID_PLUGINS" }, 400);
  if (!maxDevices) return createJsonResponse({ "success": false, "error": "INVALID_MAX_DEVICES" }, 400);
  if (!isValidExpiry(subExpiry)) return createJsonResponse({ "success": false, "error": "INVALID_EXPIRY" }, 400);
  if (paymentId && !paymentProvider) return createJsonResponse({ "success": false, "error": "MISSING_PAYMENT_PROVIDER" }, 400);

  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  if (!emailRegex.test(email)) return createJsonResponse({ "success": false, "error": "INVALID_EMAIL" }, 400);

  const ss    = SpreadsheetApp.openById(SS_ID);
  const sheet = ss.getSheetByName(SHEET_LICENSES);
  if (!sheet) return createJsonResponse({ "success": false, "error": "SERVER_ERROR" }, 500);
  ensureLicensePaymentHeaders(sheet);

  const existing = findLicenseByPayment(sheet, paymentProvider, paymentId);
  if (existing) {
    let emailSent = !!existing.emailSentAt;
    if (shouldSendEmail && !emailSent) {
      sendPaidLicenseEmail(existing.alias || alias, existing.email, existing.licenseKey, existing.tier, existing.subExpiry);
      sheet.getRange(existing.rowNumber, COL_EMAIL_SENT_AT + 1).setValue(new Date().toISOString());
      emailSent = true;
    }
    return createJsonResponse({
      "success": true,
      "license_key": existing.licenseKey,
      "email": existing.email,
      "tier": existing.tier,
      "order_date": existing.orderDate,
      "payment_duplicate": true,
      "email_sent": emailSent
    }, 200);
  }

  const newKey    = generateLicenseKey();
  const orderDate = new Date().toISOString();
  let emailSentAt = "";

  sheet.appendRow([
    newKey,
    email,
    "",
    "Valid",
    tier,
    plugins,
    subExpiry,
    maxDevices,
    alias,
    orderDate,
    paymentProvider,
    paymentId,
    paymentAmount,
    paymentCurrency,
    source,
    ""
  ]);

  if (shouldSendEmail) {
    sendPaidLicenseEmail(alias, email, newKey, tier, subExpiry);
    emailSentAt = new Date().toISOString();
    sheet.getRange(sheet.getLastRow(), COL_EMAIL_SENT_AT + 1).setValue(emailSentAt);
  }

  return createJsonResponse({
    "success": true,
    "license_key": newKey,
    "email": email,
    "tier": tier,
    "order_date": orderDate,
    "email_sent": !!emailSentAt
  }, 200);
}

function ensureLicensePaymentHeaders(sheet) {
  const headers = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), COL_EMAIL_SENT_AT + 1)).getValues()[0];
  const expected = [
    "PaymentProvider",
    "PaymentId",
    "PaymentAmount",
    "PaymentCurrency",
    "Source",
    "EmailSentAt"
  ];
  let changed = false;
  for (let i = 0; i < expected.length; i++) {
    const col = COL_PAYMENT_PROVIDER + i;
    if (!headers[col]) {
      headers[col] = expected[i];
      changed = true;
    }
  }
  if (changed) sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
}

function findLicenseByPayment(sheet, provider, paymentId) {
  if (!provider || !paymentId) return null;
  const data = sheet.getDataRange().getValues();
  const normalizedProvider = provider.toLowerCase();
  const normalizedPaymentId = paymentId.toLowerCase();
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowProvider = normalizePaymentValue(row[COL_PAYMENT_PROVIDER] || "").toLowerCase();
    const rowPaymentId = normalizePaymentValue(row[COL_PAYMENT_ID] || "").toLowerCase();
    if (rowProvider !== normalizedProvider || rowPaymentId !== normalizedPaymentId) continue;
    return {
      rowNumber: i + 1,
      licenseKey: String(row[COL_LICENSE_KEY] || "").trim(),
      email: String(row[COL_ORDER_EMAIL] || "").trim(),
      tier: String(row[COL_TIER] || "photo").trim(),
      subExpiry: String(row[COL_SUB_EXPIRY] || "never").trim(),
      alias: String(row[COL_ALIAS] || "").trim(),
      orderDate: String(row[COL_ORDER_DATE] || "").trim(),
      emailSentAt: String(row[COL_EMAIL_SENT_AT] || "").trim()
    };
  }
  return null;
}

// ─────────────────────────────────────────
// 序號驗證（含有效期檢查）
// ─────────────────────────────────────────
function verifyLicense(inputKey, inputHwid) {
  const sheet = SpreadsheetApp.openById(SS_ID).getSheetByName(SHEET_LICENSES);
  if (!sheet) return createJsonResponse({ "success": false, "message": "Sheet 'Licenses' not found." }, 500);

  const normalizedInputKey = normalizeLicenseKey(inputKey);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const row       = data[i];
    const keyInDb   = normalizeLicenseKey(row[0]);
    const emailInDb = String(row[1]).trim();
    const boundHwid = String(row[2]).trim();
    const status    = String(row[3]).trim().toLowerCase();

    if (keyInDb !== normalizedInputKey) continue;

    // 1. 撤銷檢查
    if (status === "revoked") {
      return createJsonResponse({ "success": false, "error_code": "LICENSE_REVOKED", "message": "此序號已被撤銷" }, 403);
    }

    // 2. ★ 有效期檢查
    const subExpiry = String(row[COL_SUB_EXPIRY] || "").trim();
    if (subExpiry && subExpiry !== "never") {
      if (new Date() > new Date(subExpiry)) {
        return createJsonResponse({
          "success": false,
          "error_code": "LICENSE_EXPIRED",
          "message": "試用期已結束，感謝你的使用。如需繼續，請購買正式版。"
        }, 403);
      }
    }

    // 3. HWID 綁定 / 驗證
    const tier    = normalizeTier(row[4] || "photo");
    const plugins = normalizePlugins(row[5] || "");
    if (!tier) {
      return createJsonResponse({ "success": false, "error_code": "INVALID_TIER", "message": "Invalid license tier in server row." }, 500);
    }
    if (plugins === null) {
      return createJsonResponse({ "success": false, "error_code": "INVALID_PLUGINS", "message": "Invalid license plugins in server row." }, 500);
    }

    if (boundHwid === "") {
      sheet.getRange(i + 1, COL_BOUND_HWID + 1).setValue(inputHwid);
      return createLicensedResponse("License Activated & Bound successfully.", normalizedInputKey, inputHwid, emailInDb, tier, plugins, subExpiry);
    } else if (boundHwid === inputHwid) {
      return createLicensedResponse("License Validated.", normalizedInputKey, inputHwid, emailInDb, tier, plugins, subExpiry);
    } else {
      return createJsonResponse({ "success": false, "error_code": "HWID_MISMATCH", "message": "此序號已在其他裝置上綁定。若需轉移授權，請聯絡客服解綁。" }, 403);
    }
  }

  return createJsonResponse({ "success": false, "error_code": "INVALID_KEY", "message": "無效的序號" }, 404);
}

// ─────────────────────────────────────────
// 序號產生器
// ─────────────────────────────────────────
function generateLicenseKey() {
  const chars = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789";
  const ss    = SpreadsheetApp.openById(SS_ID);
  const sheet = ss.getSheetByName(SHEET_LICENSES);
  const existing = sheet ? sheet.getDataRange().getValues().map(r => r[0].toString()) : [];
  let key;
  do {
    const seg = () => Array.from({ length: 4 }, () => chars[Math.floor(Math.random() * chars.length)]).join("");
    key = "PGO-" + seg() + "-" + seg() + "-" + seg();
  } while (existing.includes(key));
  return key;
}

// ─────────────────────────────────────────
// Beta Email 寄送
// ─────────────────────────────────────────
function sendBetaEmail(name, email, code) {
  const subject = "你的 PlugnGO Beta 序號來了";

  const DL_M     = "https://github.com/skiseiju/plugngo-website/releases/download/v1.8.2/PlugnGO-v1.8.2-Mac.dmg";
  const DL_INTEL = "https://github.com/skiseiju/plugngo-website/releases/download/v1.8.2/PlugnGO-v1.8.2-Intel.dmg";
  const GK_URL   = "https://plugngo.skiseiju.com/gatekeeper";

  const plainBody = `嗨 ${name}，

感謝你申請 PlugnGO Beta 試用！

你的啟動碼：${code}
有效期限：三個月

▼ 下載 PlugnGO
M 系列版：${DL_M}
Intel 版：${DL_INTEL}

▼ macOS 首次開啟教學
首次開啟可能被 macOS 攔截，30 秒即可解除：
1. 出現「無法打開」→ 按「完成」關閉
2. 到「系統設定 → 隱私權與安全性」→ 點「強制打開」
3. 輸入密碼確認，搞定

如果出現「已損壞」→ 開終端機執行：xattr -cr /Applications/PlugnGO.app

完整圖文教學：${GK_URL}

使用過程有任何問題，直接回覆這封信就可以找到我，或到我的 Facebook 隨意留言也可以。
${BETA_FACEBOOK_URL}

海哥
PlugnGO`;

  const htmlBody = `<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <style>
    body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif; background: #f5f5f5; margin: 0; padding: 40px 20px; }
    .container { max-width: 520px; margin: 0 auto; background: #fff; border-radius: 12px; overflow: hidden; }
    .header { background: #0a0a0a; padding: 32px; text-align: center; }
    .header h1 { color: #fff; margin: 0; font-size: 22px; letter-spacing: 0.05em; }
    .header span { color: #888; font-size: 13px; }
    .bd { padding: 36px 32px; }
    .greeting { font-size: 16px; color: #333; margin-bottom: 24px; }
    .code-box { background: #f8f8f8; border: 2px dashed #d4d4d4; border-radius: 8px; padding: 20px; text-align: center; margin: 28px 0; }
    .code-label { font-size: 11px; color: #888; letter-spacing: 0.12em; text-transform: uppercase; margin-bottom: 8px; }
    .code { font-size: 26px; font-weight: 700; color: #0a0a0a; letter-spacing: 0.08em; font-family: 'SF Mono', Menlo, monospace; }
    .validity { font-size: 13px; color: #888; margin-top: 8px; }
    .msg { font-size: 15px; color: #555; line-height: 1.7; margin-bottom: 24px; }
    .section-label { font-size: 11px; font-weight: 700; letter-spacing: 0.12em; text-transform: uppercase; color: #999; margin-bottom: 16px; }
    .dl-row { margin-bottom: 16px; }
    .dl-row a { display: inline-block; padding: 12px 24px; border-radius: 6px; font-size: 14px; font-weight: 700; text-decoration: none; text-align: center; }
    .btn-primary { background: #f97316; color: #ffffff !important; }
    .btn-ghost { background: #f5f5f5; color: #333 !important; border: 1px solid #d4d4d4; }
    .dl-meta { font-size: 12px; color: #999; margin-top: 8px; }
    .gk-box { background: #fffbf5; border: 1px solid #fed7aa; border-radius: 8px; padding: 20px 24px; margin: 28px 0; }
    .gk-title { font-size: 14px; font-weight: 700; color: #333; margin-bottom: 4px; }
    .gk-sub { font-size: 13px; color: #999; margin-bottom: 16px; }
    .gk-step { margin-bottom: 14px; line-height: 1.6; }
    .gk-num { display: inline-block; width: 22px; height: 22px; background: #f97316; color: #fff; border-radius: 50%; font-size: 11px; font-weight: 800; text-align: center; line-height: 22px; margin-right: 8px; vertical-align: top; }
    .gk-text { font-size: 13px; color: #555; }
    .gk-text strong { color: #333; }
    .gk-divider { border-top: 1px solid #fed7aa; margin: 16px 0; }
    .gk-warn { font-size: 12px; color: #dc2626; margin-bottom: 10px; font-weight: 600; }
    .gk-cmd { background: #1e1e2e; color: #cdd6f4; padding: 8px 12px; border-radius: 5px; font-family: 'SF Mono', Menlo, monospace; font-size: 12px; margin: 6px 0; display: inline-block; }
    .gk-note { font-size: 12px; color: #888; line-height: 1.6; margin-top: 16px; padding-top: 14px; border-top: 1px solid #fed7aa; }
    .gk-note li { margin-bottom: 4px; }
    .gk-link { display: inline-block; margin-top: 14px; font-size: 12px; color: #f97316; text-decoration: none; font-weight: 600; }
    .fb-link { color: #0a0a0a; font-weight: 600; text-decoration: none; }
    .ft { border-top: 1px solid #eee; padding: 20px 32px; font-size: 13px; color: #aaa; }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <h1>PlugnGO</h1>
      <span>Beta 試用邀請</span>
    </div>
    <div class="bd">
      <p class="greeting">嗨 ${name}，</p>
      <p class="msg">感謝你申請 PlugnGO Beta 試用！你的啟動碼在這裡：</p>

      <div class="code-box">
        <div class="code-label">啟動碼</div>
        <div class="code">${code}</div>
        <div class="validity">有效期限：三個月</div>
      </div>

      <div class="section-label">下載 PlugnGO</div>
      <div class="dl-row">
        <a href="${DL_M}" class="btn-primary">⬇ M 系列版下載</a>
        &nbsp;&nbsp;
        <a href="${DL_INTEL}" class="btn-ghost">⬇ Intel 版下載</a>
      </div>
      <div class="dl-meta">v1.8.2 · macOS 12 Monterey 以上</div>

      <div class="gk-box">
        <div class="gk-title">macOS 首次開啟教學</div>
        <div class="gk-sub">首次開啟可能被 macOS 攔截，30 秒搞定。</div>

        <div style="font-size:12px;color:#888;font-weight:600;margin-bottom:10px;">情況 A —「無法打開」警告</div>
        <div class="gk-step">
          <span class="gk-num">1</span>
          <span class="gk-text">出現警告後，按<strong>「完成」</strong>關閉（不要按移到垃圾桶）。</span>
        </div>
        <div class="gk-step">
          <span class="gk-num">2</span>
          <span class="gk-text">打開 <strong> → 系統設定 → 隱私權與安全性</strong>，找到 PlugnGO 被封鎖的訊息。</span>
        </div>
        <div class="gk-step">
          <span class="gk-num">3</span>
          <span class="gk-text">點<strong>「強制打開」</strong>，輸入密碼確認。搞定！</span>
        </div>

        <div class="gk-divider"></div>

        <div class="gk-warn">情況 B —「已損壞，無法打開」</div>
        <div class="gk-text" style="margin-bottom:6px;">按「取消」，然後開啟終端機，貼上以下指令：</div>
        <div class="gk-cmd">xattr -cr /Applications/PlugnGO.app</div>
        <div class="gk-text" style="margin-top:8px;">執行後再開啟 PlugnGO 即可。</div>

        <div class="gk-note">
          <strong>注意：</strong>
          <ul style="padding-left:1.2em;margin:4px 0 0;">
            <li>請先將 App 從 DMG 拖到「應用程式」資料夾，再開啟</li>
            <li>首次備份 SD 卡時，記得點「允許」存取可卸除式卷宗</li>
            <li>只需設定一次，之後不會再被攔截</li>
          </ul>
        </div>

        <a href="${GK_URL}" class="gk-link">需要更詳細的圖文教學 →</a>
      </div>

      <p class="msg">使用過程有任何問題，直接回覆這封信，或到 <a href="${BETA_FACEBOOK_URL}" class="fb-link">Facebook</a> 找我。</p>
      <p class="msg">海哥<br>PlugnGO</p>
    </div>
    <div class="ft">這是自動寄出的信件，有問題請直接回覆。</div>
  </div>
</body>
</html>`;

  GmailApp.sendEmail(email, subject, plainBody, {
    htmlBody: htmlBody,
    name: BETA_SENDER_NAME,
    replyTo: BETA_REPLY_TO
  });
}

// ─────────────────────────────────────────
// 正式購買 Email 寄送
// ─────────────────────────────────────────
function sendPaidLicenseEmail(name, email, code, tier, expiry) {
  const displayName = String(name || "").trim() || "你好";
  const planName = tier === "photo" ? "PlugnGO Photo 攝影版" : "PlugnGO";
  const expiryLabel = formatExpiryLabel(expiry);
  const subject = "你的 PlugnGO Photo 一年授權序號";
  const downloadUrl = "https://plugngo.skiseiju.com/#download";
  const gatekeeperUrl = "https://plugngo.skiseiju.com/gatekeeper";

  const plainBody = `${displayName}，

感謝你購買 ${planName}。

你的序號：${code}
授權期限：${expiryLabel}

下載 PlugnGO：
${downloadUrl}

啟用方式：
1. 安裝並開啟 PlugnGO
2. 到授權/設定畫面輸入上方序號
3. 完成啟用後即可使用 Photo 版功能

macOS 首次開啟若被系統攔截，可參考：
${gatekeeperUrl}

如果付款完成後序號無法啟用，直接回覆這封信即可。

海哥
PlugnGO`;

  const htmlBody = `<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <style>
    body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif; background: #f5f5f5; margin: 0; padding: 40px 20px; color: #222; }
    .container { max-width: 560px; margin: 0 auto; background: #fff; border-radius: 12px; overflow: hidden; }
    .header { background: #111; padding: 30px 32px; }
    .header h1 { color: #fff; margin: 0; font-size: 22px; }
    .header p { color: #aaa; margin: 8px 0 0; font-size: 13px; }
    .bd { padding: 34px 32px; }
    .msg { font-size: 15px; line-height: 1.75; color: #444; margin: 0 0 22px; }
    .code-box { background: #f8f8f8; border: 2px dashed #d4d4d4; border-radius: 8px; padding: 22px; text-align: center; margin: 26px 0; }
    .code-label { font-size: 11px; color: #888; letter-spacing: 0.12em; text-transform: uppercase; margin-bottom: 8px; }
    .code { font-size: 26px; font-weight: 800; color: #111; letter-spacing: 0.08em; font-family: 'SF Mono', Menlo, monospace; }
    .validity { font-size: 13px; color: #777; margin-top: 10px; }
    .btn { display: inline-block; background: #f97316; color: #fff !important; text-decoration: none; border-radius: 6px; padding: 12px 20px; font-weight: 700; font-size: 14px; }
    .steps { background: #fffbf5; border: 1px solid #fed7aa; border-radius: 8px; padding: 18px 20px; margin: 26px 0; }
    .steps h2 { font-size: 14px; margin: 0 0 12px; }
    .steps ol { padding-left: 1.3em; margin: 0; color: #555; font-size: 14px; line-height: 1.7; }
    .small { font-size: 13px; color: #777; line-height: 1.7; }
    .ft { border-top: 1px solid #eee; padding: 18px 32px; font-size: 12px; color: #999; }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <h1>PlugnGO Photo</h1>
      <p>一年授權序號</p>
    </div>
    <div class="bd">
      <p class="msg">${displayName}，感謝你購買 ${planName}。付款已完成，你的序號如下：</p>
      <div class="code-box">
        <div class="code-label">序號</div>
        <div class="code">${code}</div>
        <div class="validity">授權期限：${expiryLabel}</div>
      </div>
      <p><a class="btn" href="${downloadUrl}">下載 PlugnGO</a></p>
      <div class="steps">
        <h2>啟用方式</h2>
        <ol>
          <li>安裝並開啟 PlugnGO</li>
          <li>到授權/設定畫面輸入上方序號</li>
          <li>完成啟用後即可使用 Photo 版功能</li>
        </ol>
      </div>
      <p class="small">macOS 首次開啟若被系統攔截，可參考 <a href="${gatekeeperUrl}">首次開啟教學</a>。</p>
      <p class="small">如果付款完成後序號無法啟用，直接回覆這封信即可。</p>
      <p class="msg">海哥<br>PlugnGO</p>
    </div>
    <div class="ft">這是自動寄出的授權信，有問題請直接回覆。</div>
  </div>
</body>
</html>`;

  GmailApp.sendEmail(email, subject, plainBody, {
    htmlBody: htmlBody,
    name: BETA_SENDER_NAME,
    replyTo: BETA_REPLY_TO
  });
}

// ─────────────────────────────────────────
// Trial 處理
// ─────────────────────────────────────────
function handleTrial(inputHwid) {
  const ss  = SpreadsheetApp.openById(SS_ID);
  let sheet = ss.getSheetByName(SHEET_TRIALS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_TRIALS);
    sheet.appendRow(["HWID", "StartTime_UTC", "LastSeen_UTC", "RequestCount"]);
  }
  const data   = sheet.getDataRange().getValues();
  const nowUtc = new Date().toISOString();
  const nowMs  = new Date().getTime();

  for (let i = 1; i < data.length; i++) {
    const row          = data[i];
    const dbHwid       = String(row[0]).trim();
    const startTime    = String(row[1]).trim();
    let requestCount   = parseInt(row[3]) || 0;
    if (dbHwid !== inputHwid) continue;
    requestCount++;
    sheet.getRange(i + 1, 3).setValue(nowUtc);
    sheet.getRange(i + 1, 4).setValue(requestCount);
    const daysElapsed = (nowMs - new Date(startTime).getTime()) / (1000 * 60 * 60 * 24);
    return createJsonResponse({ "success": true, "is_new": false, "trial_start_utc": startTime, "server_time_utc": nowUtc, "days_elapsed": Math.floor(daysElapsed), "request_count": requestCount }, 200);
  }

  sheet.appendRow([inputHwid, nowUtc, nowUtc, 1]);
  return createJsonResponse({ "success": true, "is_new": true, "trial_start_utc": nowUtc, "server_time_utc": nowUtc, "days_elapsed": 0, "request_count": 1, "message": "15-day trial started." }, 200);
}

// ─────────────────────────────────────────
// 工具函式
// ─────────────────────────────────────────
function createJsonResponse(data, statusCode) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}

function verifyAdminSecret(postData) {
  const expected = PropertiesService.getScriptProperties().getProperty(ADMIN_SECRET_PROPERTY);
  if (!expected) return false;
  const provided = (postData.admin_secret || postData.admin_token || "").toString();
  return provided && provided === expected;
}

function normalizeTier(value) {
  const tier = String(value || "").trim().toLowerCase();
  return VALID_LICENSE_TIERS.indexOf(tier) >= 0 ? tier : "";
}

function normalizePlugins(value) {
  const raw = String(value || "").trim().toLowerCase();
  if (!raw) return "";
  const plugins = raw
    .split(",")
    .map(p => p.trim())
    .filter(p => p);
  for (let i = 0; i < plugins.length; i++) {
    if (VALID_LICENSE_PLUGINS.indexOf(plugins[i]) < 0) return null;
  }
  return plugins.join(",");
}

function normalizeMaxDevices(value) {
  const parsed = parseInt(value, 10);
  if (!Number.isFinite(parsed) || parsed < 1 || parsed > 10) return "";
  return String(parsed);
}

function normalizePaymentValue(value) {
  return String(value || "").trim().replace(/[\r\n\t]+/g, " ");
}

function formatExpiryLabel(value) {
  const expiry = String(value || "").trim();
  if (!expiry || expiry === "never") return "永久";
  return expiry;
}

function isValidExpiry(value) {
  const expiry = String(value || "").trim();
  return expiry === "never" || /^\d{4}-\d{2}-\d{2}$/.test(expiry);
}

function verifyZeroTrustSig(ts, hwid, licenseKey, clientSig) {
  const serverNow = new Date().getTime() / 1000;
  if (!ts || Math.abs(serverNow - ts) > 600) return false;
  const salt      = "PGO_BETA_2026_SALT";
  const payload   = ts + hwid + normalizeLicenseKey(licenseKey) + salt;
  const signature = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, payload);
  const hexSig    = signature.map(b => ("0" + (b & 0xFF).toString(16)).slice(-2)).join("");
  return hexSig === clientSig;
}

function normalizeLicenseKey(value) {
  return String(value || "").trim().replace(/\s+/g, "").toUpperCase();
}

function createLicensedResponse(message, licenseKey, hwid, email, tier, plugins, expiry) {
  const cert = createLicenseCertificate(licenseKey, hwid, email, tier, plugins, expiry);
  return createJsonResponse({
    "success": true,
    "message": message,
    "email": email,
    "tier": tier,
    "plugins": plugins,
    "expiry": expiry || "never",
    "license_certificate": cert
  }, 200);
}

function createLicenseCertificate(licenseKey, hwid, email, tier, plugins, expiry) {
  const now = new Date();
  const notAfter = certificateNotAfter(now, expiry);
  const payload = {
    "cert_version": 1,
    "license_key": normalizeLicenseKey(licenseKey),
    "hwid": String(hwid || "").trim(),
    "email": String(email || "").trim(),
    "tier": normalizeTier(tier),
    "plugins": String(plugins || "").trim().toLowerCase().split(",").map(p => p.trim()).filter(p => p),
    "expiry": String(expiry || "never").trim() || "never",
    "issued_at": now.toISOString(),
    "not_after": notAfter.toISOString()
  };
  const canonical = canonicalJson(payload);
  const privateKey = getLicensePrivateKey();
  if (!privateKey) throw new Error("Missing " + LICENSE_PRIVATE_KEY_PROPERTY);
  const signatureBytes = Utilities.computeRsaSha256Signature(canonical, privateKey);
  return {
    "payload": payload,
    "signature": Utilities.base64Encode(signatureBytes)
  };
}

function getLicensePrivateKey() {
  const raw = PropertiesService.getScriptProperties().getProperty(LICENSE_PRIVATE_KEY_PROPERTY);
  if (!raw) return "";
  return normalizePrivateKeyPem(raw);
}

function normalizePrivateKeyPem(value) {
  let key = String(value || "").trim();

  // The Apps Script settings UI stores property values in a single-line input.
  // Accept escaped newlines, one-line PEM, and the hex-encoded PEM format used
  // by the local Keychain backup.
  if (/^[0-9a-fA-F]+$/.test(key) && key.length % 2 === 0) {
    let decoded = "";
    for (let i = 0; i < key.length; i += 2) {
      decoded += String.fromCharCode(parseInt(key.substr(i, 2), 16));
    }
    key = decoded.trim();
  }

  key = key.replace(/\\n/g, "\n").trim();
  const begin = "-----BEGIN PRIVATE KEY-----";
  const end = "-----END PRIVATE KEY-----";
  if (key.indexOf(begin) >= 0 && key.indexOf(end) >= 0 && key.indexOf("\n") < 0) {
    const body = key
      .replace(begin, "")
      .replace(end, "")
      .replace(/\s+/g, "");
    const lines = body.match(/.{1,64}/g) || [];
    key = begin + "\n" + lines.join("\n") + "\n" + end + "\n";
  }
  return key;
}

function certificateNotAfter(now, expiry) {
  const ttl = new Date(now.getTime() + LICENSE_CERT_TTL_DAYS * 24 * 60 * 60 * 1000);
  const expiryValue = String(expiry || "").trim();
  if (!expiryValue || expiryValue === "never") return ttl;
  const subscriptionEnd = new Date(expiryValue + "T23:59:59Z");
  if (isNaN(subscriptionEnd.getTime())) return ttl;
  return subscriptionEnd.getTime() < ttl.getTime() ? subscriptionEnd : ttl;
}

function canonicalJson(value) {
  if (value === null || typeof value !== "object") return JSON.stringify(value);
  if (Array.isArray(value)) return "[" + value.map(canonicalJson).join(",") + "]";
  const keys = Object.keys(value).sort();
  return "{" + keys.map(k => JSON.stringify(k) + ":" + canonicalJson(value[k])).join(",") + "}";
}

// ─────────────────────────────────────────
// 初始化工具（只需執行一次）
// ─────────────────────────────────────────
function setupBetaSheets() {
  const ss = SpreadsheetApp.openById(SS_ID);
  let c = ss.getSheetByName(BETA_CODES_SHEET);
  if (!c) c = ss.insertSheet(BETA_CODES_SHEET);
  c.clearContents();
  c.getRange(1, 1, 1, 2).setValues([["序號", "狀態"]]);
  let s = ss.getSheetByName(BETA_SIGNUPS_SHEET);
  if (!s) s = ss.insertSheet(BETA_SIGNUPS_SHEET);
  s.clearContents();
  s.getRange(1, 1, 1, 5).setValues([["時間戳記", "姓名", "Email", "拍攝類型", "序號"]]);
  Logger.log("✅ Beta 工作表建立完成");
}

function generateBetaCodes() {
  const ss         = SpreadsheetApp.openById(SS_ID);
  const codesSheet = ss.getSheetByName(BETA_CODES_SHEET);
  if (!codesSheet) { Logger.log("❌ 請先執行 setupBetaSheets()"); return; }
  const chars = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789";
  const codes = new Set();
  while (codes.size < 300) {
    let code = "PNG-BETA-";
    for (let i = 0; i < 6; i++) code += chars[Math.floor(Math.random() * chars.length)];
    codes.add(code);
  }
  const rows = Array.from(codes).map(c => [c, ""]);
  codesSheet.getRange(2, 1, rows.length, 2).setValues(rows);
  Logger.log("✅ 已產生 300 組 Beta 序號");
}
