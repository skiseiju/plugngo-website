const PHOTO_PLAN = {
  route: "/api/payuni/notify/photo",
  tier: "photo",
  amount: 490,
  currency: "TWD",
  source: "payuni_uop_photo_490"
};

export default {
  async fetch(request, env) {
    const url = new URL(request.url);

    if (url.pathname === PHOTO_PLAN.route) {
      return handlePayuniPhotoNotify(request, env);
    }

    if (url.pathname === "/api/payuni/health") {
      return jsonResponse({ success: true, service: "plugngo-payuni", server_time_utc: new Date().toISOString() });
    }

    if (env.ASSETS) return env.ASSETS.fetch(request);
    return new Response("Not found", { status: 404 });
  }
};

export async function handlePayuniPhotoNotify(request, env) {
  if (request.method === "GET") {
    return jsonResponse({ success: true, endpoint: PHOTO_PLAN.route });
  }

  if (request.method !== "POST") {
    return jsonResponse({ success: false, error: "METHOD_NOT_ALLOWED" }, 405);
  }

  assertEnv(env, [
    "PAYUNI_HASH_KEY",
    "PAYUNI_HASH_IV",
    "PLUGNGO_ADMIN_SECRET",
    "PLUGNGO_APPS_SCRIPT_URL"
  ]);

  const form = await readRequestForm(request);
  const encryptInfo = form.EncryptInfo || "";
  const hashInfo = form.HashInfo || "";
  if (!encryptInfo || !hashInfo) {
    return jsonResponse({ success: false, error: "MISSING_PAYUNI_FIELDS" }, 400);
  }

  const expectedHash = await payuniHash(env.PAYUNI_HASH_KEY, encryptInfo, env.PAYUNI_HASH_IV);
  if (expectedHash !== hashInfo.toUpperCase()) {
    return jsonResponse({ success: false, error: "PAYUNI_HASH_MISMATCH" }, 403);
  }

  const payment = await decryptPayuniEncryptInfo(encryptInfo, env.PAYUNI_HASH_KEY, env.PAYUNI_HASH_IV);
  const status = normalizeText(payment.Status || form.Status || payment.TradeStatus || "");
  const amount = Number(payment.TradeAmt || payment.Amt || payment.Amount || 0);
  const email = normalizeText(payment.UsrMail || payment.BuyerEmail || payment.Email || payment.EmailAddress || "");
  const buyerName = normalizeText(payment.UsrName || payment.BuyerName || payment.Name || "");
  const paymentId = normalizeText(
    payment.TradeNo ||
    payment.MerTradeNo ||
    payment.UOPTradeNo ||
    payment.OrderNo ||
    await sha256Hex(encryptInfo)
  );

  if (!isPaidStatus(payment, form)) {
    return jsonResponse({ success: false, error: "PAYUNI_NOT_PAID", status }, 409);
  }

  if (amount !== PHOTO_PLAN.amount) {
    return jsonResponse({ success: false, error: "AMOUNT_MISMATCH", expected: PHOTO_PLAN.amount, received: amount }, 409);
  }

  if (!email || !isValidEmail(email)) {
    return jsonResponse({ success: false, error: "MISSING_BUYER_EMAIL" }, 409);
  }

  const licensePayload = {
    action: "create_license",
    admin_secret: env.PLUGNGO_ADMIN_SECRET,
    email,
    tier: PHOTO_PLAN.tier,
    plugins: "",
    max_devices: 1,
    alias: buyerName,
    sub_expiry: oneYearFromTodayTaipei(),
    payment_provider: "payuni",
    payment_id: paymentId,
    payment_amount: String(amount),
    payment_currency: PHOTO_PLAN.currency,
    source: PHOTO_PLAN.source,
    send_email: true
  };

  const licenseResponse = await fetch(env.PLUGNGO_APPS_SCRIPT_URL, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(licensePayload)
  });

  const licenseResult = await readJsonSafely(licenseResponse);
  if (!licenseResponse.ok || !licenseResult.success) {
    return jsonResponse({
      success: false,
      error: "LICENSE_CREATE_FAILED",
      status: licenseResponse.status,
      license_error: licenseResult.error || licenseResult.message || "UNKNOWN"
    }, 502);
  }

  return jsonResponse({
    success: true,
    provider: "payuni",
    plan: PHOTO_PLAN.tier,
    amount,
    email,
    payment_id: paymentId,
    license_created: !licenseResult.payment_duplicate,
    email_sent: !!licenseResult.email_sent
  });
}

function assertEnv(env, names) {
  const missing = names.filter((name) => !env[name]);
  if (missing.length) throw new Error("Missing Worker env: " + missing.join(", "));
}

async function readRequestForm(request) {
  const contentType = request.headers.get("content-type") || "";
  if (contentType.includes("multipart/form-data")) {
    const formData = await request.formData();
    return Object.fromEntries(Array.from(formData.entries()).map(([key, value]) => [key, String(value)]));
  }
  const body = await request.text();
  return Object.fromEntries(new URLSearchParams(body).entries());
}

async function payuniHash(key, encryptInfo, iv) {
  return sha256Hex(String(key).trim() + encryptInfo + String(iv).trim());
}

async function decryptPayuniEncryptInfo(encryptInfo, key, iv) {
  const packed = new TextDecoder().decode(hexToBytes(encryptInfo));
  const [encryptedBase64, tagBase64] = packed.split(":::");
  if (!encryptedBase64 || !tagBase64) throw new Error("Invalid PAYUNi EncryptInfo envelope");

  const encrypted = base64ToBytes(encryptedBase64);
  const tag = base64ToBytes(tagBase64);
  const cipherWithTag = concatBytes(encrypted, tag);
  const cryptoKey = await crypto.subtle.importKey(
    "raw",
    new TextEncoder().encode(String(key).trim()),
    "AES-GCM",
    false,
    ["decrypt"]
  );
  const plainBuffer = await crypto.subtle.decrypt(
    { name: "AES-GCM", iv: new TextEncoder().encode(String(iv).trim()), tagLength: 128 },
    cryptoKey,
    cipherWithTag
  );
  const plain = new TextDecoder().decode(plainBuffer);
  return Object.fromEntries(new URLSearchParams(plain).entries());
}

function isPaidStatus(payment, rawForm) {
  const candidates = [
    payment.Status,
    rawForm.Status,
    payment.TradeStatus,
    payment.PayStatus,
    payment.StatusCode
  ].map((value) => normalizeText(value).toUpperCase()).filter(Boolean);

  if (candidates.includes("SUCCESS")) return true;
  if (candidates.includes("PAID")) return true;
  if (candidates.includes("1")) return true;
  return false;
}

function oneYearFromTodayTaipei(now = new Date()) {
  const taipei = new Date(now.getTime() + 8 * 60 * 60 * 1000);
  const year = taipei.getUTCFullYear() + 1;
  const month = taipei.getUTCMonth();
  const day = taipei.getUTCDate();
  const next = new Date(Date.UTC(year, month, day));
  return [
    next.getUTCFullYear(),
    String(next.getUTCMonth() + 1).padStart(2, "0"),
    String(next.getUTCDate()).padStart(2, "0")
  ].join("-");
}

function isValidEmail(value) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(value);
}

function normalizeText(value) {
  return String(value || "").trim();
}

async function readJsonSafely(response) {
  const text = await response.text();
  try {
    return JSON.parse(text);
  } catch (_err) {
    return { success: false, message: text.slice(0, 500) };
  }
}

async function sha256Hex(value) {
  const digest = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(value));
  return bytesToHex(new Uint8Array(digest)).toUpperCase();
}

function hexToBytes(hex) {
  const clean = String(hex || "").trim();
  if (!/^[0-9a-fA-F]*$/.test(clean) || clean.length % 2 !== 0) {
    throw new Error("Invalid hex input");
  }
  const bytes = new Uint8Array(clean.length / 2);
  for (let i = 0; i < clean.length; i += 2) {
    bytes[i / 2] = parseInt(clean.slice(i, i + 2), 16);
  }
  return bytes;
}

function base64ToBytes(value) {
  const binary = atob(value);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  return bytes;
}

function concatBytes(a, b) {
  const out = new Uint8Array(a.length + b.length);
  out.set(a, 0);
  out.set(b, a.length);
  return out;
}

function bytesToHex(bytes) {
  return Array.from(bytes, (byte) => byte.toString(16).padStart(2, "0")).join("");
}

function jsonResponse(data, status = 200) {
  return new Response(JSON.stringify(data), {
    status,
    headers: { "Content-Type": "application/json; charset=utf-8" }
  });
}
