const TRIAL_SECRET = "fb_trial_secret_v1";
const TRIAL_DEFAULT = 10;
const TRIAL_VERSION = 1;

function trialPayload(remaining, version) {
  return `trialRemaining=${remaining}|v=${version}`;
}

async function hmacSha256(secret, message) {
  const enc = new TextEncoder();
  const key = await crypto.subtle.importKey(
    "raw",
    enc.encode(secret),
    { name: "HMAC", hash: "SHA-256" },
    false,
    ["sign"]
  );
  const sig = await crypto.subtle.sign("HMAC", key, enc.encode(message));
  return Array.from(new Uint8Array(sig)).map((b) => b.toString(16).padStart(2, "0")).join("");
}

async function sha256Hex(message) {
  const enc = new TextEncoder();
  const digest = await crypto.subtle.digest("SHA-256", enc.encode(message));
  return Array.from(new Uint8Array(digest)).map((b) => b.toString(16).padStart(2, "0")).join("");
}

async function signTrial(remaining, version) {
  const payload = trialPayload(remaining, version);
  try {
    return await hmacSha256(TRIAL_SECRET, payload);
  } catch (_) {
    return await sha256Hex(TRIAL_SECRET + payload);
  }
}

async function loadTrialRaw() {
  const data = await chrome.storage.local.get(["trialRemaining", "trialSig", "trialVersion"]);
  return {
    remaining: typeof data.trialRemaining === "number" ? data.trialRemaining : null,
    sig: typeof data.trialSig === "string" ? data.trialSig : null,
    version: typeof data.trialVersion === "number" ? data.trialVersion : TRIAL_VERSION,
  };
}

async function saveTrial(remaining, version) {
  const sig = await signTrial(remaining, version);
  await chrome.storage.local.set({
    trialRemaining: remaining,
    trialSig: sig,
    trialVersion: version,
  });
  return { remaining, sig, version };
}

async function verifyTrial({ remaining, sig, version }) {
  if (typeof remaining !== "number" || !sig) return false;
  const expected = await signTrial(remaining, version);
  return expected === sig;
}

async function initTrialIfNeeded() {
  const raw = await loadTrialRaw();
  if (raw.remaining === null || !raw.sig) {
    await saveTrial(TRIAL_DEFAULT, TRIAL_VERSION);
    return;
  }
  const ok = await verifyTrial(raw);
  if (!ok) {
    await saveTrial(0, raw.version);
  }
}

async function getTrialRemaining() {
  const raw = await loadTrialRaw();
  if (raw.remaining === null || !raw.sig) {
    await saveTrial(TRIAL_DEFAULT, TRIAL_VERSION);
    return TRIAL_DEFAULT;
  }
  const ok = await verifyTrial(raw);
  if (!ok) {
    await saveTrial(0, raw.version);
    return 0;
  }
  return raw.remaining;
}

async function consumeTrial(count = 1) {
  const raw = await loadTrialRaw();
  if (raw.remaining === null || !raw.sig) {
    await saveTrial(0, raw.version);
    return { remaining: 0, ok: false, reason: "invalid" };
  }
  const ok = await verifyTrial(raw);
  if (!ok) {
    await saveTrial(0, raw.version);
    return { remaining: 0, ok: false, reason: "invalid" };
  }
  if (raw.remaining <= 0) {
    return { remaining: 0, ok: false, reason: "empty" };
  }
  const next = Math.max(0, raw.remaining - count);
  await saveTrial(next, raw.version);
  return { remaining: next, ok: true };
}

async function resetTrialForDev() {
  await saveTrial(TRIAL_DEFAULT, TRIAL_VERSION);
}
