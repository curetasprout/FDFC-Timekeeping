// usage.js — private per-user usage analytics.
//
// Records the user's work email + event + timestamp to a private Google Sheet via a
// write-only Google Apps Script. NO tool data is ever sent — only email, event, time.
// Deployed-site only (never pings from file:// or localhost). Based on the shared
// usage-tracking prompt from the GL Report Builder; differences for this tool:
//   - no daily "open" ping (events are the two exports only)
//   - payload carries an `app` field so several tools can share one Sheet
//   - tracking (and the email prompt) stays off until USAGE_TOKEN is filled in

const USAGE_URL = 'https://script.google.com/macros/s/AKfycbxOgOePw_WaMZOthS9eWNQoaCb0uG4egT1yU5yHcEbjMthTLOx1k1jqER2bEQaL26oFpw/exec';
const USAGE_TOKEN = 'ghtGqe9qV_JetDg-YxC2Nr02';
const USAGE_APP = 'hris-biologs';
const USAGE_EMAIL_KEY = 'hris-biologs-usage-email';

function usageEnabled() {
  try {
    if (USAGE_TOKEN.indexOf('PASTE_') === 0) return false; // not configured yet
    const h = location.hostname;
    return (location.protocol === 'http:' || location.protocol === 'https:') && h && h !== 'localhost' && h !== '127.0.0.1';
  } catch (e) { return false; }
}

function sendUsage(event) {
  if (!usageEnabled()) return;
  let email; try { email = localStorage.getItem(USAGE_EMAIL_KEY); } catch (e) { return; }
  if (!email) return;
  try {
    const p = fetch(USAGE_URL, { method: 'POST', mode: 'no-cors', keepalive: true,
      headers: { 'Content-Type': 'text/plain' },
      body: JSON.stringify({ token: USAGE_TOKEN, app: USAGE_APP, email, event }) });
    if (p && p.catch) p.catch(() => {});
  } catch (e) {}
}

// One-time work-email prompt; resolves to the full lowercased email, or null if cancelled.
function askWorkEmail() {
  return new Promise((resolve) => {
    const ov = document.createElement('div');
    ov.style.cssText = 'position:fixed;inset:0;background:rgba(0,0,0,.45);display:flex;align-items:center;justify-content:center;z-index:9999;';
    const box = document.createElement('div');
    box.style.cssText = 'background:#fff;border-radius:12px;padding:24px;max-width:380px;width:90%;box-shadow:0 12px 40px rgba(0,0,0,.25);font-family:sans-serif;';
    box.innerHTML = '<h3 style="font-size:1.05rem;font-weight:700;margin:0 0 6px;color:#2d3748">Enter your work email to continue</h3>' +
      '<p style="font-size:.84rem;color:#718096;margin:0 0 14px">We record your email to monitor who is using the tool. Your data is never collected or uploaded — only your email and when you run an export.</p>';
    const inp = document.createElement('input'); inp.type = 'email'; inp.placeholder = 'you@company.com';
    inp.style.cssText = 'width:100%;padding:9px 11px;border:1px solid #cbd5e0;border-radius:8px;font-size:.9rem;box-sizing:border-box;';
    const err = document.createElement('div'); err.style.cssText = 'color:#e11d48;font-size:.78rem;min-height:18px;margin:6px 0 10px;';
    const row = document.createElement('div'); row.style.cssText = 'display:flex;gap:8px;justify-content:flex-end;';
    const cancel = document.createElement('button'); cancel.textContent = 'Cancel';
    const ok = document.createElement('button'); ok.textContent = 'Continue';
    [cancel, ok].forEach((b) => { b.style.cssText = 'padding:8px 14px;border-radius:8px;border:1px solid #00A651;cursor:pointer;font-size:.85rem;'; });
    cancel.style.background = '#fff'; cancel.style.color = '#00A651'; ok.style.background = '#00A651'; ok.style.color = '#fff';
    row.append(cancel, ok); box.append(inp, err, row); ov.appendChild(box); document.body.appendChild(ov);
    inp.focus();
    const close = (v) => { ov.remove(); resolve(v); };
    cancel.onclick = () => close(null);
    const submit = () => {
      const v = (inp.value || '').trim();
      if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(v)) { err.textContent = 'Please enter a valid work email.'; return; }
      close(v.toLowerCase());
    };
    ok.onclick = submit;
    inp.addEventListener('keydown', (e) => { if (e.key === 'Enter') submit(); });
  });
}

// Call before an export. Resolves true to proceed, false if the user cancelled the
// email prompt (export should abort; they can simply click again). Always true when
// tracking is off (file://, localhost, or not configured) or storage is unavailable.
async function ensureUsageEmail() {
  if (!usageEnabled()) return true;
  let email;
  try { email = localStorage.getItem(USAGE_EMAIL_KEY); } catch (e) { return true; }
  if (email) return true;
  const entered = await askWorkEmail();
  if (!entered) return false;
  try { localStorage.setItem(USAGE_EMAIL_KEY, entered); } catch (e) {}
  return true;
}
