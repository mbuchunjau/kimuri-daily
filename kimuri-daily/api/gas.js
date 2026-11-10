// Vercel serverless function: /api/gas
// The browser only ever talks to this endpoint. It checks the cashier PIN,
// then forwards the request to the Apps Script web app with a secret key.
//
// Environment variables (Vercel > Project > Settings > Environment Variables):
//   GAS_URL  = the Apps Script web app URL (ends in /exec)
//   GAS_KEY  = same value as the API_KEY script property in Apps Script
//   APP_PIN  = the PIN the cashier types to unlock the app

const crypto = require('crypto');

const ALLOWED = new Set([
  'getDay', 'saveDay', 'getSummary', 'sendSummary', 'getReport', 'listDays', 'sendUnsent'
]);

function safeEqual(a, b) {
  const ha = crypto.createHash('sha256').update(String(a)).digest();
  const hb = crypto.createHash('sha256').update(String(b)).digest();
  return crypto.timingSafeEqual(ha, hb);
}

module.exports = async function handler(req, res) {
  res.setHeader('Cache-Control', 'no-store');

  if (req.method !== 'POST') {
    return res.status(405).json({ ok: false, error: 'Method not allowed' });
  }

  const { GAS_URL, GAS_KEY, APP_PIN } = process.env;
  if (!GAS_URL || !GAS_KEY || !APP_PIN) {
    return res.status(500).json({ ok: false, error: 'Server is not configured (missing environment variables).' });
  }

  let body = req.body;
  if (typeof body === 'string') {
    try { body = JSON.parse(body); } catch (e) { body = {}; }
  }
  body = body || {};
  const { action, pin, data } = body;

  if (!pin || !safeEqual(pin, APP_PIN)) {
    return res.status(401).json({ ok: false, error: 'Wrong PIN.' });
  }
  if (action === 'ping') {
    return res.status(200).json({ ok: true, data: 'pong' });
  }
  if (!ALLOWED.has(action)) {
    return res.status(400).json({ ok: false, error: 'Unknown action.' });
  }

  try {
    const upstream = await fetch(GAS_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({ key: GAS_KEY, action, data: data || {} }),
      redirect: 'follow'
    });
    const text = await upstream.text();
    let json;
    try {
      json = JSON.parse(text);
    } catch (e) {
      return res.status(502).json({
        ok: false,
        error: 'Apps Script did not return JSON. Check that the web app is deployed with access set to "Anyone" and that GAS_URL ends in /exec.'
      });
    }
    return res.status(200).json(json);
  } catch (err) {
    return res.status(502).json({ ok: false, error: 'Could not reach Apps Script: ' + (err && err.message ? err.message : err) });
  }
};
