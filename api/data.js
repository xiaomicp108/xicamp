// Vercel serverless function: pulls the WOD workbook from SharePoint,
// parses it, and returns compact JSON. Edge-cached for 5 minutes so
// SharePoint is checked for updates at most every 5 minutes.
'use strict';
const { parseWOD } = require('../lib/parse-wod');

const SOURCE_URL = process.env.WOD_SOURCE_URL ||
  'https://teamchannelplay-my.sharepoint.com/:x:/g/personal/krishna_kumar2_channelplay_in/IQBLKWz7K79vR5fuLfm_2dUrATdb7KG_TSb2i_Ze4uzjMQs?e=msqM8W&download=1';

// Anonymous SharePoint share links redirect through a cookie-setting hop,
// so follow redirects manually and carry cookies along.
async function fetchWithCookies(url, maxHops = 10) {
  const jar = new Map();
  let current = url;
  for (let hop = 0; hop < maxHops; hop++) {
    const cookie = [...jar].map(([k, v]) => `${k}=${v}`).join('; ');
    const res = await fetch(current, {
      redirect: 'manual',
      headers: {
        'User-Agent': 'Mozilla/5.0 (WOD-Dashboard)',
        'Accept': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/octet-stream,*/*',
        ...(cookie ? { Cookie: cookie } : {})
      }
    });
    const setCookies = typeof res.headers.getSetCookie === 'function' ? res.headers.getSetCookie() : [];
    for (const c of setCookies) {
      const [pair] = c.split(';');
      const eq = pair.indexOf('=');
      if (eq > 0) jar.set(pair.slice(0, eq).trim(), pair.slice(eq + 1).trim());
    }
    if (res.status >= 300 && res.status < 400 && res.headers.get('location')) {
      current = new URL(res.headers.get('location'), current).toString();
      continue;
    }
    if (!res.ok) throw new Error(`SharePoint responded ${res.status}`);
    const buf = Buffer.from(await res.arrayBuffer());
    // xlsx files are zips: they start with "PK"
    if (buf[0] !== 0x50 || buf[1] !== 0x4b) {
      throw new Error('SharePoint returned a web page instead of the Excel file (link may require sign-in or has expired)');
    }
    return buf;
  }
  throw new Error('Too many redirects from SharePoint');
}

module.exports = async function handler(req, res) {
  try {
    const buf = await fetchWithCookies(SOURCE_URL);
    const data = parseWOD(buf);
    data.fetchedAt = new Date().toISOString();
    data.origin = 'sharepoint';
    res.setHeader('Content-Type', 'application/json; charset=utf-8');
    res.setHeader('Cache-Control', 'public, s-maxage=300, stale-while-revalidate=60, max-age=0');
    res.status(200).send(JSON.stringify(data));
  } catch (err) {
    res.setHeader('Cache-Control', 'no-store');
    res.status(502).json({ error: String(err && err.message || err) });
  }
};
