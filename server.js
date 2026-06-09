const express = require('express');
const fetch = require('node-fetch');
const XLSX = require('xlsx');
const path = require('path');

const app = express();
const PORT = process.env.PORT || 3000;
const EXCEL_URL = process.env.EXCEL_URL || 'https://water-tech.cboost.nl/excel';

const SENSOR_HEIGHTS = {
  'VDB45': 29,
  'VDB39': 100,
  'VDB37': 130,
  'VDB36': 207,
  'VDB14': 283
};

let cache = null;
let cacheTime = 0;
const CACHE_TTL = 20 * 60 * 1000;

function toDate(raw) {
  if (!raw && raw !== 0) return null;
  if (raw instanceof Date) return raw;
  if (typeof raw === 'number') {
    const d = XLSX.SSF.parse_date_code(raw);
    if (!d) return null;
    return new Date(d.y, d.m - 1, d.d, d.H || 0, d.M || 0, d.S || 0);
  }
  if (typeof raw === 'string') {
    const ts = new Date(raw);
    return isNaN(ts) ? null : ts;
  }
  return null;
}

async function fetchAndParse() {
  console.log('Fetching:', EXCEL_URL);
  const response = await fetch(EXCEL_URL, {
    headers: { 'User-Agent': 'Mozilla/5.0', 'Accept': '*/*' },
    redirect: 'follow',
    timeout: 90000
  });

  console.log('Status:', response.status);
  if (!response.ok) throw new Error(`HTTP ${response.status}`);

  const buffer = await response.buffer();
  console.log('Buffer:', buffer.length, 'bytes — parsing...');

  const wb = XLSX.read(buffer, {
    type: 'buffer',
    cellDates: false,
    cellNF: false,
    cellStyles: false,
    cellHTML: false,
    sheets: Object.keys(SENSOR_HEIGHTS)
  });

  const now = new Date();
  const cutoff = new Date(now);
  cutoff.setDate(cutoff.getDate() - 14);
  cutoff.setHours(0, 0, 0, 0);
  console.log('Cutoff:', cutoff.toISOString());

  const result = {};

  for (const [sensor, height] of Object.entries(SENSOR_HEIGHTS)) {
    if (!wb.Sheets[sensor]) { console.log('Missing:', sensor); continue; }

    const rows = XLSX.utils.sheet_to_json(wb.Sheets[sensor], { header: 1, raw: true });
    if (rows.length < 2) { console.log(sensor, ': empty'); continue; }

    const headers = (rows[0] || []).map(h => String(h || '').toLowerCase().trim());
    const tsIdx   = headers.indexOf('timestamp');
    const tempIdx = headers.indexOf('temperature');
    const humIdx  = headers.indexOf('humidity');

    if (tsIdx < 0 || tempIdx < 0 || humIdx < 0) {
      console.log(sensor, 'headers:', headers.slice(0, 5)); continue;
    }

    if (rows[1]) {
      console.log(sensor, 'row1 sample:', rows[1][tsIdx], typeof rows[1][tsIdx]);
    }

    const byDay = {};
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      if (!row || row[tsIdx] == null) continue;

      const ts = toDate(row[tsIdx]);
      if (!ts || isNaN(ts) || ts < cutoff) continue;

      const dateKey = ts.toISOString().slice(0, 10);
      const hourFrac = Math.round((ts.getHours() + ts.getMinutes() / 60) * 1000) / 1000;
      const temp = Math.round(parseFloat(row[tempIdx]) * 10) / 10;
      const hum  = Math.round(parseFloat(row[humIdx])  * 10) / 10;
      if (isNaN(temp) || isNaN(hum)) continue;

      if (!byDay[dateKey]) byDay[dateKey] = { hrs: [], temp: [], hum: [] };
      byDay[dateKey].hrs.push(hourFrac);
      byDay[dateKey].temp.push(temp);
      byDay[dateKey].hum.push(hum);
    }

    for (const [day, vals] of Object.entries(byDay)) {
      if (!result[day]) result[day] = {};
      result[day][sensor] = { height, ...vals };
    }
    console.log(sensor, 'OK:', Object.keys(byDay).length, 'days');
  }

  return result;
}

app.get('/api/debug', async (req, res) => {
  try {
    const response = await fetch(EXCEL_URL, {
      headers: { 'User-Agent': 'Mozilla/5.0' },
      redirect: 'follow', timeout: 30000
    });
    const buf = await response.buffer();
    res.json({ status: response.status, contentType: response.headers.get('content-type'), bufferSize: buf.length });
  } catch (err) {
    res.json({ error: err.message });
  }
});

app.get('/api/data', async (req, res) => {
  try {
    const now = Date.now();
    if (!cache || now - cacheTime > CACHE_TTL) {
      cache = await fetchAndParse();
      cacheTime = now;
    }
    res.json({ ok: true, fetched: new Date(cacheTime).toISOString(), data: cache });
  } catch (err) {
    console.error('Error:', err.message);
    res.status(500).json({ ok: false, error: err.message });
  }
});

app.use(express.static(path.join(__dirname, 'public')));

app.listen(PORT, () => console.log(`Port ${PORT} | EXCEL_URL=${EXCEL_URL}`));
