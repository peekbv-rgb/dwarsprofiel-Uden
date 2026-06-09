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

async function fetchAndParse() {
  console.log('Fetching:', EXCEL_URL);
  const response = await fetch(EXCEL_URL, {
    headers: { 'User-Agent': 'Mozilla/5.0', 'Accept': '*/*' },
    redirect: 'follow',
    timeout: 90000
  });

  console.log('Status:', response.status);
  if (!response.ok) throw new Error(`HTTP ${response.status}`);

  const contentType = response.headers.get('content-type') || '';
  if (contentType.includes('text/html')) {
    const text = await response.text();
    throw new Error(`Got HTML: ${text.slice(0, 200)}`);
  }

  const buffer = await response.buffer();
  console.log('Buffer:', buffer.length, 'bytes — parsing...');

  const wb = XLSX.read(buffer, {
    type: 'buffer',
    cellDates: true,
    cellNF: false,
    cellStyles: false,
    cellHTML: false,
    sheets: Object.keys(SENSOR_HEIGHTS),
    dense: true
  });

  buffer.fill(0);

  const now = new Date();
  const cutoff = new Date(now);
  cutoff.setDate(cutoff.getDate() - 14);
  cutoff.setHours(0, 0, 0, 0);
  console.log('Cutoff:', cutoff.toISOString());

  const result = {};

  for (const [sensor, height] of Object.entries(SENSOR_HEIGHTS)) {
    if (!wb.Sheets[sensor]) { console.log('Missing:', sensor); continue; }

    const ws = wb.Sheets[sensor];
    const rowsFlat = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false });
    const headers = (rowsFlat[0] || []).map(h => String(h).toLowerCase().trim());
    const tsIdx = headers.indexOf('timestamp');
    const tempIdx = headers.indexOf('temperature');
    const humIdx = headers.indexOf('humidity');

    if (tsIdx < 0 || tempIdx < 0 || humIdx < 0) {
      console.log(sensor, 'missing columns:', headers); continue;
    }

    const byDay = {};
    for (let i = 1; i < rowsFlat.length; i++) {
      const row = rowsFlat[i];
      if (!row[tsIdx]) continue;
      const ts = new Date(row[tsIdx]);
      if (isNaN(ts) || ts < cutoff) continue;

      const dateKey = ts.toISOString().slice(0, 10);
      const hourFrac = Math.round((ts.getHours() + ts.getMinutes() / 60) * 1000) / 1000;
      const temp = Math.round(parseFloat(row[tempIdx]) * 10) / 10;
      const hum = Math.round(parseFloat(row[humIdx]) * 10) / 10;
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
    const ct = response.headers.get('content-type');
    const buf = await response.buffer();
    res.json({ status: response.status, contentType: ct, bufferSize: buf.length });
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

app.listen(PORT, () => {
  console.log(`Port ${PORT} | EXCEL_URL=${EXCEL_URL}`);
});
