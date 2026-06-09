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
  console.log('Fetching Excel from:', EXCEL_URL);

  const response = await fetch(EXCEL_URL, {
    headers: {
      'User-Agent': 'Mozilla/5.0 (compatible; SensorDashboard/1.0)',
      'Accept': '*/*'
    },
    redirect: 'follow'
  });

  console.log('Response status:', response.status);
  console.log('Content-Type:', response.headers.get('content-type'));

  if (!response.ok) {
    const text = await response.text();
    throw new Error(`HTTP ${response.status}: ${text.slice(0, 200)}`);
  }

  const contentType = response.headers.get('content-type') || '';
  if (contentType.includes('text/html')) {
    const text = await response.text();
    throw new Error(`Server returned HTML instead of Excel. First 200 chars: ${text.slice(0, 200)}`);
  }

  const buffer = await response.buffer();
  console.log('Buffer size:', buffer.length, 'bytes');

  const wb = XLSX.read(buffer, { type: 'buffer', cellDates: true });
  console.log('Sheets found:', wb.SheetNames.slice(0, 10).join(', '), '...');

  const now = new Date();
  const cutoff = new Date(now - 14 * 24 * 60 * 60 * 1000);
  const result = {};

  for (const [sensor, height] of Object.entries(SENSOR_HEIGHTS)) {
    if (!wb.SheetNames.includes(sensor)) {
      console.log('Sheet not found:', sensor);
      continue;
    }

    const ws = wb.Sheets[sensor];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });
    if (rows.length < 2) continue;

    const headers = rows[0].map(h => String(h).toLowerCase().trim());
    const tsIdx = headers.indexOf('timestamp');
    const tempIdx = headers.indexOf('temperature');
    const humIdx = headers.indexOf('humidity');

    if (tsIdx < 0 || tempIdx < 0 || humIdx < 0) {
      console.log('Missing columns in', sensor, ':', headers);
      continue;
    }

    const byDay = {};
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      if (!row[tsIdx]) continue;

      let ts;
      const raw = row[tsIdx];
      if (raw instanceof Date) {
        ts = raw;
      } else if (typeof raw === 'number') {
        const parsed = XLSX.SSF.parse_date_code(raw);
        ts = new Date(parsed.y, parsed.m - 1, parsed.d, parsed.H, parsed.M, parsed.S);
      } else {
        ts = new Date(raw);
      }

      if (isNaN(ts.getTime()) || ts < cutoff) continue;

      const dateKey = ts.toISOString().slice(0, 10);
      const hourFrac = ts.getHours() + ts.getMinutes() / 60;
      const temp = parseFloat(row[tempIdx]);
      const hum = parseFloat(row[humIdx]);
      if (isNaN(temp) || isNaN(hum)) continue;

      if (!byDay[dateKey]) byDay[dateKey] = { hrs: [], temp: [], hum: [] };
      byDay[dateKey].hrs.push(Math.round(hourFrac * 1000) / 1000);
      byDay[dateKey].temp.push(Math.round(temp * 10) / 10);
      byDay[dateKey].hum.push(Math.round(hum * 10) / 10);
    }

    for (const [day, vals] of Object.entries(byDay)) {
      if (!result[day]) result[day] = {};
      result[day][sensor] = { height, ...vals };
    }

    console.log(sensor, ': processed', Object.keys(byDay).length, 'days');
  }

  return result;
}

// Debug endpoint
app.get('/api/debug', async (req, res) => {
  try {
    const response = await fetch(EXCEL_URL, {
      headers: { 'User-Agent': 'Mozilla/5.0', 'Accept': '*/*' },
      redirect: 'follow'
    });
    const contentType = response.headers.get('content-type');
    const buffer = await response.buffer();
    res.json({
      status: response.status,
      contentType,
      bufferSize: buffer.length,
      first100bytes: buffer.slice(0, 100).toString('hex')
    });
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
    console.error('Error in /api/data:', err.message);
    res.status(500).json({ ok: false, error: err.message });
  }
});

app.use(express.static(path.join(__dirname, 'public')));

app.listen(PORT, () => console.log(`Server running on port ${PORT}, EXCEL_URL=${EXCEL_URL}`));
