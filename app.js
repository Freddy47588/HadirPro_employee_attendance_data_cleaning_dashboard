/* ════════════════════════════════════════════
   HADIR Pro — app.js
   Attendance Data Cleaning System
   v2.1 — XLS Support for feb.xls format
════════════════════════════════════════════ */

'use strict';

// ══════════════════════════════════
// STATE
// ══════════════════════════════════
let state = {
  headers: [],
  rawData: [],
  cleanData: [],
  issues: [],
  charts: {},
  sourceFile: null
};


// ══════════════════════════════════
// XLS PARSER — format feb.xls
// ══════════════════════════════════

/**
 * Extract all HH:MM time patterns from a concatenated cell value.
 * e.g. "07:2016:03" → ["07:20","16:03"]
 *      "07:0916:1716:17" → ["07:09","16:17","16:17"]
 */
function extractTimes(cell) {
  const str = String(cell || '').trim();
  return str.match(/\d{2}:\d{2}/g) || [];
}

/**
 * Parse the "Lap. Log Absen" sheet (matrix format) into flat attendance rows.
 * The sheet has:
 *  - Row 2: date range header e.g. "2026-02-01 ~ 2026-02-28"
 *  - Row 3: day numbers [1, 2, ..., 28]
 *  - Then pairs of rows per employee:
 *      info row: ['ID:', '', empId, ..., 'Nama:', '', empName, ..., 'Dept.:', '', empDept]
 *      time row: [times for each day column]
 */
function parseFebLogAbsen(rawMatrix) {
  // Detect date range from row 2, col 2
  let year = 2026, month = 2;
  const dateRangeCell = String((rawMatrix[2] || [])[2] || '');
  const dmatch = dateRangeCell.match(/(\d{4})-(\d{2})/);
  if (dmatch) { year = parseInt(dmatch[1]); month = parseInt(dmatch[2]); }

  // Day count from row 3
  const dayRow = rawMatrix[3] || [];
  const dayCount = dayRow.filter(d => d !== '' && d !== null).length;

  const headers = ['ID', 'Nama', 'Departemen', 'Tanggal', 'Jam Masuk', 'Jam Keluar'];
  const rows = [];
  const issues_xls = []; // raw issues found during parse

  let i = 4;
  while (i < rawMatrix.length) {
    const infoRow = rawMatrix[i] || [];

    if (String(infoRow[0]).trim() === 'ID:') {
      const empId   = String(infoRow[2]  || '').trim();
      const empName = String(infoRow[10] || '').trim();
      const empDept = String(infoRow[20] || '').trim();

      const timeRow = (i + 1 < rawMatrix.length) ? (rawMatrix[i + 1] || []) : [];

      for (let dayIdx = 0; dayIdx < dayCount; dayIdx++) {
        const cell = String(timeRow[dayIdx] || '').trim();
        if (!cell) continue; // absent / weekend / holiday → skip

        const times = extractTimes(cell);
        if (times.length === 0) continue;

        const day = dayIdx + 1;
        const dateStr = `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;

        let checkIn  = '';
        let checkOut = '';
        let punchNote = '';

        if (times.length === 1) {
          // Single punch — flag as issue
          const t = times[0];
          const h = parseInt(t.split(':')[0]);
          if (h < 12) {
            checkIn  = t;
            punchNote = 'single-in';
          } else {
            checkOut = t;
            punchNote = 'single-out';
          }
        } else {
          // Multiple times: first = earliest check-in, last = latest check-out
          // Separate by time-of-day: <12 = check-in candidates, >=12 = check-out candidates
          const ins  = times.filter(t => parseInt(t.split(':')[0]) < 12);
          const outs = times.filter(t => parseInt(t.split(':')[0]) >= 12);
          checkIn  = ins.length  ? ins[0]               : times[0];
          checkOut = outs.length ? outs[outs.length - 1] : times[times.length - 1];
        }

        rows.push({
          _origIdx: rows.length,
          _punchNote: punchNote,
          'ID'         : empId,
          'Nama'       : empName,
          'Departemen' : empDept,
          'Tanggal'    : dateStr,
          'Jam Masuk'  : checkIn,
          'Jam Keluar' : checkOut
        });
      }

      i += 2; // skip the time row
    } else {
      i++;
    }
  }

  return { hdrs: headers, rows };
}

/**
 * Read an XLS/XLSX ArrayBuffer using SheetJS and return parsed flat rows.
 */
function parseXLSBuffer(buffer) {
  const wb = XLSX.read(new Uint8Array(buffer), { type: 'array' });

  // Prefer the attendance log sheet
  const targetName = wb.SheetNames.find(n =>
    /log.absen|lap.*log|absen.*log/i.test(n)
  ) || wb.SheetNames.find(n => /absen/i.test(n)) || wb.SheetNames[0];

  const ws  = wb.Sheets[targetName];
  const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  return parseFebLogAbsen(raw);
}

// ══════════════════════════════════
// COLUMN DETECTION
// ══════════════════════════════════
function detectCol(headers, patterns) {
  return headers.find(h => patterns.some(p => p.test(h.toLowerCase())));
}

function getColNames(headers) {
  return {
    id:   detectCol(headers, [/^id$|nip|nik|kode/]),
    name: detectCol(headers, [/nama|name/]),
    dept: detectCol(headers, [/dept|divisi|bagian/]),
    date: detectCol(headers, [/tanggal|tgl|date/]),
    in:   detectCol(headers, [/masuk|in$|clock.in|check.in/]),
    out:  detectCol(headers, [/keluar|pulang|out$|clock.out|check.out/])
  };
}

// ══════════════════════════════════
// UTILITIES
// ══════════════════════════════════
function parseCSV(text) {
  const lines = text.trim().split('\n').map(l => l.trim()).filter(Boolean);
  const hdrs = lines[0].split(',').map(h => h.trim());
  const rows = lines.slice(1).map((line, idx) => {
    const vals = line.split(',').map(v => v.trim());
    const obj  = { _origIdx: idx };
    hdrs.forEach((h, i) => obj[h] = vals[i] !== undefined ? vals[i] : '');
    return obj;
  });
  return { hdrs, rows };
}

function normalizeTime(t) {
  if (!t) return '';
  t = t.trim().replace('.', ':');
  const m = t.match(/^(\d{1,2}):(\d{2})$/);
  if (m) return m[1].padStart(2, '0') + ':' + m[2];
  return t;
}

function timeToMins(t) {
  if (!t) return null;
  const m = t.match(/^(\d{2}):(\d{2})$/);
  if (!m) return null;
  return parseInt(m[1]) * 60 + parseInt(m[2]);
}

function minsToHM(mins) {
  if (mins === null || mins < 0) return '-';
  const h = Math.floor(mins / 60);
  const m = mins % 60;
  return `${h}j ${m}m`;
}

function isValidTime(t) {
  return /^\d{2}:\d{2}$/.test(t);
}

function cfg(id) { return document.getElementById(id).value.trim(); }

// ══════════════════════════════════
// DETECT ISSUES
// ══════════════════════════════════
function detectIssues(rows, cols) {
  const issues = [];
  const seen   = new Set();

  rows.forEach((row, i) => {
    const keyFull = (row[cols.id] || '') + '§' + (row[cols.date] || '');

    // Single punch — flagged during XLS parse
    if (row._punchNote === 'single-in') {
      issues.push({ row: i, type: 'miss', msg: 'Hanya ada jam masuk (tidak ada jam keluar)', detail: `${row[cols.name] || row[cols.id]} — ${row[cols.date]}` });
    }
    if (row._punchNote === 'single-out') {
      issues.push({ row: i, type: 'miss', msg: 'Hanya ada jam keluar (tidak ada jam masuk)', detail: `${row[cols.name] || row[cols.id]} — ${row[cols.date]}` });
    }

    // Duplicates
    if (cols.id && cols.date && keyFull !== '§') {
      if (seen.has(keyFull)) {
        issues.push({ row: i, type: 'dup', msg: 'Baris duplikat', detail: `${row[cols.id]} | ${row[cols.date]}` });
      } else {
        seen.add(keyFull);
      }
    }

    // Missing jam masuk
    if (cols.in && !row[cols.in] && !row._punchNote) {
      issues.push({ row: i, type: 'miss', msg: 'Jam masuk kosong', detail: `${row[cols.name] || row[cols.id]} — ${row[cols.date]}` });
    }

    // Missing jam keluar
    if (cols.out && !row[cols.out] && !row._punchNote) {
      issues.push({ row: i, type: 'miss', msg: 'Jam keluar kosong', detail: `${row[cols.name] || row[cols.id]} — ${row[cols.date]}` });
    }

    // Format jam tidak valid
    if (cols.in && row[cols.in] && !isValidTime(normalizeTime(row[cols.in]))) {
      issues.push({ row: i, type: 'fmt', msg: 'Format jam masuk tidak valid', detail: `Nilai: "${row[cols.in]}"` });
    }
    if (cols.out && row[cols.out] && !isValidTime(normalizeTime(row[cols.out]))) {
      issues.push({ row: i, type: 'fmt', msg: 'Format jam keluar tidak valid', detail: `Nilai: "${row[cols.out]}"` });
    }
  });

  return issues;
}

// ══════════════════════════════════
// QUALITY SCORE
// ══════════════════════════════════
function calcQuality(rows, issues) {
  if (!rows.length) return 0;
  const dups  = issues.filter(i => i.type === 'dup').length;
  const miss  = issues.filter(i => i.type === 'miss').length;
  const fmt   = issues.filter(i => i.type === 'fmt').length;
  const total = rows.length;
  const penalty = Math.min(1, (dups * 1 + miss * 0.7 + fmt * 0.5) / total);
  return Math.max(0, Math.round((1 - penalty) * 100));
}

// ══════════════════════════════════
// CLEAN DATA
// ══════════════════════════════════
function runClean() {
  const { headers: hdrs, rawData: rows } = state;
  const cols = getColNames(hdrs);
  const stdIn   = cfg('cfgStdIn');
  const stdOut  = cfg('cfgStdOut');
  const otLine  = cfg('cfgOT');
  const tolMins = parseInt(cfg('cfgTol')) || 15;

  let processed = rows.map(r => ({ ...r }));

  // 1. Normalize times
  if (document.getElementById('chkFmt').checked) {
    if (cols.in)  processed.forEach(r => r[cols.in]  = normalizeTime(r[cols.in]));
    if (cols.out) processed.forEach(r => r[cols.out] = normalizeTime(r[cols.out]));
  }

  // 2. Remove duplicates
  if (document.getElementById('chkDup').checked && cols.id && cols.date) {
    const seen = new Set();
    processed = processed.filter(r => {
      const key = (r[cols.id] || '') + '§' + (r[cols.date] || '');
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    });
  }

  // 3. Durasi kerja
  if (document.getElementById('chkDur').checked && cols.in && cols.out) {
    processed.forEach(r => {
      const dur = timeToMins(r[cols.out]) !== null && timeToMins(r[cols.in]) !== null
        ? timeToMins(r[cols.out]) - timeToMins(r[cols.in])
        : null;
      r['_Durasi Kerja'] = dur !== null && dur >= 0 ? minsToHM(dur) : '-';
    });
  }

  // 4. Keterlambatan
  if (document.getElementById('chkLate').checked && cols.in) {
    const stdMins = timeToMins(stdIn);
    processed.forEach(r => {
      if (!r[cols.in]) {
        r['_Status Kehadiran'] = 'Tidak Lengkap';
      } else {
        const inMins = timeToMins(r[cols.in]);
        if (inMins === null) {
          r['_Status Kehadiran'] = 'Format Error';
        } else {
          const diff = inMins - stdMins;
          if (diff <= 0) r['_Status Kehadiran'] = 'Tepat Waktu';
          else if (diff <= tolMins) r['_Status Kehadiran'] = `Toleransi (${diff}m)`;
          else r['_Status Kehadiran'] = `Terlambat (${diff}m)`;
        }
      }
    });
  }

  // 5. Lembur
  if (document.getElementById('chkOT').checked && cols.out) {
    const otMins     = timeToMins(otLine);
    const stdOutMins = timeToMins(stdOut);
    const baseline   = otMins || stdOutMins;
    processed.forEach(r => {
      if (!r[cols.out]) {
        r['_Jam Lembur'] = '-';
      } else {
        const outMins = timeToMins(r[cols.out]);
        if (outMins === null) {
          r['_Jam Lembur'] = '-';
        } else {
          const ot = Math.max(0, outMins - baseline);
          r['_Jam Lembur'] = ot > 0 ? minsToHM(ot) : '0j 0m';
        }
      }
    });
  }

  state.cleanData = processed;

  renderCleanTable(processed, hdrs, cols);
  renderCharts(processed, hdrs, cols);
  updateOverview(rows, state.issues, processed);
  updateTopbar(rows, state.issues);

  document.getElementById('btnDownload').disabled = false;
  document.getElementById('cleanCount').textContent = processed.length;
  switchSection('clean');
}

// ══════════════════════════════════
// RENDER RAW TABLE
// ══════════════════════════════════
function renderRawTable(rows, hdrs, issues) {
  const issueSet = new Set(issues.map(i => i.row));
  const dupSet   = new Set(issues.filter(i => i.type === 'dup').map(i => i.row));
  const missSet  = new Set(issues.filter(i => i.type === 'miss').map(i => i.row));
  const cols = getColNames(hdrs);

  let html = `<table><thead><tr><th>#</th>`;
  hdrs.forEach(h => html += `<th>${h}</th>`);
  html += `</tr></thead><tbody>`;

  rows.forEach((row, i) => {
    const cls = dupSet.has(i) ? 'row-dup' : missSet.has(i) ? 'row-miss' : '';
    html += `<tr class="${cls}">`;
    html += `<td style="color:var(--muted);font-size:11px">${i + 1}</td>`;
    hdrs.forEach(h => {
      const isMissIn  = (h === cols.in  && !row[h]);
      const isMissOut = (h === cols.out && !row[h]);
      const isBadFmt  = (h === cols.in  || h === cols.out) && row[h] && !isValidTime(normalizeTime(row[h]));
      const isName    = h === cols.name;
      const val       = row[h] || '';
      const extra     = isMissIn || isMissOut ? ' <span class="chip chip-err">kosong</span>'
                      : isBadFmt             ? ' <span class="chip chip-warn">format</span>' : '';
      html += `<td class="${isName ? 'name-col' : ''}">${val}${extra}</td>`;
    });
    html += `</tr>`;
  });

  html += `</tbody></table>`;
  return html;
}

// ══════════════════════════════════
// RENDER CLEAN TABLE
// ══════════════════════════════════
function renderCleanTable(rows, hdrs, cols) {
  const extraKeys = rows.length ? Object.keys(rows[0]).filter(k => k.startsWith('_') && k !== '_origIdx' && k !== '_punchNote') : [];
  const allHdrs   = [...hdrs, ...extraKeys];

  let html = `<table><thead><tr><th>#</th>`;
  allHdrs.forEach(h => html += `<th>${h.replace(/^_/, '')}</th>`);
  html += `</tr></thead><tbody>`;

  rows.forEach((row, i) => {
    html += `<tr>`;
    html += `<td style="color:var(--muted);font-size:11px">${i + 1}</td>`;
    allHdrs.forEach(h => {
      const isName = h === cols.name;
      const val    = row[h] || '';
      let cell     = val;

      if (h === '_Status Kehadiran') {
        if (val.includes('Tepat'))     cell = `<span class="chip chip-ok">${val}</span>`;
        else if (val.includes('Toleransi')) cell = `<span class="chip chip-info">${val}</span>`;
        else if (val.includes('Terlambat')) cell = `<span class="chip chip-warn">${val}</span>`;
        else                           cell = `<span class="chip chip-err">${val}</span>`;
      }
      if (h === '_Jam Lembur' && val && val !== '-' && val !== '0j 0m') {
        cell = `<span class="chip chip-ot">${val}</span>`;
      }

      html += `<td class="${isName ? 'name-col' : ''}">${cell}</td>`;
    });
    html += `</tr>`;
  });

  html += `</tbody></table>`;

  document.getElementById('cleanTableWrap').innerHTML = html;
  document.getElementById('cleanTableInfo').textContent = `${rows.length} baris`;
}

// ══════════════════════════════════
// RENDER ISSUES
// ══════════════════════════════════
function renderIssues(issues) {
  const dups = issues.filter(i => i.type === 'dup').length;
  const miss = issues.filter(i => i.type === 'miss').length;
  const fmt  = issues.filter(i => i.type === 'fmt').length;

  let html = `<div class="issue-summary-bar">
    <div class="iss-sum-card">
      <div class="kpi-icon yellow" style="width:28px;height:28px"><svg viewBox="0 0 20 20" width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><path d="M10 2l1.5 3h3l-2.5 2 1 3-3-2-3 2 1-3L5.5 5h3z"/></svg></div>
      <div><div class="iss-sum-num" style="color:var(--yellow)">${dups}</div><div class="iss-sum-lbl">Duplikat</div></div>
    </div>
    <div class="iss-sum-card">
      <div class="kpi-icon red" style="width:28px;height:28px"><svg viewBox="0 0 20 20" width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><circle cx="10" cy="10" r="8"/><path d="M10 6v5M10 14v.5"/></svg></div>
      <div><div class="iss-sum-num" style="color:var(--red)">${miss}</div><div class="iss-sum-lbl">Data Kosong</div></div>
    </div>
    <div class="iss-sum-card">
      <div class="kpi-icon blue" style="width:28px;height:28px"><svg viewBox="0 0 20 20" width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><path d="M4 6h12M4 10h8M4 14h10"/></svg></div>
      <div><div class="iss-sum-num" style="color:var(--blue)">${fmt}</div><div class="iss-sum-lbl">Format Salah</div></div>
    </div>
  </div>`;

  if (!issues.length) {
    html += `<div class="issue-item"><div class="issue-icon" style="background:var(--green-dim);color:var(--green)">✓</div><div class="issue-body"><div class="issue-title">Tidak ada masalah ditemukan</div><div class="issue-detail">Seluruh data terlihat bersih</div></div></div>`;
  } else {
    const typeMap  = { dup: 'dup', miss: 'miss', fmt: 'fmt' };
    const iconMap  = { dup: '⊘', miss: '!', fmt: '~' };
    issues.forEach(iss => {
      html += `<div class="issue-item">
        <div class="issue-icon ${typeMap[iss.type] || ''}">${iconMap[iss.type] || '?'}</div>
        <div class="issue-body">
          <div class="issue-title">${iss.msg}</div>
          <div class="issue-detail">${iss.detail}</div>
        </div>
        <div class="issue-row-badge">Baris ${iss.row + 1}</div>
      </div>`;
    });
  }

  document.getElementById('issueList').innerHTML = html;
}

// ══════════════════════════════════
// UPDATE OVERVIEW DASHBOARD
// ══════════════════════════════════
function updateOverview(rows, issues, cleanRows) {
  const cols = getColNames(state.headers);
  const q    = calcQuality(rows, issues);

  const scoreEl = document.getElementById('qualScoreBig');
  const fillEl  = document.getElementById('qualBarFill');
  scoreEl.textContent = q + '%';
  scoreEl.style.color = q >= 80 ? 'var(--green)' : q >= 50 ? 'var(--yellow)' : 'var(--red)';
  fillEl.style.width  = q + '%';

  const dups = issues.filter(i => i.type === 'dup').length;
  const miss = issues.filter(i => i.type === 'miss').length;
  const fmt  = issues.filter(i => i.type === 'fmt').length;
  const ok   = rows.length - new Set(issues.map(i => i.row)).size;

  document.getElementById('qbDup').textContent  = dups;
  document.getElementById('qbMiss').textContent = miss;
  document.getElementById('qbFmt').textContent  = fmt;
  document.getElementById('qbOK').textContent   = ok;

  document.getElementById('kpiTotal').textContent  = rows.length;
  document.getElementById('kpiIssues').textContent = issues.length;

  if (cleanRows && cleanRows.length && cleanRows[0]['_Status Kehadiran']) {
    const lateRows    = cleanRows.filter(r => r['_Status Kehadiran'] && r['_Status Kehadiran'].includes('Terlambat'));
    const lateMins    = lateRows.map(r => { const m = r['_Status Kehadiran'].match(/(\d+)m/); return m ? parseInt(m[1]) : 0; });
    const avg         = lateMins.length ? Math.round(lateMins.reduce((a, b) => a + b, 0) / lateMins.length) : 0;
    document.getElementById('kpiLate').textContent = avg ? avg + ' mnt' : '—';
  } else {
    document.getElementById('kpiLate').textContent = '—';
  }

  if (cleanRows && cleanRows.length && cleanRows[0]['_Jam Lembur'] !== undefined) {
    let totalOT = 0;
    cleanRows.forEach(r => {
      if (!r['_Jam Lembur'] || r['_Jam Lembur'] === '-') return;
      const hm = r['_Jam Lembur'].match(/(\d+)j (\d+)m/);
      if (hm) totalOT += parseInt(hm[1]) * 60 + parseInt(hm[2]);
    });
    const h = Math.floor(totalOT / 60);
    document.getElementById('kpiOT').textContent = h ? h + ' jam' : '0 jam';
  } else {
    document.getElementById('kpiOT').textContent = '—';
  }

  renderIssueDonut(dups, miss, fmt, ok);
  renderArrivalChart(rows, cols);

  document.getElementById('overviewEmpty').style.display = 'none';
  document.getElementById('overviewDash').style.display  = 'block';
}

function updateTopbar(rows, issues) {
  const q = calcQuality(rows, issues);
  document.getElementById('qpScore').textContent = q + '%';
  document.getElementById('qpScore').style.color = q >= 80 ? 'var(--green)' : q >= 50 ? 'var(--yellow)' : 'var(--red)';
  const fname = state.sourceFile ? ` · ${state.sourceFile}` : '';
  document.getElementById('pageSub').textContent = `${rows.length} baris dimuat · ${issues.length} masalah ditemukan${fname}`;
}

// ══════════════════════════════════
// CHARTS
// ══════════════════════════════════
const CHART_COLORS = {
  green:  '#10B981',
  blue:   '#3B82F6',
  yellow: '#F59E0B',
  red:    '#EF4444',
  orange: '#F97316',
  muted:  '#536070',
  border: '#253550',
  text2:  '#8fa3c0',
  bg3:    '#162035'
};

function destroyChart(id) {
  if (state.charts[id]) { state.charts[id].destroy(); delete state.charts[id]; }
}

function renderIssueDonut(dups, miss, fmt, ok) {
  destroyChart('issueDonut');
  const canvas = document.getElementById('chartIssueDonut');
  if (!canvas) return;
  state.charts['issueDonut'] = new Chart(canvas, {
    type: 'doughnut',
    data: {
      labels: ['Duplikat', 'Data Kosong', 'Format Salah', 'Valid'],
      datasets: [{
        data: [dups, miss, fmt, ok],
        backgroundColor: [CHART_COLORS.yellow, CHART_COLORS.red, CHART_COLORS.orange, CHART_COLORS.green],
        borderColor: CHART_COLORS.border,
        borderWidth: 2
      }]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      cutout: '65%',
      plugins: {
        legend: { display: true, position: 'right', labels: { color: CHART_COLORS.text2, font: { size: 11, family: 'DM Sans' }, boxWidth: 10, padding: 10 } }
      }
    }
  });
}

function renderArrivalChart(rows, cols) {
  destroyChart('arrival');
  const canvas = document.getElementById('chartArrival');
  if (!canvas || !cols.in) return;

  const buckets = {};
  rows.forEach(r => {
    const t = normalizeTime(r[cols.in]);
    if (!isValidTime(t)) return;
    const h = t.slice(0, 2) + ':00';
    buckets[h] = (buckets[h] || 0) + 1;
  });
  const labels = Object.keys(buckets).sort();
  const vals   = labels.map(l => buckets[l]);

  state.charts['arrival'] = new Chart(canvas, {
    type: 'bar',
    data: {
      labels,
      datasets: [{ label: 'Jumlah karyawan', data: vals, backgroundColor: CHART_COLORS.blue + '99', borderColor: CHART_COLORS.blue, borderWidth: 1, borderRadius: 4 }]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: { display: false } },
      scales: {
        x: { grid: { color: CHART_COLORS.border }, ticks: { color: CHART_COLORS.text2, font: { size: 11 } } },
        y: { grid: { color: CHART_COLORS.border }, ticks: { color: CHART_COLORS.text2, stepSize: 1, font: { size: 11 } } }
      }
    }
  });
}

function renderCharts(cleanRows, hdrs, cols) {
  document.getElementById('chartsEmpty').style.display = 'none';
  document.getElementById('chartsDash').style.display  = 'block';

  // 1. Tren harian
  destroyChart('trend');
  if (cols.date) {
    const dateCounts = {};
    cleanRows.forEach(r => { const d = r[cols.date] || 'N/A'; dateCounts[d] = (dateCounts[d] || 0) + 1; });
    const dLabels = Object.keys(dateCounts).sort();
    const dVals   = dLabels.map(d => dateCounts[d]);
    state.charts['trend'] = new Chart(document.getElementById('chartTrend'), {
      type: 'line',
      data: {
        labels: dLabels,
        datasets: [{ label: 'Jumlah Hadir', data: dVals, borderColor: CHART_COLORS.green, backgroundColor: CHART_COLORS.green + '22', fill: true, tension: 0.35, pointBackgroundColor: CHART_COLORS.green, pointRadius: 4 }]
      },
      options: {
        responsive: true, maintainAspectRatio: false,
        plugins: { legend: { display: false } },
        scales: {
          x: { grid: { color: CHART_COLORS.border }, ticks: { color: CHART_COLORS.text2, font: { size: 11 }, maxRotation: 45, maxTicksLimit: 14 } },
          y: { grid: { color: CHART_COLORS.border }, ticks: { color: CHART_COLORS.text2, stepSize: 1, font: { size: 11 } }, beginAtZero: true }
        }
      }
    });
  }

  // 2. Lembur per karyawan
  destroyChart('ot');
  if (cleanRows[0] && cleanRows[0]['_Jam Lembur'] !== undefined && cols.name) {
    const otMap = {};
    cleanRows.forEach(r => {
      const name = r[cols.name] || r[cols.id] || 'N/A';
      if (!r['_Jam Lembur'] || r['_Jam Lembur'] === '-') return;
      const hm = r['_Jam Lembur'].match(/(\d+)j (\d+)m/);
      if (hm) otMap[name] = (otMap[name] || 0) + parseInt(hm[1]) * 60 + parseInt(hm[2]);
    });
    const otNames = Object.keys(otMap).sort((a, b) => otMap[b] - otMap[a]).slice(0, 10);
    const otVals  = otNames.map(n => +(otMap[n] / 60).toFixed(1));
    state.charts['ot'] = new Chart(document.getElementById('chartOT'), {
      type: 'bar',
      data: {
        labels: otNames,
        datasets: [{ label: 'Jam Lembur', data: otVals, backgroundColor: CHART_COLORS.orange + '99', borderColor: CHART_COLORS.orange, borderWidth: 1, borderRadius: 4 }]
      },
      options: {
        indexAxis: 'y', responsive: true, maintainAspectRatio: false,
        plugins: { legend: { display: false } },
        scales: {
          x: { grid: { color: CHART_COLORS.border }, ticks: { color: CHART_COLORS.text2, font: { size: 11 } }, title: { display: true, text: 'Jam', color: CHART_COLORS.muted } },
          y: { grid: { color: 'transparent' }, ticks: { color: CHART_COLORS.text2, font: { size: 11 } } }
        }
      }
    });
  }

  // 3. Frekuensi keterlambatan per orang
  destroyChart('late');
  if (cleanRows[0] && cleanRows[0]['_Status Kehadiran'] !== undefined && cols.name) {
    const lateMap = {};
    cleanRows.forEach(r => {
      if (r['_Status Kehadiran'] && r['_Status Kehadiran'].includes('Terlambat')) {
        const name = r[cols.name] || r[cols.id] || 'N/A';
        lateMap[name] = (lateMap[name] || 0) + 1;
      }
    });
    const lNames = Object.keys(lateMap).sort((a, b) => lateMap[b] - lateMap[a]);
    const lVals  = lNames.map(n => lateMap[n]);
    state.charts['late'] = new Chart(document.getElementById('chartLate'), {
      type: 'bar',
      data: {
        labels: lNames,
        datasets: [{ label: 'Frekuensi Terlambat', data: lVals, backgroundColor: CHART_COLORS.yellow + '99', borderColor: CHART_COLORS.yellow, borderWidth: 1, borderRadius: 4 }]
      },
      options: {
        responsive: true, maintainAspectRatio: false,
        plugins: { legend: { display: false } },
        scales: {
          x: { grid: { color: CHART_COLORS.border }, ticks: { color: CHART_COLORS.text2, font: { size: 11 }, maxRotation: 30 } },
          y: { grid: { color: CHART_COLORS.border }, ticks: { color: CHART_COLORS.text2, stepSize: 1, font: { size: 11 } }, beginAtZero: true }
        }
      }
    });
  }

  // 4. Distribusi durasi kerja
  destroyChart('durDist');
  if (cleanRows[0] && cleanRows[0]['_Durasi Kerja'] !== undefined && cols.in && cols.out) {
    const durBuckets = { '<6j': 0, '6–7j': 0, '7–8j': 0, '8–9j': 0, '9–10j': 0, '>10j': 0 };
    cleanRows.forEach(r => {
      const dur = timeToMins(r[cols.out]) !== null && timeToMins(r[cols.in]) !== null
        ? timeToMins(r[cols.out]) - timeToMins(r[cols.in]) : null;
      if (dur === null || dur < 0) return;
      const h = dur / 60;
      if (h < 6) durBuckets['<6j']++;
      else if (h < 7) durBuckets['6–7j']++;
      else if (h < 8) durBuckets['7–8j']++;
      else if (h < 9) durBuckets['8–9j']++;
      else if (h < 10) durBuckets['9–10j']++;
      else durBuckets['>10j']++;
    });
    state.charts['durDist'] = new Chart(document.getElementById('chartDurDist'), {
      type: 'bar',
      data: {
        labels: Object.keys(durBuckets),
        datasets: [{ label: 'Jumlah Karyawan', data: Object.values(durBuckets), backgroundColor: CHART_COLORS.blue + '99', borderColor: CHART_COLORS.blue, borderWidth: 1, borderRadius: 4 }]
      },
      options: {
        responsive: true, maintainAspectRatio: false,
        plugins: { legend: { display: false } },
        scales: {
          x: { grid: { color: CHART_COLORS.border }, ticks: { color: CHART_COLORS.text2, font: { size: 11 } } },
          y: { grid: { color: CHART_COLORS.border }, ticks: { color: CHART_COLORS.text2, stepSize: 1, font: { size: 11 } }, beginAtZero: true }
        }
      }
    });
  }

  // 5. Kualitas data per kolom
  destroyChart('colQuality');
  const colQuality = {};
  state.headers.forEach(h => {
    const total  = state.rawData.length;
    const filled = state.rawData.filter(r => r[h] && String(r[h]).trim() !== '').length;
    colQuality[h] = Math.round((filled / total) * 100);
  });
  const cqLabels = Object.keys(colQuality);
  const cqVals   = cqLabels.map(k => colQuality[k]);
  const cqColors = cqVals.map(v => v >= 90 ? CHART_COLORS.green + '99' : v >= 70 ? CHART_COLORS.yellow + '99' : CHART_COLORS.red + '99');
  const cqBorder = cqVals.map(v => v >= 90 ? CHART_COLORS.green : v >= 70 ? CHART_COLORS.yellow : CHART_COLORS.red);
  state.charts['colQuality'] = new Chart(document.getElementById('chartColQuality'), {
    type: 'bar',
    data: {
      labels: cqLabels,
      datasets: [{ label: 'Kelengkapan (%)', data: cqVals, backgroundColor: cqColors, borderColor: cqBorder, borderWidth: 1, borderRadius: 4 }]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: { display: false } },
      scales: {
        x: { grid: { color: CHART_COLORS.border }, ticks: { color: CHART_COLORS.text2, font: { size: 11 }, maxRotation: 30 } },
        y: { grid: { color: CHART_COLORS.border }, ticks: { color: CHART_COLORS.text2, font: { size: 11 } }, min: 0, max: 100, title: { display: true, text: '%', color: CHART_COLORS.muted } }
      }
    }
  });
}

// ══════════════════════════════════
// DOWNLOAD CSV
// ══════════════════════════════════
function downloadCSV() {
  if (!state.cleanData.length) return;
  const extraKeys = Object.keys(state.cleanData[0]).filter(k => k.startsWith('_') && k !== '_origIdx' && k !== '_punchNote');
  const allHdrs   = [...state.headers, ...extraKeys];

  let csv = allHdrs.map(h => `"${h.replace(/^_/, '')}"`).join(',') + '\n';
  state.cleanData.forEach(row => {
    csv += allHdrs.map(h => `"${(row[h] || '').toString().replace(/"/g, '""')}"`).join(',') + '\n';
  });

  const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
  const a    = document.createElement('a');
  a.href     = URL.createObjectURL(blob);
  a.download = `hadir_pro_bersih_${new Date().toISOString().slice(0, 10)}.csv`;
  a.click();
}

// ══════════════════════════════════
// FILTER TABLE
// ══════════════════════════════════
function filterTable(which) {
  const q    = document.getElementById(which + 'Search').value.toLowerCase();
  const rows = document.querySelectorAll(`#${which}TableWrap table tbody tr`);
  let visible = 0;
  rows.forEach(tr => {
    const match = tr.textContent.toLowerCase().includes(q);
    tr.style.display = match ? '' : 'none';
    if (match) visible++;
  });
  document.getElementById(which + 'TableInfo').textContent = `${visible} baris`;
}

// ══════════════════════════════════
// LOAD DATA (common entry point)
// ══════════════════════════════════
function loadData(hdrs, rows) {
  state.headers   = hdrs;
  state.rawData   = rows;
  state.cleanData = [];
  state.issues    = [];

  const cols   = getColNames(hdrs);
  const issues = detectIssues(rows, cols);
  state.issues = issues;

  document.getElementById('rawTableWrap').innerHTML = renderRawTable(rows, hdrs, issues);
  document.getElementById('rawTableInfo').textContent = `${rows.length} baris`;
  document.getElementById('rawCount').textContent     = rows.length;

  renderIssues(issues);
  document.getElementById('issueCount').textContent = issues.length;

  updateOverview(rows, issues, []);
  updateTopbar(rows, issues);

  document.getElementById('btnClean').disabled    = false;
  document.getElementById('btnDownload').disabled = true;
  document.getElementById('cleanCount').textContent = '—';
  document.getElementById('cleanTableWrap').innerHTML = `<div class="empty-table">Klik "Bersihkan" untuk memproses data.</div>`;

  switchSection('overview');
}

function loadSample() {
  state.sourceFile = 'data-contoh.csv';
  const { hdrs, rows } = parseCSV(SAMPLE_CSV);
  loadData(hdrs, rows);
}

// ══════════════════════════════════
// SECTION SWITCHING
// ══════════════════════════════════
function switchSection(name) {
  document.querySelectorAll('.section').forEach(s => s.classList.toggle('active', s.id === 'sec-' + name));
  document.querySelectorAll('.nav-item').forEach(a => a.classList.toggle('active', a.dataset.section === name));

  const titles = { overview: 'Ringkasan', raw: 'Data Mentah', issues: 'Masalah Data', clean: 'Data Bersih', charts: 'Grafik & Analitik' };
  document.getElementById('pageTitle').textContent = titles[name] || name;
}

// ══════════════════════════════════
// FILE READING (CSV & XLS/XLSX)
// ══════════════════════════════════
function readFile(file) {
  state.sourceFile = file.name;
  const ext = file.name.split('.').pop().toLowerCase();

  if (ext === 'xls' || ext === 'xlsx') {
    // Check if SheetJS is available
    if (typeof XLSX === 'undefined') {
      alert('Library SheetJS belum dimuat. Pastikan koneksi internet aktif dan muat ulang halaman.');
      return;
    }
    const reader = new FileReader();
    reader.onload = ev => {
      try {
        const { hdrs, rows } = parseXLSBuffer(ev.target.result);
        if (!rows.length) {
          alert('Tidak ada data absensi yang ditemukan di file XLS ini. Pastikan file memiliki sheet "Lap. Log Absen".');
          return;
        }
        loadData(hdrs, rows);
      } catch (e) {
        alert('Gagal membaca file XLS: ' + e.message);
        console.error(e);
      }
    };
    reader.readAsArrayBuffer(file);
  } else {
    const reader = new FileReader();
    reader.onload = ev => {
      const { hdrs, rows } = parseCSV(ev.target.result);
      loadData(hdrs, rows);
    };
    reader.readAsText(file);
  }
}

// ══════════════════════════════════
// DRAG & DROP
// ══════════════════════════════════
const zone = document.getElementById('uploadZone');
zone.addEventListener('dragover',  e => { e.preventDefault(); zone.classList.add('drag-active'); });
zone.addEventListener('dragleave', ()  => zone.classList.remove('drag-active'));
zone.addEventListener('drop', e => {
  e.preventDefault();
  zone.classList.remove('drag-active');
  const file = e.dataTransfer.files[0];
  if (file) readFile(file);
});

document.getElementById('fileInput').addEventListener('change', e => {
  if (e.target.files[0]) readFile(e.target.files[0]);
});
