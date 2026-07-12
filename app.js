// ═══════════════════════════════════════════════
//  CASEOPS — COMMAND CENTER v2.0
//  Local-only. All data in localStorage.
// ═══════════════════════════════════════════════

let allData = [];
let currentData = [];
let donutChart = null;
let barChart = null;
let pendingWithChart = null;
let activeDrawerCase = null;
let envChart = null;

// ─── STORAGE HELPERS ───────────────────────────
const store = {
  get: k => { try { return JSON.parse(localStorage.getItem(k)); } catch { return null; } },
  set: (k, v) => localStorage.setItem(k, JSON.stringify(v)),
  getCase: (caseNum, key) => {
    const all = store.get('caseData') || {};
    return (all[caseNum] || {})[key];
  },
  setCase: (caseNum, key, val) => {
    const all = store.get('caseData') || {};
    if (!all[caseNum]) all[caseNum] = {};
    all[caseNum][key] = val;
    store.set('caseData', all);
  }
};

// ─── WATCHLIST HELPERS ─────────────────────────
const watchlist = {
  get: () => store.get('watchlist') || {},
  getCase: (caseNum) => (watchlist.get()[caseNum] || null),
  isStarred: (caseNum) => !!(watchlist.get()[caseNum]?.starred),
  star: (caseNum, note = '') => {
    const wl = watchlist.get();
    wl[caseNum] = { starred: true, note, starredAt: new Date().toISOString() };
    store.set('watchlist', wl);
  },
  unstar: (caseNum) => {
    const wl = watchlist.get();
    delete wl[caseNum];
    store.set('watchlist', wl);
  },
  updateNote: (caseNum, note) => {
    const wl = watchlist.get();
    if (wl[caseNum]) { wl[caseNum].note = note; store.set('watchlist', wl); }
  },
  activeItems: () => {
    const wl = watchlist.get();
    const starred = Object.keys(wl).filter(k => wl[k]?.starred);
    return starred.sort((a, b) => new Date(wl[a].starredAt) - new Date(wl[b].starredAt));
  }
};

// ─── BOOT ──────────────────────────────────────
window.addEventListener('load', () => {
  const msgs = ['INITIALIZING...', 'LOADING MODULES...', 'SYNCING LOCAL DATA...', 'READY'];
  let i = 0;
  const el = document.getElementById('bootStatus');
  const tick = () => { el.textContent = msgs[Math.min(i++, msgs.length-1)]; };
  tick();
  const iv = setInterval(tick, 450);

  setTimeout(() => {
    clearInterval(iv);
    el.textContent = 'READY';
    const boot = document.getElementById('boot');
    boot.style.opacity = '0';
    setTimeout(() => {
      boot.style.display = 'none';
      document.getElementById('app').style.display = 'block';
      const saved = store.get('cases');
      if (saved && saved.length) {
        processData(saved);
        showBriefing();
      }
    }, 600);
  }, 1900);
});

// ─── CLOCK ─────────────────────────────────────
function updateClock() {
  const now = new Date();
  document.getElementById('clock').textContent =
    now.toLocaleTimeString('en-IN', { hour: '2-digit', minute: '2-digit', second: '2-digit' });
  document.getElementById('dateDisplay').textContent =
    now.toLocaleDateString('en-IN', { weekday: 'short', day: '2-digit', month: 'short', year: 'numeric' }).toUpperCase();
}
setInterval(updateClock, 1000);
updateClock();

// ── PROD WINDOW BANNER ─────────────────────────
function updateProdWindow() {
  const now    = new Date();
  const hour   = now.getHours();
  const minute = now.getMinutes();
  const mins   = hour * 60 + minute;

  const PRE_OPEN    = 9  * 60;
  const CLOSE       = 15 * 60 + 30;
  const WARN_BEFORE = 30;

  let html = '', cls = '';

  if (mins >= PRE_OPEN && mins < CLOSE) {
    const remaining = CLOSE - mins;
    const h = Math.floor(remaining / 60), m = remaining % 60;
    cls  = 'prod-banner prod-live';
    html = `🔴 MARKET LIVE — PROD changes blocked &nbsp;|&nbsp; Window closes in <strong>${h}h ${m}m</strong> (15:30)`;
  } else if (mins >= (PRE_OPEN - WARN_BEFORE) && mins < PRE_OPEN) {
    const remaining = PRE_OPEN - mins;
    cls  = 'prod-banner prod-warn';
    html = `🟡 Market opens in <strong>${remaining} min</strong> — wrap up PROD changes before 9:00`;
  } else {
    cls  = 'prod-banner prod-safe';
    if (mins < PRE_OPEN) {
      const remaining = PRE_OPEN - mins;
      const h = Math.floor(remaining / 60), m = remaining % 60;
      html = `🟢 PROD window open — market opens in ${h}h ${m}m &nbsp;|&nbsp; UAT: anytime`;
    } else {
      html = `🟢 PROD window open (post-market) &nbsp;|&nbsp; UAT: anytime`;
    }
  }

  let banner = document.getElementById('prodWindowBanner');
  if (!banner) return;
  banner.className = cls;
  banner.innerHTML = html;
}
setInterval(updateProdWindow, 60000);

// ─── FILE UPLOAD ───────────────────────────────
document.getElementById('upload').addEventListener('change', e => {
  const file = e.target.files[0];
  if (!file) return;

  // ── Reset input so same file can be re-uploaded
  e.target.value = '';

  const reader = new FileReader();

  reader.onerror = () => {
    alert('Failed to read file. Please try again.');
  };

  reader.onload = ev => {
    try {
      const wb = XLSX.read(new Uint8Array(ev.target.result), {
        type: 'array',
        cellDates: true,      // ✅ parse dates properly
        dateNF: 'dd/mm/yyyy', // ✅ consistent date format
        raw: false            // ✅ get formatted strings not raw numbers
      });

      // ── Try first sheet, or find sheet with case data
      let sheetName = wb.SheetNames[0];

      // If multiple sheets, try to find one with 'case' in the name
      if (wb.SheetNames.length > 1) {
        const caseSheet = wb.SheetNames.find(n => n.toLowerCase().includes('case'));
        if (caseSheet) sheetName = caseSheet;
      }

      const sheet = wb.Sheets[sheetName];
      const rows  = XLSX.utils.sheet_to_json(sheet, {
        header: 1,
        defval: '',    // ✅ empty string instead of undefined for missing cells
        blankrows: false
      });

      if (!rows || rows.length === 0) {
        alert('Sheet appears to be empty.');
        return;
      }

      const cleaned = extractTable(rows);

      if (!cleaned.length) {
        alert(`No valid cases found in sheet "${sheetName}". Check that your file has a "Case Number" column.`);
        return;
      }

      store.set('cases', cleaned);
      processData(cleaned);
      showBriefing();

    } catch (err) {
      console.error('Excel parse error:', err);
      alert('Failed to parse Excel file: ' + err.message);
    }
  };

  reader.readAsArrayBuffer(file);
});

function extractTable(rows) {
  // ── Find header row: look for row containing 'case number' (case-insensitive)
  let start = -1;
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    if (!row) continue;
    const joined = row.map(c => String(c || '').toLowerCase()).join(' ');
    if (joined.includes('case number')) {
      start = i;
      break;
    }
  }

  if (start === -1) {
    alert('Could not find header row. Make sure your Excel has a "Case Number" column.');
    return [];
  }

  // ── Clean headers: trim spaces, normalize
  const rawHeaders = rows[start];
  const headers = rawHeaders.map(h => String(h || '').trim());

  // ── Build rows
  const data = rows
    .slice(start + 1)
    .map(r => {
      if (!r || r.every(cell => cell === null || cell === undefined || String(cell).trim() === '')) return null;
      const obj = {};
      headers.forEach((h, i) => {
        obj[h] = (r[i] !== undefined && r[i] !== null) ? r[i] : '';
      });
      return obj;
    })
    .filter(r => {
      if (!r) return false;
      // Accept rows where Case Number starts with 00 or is a number
      const caseNum = String(r['Case Number'] || '').trim();
      return caseNum !== '' && caseNum.toLowerCase() !== 'total';
    });

  console.log(`✅ Loaded ${data.length} rows from row index ${start}`);
  return data;
}
// ─── DATE HELPERS ──────────────────────────────
function parseDDMMYYYY(dateStr) {
  if (!dateStr) return null;
  const parts = String(dateStr).split('/');
  if (parts.length !== 3) return null;
  const [day, month, year] = parts;
  return new Date(year, month - 1, day);
}

function parseDeliveryDate(raw) {
  if (!raw || String(raw).trim() === '') return null;
  if (typeof raw === 'number') {
    return new Date(Math.round((raw - 25569) * 86400 * 1000));
  }
  if (String(raw).includes('/')) {
    return parseDDMMYYYY(String(raw));
  }
  const d = new Date(raw);
  return isNaN(d.getTime()) ? null : d;
}

// ─── PROCESS DATA ──────────────────────────────
function processData(data) {
  data.forEach(d => {
    allData = data;
    window.allData = allData; // for debugging
    console.log ("Loaded Data:", allData);
    d.owner        = d['Case Owner: Full Name'] ||d['Case Owner']|| 'Unknown';
    d.pendingWith  = d['Cases pending with'] || 'Unknown';
    d.status       = d['Status'] || 'Pending';
    d.inc          = d['Related Issue: Incident Number'] || '';
    d.Server       = d['Server (UAT/PROD)'] || 'Unknown';
    d.deliveryDate = d['Actual Delivery Date'] || null;
    d.jirastatus = d['JIRA Ticket Status'] || '';
    d.incAge       = d['INC Age'] || d['Incident Age'] || null;
    d.incOwner     = d['INC Owner'] || d['Incident owner'] || '—';
    d.incOwnerQueue = d['INC Owner Queue'] || d['Incident Owner Queue'] || '—';
    d.nextUpdate   = d['Next Update'] || d['Next Update Date'] || null;
    d.latestComment = d['Latest Comments'] || '';
    d.modifiedBy    = d['Latest Comments Modified By: Full Name'] || '';
    d.modifiedDate  = d['Latest Comments Modified Date'] || '';
    d.emailLastSent = d['Email Last Sent Date and Time'] || d['Client Email Last Sent Date and Time'] ||null;

    const rawAccount = d['Account Name: Account Name'] || d['Account Name'] || d['Account'] || 'Unknown';
    const fullAccount = String(rawAccount || '').trim() || 'Unknown';
    d.account = fullAccount.split(' ')[0];

    const ageRaw = d['Case Age(in number)'] || d['Case Age(in number) '] || d['Age'] || null;
    let age = parseInt(ageRaw);
    if (isNaN(age)) age = null;
    d.age = age;

    if (age === null)   d.bucket = 'Unknown';
    else if (age <= 10) d.bucket = '1-10';
    else if (age <= 20) d.bucket = '10-20';
    else if (age <= 30) d.bucket = '20-30';
    else                d.bucket = '30+';
  });

  allData = data;
  currentData = data;

  renderKPI();
  renderCharts();
  renderFilters();
  renderWatchlist();
  renderTable(data);
  renderAlertStrip();
  renderDeliveryAlertBanner();
  renderINCAlertBanner();
  renderCommentAlertBanner();
  renderNextUpdateAlertBanner();
  updateBellBadge();
}

// ─── VALIDATE CASE NUMBER ──────────────────────
function isValidCaseNumber(value) {
  return String(value).trim().startsWith("00");
}

// ─── KPI ───────────────────────────────────────
function renderKPI() {
  const validCases = allData.filter(row => isValidCaseNumber(row["Case Number"]));
  const total   = validCases.length;
  const pending = validCases.filter(d => d.Status !== 'Closed' && d.status !== 'Closed').length;
  const incYes  = validCases.filter(d => d.inc).length;
  const incNo   = total - incYes;

  document.getElementById('kpi').innerHTML = `
    <div class="kpi-card k-total">
      <div class="kpi-label">TOTAL CASES</div>
      <div class="kpi-value">${total}</div>
      <div class="kpi-sub">based on Case Number column</div>
    </div>
    <div class="kpi-card k-pending">
      <div class="kpi-label">OPEN / PENDING</div>
      <div class="kpi-value">${pending}</div>
      <div class="kpi-sub">${Math.round(pending / total * 100) || 0}% of total</div>
    </div>
    <div class="kpi-card k-inc">
      <div class="kpi-label">INC CREATED</div>
      <div class="kpi-value">${incYes}</div>
    </div>
    <div class="kpi-card k-noinc">
      <div class="kpi-label">NO INC</div>
      <div class="kpi-value">${incNo}</div>
    </div>
  `;
}

// ─── CHARTS ────────────────────────────────────
const CHART_COLORS = [
  '#3b82f6','#06b6d4','#8b5cf6','#10b981',
  '#f59e0b','#ef4444','#ec4899','#14b8a6','#a78bfa'
];

function renderCharts(data= allData) {
  const acc = {};
  data.forEach(d => acc[d.account] = (acc[d.account] || 0) + 1);
  const aL = Object.keys(acc), aV = Object.values(acc);

  if (donutChart) donutChart.destroy();
  donutChart = new Chart(document.getElementById('donut'), {
    type: 'doughnut',
    data: { labels: aL, datasets: [{ data: aV, backgroundColor: CHART_COLORS, borderWidth: 0 }] },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: { display: false } },
      onClick: (e, el) => { if (el.length) quickFilter('accountFilter', aL[el[0].index]); }
    }
  });

  const bucketOrder = ['1-10', '10-20', '20-30', '30+', 'Unknown'];
  const bkt = {};
  data.forEach(d => bkt[d.bucket] = (bkt[d.bucket] || 0) + 1);
  const bL = bucketOrder.filter(b => bkt[b]);
  const bV = bL.map(b => bkt[b]);
  const barColors = bL.map(b =>
    b === '30+' ? '#ef4444' : b === '20-30' ? '#f59e0b' : b === '10-20' ? '#3b82f6' : '#10b981'
  );

  if (barChart) barChart.destroy();
  barChart = new Chart(document.getElementById('bar'), {
    type: 'bar',
    data: { labels: bL, datasets: [{ data: bV, backgroundColor: barColors, borderRadius: 4, borderSkipped: false }] },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: { display: false } },
      scales: {
        x: { grid: { color: 'rgba(255,255,255,0.05)' }, ticks: { color: '#64748b', font: { family: 'Space Mono', size: 10 } } },
        y: { grid: { color: 'rgba(255,255,255,0.05)' }, ticks: { color: '#64748b', font: { family: 'Space Mono', size: 10 } } }
      },
      onClick: (e, el) => { if (el.length) quickFilter('bucketFilter', bL[el[0].index]); }
    }
  });

  const pwTotal = {}, pwDelivery = {};
  data.filter(d => d.status !== 'Closed').forEach(d => {
    const key = d.pendingWith || 'Unknown';
    pwTotal[key] = (pwTotal[key] || 0) + 1;
    if (key === 'Developers' && d.deliveryDate && String(d.deliveryDate).trim() !== '') {
      pwDelivery[key] = (pwDelivery[key] || 0) + 1;
    }
  });

  const pL = Object.keys(pwTotal);
  const pV = Object.values(pwTotal);
  const pwbarColors = pL.map(label => label === 'Developers' ? '#22c55e' : '#3b82f6');

  if (pendingWithChart) pendingWithChart.destroy();
  if (!pL.length) return;

  pendingWithChart = new Chart(document.getElementById('pendingWithChart'), {
    type: 'bar',
    data: { labels: pL, datasets: [{ data: pV, backgroundColor: pwbarColors, borderRadius: 4, borderSkipped: false }] },
    options: {
      responsive: true, maintainAspectRatio: false, indexAxis: 'y',
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: {
            label: function(context) {
              const key = context.label;
              const total = pwTotal[key] || 0;
              if (key === 'Developers') return `Total: ${total} | ETA: ${pwDelivery[key] || 0}`;
              return `Total: ${total}`;
            }
          }
        }
      },
      scales: {
        x: { grid: { color: 'rgba(255,255,255,0.05)' }, ticks: { color: '#64748b', font: { family: 'Space Mono', size: 10 } } },
        y: { grid: { display: false }, ticks: { color: '#64748b', font: { family: 'Space Mono', size: 10 } } }
      },
      onClick: (e, elements) => {
        if (!elements.length) return;
        const clicked = pL[elements[0].index];
        let filtered;
        if (clicked === 'Developers') {
          filtered = e.native.shiftKey
            ? allData.filter(d => d.pendingWith === 'Developers' && d.status !== 'Closed' && d.deliveryDate && String(d.deliveryDate).trim() !== '')
            : allData.filter(d => d.pendingWith === 'Developers' && d.status !== 'Closed');
        } else {
          filtered = allData.filter(d => d.pendingWith === clicked && d.status !== 'Closed');
        }
        currentData = filtered;
        renderTable(filtered);
      }
    }
  });

  const ser = {};
  data.forEach(d => {
    const key = String(d.Server || 'Unknown').trim().toUpperCase();
    ser[key] = (ser[key] || 0) + 1;
  });
  const sL = Object.keys(ser), sV = Object.values(ser);

  if (envChart) envChart.destroy();
  envChart = new Chart(document.getElementById('envChart'), {
    type: 'doughnut',
    data: { labels: sL, datasets: [{ data: sV, backgroundColor: ['#ef4444', '#3b82f6', '#64748b'], borderWidth: 0 }] },
    options: {
      responsive: true, maintainAspectRatio: false, cutout: '65%',
      plugins: { legend: { position: 'bottom', labels: { color: '#64748b', font: { family: 'Space Mono', size: 10 } } } },
      onClick: (e, elements) => {
        if (!elements.length) return;
        const clicked = sL[elements[0].index];
        const filtered = allData.filter(d => String(d.Server || '').trim().toUpperCase() === clicked);
        currentData = filtered;
        renderTable(filtered);
      }
    }
  });
}

// ─── PROMISE HELPERS ───────────────────────────
function getPromises(caseNum) { return store.getCase(caseNum, 'promises') || []; }
function getEmails(caseNum)   { return store.getCase(caseNum, 'emails')   || []; }

function getCasePromiseStatus(caseNum) {
  const today = new Date(); today.setHours(0,0,0,0);
  const promises = getPromises(caseNum).filter(p => !p.done);
  if (!promises.length) return null;
  for (const p of promises) {
    const d = new Date(p.date); d.setHours(0,0,0,0);
    const diff = Math.floor((d - today) / 86400000);
    if (diff < 0)   return 'overdue';
    if (diff === 0) return 'today';
  }
  return 'upcoming';
}

// ─── ALERT STRIP ───────────────────────────────
function renderAlertStrip() {
  const today = new Date(); today.setHours(0,0,0,0);
  const alerts = [];

  allData.forEach(d => {
    const ps = getCasePromiseStatus(d['Case Number']);
    if (ps === 'overdue' || ps === 'today') {
      getPromises(d['Case Number']).filter(p => !p.done).forEach(p => {
        const pd = new Date(p.date); pd.setHours(0,0,0,0);
        const diff = Math.floor((pd - today) / 86400000);
        if (diff <= 0) alerts.push({ caseNum: d['Case Number'], account: d.account, text: p.text, diff, type: diff < 0 ? 'overdue' : 'today' });
      });
    }
  });

  const strip = document.getElementById('alertStrip');
  if (!alerts.length) { strip.innerHTML = ''; return; }

  strip.innerHTML = alerts.map(a => `
    <div class="strip-item ${a.type}" onclick="openDrawer('${a.caseNum}')">
      <span class="strip-icon">${a.type === 'overdue' ? '🔴' : '🟡'}</span>
      <span class="strip-text"><strong>${a.caseNum}</strong> — ${a.text}</span>
      <span class="strip-case">${a.type === 'overdue' ? Math.abs(a.diff)+' DAYS OVERDUE' : 'DUE TODAY'}</span>
    </div>
  `).join('');

  const alertList = document.getElementById('alertList');
  alertList.innerHTML = alerts.map(a => `
    <div class="alert-item ${a.type}" onclick="openDrawer('${a.caseNum}')">
      <div class="alert-item-case">${a.caseNum} · ${a.account}</div>
      <div class="alert-item-text">${a.text}</div>
      <div class="alert-item-due">${a.type === 'overdue' ? '⚠ ' + Math.abs(a.diff) + ' days overdue' : '⏰ Due today'}</div>
    </div>
  `).join('');
}

function updateBellBadge() {
  const today = new Date(); today.setHours(0,0,0,0);
  let count = 0;
  allData.forEach(d => {
    const ps = getCasePromiseStatus(d['Case Number']);
    if (ps === 'overdue' || ps === 'today') count++;
  });
  document.getElementById('bellBadge').textContent = count;
  document.getElementById('bellBadge').style.background = count ? '#ef4444' : '#334155';
}

// ─── ALERT PANEL ───────────────────────────────
function toggleAlertPanel() {
  document.getElementById('alertPanel').classList.toggle('open');
}

// ─── DELIVERY ALERT BANNER ─────────────────────
function renderDeliveryAlertBanner() {
  const today = new Date(); today.setHours(0, 0, 0, 0);
  const overdue = [], dueToday = [];

  allData.forEach(d => {

    if (d.jirastatus && String(d.jirastatus).trim().toLowerCase() === 'released') return; // ignore if JIRA done
    if (!d.deliveryDate || String(d.deliveryDate).trim() === '' || d.status === 'Closed') return;
   // ignore if JIRA already done
    const dDate = parseDeliveryDate(d.deliveryDate);
    if (!dDate || isNaN(dDate.getTime())) return;
    dDate.setHours(0, 0, 0, 0);
    const diff = Math.floor((dDate - today) / 86400000);
    if (diff < 0) overdue.push({ ...d, daysLate: Math.abs(diff) });
    else if (diff === 0) dueToday.push(d);
  });

  const bar = document.getElementById('deliveryAlertBar');
  if (!bar) return;

  const total = overdue.length + dueToday.length;
  if (!total) {
    bar.innerHTML = '';
    bar.style.display = 'none';
    return;
  }

  bar.style.display = 'flex';
  bar.innerHTML = `
    <div class="da-summary" onclick="filterByDeliveryAlert()">
      <span>📦</span>
      <span>ETA</span>
      <span class="da-count-badge">${total}</span>
    </div>
  `;
}


/*--- Release alert---------
function renderReleasedAlert() {
  const today = newDate(); today.setHours(0,0,0,0);

  allData= forEach (d => {
    if (d.jirastatus && string(d.jirastatus).trim().toLowerCase() === 'released'){
      released.push(d);
    }
  });
 const bar = document.getElementById("rleAlertBar");
  if (!bar) return;

  const details = released
  .map(d => `${d.inc} - ${d.incOwner}`)
  .join("<br>");

  const total = released.length;

  if (!total) {
    bar.innerHTML = "";
    bar.style.display = "none";
    return;
  }

  bar.innerHTML = `
<div class="da-summary" title="${released
  .map(d => `${d.inc} - ${d.incOwner}`)
  .join('\n')}">
    🚀 RLE
    <span class="da-count-badge">${released.length}</span>
</div>
`;

}
*/
function renderINCAlertBanner() {
  const incCases = allData.filter(d =>
    d.inc &&
    String(d.inc).trim() !== '' &&
    d.jirastatus &&
    String(d.jirastatus).trim().toLowerCase() === 'released' &&
    d.status !== 'Closed'
  );


  const bar = document.getElementById('incAlertBar');
  if (!bar) return;

  const total = incCases.length;

  if (!total) {
    bar.innerHTML = '';
    bar.style.display = 'none';
    return;
  }

  bar.style.display = 'flex';
  bar.innerHTML = `
    <div class="da-summary" onclick="openINCModal()">
      <span>🔥</span>
      <span>INC</span>
      <span class="da-count-badge">${total}</span>
    </div>
  `;
}

function openINCModal() {
  const incCases = allData.filter(d =>
   d.inc &&
   String(d.inc).trim() !== '' &&
   d.jirastatus &&
  String(d.jirastatus).trim().toLowerCase() === 'released' &&
  d.status !== 'Closed'
  );



  const buildRow = (d) => {
  const incAge = parseInt(d.incAge);
  const incAgeClass = !isNaN(incAge)
  ? incAge > 30 ? 'age-high' : incAge > 20 ? 'age-med' : 'age-low'
  : '';



  // ✅ inc might be an array or comma-separated — handle both
    const incNum = String(d.inc || '').trim();



    return `
      <tr class="dm-row"
        onclick="closeINCModal(); setTimeout(() => openDrawer('${d['Case Number']}'), 200)">
        <td style="white-space:nowrap;color:#a78bfa;font-family:var(--mono);font-size:11px">${incNum || '—'}</td>
        <td style="white-space:nowrap">${d.incOwner || '—'}</td>
        <td>${!isNaN(incAge) ? `<span class="age-num ${incAgeClass}">${incAge}d</span>` : '—'}</td>
        <td>
          <button class="open-btn"
            onclick="event.stopPropagation(); closeINCModal(); setTimeout(() => openDrawer('${d['Case Number']}'), 200)">
            OPEN ›
          </button>
        </td>
      </tr>
    `;
  };



  const html = `
    <div class="dm-section-label" style="color:#22c55e;border-color:rgba(34,197,94,0.2);margin-bottom:16px">
      🚀 RELEASED INC — ${incCases.length} CASES
    </div>
    <div class="table-wrap">
      <table class="dm-table" style="min-width:500px">
        <thead>
          <tr>
            <th style="width:200px">INC NUMBER</th>
            <th style="width:200px">INC OWNER</th>
            <th style="width:80px">INC AGE</th>
            <th style="width:80px"></th>
          </tr>
        </thead>
        <tbody>${incCases.map(d => buildRow(d)).join('')}</tbody>
      </table>
    </div>
    <div style="padding:20px 0 4px;text-align:right">
      <button class="modal-cta" style="font-size:11px;padding:8px 20px"
        onclick="closeINCModal(); filterINCCases()">
        VIEW ALL ${incCases.length} CASES IN TABLE ›
      </button>
    </div>
  `;



  document.getElementById('incModalContent').innerHTML = html;
  document.getElementById('incModal').classList.remove('hidden');
}

function closeINCModal() {
  document.getElementById('incModal').classList.add('hidden');
}

function filterINCCases() {
  const filtered = allData.filter(d =>
    d.inc && String(d.inc).trim() !== '' && 
    d.jirastatus && String(d.jirastatus).trim().toLowerCase() === 'released' &&
    d.status !== 'Closed'
  );
  currentData = filtered;
  renderTable(filtered);
  renderWatchlist();
  document.getElementById('caseTable').scrollIntoView({ behavior: 'smooth', block: 'start' });
}

// ─── EMAIL LAST SENT ALERT BANNER ──────────────


function parseDate(dateStr) {
  if (!dateStr) return null;

  // Excel numeric date — correct epoch
  if (typeof dateStr === 'number') {
    // ✅ Excel's epoch is Dec 30 1899, and has a leap year bug for 1900
    const date = new Date(Date.UTC(1899, 11, 30) + dateStr * 86400000);
    return isNaN(date.getTime()) ? null : date;
  }

  if (dateStr instanceof Date) {
    return isNaN(dateStr.getTime()) ? null : dateStr;
  }

  if (typeof dateStr === 'string') {
    const str = dateStr.trim();
    if (!str) return null;

    if (str.includes('/')) {
      const parts = str.split('/');
      if (parts.length === 3) {
        let d, m, y;
        if (parts[0].length === 4) {
          [y, m, d] = parts; // YYYY/MM/DD
        } else {
          [d, m, y] = parts; // DD/MM/YYYY
          if (String(y).length === 2) y = 2000 + parseInt(y);
        }
        const parsed = new Date(parseInt(y), parseInt(m) - 1, parseInt(d));
        return isNaN(parsed.getTime()) ? null : parsed;
      }
    }

    if (str.includes('-')) {
      const parts = str.split('-');
      if (parts.length === 3) {
        let d, m, y;
        if (parts[0].length === 4) {
          [y, m, d] = parts; // YYYY-MM-DD
        } else {
          [d, m, y] = parts; // DD-MM-YYYY
        }
        const parsed = new Date(parseInt(y), parseInt(m) - 1, parseInt(d));
        return isNaN(parsed.getTime()) ? null : parsed;
      }
    }

    const fallback = new Date(dateStr);
    return isNaN(fallback.getTime()) ? null : fallback;
  }

  return null;
}


function renderCommentAlertBanner() {
  const now = new Date();
  now.setHours(0, 0, 0, 0);

  const staleCases = allData.filter(d => {
    if (d.status === 'Closed') return false;
    if (!d.emailLastSent) return true;

    const sentDate = parseDate(d.emailLastSent);
    if (!sentDate || isNaN(sentDate.getTime())) return true;

    sentDate.setHours(0, 0, 0, 0);
    const diff = Math.floor((now - sentDate) / 86400000);
    return diff > 3;
  });

  const bar = document.getElementById('commentAlertBar');
  if (!bar) return;

  if (!staleCases.length) {
    bar.style.display = 'none';
    bar.innerHTML = '';
    return;
  }

  bar.style.display = 'flex';
  bar.innerHTML = `
    <div class="da-summary" onclick="filterOldCommentCases()">
      <span>📧</span>
      <span>Last EMAIL</span>
      <span class="da-count-badge">${staleCases.length}</span>
    </div>
  `;
}

function filterOldCommentCases() {
  const today = new Date(); today.setHours(0,0,0,0);

  const filtered = allData.filter(d => {
    if (d.status === 'Closed') return false;
    if (!d.emailLastSent) return true;
    const sentDate = parseDate(d.emailLastSent);
    if (isNaN(sentDate.getTime())) return true;
    sentDate.setHours(0,0,0,0);
    const diff = Math.floor((today - sentDate) / 86400000);
    return diff > 3;
  });

  currentData = filtered;
  renderTable(filtered);
  document.getElementById('caseTable').scrollIntoView({ behavior: 'smooth', block: 'start' });
}

// ─── NEXT UPDATE ALERT BANNER ──────────────────
function renderNextUpdateAlertBanner() {
  const today = new Date(); today.setHours(0,0,0,0);

  const overdueCases = allData.filter(d => {
    if (d.status === 'Closed') return false;
    if (!d.nextUpdate) return false;
    const nDate = parseDDMMYYYY(String(d.nextUpdate));
    if (!nDate || isNaN(nDate.getTime())) return false;
    nDate.setHours(0,0,0,0);
    return nDate < today;
  });

  // reuse or create a third bar slot
  let bar = document.getElementById('nextUpdateAlertBar');
  if (!bar) {
    bar = document.createElement('div');
    bar.id = 'nextUpdateAlertBar';
    bar.className = 'delivery-alert-bar';
    // insert after commentAlertBar
    const commentBar = document.getElementById('commentAlertBar');
    commentBar.parentNode.insertBefore(bar, commentBar.nextSibling);
  }

  if (!overdueCases.length) {
    bar.style.display = 'none';
    bar.innerHTML = '';
    return;
  }

  bar.style.display = 'flex';
  bar.innerHTML = `
    <div class="da-summary" style="color:#f59e0b;background:rgba(245,158,11,0.08);border-color:rgba(245,158,11,0.25);"
      onclick="filterOverdueNextUpdate()">
      <span>📅</span>
      <span>NU</span>
      <span class="da-count-badge" style="background:#f59e0b;">${overdueCases.length}</span>
    </div>
  `;
}

function filterOverdueNextUpdate() {
  const today = new Date(); today.setHours(0,0,0,0);

  const filtered = allData.filter(d => {
    if (d.status === 'Closed') return false;
    if (!d.nextUpdate) return false;
    const nDate = parseDDMMYYYY(String(d.nextUpdate));
    if (!nDate || isNaN(nDate.getTime())) return false;
    nDate.setHours(0,0,0,0);
    return nDate < today;
  });

  currentData = filtered;
  renderTable(filtered);
  document.getElementById('caseTable').scrollIntoView({ behavior: 'smooth', block: 'start' });
}
// ─── DELIVERY MODAL ────────────────────────────
function openDeliveryModal() {
  const today = new Date(); today.setHours(0, 0, 0, 0);
  const overdue = [], dueToday = [];

  allData.forEach(d => {
    if (!d.deliveryDate || String(d.deliveryDate).trim() === '' || d.status === 'Closed') return;
    const dDate = parseDeliveryDate(d.deliveryDate);
    if (!dDate || isNaN(dDate.getTime())) return;
    dDate.setHours(0, 0, 0, 0);
    const diff = Math.floor((dDate - today) / 86400000);
    if (diff < 0) overdue.push({ ...d, daysLate: Math.abs(diff), parsedDelivery: dDate });
    else if (diff === 0) dueToday.push({ ...d, daysLate: 0, parsedDelivery: dDate });
  });

  overdue.sort((a, b) => b.daysLate - a.daysLate);

  const buildRow = (d, type) => {
    const daysLabel = type === 'overdue'
      ? `<span class="da-row-badge overdue">${d.daysLate}d late</span>`
      : `<span class="da-row-badge today">due today</span>`;

    const deliveryFormatted = d.parsedDelivery
      ? d.parsedDelivery.toLocaleDateString('en-IN', { day: '2-digit', month: 'short', year: 'numeric' })
      : '—';

    const age = d.age !== null
      ? `<span class="age-num ${d.age > 30 ? 'age-high' : d.age > 20 ? 'age-med' : 'age-low'}">${d.age}d</span>`
      : '—';

    const status  = d.status || '';
    const sLower  = status.toLowerCase();
    const pillClass = sLower.includes('closed') ? 'pill-closed'
      : sLower.includes('progress') ? 'pill-open' : 'pill-pending';

    const incAge      = parseInt(d.incAge);
    const incAgeClass = !isNaN(incAge)
      ? incAge > 30 ? 'age-high' : incAge > 20 ? 'age-med' : 'age-low' : '';

    return `
      <tr class="dm-row"
        onclick="closeDeliveryModal(); filterByDeliveryAlert(); setTimeout(() => openDrawer('${d['Case Number']}'), 300)">
        <td><span class="status-pill ${pillClass}">${status}</span></td>
        <td><span class="case-num">${d['Case Number']}</span></td>
        <td>${d.account}</td>
        <td style="white-space:nowrap">${d.owner}</td>
        <td>${age}</td>
        <td>${d.pendingWith}</td>
        <td>${d.inc ? `<span class="inc-badge">${d.inc}</span>` : '<span style="color:var(--text-muted)">—</span>'}</td>
        <td>${!isNaN(incAge) ? `<span class="age-num ${incAgeClass}">${incAge}d</span>` : '—'}</td>
        <td style="color:#94a3b8;white-space:nowrap">${deliveryFormatted}</td>
        <td>${daysLabel}</td>
        <td>
          <button class="open-btn"
            onclick="event.stopPropagation(); closeDeliveryModal(); filterByDeliveryAlert(); setTimeout(() => openDrawer('${d['Case Number']}'), 300)">
            OPEN ›
          </button>
        </td>
      </tr>
    `;
  };

  let html = '';

  if (overdue.length) {
    html += `
      <div class="dm-section-label">⚠ OVERDUE — ${overdue.length} CASES</div>
      <div class="table-wrap">
        <table class="dm-table">
          <thead><tr>
            <th>STATUS</th><th>CASE #</th><th>ACCOUNT</th><th>OWNER</th>
            <th>AGE</th><th>PENDING WITH</th><th>INC</th><th>INC AGE</th>
            <th>DELIVERY DATE</th><th>OVERDUE BY</th><th></th>
          </tr></thead>
          <tbody>${overdue.map(d => buildRow(d, 'overdue')).join('')}</tbody>
        </table>
      </div>
    `;
  }

  if (dueToday.length) {
    html += `
      <div class="dm-section-label" style="margin-top:24px;color:#f59e0b;border-color:rgba(245,158,11,0.2)">
        📅 DUE TODAY — ${dueToday.length} CASES
      </div>
      <div class="table-wrap">
        <table class="dm-table">
          <thead><tr>
            <th>STATUS</th><th>CASE #</th><th>ACCOUNT</th><th>OWNER</th>
            <th>AGE</th><th>PENDING WITH</th><th>INC</th><th>INC AGE</th>
            <th>DELIVERY DATE</th><th></th><th></th>
          </tr></thead>
          <tbody>${dueToday.map(d => buildRow(d, 'today')).join('')}</tbody>
        </table>
      </div>
    `;
  }

  if (!overdue.length && !dueToday.length) {
    html = `<div style="color:var(--text-dim);padding:48px;text-align:center;font-family:var(--mono)">✅ No delivery alerts right now.</div>`;
  }

  // VIEW ALL button
  if (overdue.length || dueToday.length) {
    html += `
      <div style="padding:20px 0 4px;text-align:right">
        <button class="modal-cta" style="font-size:11px;padding:8px 20px"
          onclick="closeDeliveryModal(); filterByDeliveryAlert()">
          VIEW ALL ${overdue.length + dueToday.length} CASES IN TABLE ›
        </button>
      </div>
    `;
  }

  document.getElementById('deliveryModalContent').innerHTML = html;
  document.getElementById('deliveryModal').classList.remove('hidden');
}

function closeDeliveryModal() {
  document.getElementById('deliveryModal').classList.add('hidden');
}

// ─── FILTER BY DELIVERY ALERT ──────────────────
function filterByDeliveryAlert() {
  const today = new Date(); today.setHours(0, 0, 0, 0);

  const filtered = allData.filter(d => {
    if (d.jirastatus && String(d.jirastatus).trim().toLowerCase() === 'released') 
      {
        return false;
       } // ignore if JIRA already done
    if (!d.deliveryDate || String(d.deliveryDate).trim() === '' || d.status === 'Closed')
      {

       return false;
      }
    const dDate = parseDeliveryDate(d.deliveryDate);
    if (!dDate || isNaN(dDate.getTime())) return false;
    dDate.setHours(0, 0, 0, 0);
    return Math.floor((dDate - today) / 86400000) <= 0;
  });

  currentData = filtered;
  renderTable(filtered);
  renderWatchlist();

  document.getElementById('caseTable').scrollIntoView({ behavior: 'smooth', block: 'start' });
}

// ─── DAILY BRIEFING ────────────────────────────
function showBriefing() {
  const today = new Date(); today.setHours(0,0,0,0);
  const todayKey = today.toDateString();
  if (store.get('lastBriefing') === todayKey) return;

  const total    = allData.length;
  const pending  = allData.filter(d => d.status !== 'Closed').length;
  const critical = allData.filter(d => d.age > 30).length;

  const promises = [];
  allData.forEach(d => {
    getPromises(d['Case Number']).filter(p => !p.done).forEach(p => {
      const pd = new Date(p.date); pd.setHours(0,0,0,0);
      const diff = Math.floor((pd - today) / 86400000);
      promises.push({ caseNum: d['Case Number'], text: p.text, diff, account: d.account });
    });
  });

  const overdue  = promises.filter(p => p.diff < 0);
  const dueToday = promises.filter(p => p.diff === 0);
  const upcoming = promises.filter(p => p.diff > 0 && p.diff <= 3);

  const noEmailCases = allData.filter(d => {
    return !getEmails(d['Case Number']).length && d.status !== 'Closed';
  }).slice(0, 5);

  let html = `
    <div class="briefing-stats-row">
      <div class="b-stat-card"><div class="b-stat-val">${total}</div><div class="b-stat-label">Total Cases</div></div>
      <div class="b-stat-card"><div class="b-stat-val" style="color:var(--warning)">${pending}</div><div class="b-stat-label">Still Open</div></div>
      <div class="b-stat-card"><div class="b-stat-val" style="color:var(--danger)">${critical}</div><div class="b-stat-label">Age &gt; 30 days</div></div>
    </div>
  `;

  const starredItems = watchlist.activeItems();
  const starredRows  = starredItems
    .map(cn => allData.find(d => String(d['Case Number']) === String(cn)))
    .filter(Boolean);

  if (starredRows.length) {
    html += `<div class="briefing-section"><div class="briefing-label">⭐ YOUR WATCHLIST TODAY (${starredRows.length})</div>`;
    starredRows.forEach(d => {
      const wlData = watchlist.getCase(d['Case Number']);
      const note = wlData?.note ? ` — <em style="color:var(--text-dim)">${wlData.note}</em>` : '';
      html += `<div class="briefing-item"><span class="briefing-dot b-blue"></span><span class="briefing-item-text"><strong>${d['Case Number']}</strong> (${d.account})${note}</span></div>`;
    });
    html += `</div>`;
  }

  if (overdue.length) {
    html += `<div class="briefing-section"><div class="briefing-label">⚠ OVERDUE PROMISES</div>`;
    overdue.forEach(p => {
      html += `<div class="briefing-item"><span class="briefing-dot b-red"></span><span class="briefing-item-text"><strong>${p.caseNum}</strong> (${p.account}) — ${p.text} <em style="color:var(--danger)">[${Math.abs(p.diff)} days late]</em></span></div>`;
    });
    html += `</div>`;
  }

  if (dueToday.length) {
    html += `<div class="briefing-section"><div class="briefing-label">📅 DUE TODAY</div>`;
    dueToday.forEach(p => {
      html += `<div class="briefing-item"><span class="briefing-dot b-yellow"></span><span class="briefing-item-text"><strong>${p.caseNum}</strong> — ${p.text}</span></div>`;
    });
    html += `</div>`;
  }

  if (upcoming.length) {
    html += `<div class="briefing-section"><div class="briefing-label">🔜 UPCOMING (NEXT 3 DAYS)</div>`;
    upcoming.forEach(p => {
      html += `<div class="briefing-item"><span class="briefing-dot b-blue"></span><span class="briefing-item-text"><strong>${p.caseNum}</strong> — ${p.text} <em style="color:var(--text-dim)">in ${p.diff} day(s)</em></span></div>`;
    });
    html += `</div>`;
  }

  if (noEmailCases.length) {
    html += `<div class="briefing-section"><div class="briefing-label">📭 NO EMAIL LOGGED YET</div>`;
    noEmailCases.forEach(d => {
      html += `<div class="briefing-item"><span class="briefing-dot b-green"></span><span class="briefing-item-text"><strong>${d['Case Number']}</strong> (${d.account}) — no client email tracked</span></div>`;
    });
    html += `</div>`;
  }

  if (!overdue.length && !dueToday.length) {
    html += `<div class="briefing-item"><span class="briefing-dot b-green"></span><span class="briefing-item-text">All clear! No overdue promises today.</span></div>`;
  }

  document.getElementById('briefingContent').innerHTML = html;
  document.getElementById('briefingModal').classList.remove('hidden');
}

function closeBriefing() {
  document.getElementById('briefingModal').classList.add('hidden');
  store.set('lastBriefing', new Date().setHours(0,0,0,0).toString());
}

/*
// ─── FILTERS ───────────────────────────────────
function renderFilters() {
  createDropdown('accountFilter', 'account');
  createDropdown('ownerFilter',   'owner');
  createDropdown('statusFilter',  'status');
  createDropdown('pendingFilter', 'pendingWith');
  createDropdown('bucketFilter',  'bucket');
  createDropdown('INCFilter',  'incOwner');
  createDropdown('INCFilterqueue',  'incOwnerQueue');

}

const choiceInstances = {};
function createDropdown(id, key) {
  const vals = [...new Set(allData.map(d => d[key]).filter(Boolean))].sort();
  const sel  = document.getElementById(id);
  sel.innerHTML = "";
  vals.forEach(v => {
    sel.innerHTML += `<option value="${v}">${v}</option>`;
  });
 
  if (choiceInstances[id]) {
    choiceInstances[id].destroy();
  }
  choiceInstances[id] = new Choices(sel, {
    searchEnabled: true,
    shouldSort: false,
    removeItemButton: true,
    placeholder: true,
    placeholderValue: 'All',
    searchPlaceholderValue: 'Search...',
    itemSelectText: '',
  });
  sel.addEventListener('change', applyFilters);
}
function quickFilter(filterId, value) {
  document.getElementById(filterId).value = value;
  applyFilters();
}

function getSelectedValues(id) {

    return [...document.getElementById(id).selectedOptions]
        .map(option => option.value);

}

function applyFilters() {
  const f = {
    account:     document.getSelectedValues('accountFilter').value,
    owner:       document.getSelectedValues('ownerFilter').value,
    status:      document.getSelectedValues('statusFilter').value,
    pendingWith: document.getSelectedValues('pendingFilter').value,
    bucket:      document.getSelectedValues('bucketFilter').value,
    incOwner:    document.getSelectedValues('INCFilter').value,
    incOwnerQueue: document.getSelectedValues('INCFilterqueue').value
  };

  if (account.length && !account.includes(row.account))
    return false;

  if (owner.length && !owner.includes(row.owner))
    return false;

  if (status.length && !status.includes(row.status))
    return false;

  if (pendingWith.length && !pendingWith.includes(row.pendingWith))
    return false;

  if (bucket.length && !bucket.includes(row.bucket))
    return false;

  if (incOwner.length && !incOwner.include(row.incOwner))
    return false;

  if (incOwnerQueue.length && !incOwnerQueue.include(row.incOwnerQueue))
    return false;


  const promiseF = document.getElementById('promiseFilter').value;
   
  let filtered = allData;
  Object.keys(f).forEach(k => { if (f[k]) filtered = filtered.filter(d => d[k] === f[k]); });

  if (promiseF) {
    filtered = filtered.filter(d => {
      const ps = getCasePromiseStatus(d['Case Number']);
      if (promiseF === 'overdue') return ps === 'overdue';
      if (promiseF === 'today')   return ps === 'today';
      if (promiseF === 'pending') return !!ps;
      if (promiseF === 'none')    return !ps;
      return true;
    });
  }

  currentData = filtered;
  renderCharts(filtered);
  renderTable(filtered);
  resetFilters(filtered);
  renderWatchlist();
}*/



/*function resetFilters(data = allData) {
  ['accountFilter','ownerFilter','statusFilter','pendingFilter','bucketFilter','promiseFilter','INCFilter' ,'INCFilterqueue']
    .forEach(id => document.getElementById(id).value = '');
  currentData = allData;
  renderTable(allData);
  renderWatchlist();
  renderCharts(data);
}
*/




// state per filter id: which values are checked
const xfilterState = {};
// maps filter id -> data field key (e.g. accountFilter -> 'account')
const FILTER_KEYS  = {};

// ─── FILTERS: build all excel-style dropdowns ──
function renderFilters() {
  setupExcelFilter('accountFilter',   'account');
  setupExcelFilter('ownerFilter',     'owner');
  setupExcelFilter('statusFilter',    'status');
  setupExcelFilter('pendingFilter',   'pendingWith');
  setupExcelFilter('bucketFilter',    'bucket');
  setupExcelFilter('INCFilter',       'incOwner');
  setupExcelFilter('INCFilterqueue',  'incOwnerQueue');
}

function setupExcelFilter(id, key) {
  const sel = document.getElementById(id);
  if (!sel) return;

  sel.style.display = 'none'; // hide the original <select multiple>
  FILTER_KEYS[id] = key;
  xfilterState[id] = new Set(); // reset selection whenever data reloads

  // remove old widget if this ran before (e.g. re-uploaded file)
  const old = sel.parentNode.querySelector(`.xfilter[data-for="${id}"]`);
  if (old) old.remove();

  const wrapper = document.createElement('div');
  wrapper.className = 'xfilter';
  wrapper.dataset.for = id;
  wrapper.innerHTML = `
    <button type="button" class="xfilter-btn" data-id="${id}">
      <span class="btn-label">All</span>
      <span class="caret">▾</span>
    </button>
    <div class="xfilter-panel" data-id="${id}">
      <div class="xfilter-search"><input type="text" placeholder="Search..." /></div>
      <div class="xfilter-actions">
        <button type="button" data-act="all">Select all</button>
        <button type="button" data-act="clear">Clear</button>
      </div>
      <div class="xfilter-list"></div>
      <div class="xfilter-footer">
        <button type="button" class="cancel">Cancel</button>
        <button type="button" class="apply">Apply</button>
      </div>
    </div>
  `;
  sel.insertAdjacentElement('afterend', wrapper);

  // force-hide immediately, regardless of whether the CSS loaded correctly
  wrapper.querySelector('.xfilter-panel').style.display = 'none';

  wireExcelFilter(id, key, wrapper);
}

// values still available for this filter, given what's checked in the OTHER filters
function getAvailableValues(id, key) {
  const otherIds = Object.keys(xfilterState).filter(k => k !== id);
  const scoped = allData.filter(row =>
    otherIds.every(oid => {
      const s = xfilterState[oid];
      if (!s || s.size === 0) return true;
      return s.has(row[FILTER_KEYS[oid]]);
    })
  );
  return [...new Set(scoped.map(r => r[key]).filter(Boolean))].sort();
}

function wireExcelFilter(id, key, root) {
  const btn    = root.querySelector('.xfilter-btn');
  const panel  = root.querySelector('.xfilter-panel');
  const list   = root.querySelector('.xfilter-list');
  const search = root.querySelector('.xfilter-search input');

  panel.addEventListener('click', e => e.stopPropagation());

  let draft = new Set(xfilterState[id]);

  function renderList(filterText = '') {
    const values = getAvailableValues(id, key);
    const q = filterText.trim().toLowerCase();
    const shown = q ? values.filter(v => String(v).toLowerCase().includes(q)) : values;

    list.innerHTML = '';
    if (!shown.length) {
      list.innerHTML = `<div class="xfilter-empty">No matches</div>`;
      return;
    }
    shown.forEach(val => {
      const count = allData.filter(r => r[key] === val).length;
      const item = document.createElement('label');
      item.className = 'xfilter-item';
      item.innerHTML = `
        <input type="checkbox" value="${val}" ${draft.has(val) ? 'checked' : ''}/>
        <span class="lbl">${val}</span>
        <span class="n">${count}</span>
      `;
      item.querySelector('input').addEventListener('change', e => {
        e.target.checked ? draft.add(val) : draft.delete(val);
      });
      list.appendChild(item);
    });
  }

  function openPanel() {
    document.querySelectorAll('.xfilter-panel.open').forEach(p => {
      if (p !== panel) { p.classList.remove('open'); p.style.display = 'none'; }
    });
    draft = new Set(xfilterState[id]);
    search.value = '';
    renderList();
    panel.classList.add('open');
    panel.style.display = 'block';   // <-- forces it open regardless of CSS
    search.focus();
  }
  function closePanel() {
    panel.classList.remove('open');
    panel.style.display = 'none';    // <-- forces it closed regardless of CSS
  }

  btn.addEventListener('click', e => {
    e.stopPropagation();
    panel.classList.contains('open') ? closePanel() : openPanel();
  });

  search.addEventListener('input', () => renderList(search.value));

  panel.querySelector('[data-act="all"]').onclick = () => {
    getAvailableValues(id, key).forEach(v => draft.add(v));
    renderList(search.value);
  };
  panel.querySelector('[data-act="clear"]').onclick = () => {
    draft.clear();
    renderList(search.value);
  };
  panel.querySelector('.cancel').onclick = closePanel;
  panel.querySelector('.apply').onclick = () => {
    xfilterState[id] = new Set(draft);
    updateXFilterLabel(id);
    closePanel();
    applyFilters();
  };
}

function updateXFilterLabel(id) {
  const btn = document.querySelector(`.xfilter-btn[data-id="${id}"]`);
  if (!btn) return;
  const s = xfilterState[id];
  const label = btn.querySelector('.btn-label');
  if (!s || s.size === 0) {
    label.textContent = 'All';
    btn.classList.remove('active-filter');
  } else if (s.size === 1) {
    label.textContent = [...s][0];
    btn.classList.add('active-filter');
  } else {
    label.textContent = `${s.size} selected`;
    btn.classList.add('active-filter');
  }
}

function refreshOpenXFilterPanels() {
  document.querySelectorAll('.xfilter-panel.open').forEach(p => {
    p.querySelector('.xfilter-search input').dispatchEvent(new Event('input'));
  });
}

document.addEventListener('click', () => {
  document.querySelectorAll('.xfilter-panel.open').forEach(p => {
    p.classList.remove('open');
    p.style.display = 'none';
  });
});

function quickFilter(filterId, value) {
  if (!(filterId in xfilterState)) return;
  xfilterState[filterId] = new Set([value]);
  updateXFilterLabel(filterId);
  applyFilters();
}

// ─── APPLY FILTERS (same name, correct logic) ──
function applyFilters() {
  let filtered = allData.filter(row =>
    Object.keys(xfilterState).every(id => {
      const s = xfilterState[id];
      if (!s || s.size === 0) return true;
      return s.has(row[FILTER_KEYS[id]]);
    })
  );

  const promiseF = document.getElementById('promiseFilter').value;
  if (promiseF) {
    filtered = filtered.filter(d => {
      const ps = getCasePromiseStatus(d['Case Number']);
      if (promiseF === 'overdue') return ps === 'overdue';
      if (promiseF === 'today')   return ps === 'today';
      if (promiseF === 'pending') return !!ps;
      if (promiseF === 'none')    return !ps;
      return true;
    });
  }

  currentData = filtered;
  renderCharts(filtered);
  renderTable(filtered);
  renderWatchlist();
  refreshOpenXFilterPanels();
}

// ─── RESET FILTERS (same name, correct behavior) ──
function resetFilters(data = allData) {
  Object.keys(xfilterState).forEach(id => {
    xfilterState[id].clear();
    updateXFilterLabel(id);
  });
  const promiseEl = document.getElementById('promiseFilter');
  if (promiseEl) promiseEl.value = '';

  currentData = allData;
  renderTable(allData);
  renderWatchlist();
  renderCharts(allData);
}

//--end of new  code 

// ─── WATCHLIST ─────────────────────────────────
function renderWatchlist() {
  const el = document.getElementById('watchlistSection');
  if (!el) return;

  const starred = watchlist.activeItems();
  const rows = starred
    .map(cn => currentData.find(d => String(d['Case Number']) === String(cn)))
    .filter(Boolean);

  const kpiEl = document.getElementById('kpiWatchCount');
  if (kpiEl) kpiEl.textContent = rows.length;

  if (!rows.length) {
    el.innerHTML = `
      <div class="watchlist-empty">
        <span class="wl-empty-icon">☆</span>
        <span>No cases starred. Hit ★ on any row to watch it for tomorrow.</span>
      </div>`;
    return;
  }

  el.innerHTML = rows.map(row => {
    const caseNum  = row['Case Number'];
    const wlData   = watchlist.getCase(caseNum);
    const note     = wlData?.note || '';
    const ps       = getCasePromiseStatus(caseNum);
    const emails   = getEmails(caseNum);
    const age      = row.age;
    const ageClass = age > 30 ? 'age-high' : age > 20 ? 'age-med' : 'age-low';

    let promiseBadge = '';
    if (ps === 'overdue') promiseBadge = `<span class="wl-badge wl-badge-red">PROMISE OVERDUE</span>`;
    else if (ps === 'today') promiseBadge = `<span class="wl-badge wl-badge-yellow">PROMISE DUE TODAY</span>`;

    const lastEmail = emails[emails.length - 1];
    const emailLine = lastEmail
      ? `<span class="wl-email">${lastEmail.direction === 'received' ? '📥' : '📤'} ${lastEmail.subject}</span>`
      : `<span class="wl-email wl-email-none">no email logged</span>`;

    const starredDate = wlData?.starredAt
      ? new Date(wlData.starredAt).toLocaleDateString('en-IN', { day: '2-digit', month: 'short' })
      : '';

    return `
      <div class="wl-card" id="wlcard-${caseNum}">
        <div class="wl-card-top">
          <div class="wl-card-left">
            <span class="wl-case-num">${caseNum}</span>
            <span class="wl-account">${row.account}</span>
            ${promiseBadge}
          </div>
          <div class="wl-card-right">
            <span class="wl-age ${ageClass}">${age !== null ? age+'d' : '—'}</span>
            <span class="wl-pending">${row.pendingWith}</span>
            <span class="wl-starred-date">starred ${starredDate}</span>
          </div>
        <div class="wl-card-bottom">
          ${emailLine}
          <div class="wl-note-row">
            <input
              class="wl-note-input"
              type="text"
              value="${note.replace(/"/g, '&quot;')}"
              placeholder="Follow-up note…"
              onchange="watchlist.updateNote('${caseNum}', this.value)"
            />
            <button class="wl-open-btn" onclick="openDrawer('${caseNum}')">OPEN ›</button>
            <button class="wl-done-btn" onclick="unstarCase('${caseNum}')" title="Mark done & remove from watchlist">✓ DONE</button>
          </div>
        </div>
      </div>
    `;
  }).join('');
}

function toggleStar(caseNum, btn) {
  if (watchlist.isStarred(caseNum)) {
    watchlist.unstar(caseNum);
    btn.classList.remove('starred');
    btn.title = 'Watch this case';
  } else {
    watchlist.star(caseNum);
    btn.classList.add('starred');
    btn.title = 'Remove from watchlist';
  }
  renderWatchlist();
}

function unstarCase(caseNum) {
  watchlist.unstar(caseNum);
  const card = document.getElementById(`wlcard-${caseNum}`);
  if (card) {
    card.style.transition = 'opacity 0.3s, transform 0.3s';
    card.style.opacity = '0';
    card.style.transform = 'translateX(20px)';
    setTimeout(() => renderWatchlist(), 320);
  } else {
    renderWatchlist();
  }
  const starBtn = document.getElementById(`star-${caseNum}`);
  if (starBtn) starBtn.classList.remove('starred');
}

// ─── TABLE ─────────────────────────────────────
function renderTable(data) {
  const starredInView = data.filter(r => watchlist.isStarred(r['Case Number'])).length;
  const countEl = document.getElementById('tableCount');
  countEl.textContent = starredInView
    ? `${data.length} CASES · ${starredInView} ★ PINNED`
    : `${data.length} CASES`;

  const today = new Date(); today.setHours(0,0,0,0);
  const tbody = document.getElementById('tableBody');

  const starred = data
    .filter(r => watchlist.isStarred(r['Case Number']))
    .sort((a, b) => {
      const wa = watchlist.getCase(a['Case Number'])?.starredAt || '';
      const wb = watchlist.getCase(b['Case Number'])?.starredAt || '';
      return new Date(wa) - new Date(wb);
    });
  const unstarred = data.filter(r => !watchlist.isStarred(r['Case Number']));

  const buildRow = (row, addDivider) => {
    const caseNum   = row['Case Number'];
    const status    = row.status || '';
    const age       = row.age;
    const ps        = getCasePromiseStatus(caseNum);
    const promises  = getPromises(caseNum).filter(p => !p.done);
    const isStarred = watchlist.isStarred(caseNum);

    const sLower    = status.toLowerCase();
    const pillClass = sLower.includes('closed') ? 'pill-closed'
      : sLower.includes('progress') ? 'pill-open' : 'pill-pending';

    const ageClass  = age > 30 ? 'age-high' : age > 20 ? 'age-med' : 'age-low';
    const incAge    = parseInt(row.incAge);
    const incAgeClass = !isNaN(incAge)
      ? incAge > 30 ? 'age-high' : incAge > 20 ? 'age-med' : 'age-low' : '';

    const ageDisplay = age !== null ? `<span class="age-num ${ageClass}">${age}d</span>` : '—';

    let promiseCell = `<span class="promise-indicator promise-none">—</span>`;
    if (ps === 'overdue') {
      const p = promises[0];
      const d = new Date(p.date); d.setHours(0,0,0,0);
      const diff = Math.abs(Math.floor((d - today) / 86400000));
      promiseCell = `<span class="promise-indicator promise-overdue">🔴 ${diff}d overdue</span>`;
    } else if (ps === 'today') {
      promiseCell = `<span class="promise-indicator promise-today">🟡 due today</span>`;
    } else if (ps === 'upcoming') {
      const p = promises[0];
      const d = new Date(p.date); d.setHours(0,0,0,0);
      const diff = Math.ceil((d - today) / 86400000);
      promiseCell = `<span class="promise-indicator promise-ok">🟢 in ${diff}d</span>`;
    }

    let rowClass = isStarred ? 'row-starred' : '';
    if (ps === 'overdue') rowClass += ' has-alert';
    else if (ps === 'today') rowClass += ' has-today';

    let nextUpdateClass = '';
    if (row.nextUpdate) {
      const nDate = parseDDMMYYYY(row.nextUpdate);
      if (nDate) {
        nDate.setHours(0,0,0,0);
        if (nDate < today) nextUpdateClass = 'overdue';
        else if (nDate.getTime() === today.getTime()) nextUpdateClass = 'today';
      }
    }

    let deliveryClass = '', deliveryDisplay = '—';
    if (row.deliveryDate) {
      const dDate = parseDeliveryDate(row.deliveryDate);
      if (dDate && !isNaN(dDate.getTime())) {
        dDate.setHours(0,0,0,0);
        if (dDate < today) deliveryClass = 'overdue';
        else if (dDate.getTime() === today.getTime()) deliveryClass = 'today';
        deliveryDisplay = dDate.toLocaleDateString('en-IN');
      }
    }

    const starClass = isStarred ? 'star-btn starred' : 'star-btn';
    const starTitle = isStarred ? 'Unstar — remove from top' : 'Star — pin to top';

    const divider = addDivider
      ? `<tr class="table-divider"><td colspan="14"><span>── WATCHLIST ──────────────────── REST OF CASES ──</span></td></tr>`
      : '';

    return divider + `
      <tr class="${rowClass.trim()}">
        <td><button id="star-${caseNum}" class="${starClass}" onclick="toggleStar('${caseNum}', this)" title="${starTitle}">★</button></td>
        <td><span class="status-pill ${pillClass}">${status}</span></td>
        <td><span class="case-num">${caseNum}</span></td>
        <td style="max-width:90px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${row.account}</td>
        <td style="white-space:nowrap">${row.owner}</td>
        <td>${ageDisplay}</td>
        <td>${row.pendingWith}</td>
        <td>${row.inc ? `<span class="inc-badge">${row.inc}</span>` : '<span style="color:var(--text-muted)">—</span>'}</td>
        <td>${!isNaN(incAge) ? `<span class="age-num ${incAgeClass}">${incAge}d</span>` : '—'}</td>
        <td>${row.incOwner || '—'}</td>
        <td class="${deliveryClass}">${deliveryDisplay}</td>
        <td class="${nextUpdateClass}">
  ${row.nextUpdate ? (parseDDMMYYYY(row.nextUpdate)?.toLocaleDateString('en-IN') || '—') : '—'}
</td>
<td>${(() => {
  if (!row.emailLastSent) return '<span style="color:var(--danger);font-family:var(--mono);font-size:10px;">NEVER</span>';
  const d =  parseDate(row.emailLastSent);
  if (!d ||isNaN(d.getTime())) return '—';
  const today = new Date(); today.setHours(0,0,0,0); d.setHours(0,0,0,0);
  const diff = Math.floor((today - d) / 86400000);
  const label = d.toLocaleDateString('en-IN');
  if (diff > 3) return `<span style="color:var(--danger);font-family:var(--mono);font-size:10px;">${label} <em>(${diff}d ago)</em></span>`;
  if (diff > 1) return `<span style="color:var(--warning);font-family:var(--mono);font-size:10px;">${label}</span>`;
  return `<span style="color:var(--success);font-family:var(--mono);font-size:10px;">${label}</span>`;
})()} </td>
<td>${promiseCell}</td>
        <td><button class="open-btn" onclick="openDrawer('${caseNum}')">OPEN ›</button></td>
      </tr>
    `;
  };

  const starredHTML   = starred.map(r => buildRow(r, false)).join('');
  const unstarredHTML = unstarred.map((r, i) => buildRow(r, i === 0 && starred.length > 0)).join('');
  tbody.innerHTML = starredHTML + unstarredHTML;
}

// ─── DRAWER ────────────────────────────────────
function openDrawer(caseNum) {
  const caseData = allData.find(d => d['Case Number'] == caseNum);
  if (!caseData) return;
  activeDrawerCase = caseNum;

  document.getElementById('drawerCaseNum').textContent = caseNum;
  document.getElementById('drawerAccount').textContent = caseData.account;

  const age    = caseData.age;
  const status = caseData.status || '';
  const sClass = status.toLowerCase().includes('closed') ? 'status-closed'
    : status.toLowerCase().includes('progress') ? 'status-open' : 'status-pending';

  document.getElementById('drawerMeta').innerHTML = `
    <span class="meta-tag ${sClass}">${status}</span>
    <span class="meta-tag age-tag">${age !== null ? age + ' DAYS OLD' : 'AGE UNKNOWN'}</span>
    <span class="meta-tag age-tag">${caseData.pendingWith}</span>
    ${caseData.inc ? `<span class="meta-tag" style="border-color:var(--accent3);color:var(--accent3)">${caseData.inc}</span>` : ''}
    <span class="meta-tag age-tag">${caseData.Subject}</span>
  `;
  const commentSection = `
  <div class="latest-comment-box">
    <div class="lc-title">LATEST COMMENT</div>

    <div class="lc-comment">
      ${caseData.latestComment || 'No comments available'}
    </div>

    <div class="lc-meta">
  👤 ${caseData.modifiedBy || '—'}
  <span style="margin-left:12px">📅 ${caseData.modifiedDate || '—'}</span>
  <span style="margin-left:12px">📧 Last email: ${caseData.emailLastSent ? (parseDate(caseData.emailLastSent) || new Date()).toLocaleDateString('en-IN') : '<span style="color:var(--danger)">Never sent</span>'}</span>
</div>
  </div>
`;
document.getElementById('drawerMeta').insertAdjacentHTML(
  'beforeend',
  commentSection
);



  document.getElementById('caseNotes').value = store.getCase(caseNum, 'notes') || '';
  renderEmailLog(caseNum);
  renderPromiseLog(caseNum);

  document.getElementById('drawerOverlay').classList.add('open');
  document.getElementById('drawer').classList.add('open');
}

function closeDrawer() {
  document.getElementById('drawerOverlay').classList.remove('open');
  document.getElementById('drawer').classList.remove('open');
  activeDrawerCase = null;
}

// ─── EMAIL LOG ─────────────────────────────────
function renderEmailLog(caseNum) {
  const emails = getEmails(caseNum);
  const log = document.getElementById('emailLog');
  if (!emails.length) {
    log.innerHTML = `<div style="color:var(--text-dim);font-size:12px;padding:8px 0">No emails logged yet. Add the first one below.</div>`;
    return;
  }
  log.innerHTML = emails.map((e, i) => `
    <div class="email-entry ${e.direction}">
      <span class="email-dir-icon">${e.direction === 'received' ? '📥' : '📤'}</span>
      <div class="email-entry-body">
        <div class="email-subject">${e.subject}</div>
        <div class="email-time">
          ${e.direction.toUpperCase()} · ${e.timestamp}
          ${e.deliveryDate ? `<br>📅 ETA: ${e.deliveryDate}` : ''}
        </div>
      </div>
      <button class="email-delete" onclick="deleteEmail('${caseNum}', ${i})">✕</button>
    </div>
  `).join('');
}

function addEmail() {
  const subject      = document.getElementById('emailSubject').value.trim();
  const direction    = document.getElementById('emailDir').value;
  const deliveryDate = document.getElementById('actualDeliveryDate').value;
  if (!subject || !activeDrawerCase) return;

  const emails = getEmails(activeDrawerCase);
  emails.push({
    subject, direction, deliveryDate,
    timestamp: new Date().toLocaleString('en-IN', { day: '2-digit', month: 'short', hour: '2-digit', minute: '2-digit' })
  });

  store.setCase(activeDrawerCase, 'emails', emails);
  document.getElementById('emailSubject').value = '';
  document.getElementById('actualDeliveryDate').value = '';
  renderEmailLog(activeDrawerCase);
  renderTable(currentData);
}

function deleteEmail(caseNum, index) {
  const emails = getEmails(caseNum);
  emails.splice(index, 1);
  store.setCase(caseNum, 'emails', emails);
  renderEmailLog(caseNum);
  renderTable(currentData);
}

// ─── PROMISE LOG ───────────────────────────────
function renderPromiseLog(caseNum) {
  const promises = getPromises(caseNum);
  const log = document.getElementById('promiseLog');
  const today = new Date(); today.setHours(0,0,0,0);

  if (!promises.length) {
    log.innerHTML = `<div style="color:var(--text-dim);font-size:12px;padding:8px 0">No promises tracked. Add commitments below to get alerts.</div>`;
    return;
  }

  log.innerHTML = promises.map((p, i) => {
    const pd = new Date(p.date); pd.setHours(0,0,0,0);
    const diff = Math.floor((pd - today) / 86400000);
    let cls = 'upcoming', dueLabel = '';
    if (p.done)        { cls = 'done';      dueLabel = 'DONE'; }
    else if (diff < 0)  { cls = 'overdue';   dueLabel = `${Math.abs(diff)}D OVERDUE`; }
    else if (diff === 0){ cls = 'due-today'; dueLabel = 'DUE TODAY'; }
    else dueLabel = `in ${diff}d`;

    const dueColor = p.done ? 'var(--text-dim)'
      : diff < 0  ? 'var(--danger)'
      : diff === 0 ? 'var(--warning)'
      : 'var(--success)';

    return `
      <div class="promise-entry ${cls}">
        <div class="promise-body">
          <div class="promise-text-main">${p.text}</div>
          <div class="promise-meta">
            <span class="promise-due" style="color:${dueColor};font-family:var(--mono)">${dueLabel}</span>
            <span class="promise-by">BY: ${p.owner}</span>
            <span class="promise-by">${p.date}</span>
          </div>
        </div>
        <div class="promise-actions">
          ${!p.done ? `<button class="promise-done-btn" onclick="markPromiseDone('${caseNum}', ${i})">✓ DONE</button>` : ''}
          <button class="email-delete" onclick="deletePromise('${caseNum}', ${i})">✕</button>
        </div>
      </div>
    `;
  }).join('');
}

function addPromise() {
  const text  = document.getElementById('promiseText').value.trim();
  const date  = document.getElementById('promiseDate').value;
  const owner = document.getElementById('promiseOwner').value;
  if (!text || !date || !activeDrawerCase) return;

  const promises = getPromises(activeDrawerCase);
  promises.push({ text, date, owner, done: false, created: new Date().toISOString() });
  store.setCase(activeDrawerCase, 'promises', promises);
  document.getElementById('promiseText').value = '';
  document.getElementById('promiseDate').value = '';
  renderPromiseLog(activeDrawerCase);
  renderAlertStrip();
  updateBellBadge();
  renderTable(currentData);
}

function markPromiseDone(caseNum, index) {
  const promises = getPromises(caseNum);
  promises[index].done = true;
  store.setCase(caseNum, 'promises', promises);
  renderPromiseLog(caseNum);
  renderAlertStrip();
  updateBellBadge();
  renderTable(currentData);
}

function deletePromise(caseNum, index) {
  const promises = getPromises(caseNum);
  promises.splice(index, 1);
  store.setCase(caseNum, 'promises', promises);
  renderPromiseLog(caseNum);
  renderAlertStrip();
  updateBellBadge();
  renderTable(currentData);
}

// ─── NOTES ─────────────────────────────────────
function saveNotes() {
  if (!activeDrawerCase) return;
  store.setCase(activeDrawerCase, 'notes', document.getElementById('caseNotes').value);
}