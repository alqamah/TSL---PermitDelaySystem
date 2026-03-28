/* ═══════════════════════════════════════════
   PERMIT DELAY SYSTEM — Core Logic
   ═══════════════════════════════════════════
   Reads multi-sheet Excel files, extracts crane
   permit data, computes delays, and renders a
   searchable table.

   KEY:  cellDates is OFF so SheetJS returns raw
         Excel serial numbers — no timezone shift.
         cell.w is used for display, raw fractions
         for delay math.
   ═══════════════════════════════════════════ */

// ─── DOM References ──────────────────────────
const dropZone     = document.getElementById('dropZone');
const fileInput    = document.getElementById('fileInput');
const browseBtn    = document.getElementById('browseBtn');
const fileNameEl   = document.getElementById('fileName');
const statsRow     = document.getElementById('statsRow');
const tableSection = document.getElementById('tableSection');
const tableBody    = document.getElementById('tableBody');
const searchInput  = document.getElementById('searchInput');

// ─── State ───────────────────────────────────
let allRecords = [];   // master flat list of parsed rows

// ─── File Upload Events ──────────────────────
browseBtn.addEventListener('click', (e) => {
  e.stopPropagation();
  fileInput.click();
});

dropZone.addEventListener('click', () => fileInput.click());

fileInput.addEventListener('change', (e) => {
  if (e.target.files.length) handleFile(e.target.files[0]);
});

dropZone.addEventListener('dragover', (e) => {
  e.preventDefault();
  dropZone.classList.add('drag-over');
});

dropZone.addEventListener('dragleave', () => {
  dropZone.classList.remove('drag-over');
});

dropZone.addEventListener('drop', (e) => {
  e.preventDefault();
  dropZone.classList.remove('drag-over');
  if (e.dataTransfer.files.length) handleFile(e.dataTransfer.files[0]);
});

// ─── Search / Filter ─────────────────────────
searchInput.addEventListener('input', () => {
  const term = searchInput.value.toLowerCase().trim();
  renderTable(
    term
      ? allRecords.filter(r =>
          r.craneName.toLowerCase().includes(term) ||
          r.department.toLowerCase().includes(term) ||
          r.requester.toLowerCase().includes(term) ||
          r.dateStr.toLowerCase().includes(term)
        )
      : allRecords
  );
});

// ═══════════════════════════════════════════════
//  CORE:  Read Excel → Parse Sheets → Build Data
// ═══════════════════════════════════════════════
function handleFile(file) {
  if (!file) return;
  fileNameEl.textContent = `📄 ${file.name}`;

  const reader = new FileReader();
  reader.onload = (e) => {
    const data = new Uint8Array(e.target.result);
    // ⚠ cellDates:false (default) — keeps raw serial numbers,
    //   avoids the timezone-shift bug with JS Date objects.
    const workbook = XLSX.read(data, { type: 'array' });
    allRecords = parseWorkbook(workbook);
    showResults();
  };
  reader.readAsArrayBuffer(file);
}

// ─── Parse every sheet in the workbook ────────
function parseWorkbook(wb) {
  const craneMap = {};  // craneName → [ records ]

  wb.SheetNames.forEach((sheetName) => {
    const ws = wb.Sheets[sheetName];
    if (!ws) return;

    // --- Crane name: prefer cell F2, fallback to sheet name ---
    let craneName = sheetName;
    const cellF2 = ws['F2'];
    if (cellF2) {
      const raw = String(cellF2.v || '').trim();
      // Format is "CRANE :- 40T-1", extract after ":-"
      const match = raw.match(/CRANE\s*[:\-]+\s*(.+)/i);
      craneName = match ? match[1].trim() : (raw || sheetName);
    }

    if (!craneMap[craneName]) craneMap[craneName] = [];

    // --- Walk rows starting from row 4 (index 3) ---
    const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
    for (let R = 3; R <= range.e.r; R++) {
      // Column B (1) = Date
      const cellDate   = ws[XLSX.utils.encode_cell({ r: R, c: 1 })];
      // Column D (3) = Department
      const cellDept   = ws[XLSX.utils.encode_cell({ r: R, c: 3 })];
      // Column E (4) = Requester Name
      const cellReq    = ws[XLSX.utils.encode_cell({ r: R, c: 4 })];
      // Column G (6) = Crane Reporting Time at Site
      const cellReport = ws[XLSX.utils.encode_cell({ r: R, c: 6 })];
      // Column I (8) = Permit Handover Time
      const cellPermit = ws[XLSX.utils.encode_cell({ r: R, c: 8 })];

      // Skip empty rows (require at least a date)
      if (!cellDate) continue;

      const dateVal = parseDateValue(cellDate);
      if (!dateVal) continue;  // truly empty / invalid

      const department  = cellDept ? String(cellDept.v || '').trim() : '';
      const requester   = cellReq  ? String(cellReq.v  || '').trim() : '';

      // Extract raw fractional values for time math
      const reportFrac  = getRawTimeFraction(cellReport);
      const permitFrac  = getRawTimeFraction(cellPermit);

      // Display strings: use cell.w (Excel-formatted) when available
      const reportDisplay = getTimeDisplay(cellReport);
      const permitDisplay = getTimeDisplay(cellPermit);

      // Delay = (permitFrac − reportFrac) × 1440 minutes
      const delayMin = computeDelay(reportFrac, permitFrac);

      // Hourly Rate and Total Amount
      const craneRate = typeof getCraneRate === 'function' ? getCraneRate(craneName) : 0;
      const amount    = delayMin > 0 ? (delayMin / 60) * craneRate : 0;

      craneMap[craneName].push({
        craneRate,
        amount,
        craneName,
        date:       dateVal,
        dateStr:    formatDate(dateVal),
        department,
        requester,
        reportTime: reportDisplay,
        permitTime: permitDisplay,
        delayMin,
        delayStr:   delayMin !== null ? formatDelay(delayMin) : '—'
      });
    }
  });

  // Flatten to a single sorted array
  const flat = [];
  Object.values(craneMap).forEach(arr => flat.push(...arr));
  flat.sort((a, b) => a.date - b.date);
  return flat;
}

// ═══════════════════════════════════════════════
//  HELPERS:  Date parsing
// ═══════════════════════════════════════════════

/**
 * Parse a date cell value into a JS Date.
 * With cellDates OFF, dates arrive as Excel serial numbers.
 */
function parseDateValue(cell) {
  if (!cell) return null;
  const v = cell.v;
  if (v instanceof Date) return v;
  if (typeof v === 'number') {
    return excelSerialToDate(v);
  }
  if (typeof v === 'string') {
    const d = new Date(v);
    return isNaN(d) ? null : d;
  }
  return null;
}

/** Convert an Excel serial date (days since 1900-01-01) to JS Date */
function excelSerialToDate(serial) {
  // Excel epoch is 1900-01-01, but has the Lotus 1-2-3 leap-year bug
  const epoch = new Date(Date.UTC(1899, 11, 30));
  return new Date(epoch.getTime() + Math.floor(serial) * 86400000);
}

// ═══════════════════════════════════════════════
//  HELPERS:  Time extraction (NO Date objects!)
// ═══════════════════════════════════════════════

/**
 * Get the raw Excel fractional-day value (0–1) from a time cell.
 * • If cell.v is a number, extract fractional part.
 * • If cell.v is a string like "1:27:23 AM", parse to fraction.
 * Returns a fraction (0–1) or null.
 */
function getRawTimeFraction(cell) {
  if (!cell) return null;
  const v = cell.v;

  if (typeof v === 'number') {
    // Could be a pure time (< 1) or a date+time serial (> 1)
    const frac = v % 1;
    // If fraction is essentially zero and the value is > 1,
    // it's a date-only cell with no time component
    if (frac < 0.00001 && v > 1) return null;
    return frac;
  }

  if (typeof v === 'string') {
    return parseTimeStringToFraction(v.trim());
  }

  return null;
}

/**
 * Get the display string for a time cell.
 * Prefers cell.w (Excel's formatted text), falls back to
 * converting the raw fraction ourselves.
 */
function getTimeDisplay(cell) {
  if (!cell) return '';

  // Prefer Excel's own formatted string — guaranteed accurate
  if (cell.w && cell.w.trim()) return cell.w.trim();

  // Fallback: convert raw fraction to readable string
  const frac = getRawTimeFraction(cell);
  if (frac === null) return '';
  return fractionToTimeString(frac);
}

/**
 * Parse "1:27:23 AM" / "10:42:20 PM" → fraction of day.
 */
function parseTimeStringToFraction(str) {
  if (!str) return null;
  const match = str.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(AM|PM)?$/i);
  if (!match) return null;
  let h = parseInt(match[1], 10);
  const m = parseInt(match[2], 10);
  const s = match[3] ? parseInt(match[3], 10) : 0;
  const ampm = (match[4] || '').toUpperCase();
  if (ampm === 'PM' && h < 12) h += 12;
  if (ampm === 'AM' && h === 12) h = 0;
  return (h * 3600 + m * 60 + s) / 86400;
}

/**
 * Convert a fraction of day (0–1) to "h:mm:ss AM/PM" string.
 */
function fractionToTimeString(frac) {
  const totalSecs = Math.round(frac * 86400);
  let h = Math.floor(totalSecs / 3600);
  const m = Math.floor((totalSecs % 3600) / 60);
  const s = totalSecs % 60;
  const ampm = h >= 12 ? 'PM' : 'AM';
  h = h % 12 || 12;
  return `${h}:${String(m).padStart(2, '0')}:${String(s).padStart(2, '0')} ${ampm}`;
}

// ═══════════════════════════════════════════════
//  HELPERS:  Delay computation & formatting
// ═══════════════════════════════════════════════

/**
 * Delay = (permitFrac − reportFrac) × 1440  →  minutes.
 */
function computeDelay(reportFrac, permitFrac) {
  if (reportFrac === null || permitFrac === null) return null;
  const diffMinutes = Math.round((permitFrac - reportFrac) * 1440);
  return diffMinutes;
}

function formatDate(d) {
  if (!d) return '';
  const day   = String(d.getUTCDate()).padStart(2, '0');
  const months = ['Jan','Feb','Mar','Apr','May','Jun',
                  'Jul','Aug','Sep','Oct','Nov','Dec'];
  const mon   = months[d.getUTCMonth()];
  const year  = String(d.getUTCFullYear()).slice(-2);
  return `${day}-${mon}-${year}`;
}

function formatDelay(mins) {
  if (mins === null) return '—';
  const sign = mins < 0 ? '-' : '';
  const abs  = Math.abs(mins);
  const h    = Math.floor(abs / 60);
  const m    = abs % 60;
  if (h > 0) return `${sign}${h}h ${String(m).padStart(2,'0')}m`;
  return `${sign}${m}m`;
}

// ═══════════════════════════════════════════════
//  RENDER:  Stats + Table
// ═══════════════════════════════════════════════
function showResults() {
  const totalRecords = allRecords.length;
  const delays = allRecords.map(r => r.delayMin).filter(d => d !== null && d >= 0);
  const totalDelayMin = delays.reduce((a, b) => a + b, 0);
  const avgDelay = delays.length ? Math.round(totalDelayMin / delays.length) : null;

  document.getElementById('statTotal').textContent = totalRecords;
  document.getElementById('statAvgDelay').textContent = avgDelay !== null
    ? formatDelay(avgDelay) : '—';
  document.getElementById('statTotalDelay').textContent = totalDelayMin > 0
    ? formatDelay(totalDelayMin) : '—';

  const totalAmount = allRecords.reduce((acc, curr) => acc + (curr.amount || 0), 0);
  document.getElementById('statTotalAmount').textContent = totalAmount > 0
    ? '₹ ' + totalAmount.toLocaleString('en-IN', { minimumFractionDigits: 0, maximumFractionDigits: 0 })
    : '₹ 0';

  statsRow.classList.remove('hidden');
  tableSection.classList.remove('hidden');
  document.getElementById('navPanelTop').classList.remove('hidden');

  renderTable(allRecords);

  // Department summary
  const deptSummary = buildDeptSummary(allRecords);
  renderDeptTable(deptSummary);
  renderDeptChart(deptSummary);
  document.getElementById('deptSection').classList.remove('hidden');
}

function renderTable(records) {
  tableBody.innerHTML = '';

  records.forEach((r, i) => {
    const tr = document.createElement('tr');

    // Determine delay severity for pill styling
    let pillClass = '';
    if (r.delayMin !== null) {
      if      (r.delayMin < 0)    pillClass = 'red';
      else if (r.delayMin < 60)   pillClass = 'green';
      else if (r.delayMin >= 480) pillClass = 'orange';
      else                        pillClass = 'yellow';
    }

    tr.innerHTML = `
      <td>${i + 1}</td>
      <td>${r.dateStr}</td>
      <td>${r.craneName}</td>
      <td>${r.department}</td>
      <td>${r.requester}</td>
      <td>${r.reportTime}</td>
      <td>${r.permitTime}</td>
      <td><span class="delay-pill ${pillClass}">${r.delayStr}</span></td>
      <td class="amount-cell">
        ${r.amount > 0 ? '₹ ' + r.amount.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) : '—'}
      </td>
    `;
    tableBody.appendChild(tr);
  });
}
/**
 * Lookup the hourly rate for a crane from CRANE_RATES.
 */
function getCraneRate(name) {
  if (typeof CRANE_RATES === 'undefined' || !name) return 0;
  
  // Try exact match first (e.g., "160T-1")
  if (CRANE_RATES[name]) return CRANE_RATES[name];
  
  // Try prefix match (e.g., "40T-1" matches "40T")
  // Sort keys by length descending to match most specific first
  const keys = Object.keys(CRANE_RATES).sort((a, b) => b.length - a.length);
  for (const key of keys) {
    if (name.toUpperCase().startsWith(key.toUpperCase())) {
      return CRANE_RATES[key];
    }
  }
  return 0;
}

// ═══════════════════════════════════════════════
//  DEPARTMENT SUMMARY:  Aggregate by Dept
// ═══════════════════════════════════════════════

/** Crane types to show in the summary (column order) */
const CRANE_TYPES = ['160T', '100T', '80T', '55T', '40T', '300T'];

/**
 * Extract the tonnage type from a crane name.
 * e.g. "40T-1" → "40T", "160T" → "160T", "300T" → "300T"
 */
function getCraneType(craneName) {
  if (!craneName) return null;
  // Match digits followed by 'T' (case insensitive)
  const match = craneName.toUpperCase().match(/(\d+T)/);
  return match ? match[1] : null;
}

/**
 * Build department summary from all records.
 * Returns an array of dept objects sorted by DEPT. TOTAL descending.
 */
function buildDeptSummary(records) {
  const deptMap = {}; // dept → { contact, cranes: { type → { hours, amount } }, total }

  records.forEach(r => {
    if (!r.department) return;
    const dept = r.department;

    if (!deptMap[dept]) {
      deptMap[dept] = {
        department: dept,
        contactCounts: {},  // requester → count (to find most frequent)
        cranes: {},
        total: 0
      };
      CRANE_TYPES.forEach(t => {
        deptMap[dept].cranes[t] = { hours: 0, amount: 0 };
      });
    }

    const entry = deptMap[dept];

    // Track requester frequency
    if (r.requester) {
      entry.contactCounts[r.requester] = (entry.contactCounts[r.requester] || 0) + 1;
    }

    // Aggregate by crane type
    const craneType = getCraneType(r.craneName);
    if (craneType && entry.cranes[craneType]) {
      const delayHrs = (r.delayMin !== null && r.delayMin > 0) ? r.delayMin / 60 : 0;
      entry.cranes[craneType].hours  += delayHrs;
      entry.cranes[craneType].amount += (r.amount || 0);
    }
  });

  // Compute totals and contact person
  const result = Object.values(deptMap).map(entry => {
    // Most frequent requester = contact person
    let maxCount = 0, contactPerson = '';
    for (const [name, count] of Object.entries(entry.contactCounts)) {
      if (count > maxCount) { maxCount = count; contactPerson = name; }
    }

    // Dept total = sum of all crane amounts
    let deptTotal = 0;
    CRANE_TYPES.forEach(t => { deptTotal += entry.cranes[t].amount; });

    return {
      department: entry.department,
      contactPerson,
      cranes: entry.cranes,
      total: deptTotal
    };
  });

  // Sort by total descending
  result.sort((a, b) => b.total - a.total);
  return result;
}

/**
 * Render the department summary table.
 */
function renderDeptTable(summary) {
  const tbody = document.getElementById('deptTableBody');
  tbody.innerHTML = '';

  let netLoss = 0;

  summary.forEach((row, i) => {
    const tr = document.createElement('tr');
    netLoss += row.total;

    let cells = `
      <td>${i + 1}</td>
      <td class="dept-name-cell">${row.department}</td>
      <td>${row.contactPerson}</td>
    `;

    CRANE_TYPES.forEach(type => {
      const c = row.cranes[type];
      const hrs = c.hours > 0 ? c.hours.toFixed(2) : '0';
      const amt = c.amount > 0
        ? c.amount.toLocaleString('en-IN', { minimumFractionDigits: 0, maximumFractionDigits: 0 })
        : '0';
      cells += `<td class="hrs-cell">${hrs}</td>`;
      cells += `<td class="amt-cell">${amt}</td>`;
    });

    cells += `<td class="dept-total-cell">₹ ${row.total.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</td>`;

    tr.innerHTML = cells;
    tbody.appendChild(tr);
  });

  // NET LOSS footer
  const netLossCell = document.getElementById('netLossCell');
  netLossCell.innerHTML = `
    <span class="net-loss-label">NET LOSS:</span>
    <span class="net-loss-value">₹ ${netLoss.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</span>
  `;
}

// ═══════════════════════════════════════════════
//  CHART:  Interactive Department Graph
// ═══════════════════════════════════════════════

let deptChartInstance = null;
let currentChartType = 'bar';
let cachedDeptSummary = null;

/** Color palette for departments */
const CHART_COLORS = [
  '#3b82f6', '#22c55e', '#eab308', '#f97316', '#ef4444',
  '#a855f7', '#06b6d4', '#ec4899', '#14b8a6', '#f43f5e',
  '#8b5cf6', '#10b981', '#fbbf24', '#fb923c', '#f87171'
];

/**
 * Render the department chart (called from showResults).
 */
function renderDeptChart(summary) {
  cachedDeptSummary = summary;
  document.getElementById('chartSection').classList.remove('hidden');

  // Truncate long department names for chart labels
  const labels = summary.map(d => {
    const name = d.department;
    return name.length > 18 ? name.slice(0, 16) + '…' : name;
  });

  if (currentChartType === 'bar') {
    renderBarChart(summary, labels);
  } else {
    renderDoughnutChart(summary, labels);
  }
}

function renderBarChart(summary, labels) {
  destroyChart();

  const totalHours = summary.map(d => {
    let hrs = 0;
    CRANE_TYPES.forEach(t => { hrs += d.cranes[t].hours; });
    return parseFloat(hrs.toFixed(2));
  });

  const totalAmounts = summary.map(d => Math.round(d.total));

  const ctx = document.getElementById('deptChart').getContext('2d');
  deptChartInstance = new Chart(ctx, {
    type: 'bar',
    data: {
      labels,
      datasets: [
        {
          label: 'Delay Hours',
          data: totalHours,
          backgroundColor: 'rgba(59, 130, 246, 0.7)',
          borderColor: '#3b82f6',
          borderWidth: 1,
          borderRadius: 6,
          yAxisID: 'yHours',
          order: 2
        },
        {
          label: 'Amount (₹)',
          data: totalAmounts,
          backgroundColor: 'rgba(239, 68, 68, 0.55)',
          borderColor: '#ef4444',
          borderWidth: 1,
          borderRadius: 6,
          yAxisID: 'yAmount',
          order: 1
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      interaction: { mode: 'index', intersect: false },
      plugins: {
        legend: {
          labels: {
            color: '#94a3b8',
            font: { family: "'Inter', sans-serif", size: 12, weight: 600 },
            padding: 16,
            usePointStyle: true,
            pointStyle: 'rectRounded'
          }
        },
        tooltip: {
          backgroundColor: '#1a2233',
          titleColor: '#f1f5f9',
          bodyColor: '#94a3b8',
          borderColor: '#2a3650',
          borderWidth: 1,
          cornerRadius: 8,
          padding: 12,
          titleFont: { family: "'Inter', sans-serif", weight: 700 },
          bodyFont: { family: "'Inter', sans-serif" },
          callbacks: {
            label: function(ctx) {
              if (ctx.dataset.yAxisID === 'yAmount') {
                return `  Amount: ₹ ${ctx.parsed.y.toLocaleString('en-IN')}`;
              }
              return `  Hours: ${ctx.parsed.y}h`;
            }
          }
        }
      },
      scales: {
        x: {
          ticks: {
            color: '#64748b',
            font: { size: 11, weight: 500 },
            maxRotation: 45,
            minRotation: 0
          },
          grid: { display: false }
        },
        yHours: {
          type: 'linear',
          position: 'left',
          title: {
            display: true,
            text: 'Delay Hours',
            color: '#3b82f6',
            font: { size: 12, weight: 600 }
          },
          ticks: { color: '#3b82f6', font: { size: 11 } },
          grid: { color: 'rgba(59, 130, 246, 0.08)' },
          beginAtZero: true
        },
        yAmount: {
          type: 'linear',
          position: 'right',
          title: {
            display: true,
            text: 'Amount (₹)',
            color: '#ef4444',
            font: { size: 12, weight: 600 }
          },
          ticks: {
            color: '#ef4444',
            font: { size: 11 },
            callback: function(val) {
              return '₹' + val.toLocaleString('en-IN');
            }
          },
          grid: { drawOnChartArea: false },
          beginAtZero: true
        }
      },
      animation: {
        duration: 800,
        easing: 'easeOutQuart'
      }
    }
  });
}

function renderDoughnutChart(summary, labels) {
  destroyChart();

  const amounts = summary.map(d => Math.round(d.total));
  const bgColors = summary.map((_, i) => CHART_COLORS[i % CHART_COLORS.length]);
  const borderColors = summary.map((_, i) => {
    const c = CHART_COLORS[i % CHART_COLORS.length];
    return c;
  });

  const ctx = document.getElementById('deptChart').getContext('2d');
  deptChartInstance = new Chart(ctx, {
    type: 'doughnut',
    data: {
      labels,
      datasets: [{
        data: amounts,
        backgroundColor: bgColors.map(c => c + 'cc'),
        borderColor: bgColors,
        borderWidth: 2,
        hoverOffset: 12
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      cutout: '55%',
      plugins: {
        legend: {
          position: 'right',
          labels: {
            color: '#94a3b8',
            font: { family: "'Inter', sans-serif", size: 12, weight: 500 },
            padding: 12,
            usePointStyle: true,
            pointStyle: 'circle'
          }
        },
        tooltip: {
          backgroundColor: '#1a2233',
          titleColor: '#f1f5f9',
          bodyColor: '#94a3b8',
          borderColor: '#2a3650',
          borderWidth: 1,
          cornerRadius: 8,
          padding: 12,
          callbacks: {
            label: function(ctx) {
              const total = ctx.dataset.data.reduce((a, b) => a + b, 0);
              const pct = total > 0 ? ((ctx.parsed / total) * 100).toFixed(1) : 0;
              return `  ₹ ${ctx.parsed.toLocaleString('en-IN')}  (${pct}%)`;
            }
          }
        }
      },
      animation: {
        animateRotate: true,
        duration: 900,
        easing: 'easeOutQuart'
      }
    }
  });
}

function destroyChart() {
  if (deptChartInstance) {
    deptChartInstance.destroy();
    deptChartInstance = null;
  }
}

// ─── Chart Toggle Buttons ───────────────────
document.getElementById('btnBarChart').addEventListener('click', function() {
  if (currentChartType === 'bar') return;
  currentChartType = 'bar';
  toggleChartButtons(this);
  if (cachedDeptSummary) renderDeptChart(cachedDeptSummary);
});

document.getElementById('btnDoughnutChart').addEventListener('click', function() {
  if (currentChartType === 'doughnut') return;
  currentChartType = 'doughnut';
  toggleChartButtons(this);
  if (cachedDeptSummary) renderDeptChart(cachedDeptSummary);
});

function toggleChartButtons(activeBtn) {
  document.querySelectorAll('.chart-toggle').forEach(b => b.classList.remove('active'));
  activeBtn.classList.add('active');
}

