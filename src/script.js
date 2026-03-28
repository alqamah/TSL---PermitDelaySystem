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
const dropZone = document.getElementById('dropZone');
const fileInput = document.getElementById('fileInput');
const browseBtn = document.getElementById('browseBtn');
const fileNameEl = document.getElementById('fileName');
const statsRow = document.getElementById('statsRow');
const tableSection = document.getElementById('tableSection');
const tableBody = document.getElementById('tableBody');
const searchInput = document.getElementById('searchInput');
const downloadBtn = document.getElementById('downloadBtn');
const deptTableBody = document.getElementById('deptTableBody');

// ─── State ───────────────────────────────────
let allRecords = [];   // master flat list of parsed rows

// ─── File Upload Events ──────────────────────
browseBtn.addEventListener('click', (e) => {
  e.stopPropagation();
  fileInput.click();
});

dropZone.addEventListener('click', () => fileInput.click());

fileInput.addEventListener('change', (e) => {
  if (e.target.files.length) handleFiles(e.target.files);
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
  if (e.dataTransfer.files.length) handleFiles(e.dataTransfer.files);
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
/**
 * handleFiles: processes multiple Excel files, aggregates all parsed records,
 * and refreshes the dashboard UI.
 */
async function handleFiles(fileList) {
  if (!fileList || fileList.length === 0) return;

  // 1. Update File Display Label
  if (fileList.length === 1) {
    fileNameEl.textContent = `✅ ${fileList[0].name}`;
  } else {
    fileNameEl.textContent = `📚 ${fileList.length} Files Selected`;
  }

  // 2. Read and parse all files in parallel
  const filesArray = Array.from(fileList);
  const allParsedRecords = [];

  const readPromises = filesArray.map(file => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();

      reader.onload = (e) => {
        try {
          const data = new Uint8Array(e.target.result);
          const workbook = XLSX.read(data, { type: 'array' });
          const records = parseWorkbook(workbook);
          allParsedRecords.push(...records);
          resolve();
        } catch (err) {
          console.error(`Error parsing ${file.name}:`, err);
          // Don't fail the whole promise if one file is bad? 
          // For now, let's just resolve to continue with the others.
          resolve();
        }
      };

      reader.onerror = (err) => {
        console.error(`FileReader error on ${file.name}:`, err);
        resolve(); // Continue with others
      };

      reader.readAsArrayBuffer(file);
    });
  });

  // 3. Wait for all files to be processed
  await Promise.all(readPromises);

  // 4. Update the global state and UI
  // Combine with existing or replace? User's "update" usually means replace current view.
  allRecords = allParsedRecords;
  allRecords.sort((a, b) => a.date - b.date);

  showResults();
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
      const cellDate = ws[XLSX.utils.encode_cell({ r: R, c: 1 })];
      // Column D (3) = Department
      const cellDept = ws[XLSX.utils.encode_cell({ r: R, c: 3 })];
      // Column E (4) = Requester Name
      const cellReq = ws[XLSX.utils.encode_cell({ r: R, c: 4 })];
      // Column G (6) = Crane Reporting Time at Site
      const cellReport = ws[XLSX.utils.encode_cell({ r: R, c: 6 })];
      // Column I (8) = Permit Handover Time
      const cellPermit = ws[XLSX.utils.encode_cell({ r: R, c: 8 })];

      // Skip empty rows (require at least a date)
      if (!cellDate) continue;

      const dateVal = parseDateValue(cellDate);
      if (!dateVal) continue;  // truly empty / invalid

      const department = cellDept ? String(cellDept.v || '').trim() : '';
      const requester = cellReq ? String(cellReq.v || '').trim() : '';

      // Extract raw fractional values for time math
      const reportFrac = getRawTimeFraction(cellReport);
      const permitFrac = getRawTimeFraction(cellPermit);

      // Display strings: use cell.w (Excel-formatted) when available
      const reportDisplay = getTimeDisplay(cellReport);
      const permitDisplay = getTimeDisplay(cellPermit);

      // Delay = (permitFrac − reportFrac) × 1440 minutes
      const delayMin = computeDelay(reportFrac, permitFrac);

      // Hourly Rate and Total Amount
      const craneRate = typeof getCraneRate === 'function' ? getCraneRate(craneName) : 0;
      const amount = delayMin > 0 ? (delayMin / 60) * craneRate : 0;

      craneMap[craneName].push({
        craneRate,
        amount,
        craneName,
        date: dateVal,
        dateStr: formatDate(dateVal),
        department,
        requester,
        reportTime: reportDisplay,
        permitTime: permitDisplay,
        delayMin,
        delayStr: delayMin !== null ? formatDelay(delayMin) : '—'
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
  const diffMinutes = Math.round((permitFrac - reportFrac) * 1440) - 30;
  return diffMinutes;
}

function formatDate(d) {
  if (!d) return '';
  const day = String(d.getUTCDate()).padStart(2, '0');
  const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
    'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  const mon = months[d.getUTCMonth()];
  const year = String(d.getUTCFullYear()).slice(-2);
  return `${day}-${mon}-${year}`;
}

function formatDelay(mins) {
  if (mins === null) return '—';
  const sign = mins < 0 ? '-' : '';
  const abs = Math.abs(mins);
  const h = Math.floor(abs / 60);
  const m = abs % 60;
  if (h > 0) return `${sign}${h}h ${String(m).padStart(2, '0')}m`;
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

  // Remove logo2.svg once results are displayed
  const rightLogo = document.getElementById('headerLogoRight');
  if (rightLogo) rightLogo.remove();

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
      if (r.delayMin > 480) pillClass = 'red';
      else if (r.delayMin > 240) pillClass = 'orange';
      else if (r.delayMin > 120) pillClass = 'yellow';
      else if (r.delayMin < -30) pillClass = 'grey';
      else pillClass = 'green';
    }

    tr.innerHTML = `
      <td>${i + 1}</td>
      <td>${r.dateStr}</td>
      <td>${r.craneName}</td>
      <td>${r.department}</td>
      <td>${r.requester}</td>
      <td>${r.reportTime}</td>
      <td>${r.permitTime}</td>
      <td title= "30min overhead reduced"><span class="delay-pill ${pillClass}">${r.delayStr}</span></td>
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
      entry.cranes[craneType].hours += delayHrs;
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

/** Swiss Color Palette for departments */
const CHART_COLORS = [
  '#000000', '#ff0000', '#666666', '#333333', '#999999',
  '#000000', '#ff0000', '#666666', '#333333', '#999999'
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
          backgroundColor: '#000000',
          borderColor: '#000000',
          borderWidth: 0,
          yAxisID: 'yHours',
          order: 2
        },
        {
          label: 'Amount (₹)',
          data: totalAmounts,
          backgroundColor: '#ff0000',
          borderColor: '#ff0000',
          borderWidth: 0,
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
            color: '#000000',
            font: { family: "'Inter', sans-serif", size: 10, weight: 800 },
            padding: 20,
            usePointStyle: true,
            pointStyle: 'rect'
          }
        },
        tooltip: {
          backgroundColor: '#000000',
          titleColor: '#ffffff',
          bodyColor: '#ffffff',
          borderColor: '#000000',
          borderWidth: 0,
          cornerRadius: 0,
          padding: 12,
          titleFont: { family: "'Inter', sans-serif", weight: 900, size: 14 },
          bodyFont: { family: "'Inter', sans-serif", weight: 700 },
          callbacks: {
            label: function (ctx) {
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
            color: '#000000',
            font: { size: 10, weight: 700 },
            maxRotation: 0
          },
          grid: { display: false }
        },
        yHours: {
          type: 'linear',
          position: 'left',
          title: {
            display: true,
            text: 'HOURS',
            color: '#000000',
            font: { size: 10, weight: 900 }
          },
          ticks: { color: '#000000', font: { size: 10, weight: 700 } },
          grid: { color: '#e0e0e0' },
          beginAtZero: true
        },
        yAmount: {
          type: 'linear',
          position: 'right',
          title: {
            display: true,
            text: 'AMOUNT',
            color: '#ff0000',
            font: { size: 10, weight: 900 }
          },
          ticks: {
            color: '#ff0000',
            font: { size: 10, weight: 700 },
            callback: function (val) {
              return '₹' + val.toLocaleString('en-IN');
            }
          },
          grid: { display: false },
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
        backgroundColor: bgColors,
        borderColor: '#ffffff',
        borderWidth: 2,
        hoverOffset: 0
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
            color: '#000000',
            font: { family: "'Inter', sans-serif", size: 10, weight: 700 },
            padding: 12,
            usePointStyle: true,
            pointStyle: 'rect'
          }
        },
        tooltip: {
          backgroundColor: '#000000',
          titleColor: '#ffffff',
          bodyColor: '#ffffff',
          borderColor: '#000000',
          borderWidth: 0,
          cornerRadius: 0,
          padding: 12,
          callbacks: {
            label: function (ctx) {
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
document.getElementById('btnBarChart').addEventListener('click', function () {
  if (currentChartType === 'bar') return;
  currentChartType = 'bar';
  toggleChartButtons(this);
  if (cachedDeptSummary) renderDeptChart(cachedDeptSummary);
});

document.getElementById('btnDoughnutChart').addEventListener('click', function () {
  if (currentChartType === 'doughnut') return;
  currentChartType = 'doughnut';
  toggleChartButtons(this);
  if (cachedDeptSummary) renderDeptChart(cachedDeptSummary);
});

function toggleChartButtons(activeBtn) {
  document.querySelectorAll('.chart-toggle').forEach(b => b.classList.remove('active'));
  activeBtn.classList.add('active');
}

// ─── Export to Excel ────────────────────────
downloadBtn.addEventListener('click', () => {
  if (allRecords.length === 0) {
    alert('No data to export!');
    return;
  }
  downloadExcel();
});

function downloadExcel() {
  const wb = XLSX.utils.book_new();

  // --- Sheet 1: DelayRecords -> delay-records ---
  const sheet1Data = allRecords.map((r, i) => ({
    'SL': i + 1,
    'Date': r.dateStr,
    'Crane Name': r.craneName,
    'Dept. Name': r.department,
    'Requester Name': r.requester,
    'Reporting Time': r.reportTime,
    'Permit Time': r.permitTime,
    'Delay': r.delayStr,
    'Amount (₹)': r.amount
  }));
  const ws1 = XLSX.utils.json_to_sheet(sheet1Data);

  // Basic styling for Sheet 1
  const range1 = XLSX.utils.decode_range(ws1['!ref']);
  for (let R = range1.s.r; R <= range1.e.r; ++R) {
    for (let C = range1.s.c; C <= range1.e.c; ++C) {
      const cell_ref = XLSX.utils.encode_cell({ c: C, r: R });
      let cell = ws1[cell_ref];
      if (!cell) continue;

      let style = {
        font: { name: 'Calibri', sz: 11, color: { rgb: '000000' } },
        border: {
          top: { style: 'thin', color: { rgb: '000000' } },
          bottom: { style: 'thin', color: { rgb: '000000' } },
          left: { style: 'thin', color: { rgb: '000000' } },
          right: { style: 'thin', color: { rgb: '000000' } }
        },
        alignment: { vertical: 'center', horizontal: 'left' }
      };

      if (R === 0) {
        style.font.bold = true;
        style.fill = { fgColor: { rgb: 'DDDDDD' } };
        style.alignment.horizontal = 'center';
      }
      if (C === 8 && R > 0) {
        style.numFmt = '#,##0.00';
      }

      cell.s = style;
    }
  }

  // Column widths for sheet 1
  ws1['!cols'] = [
    { wch: 5 }, { wch: 12 }, { wch: 15 }, { wch: 20 },
    { wch: 20 }, { wch: 15 }, { wch: 15 }, { wch: 15 }, { wch: 15 }
  ];

  XLSX.utils.book_append_sheet(wb, ws1, 'delay-records');

  // --- Sheet 2: Department Permit Delay Summary ---
  const deptSummary = buildDeptSummary(allRecords);
  const sheet2Data = [];

  // Table Headers (match image styling if possible)
  const headers = [
    'S.NO', 'DEPARTMENT', 'CONTACT PERSON',
    '160T Hrs', '160T Amount',
    '100T Hrs', '100T Amount',
    '80T Hrs', '80T Amount',
    '55T Hrs', '55T Amount',
    '40T Hrs', '40T Amount',
    '300T Hrs', '300T Amount',
    'DEPT. TOTAL'
  ];
  sheet2Data.push(headers);

  // Table Data
  let netLoss = 0;
  deptSummary.forEach((row, i) => {
    netLoss += row.total;
    const rowData = [
      i + 1,
      row.department,
      row.contactPerson || ''
    ];

    CRANE_TYPES.forEach(type => {
      const c = row.cranes[type];
      rowData.push(c.hours > 0 ? parseFloat(c.hours.toFixed(2)) : 0);
      rowData.push(c.amount > 0 ? parseFloat(c.amount.toFixed(2)) : 0);
    });

    rowData.push(parseFloat(row.total.toFixed(2)));
    sheet2Data.push(rowData);
  });

  // Footer / Net Loss
  const footerRow = Array(16).fill('');
  footerRow[0] = 'NET LOSS:';
  footerRow[15] = parseFloat(netLoss.toFixed(2));
  sheet2Data.push(footerRow);

  const ws2 = XLSX.utils.aoa_to_sheet(sheet2Data);

  // Style Sheet 2
  const range2 = XLSX.utils.decode_range(ws2['!ref']);
  for (let R = range2.s.r; R <= range2.e.r; ++R) {
    for (let C = range2.s.c; C <= range2.e.c; ++C) {
      const cell_ref = XLSX.utils.encode_cell({ c: C, r: R });
      let cell = ws2[cell_ref];

      if (!cell) {
        cell = { t: 's', v: '' }; // blank cell insert
        ws2[cell_ref] = cell;
      }

      let style = {
        font: { name: 'Calibri', sz: 11, color: { rgb: '000000' } },
        border: {
          top: { style: 'thin', color: { rgb: '000000' } },
          bottom: { style: 'thin', color: { rgb: '000000' } },
          left: { style: 'thin', color: { rgb: '000000' } },
          right: { style: 'thin', color: { rgb: '000000' } }
        },
        alignment: { vertical: 'center', horizontal: 'center' }
      };

      if (R === 0) {
        // Header row
        style.font.bold = true;
        style.fill = { fgColor: { rgb: 'DDDDDD' } };
      } else if (R === range2.e.r) {
        // Net Loss row
        style.font.bold = true;
        style.font.color = { rgb: 'FFFFFF' };
        style.fill = { fgColor: { rgb: 'C00000' } }; // Red

        if (C === 15) {
          style.numFmt = '#,##0.00';
          style.alignment.horizontal = 'right';
        } else if (C === 0) {
          style.alignment.horizontal = 'right';
        }
      } else {
        // Data rows
        if (C === 1 || C === 2) {
          style.alignment.horizontal = 'left';
        }

        // Match numbers formatting
        if (C >= 3 && C <= 14) {
          if (C % 2 === 0) {
            // Amount columns: 4, 6, 8, 10, 12, 14
            style.numFmt = '#,##0.00';
            if (cell.v === 0) {
              cell.t = 's';
              cell.v = '-';
              style.alignment.horizontal = 'center';
            } else {
              style.alignment.horizontal = 'right';
            }
          }
        }

        // FFD966 is the gold background from image, applied to Total col.
        if (C === 15) {
          style.fill = { fgColor: { rgb: 'FFD966' } };
          style.numFmt = '#,##0.00';
          style.alignment.horizontal = 'right';
        }

        // FFF2CC is the pale yellow from the image, we apply it to Col 4 (first amount col) to recreate screenshot feel.
        if (C === 4) {
          style.fill = { fgColor: { rgb: 'FFF2CC' } };
        }
      }

      cell.s = style;
    }
  }

  // Merge the "NET LOSS:" cells
  if (!ws2['!merges']) ws2['!merges'] = [];
  ws2['!merges'].push({
    s: { r: range2.e.r, c: 0 },
    e: { r: range2.e.r, c: 14 }
  });

  // Basic styling (column widths)
  ws2['!cols'] = [
    { wch: 6 },  // SL
    { wch: 25 }, // Dept
    { wch: 15 }, // Contact
    { wch: 8 }, { wch: 12 }, // 160T
    { wch: 8 }, { wch: 12 }, // 100T
    { wch: 8 }, { wch: 12 }, // 80T
    { wch: 8 }, { wch: 12 }, // 55T
    { wch: 8 }, { wch: 12 }, // 40T
    { wch: 8 }, { wch: 12 }, // 300T
    { wch: 15 }  // Total
  ];

  XLSX.utils.book_append_sheet(wb, ws2, 'Department Permit Delay Summary');

  // --- Write File ---
  const timestamp = new Date().toISOString().split('T')[0];
  XLSX.writeFile(wb, `Permit_Delay_Report_${timestamp}.xlsx`);
}

