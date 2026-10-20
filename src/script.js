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
const deptSearchInput = document.getElementById('deptSearchInput');
const downloadbtn = document.getElementById('downloadbtn');
const deptTableBody = document.getElementById('deptTableBody');
const toggleDeptHeadersBtn = document.getElementById('toggleDeptHeadersBtn');

// ─── State ───────────────────────────────────
let allRecords = [];            // master flat list of parsed rows
let allDeptSummary = [];        // master list of parsed department summary rows
let currentDeptHeaderMode = 'tonnage'; // 'tonnage' (persistent) | 'vehicle' (sheet-mapped)
let detectedSheetVehicles = []; // unique vehicle names extracted from input sheet names

// ─── File Upload Events ──────────────────────
browseBtn.addEventListener('click', (e) => {
  e.stopPropagation();
  fileInput.click();
});

// browseBtn.addEventListener('click', (e) => { ... }) handles the manual click.
// dropZone drag events handle the drop.
// The direct click on dropZone is removed as per user request.

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

deptSearchInput.addEventListener('input', () => {
  filterAndRenderDeptTable();
});

// Toggle button event listener
if (toggleDeptHeadersBtn) {
  toggleDeptHeadersBtn.addEventListener('click', (e) => {
    const optVeh = e.target.closest('#optVehicle');
    const optTon = e.target.closest('#optTonnage');
    if (optVeh) {
      setDeptHeaderMode('vehicle');
    } else if (optTon) {
      setDeptHeaderMode('tonnage');
    } else {
      toggleDeptHeaderMode();
    }
  });
}

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
  const fileResults = new Array(filesArray.length);

  const readPromises = filesArray.map((file, fileIdx) => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();

      reader.onload = (e) => {
        try {
          const data = new Uint8Array(e.target.result);
          const workbook = XLSX.read(data, { type: 'array' });
          const parsed = parseWorkbook(workbook);
          fileResults[fileIdx] = parsed;
          resolve();
        } catch (err) {
          console.error(`Error parsing ${file.name}:`, err);
          fileResults[fileIdx] = null;
          resolve();
        }
      };

      reader.onerror = (err) => {
        console.error(`FileReader error on ${file.name}:`, err);
        fileResults[fileIdx] = null;
        resolve();
      };

      reader.readAsArrayBuffer(file);
    });
  });

  // 3. Wait for all files to be processed
  await Promise.all(readPromises);

  // 4. Update the global state and UI
  const allParsedRecords = [];
  detectedSheetVehicles = [];

  fileResults.forEach(res => {
    if (!res) return;
    const records = Array.isArray(res) ? res : (res.records || []);
    const vehicles = Array.isArray(res) ? [] : (res.vehicles || []);

    if (records) allParsedRecords.push(...records);
    if (vehicles) {
      vehicles.forEach(v => {
        if (!detectedSheetVehicles.includes(v)) {
          detectedSheetVehicles.push(v);
        }
      });
    }
  });

  // Fallback: collect unique crane names from records if vehicles array is empty
  if (detectedSheetVehicles.length === 0 && allParsedRecords.length > 0) {
    allParsedRecords.forEach(r => {
      if (r.craneName && !detectedSheetVehicles.includes(r.craneName)) {
        detectedSheetVehicles.push(r.craneName);
      }
    });
  }

  allRecords = allParsedRecords;
  allRecords.sort((a, b) => a.date - b.date);

  // Clear existing search filter inputs
  searchInput.value = '';
  deptSearchInput.value = '';

  showResults();
}

// ─── Parse every sheet in the workbook ────────
function parseWorkbook(wb) {
  const craneMap = {};  // craneName → [ records ]
  const detectedVehicles = [];

  wb.SheetNames.forEach((sheetName) => {
    const ws = wb.Sheets[sheetName];
    if (!ws) return;

    // --- Crane name: Sheetname ---
    let craneName = sheetName;

    if (!craneMap[craneName]) craneMap[craneName] = [];

    // --- Walk rows starting from row 5 (index 4) ---
    const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
    for (let R = 4; R <= range.e.r; R++) {
      // Column B (1) = Date
      const cellDate = ws[XLSX.utils.encode_cell({ r: R, c: 1 })];
      // Column C (2) = Department
      const cellDept = ws[XLSX.utils.encode_cell({ r: R, c: 2 })];
      // Column E (4) = Requester Name
      const cellReq = ws[XLSX.utils.encode_cell({ r: R, c: 4 })];
      // Column J (9) = Hiring Cost
      const cellCost = ws[XLSX.utils.encode_cell({ r: R, c: 9 })];
      // Column L (11) = Permit Create Time
      const cellCreate = ws[XLSX.utils.encode_cell({ r: R, c: 11 })];
      // Column O (14) = Permit Taken Time
      const cellTaken = ws[XLSX.utils.encode_cell({ r: R, c: 14 })];

      // Skip empty rows (require at least a date)
      if (!cellDate) continue;

      const dateVal = parseDateValue(cellDate);
      if (!dateVal) continue;  // truly empty / invalid

      const department = cellDept ? String(cellDept.v || '').trim() : '';
      const requester = cellReq ? String(cellReq.v || '').trim() : '';

      // Skip if it's a completely empty record (missing department and requester)
      if (!department && !requester) continue;

      // Extract raw fractional values for time math
      const createFrac = getRawTimeFraction(cellCreate);
      const takenFrac = getRawTimeFraction(cellTaken);

      // Display strings: 24h format
      const createDisplay = getTimeDisplay(cellCreate);
      const takenDisplay = getTimeDisplay(cellTaken);

      // Delay = Permit Taken Time - Permit Create Time - 30 minutes grace
      const delayMin = computeDelay(createFrac, takenFrac);

      // Hourly Rate and Total Amount
      let craneRate = 0;
      if (cellCost) {
        if (typeof cellCost.v === 'number') craneRate = cellCost.v;
        else if (typeof cellCost.v === 'string') craneRate = parseFloat(cellCost.v.replace(/[^\d.-]/g, '')) || 0;
      }
      if (!craneRate) {
        craneRate = getCraneRate(craneName);
      }

      const amount = delayMin > 0 ? (delayMin / 60) * craneRate : 0;

      craneMap[craneName].push({
        craneRate,
        amount,
        craneName,
        date: dateVal,
        dateStr: formatDate(dateVal),
        department,
        requester,
        reportTime: createDisplay,
        permitTime: takenDisplay,
        delayMin,
        delayStr: delayMin !== null ? formatDelay(delayMin) : '—'
      });
    }

    // If this sheet has parsed records, record it as a valid vehicle
    if (craneMap[craneName] && craneMap[craneName].length > 0) {
      detectedVehicles.push(craneName);
    }
  });

  // Flatten to a single sorted array
  const flat = [];
  Object.values(craneMap).forEach(arr => flat.push(...arr));
  flat.sort((a, b) => a.date - b.date);
  return { records: flat, vehicles: detectedVehicles };
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

function getTimeDisplay(cell) {
  if (!cell) return '';

  // Always use our formatter to ensure hh:mm 24h format
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
 * Convert a fraction of day (0–1) to "HH:mm" 24h string.
 */
function fractionToTimeString(frac) {
  const totalSecs = Math.round(frac * 86400);
  let h = Math.floor(totalSecs / 3600);
  const m = Math.floor((totalSecs % 3600) / 60);
  return `${String(h).padStart(2, '0')}:${String(m).padStart(2, '0')}`;
}

// ═══════════════════════════════════════════════
//  HELPERS:  Delay computation & formatting
// ═══════════════════════════════════════════════

/**
 * Delay = (permitFrac − reportFrac) × 1440  →  minutes.
 */
function computeDelay(createFrac, takenFrac) {
  if (createFrac === null || takenFrac === null) return null;

  // If taken time is less than create time, assume it rolled over to the next day
  if (takenFrac < createFrac) {
    takenFrac += 1;
  }

  const diffMinutes = Math.round((takenFrac - createFrac) * 1440) - 30;
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
    ? '₹' + totalAmount.toLocaleString('en-IN', { minimumFractionDigits: 0, maximumFractionDigits: 0 })
    : '₹0';

  statsRow.classList.remove('hidden');
  tableSection.classList.remove('hidden');
  document.getElementById('navPanelTop').classList.remove('hidden');

  // Remove logo2.svg once results are displayed
  const rightLogo = document.getElementById('headerLogoRight');
  if (rightLogo) rightLogo.remove();

  renderTable(allRecords);

  // Department summary
  updateToggleUI();
  renderDeptTableHeaders();
  allDeptSummary = buildDeptSummary(allRecords, currentDeptHeaderMode);
  renderDeptTable(allDeptSummary);
  renderDeptChart(allDeptSummary);
  document.getElementById('deptSection').classList.remove('hidden');
}

function renderTable(records) {
  tableBody.innerHTML = '';

  records.forEach((r, i) => {
    const tr = document.createElement('tr');

    // Determine delay severity for pill styling
    let pillClass = '';
    if (r.delayMin !== null) {
      if (r.delayMin > 360) pillClass = 'red';
      else if (r.delayMin > 120) pillClass = 'orange';
      else if (r.delayMin > 0) pillClass = 'yellow';
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
        ${r.amount > 0 ? '₹' + r.amount.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) : '—'}
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

/** Crane types to show in the summary (column order) derived dynamically from CRANE_RATES */
const CRANE_TYPES = typeof CRANE_RATES !== 'undefined'
  ? Object.keys(CRANE_RATES).sort((a, b) => {
    const numA = parseInt(a, 10) || 0;
    const numB = parseInt(b, 10) || 0;
    return numB - numA;
  })
  : [];

/**
 * Returns active column list for the department summary table:
 * If in 'vehicle' mode, returns detectedSheetVehicles extracted from input file sheets.
 * Otherwise returns CRANE_TYPES from persistent data.
 */
function getActiveDeptColumns(mode = currentDeptHeaderMode) {
  if (mode === 'vehicle' && detectedSheetVehicles && detectedSheetVehicles.length > 0) {
    return detectedSheetVehicles;
  }
  return CRANE_TYPES;
}

/**
 * Dynamically render headers for the department summary table based on active columns.
 */
function renderDeptTableHeaders() {
  const deptTable = document.getElementById('deptTable');
  if (!deptTable) return;
  const thead = deptTable.querySelector('thead');
  if (!thead) return;

  const cols = getActiveDeptColumns();

  let row1 = `
    <tr>
      <th rowspan="2">SL</th>
      <th rowspan="2">DEPARTMENT</th>
      <th rowspan="2">CONTACT PERSON</th>
  `;
  cols.forEach(t => {
    row1 += `<th colspan="2" class="crane-group">${t}</th>`;
  });
  row1 += `
      <th colspan="2" class="dept-total-col">DEPT. TOTAL</th>
    </tr>
  `;

  let row2 = `<tr>`;
  cols.forEach(() => {
    row2 += `
      <th class="sub-col">Delay (hh)</th>
      <th class="sub-col">Amount (₹)</th>
    `;
  });
  row2 += `
      <th class="sub-col dept-total-sub">Delay (hh)</th>
      <th class="sub-col dept-total-sub">Amount (₹)</th>
  `;
  row2 += `</tr>`;

  thead.innerHTML = row1 + row2;
}

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
 * Supports grouping by persistent CRANE_TYPES or directly by sheet vehicle names.
 * Computes both total amount and total delay hours for each department.
 * Returns an array of dept objects sorted by DEPT. TOTAL descending.
 */
function buildDeptSummary(records, mode = currentDeptHeaderMode) {
  const isVehicleMode = mode === 'vehicle' && detectedSheetVehicles && detectedSheetVehicles.length > 0;
  const cols = getActiveDeptColumns(mode);
  const deptMap = {}; // dept → { contact, cranes: { col → { hours, amount } }, total, totalHours }

  records.forEach(r => {
    if (!r.department) return;
    const dept = r.department;

    if (!deptMap[dept]) {
      deptMap[dept] = {
        department: dept,
        contactCounts: {},  // requester → count (to find most frequent)
        cranes: {},
        total: 0,
        totalHours: 0
      };
      cols.forEach(t => {
        deptMap[dept].cranes[t] = { hours: 0, amount: 0 };
      });
    }

    const entry = deptMap[dept];

    // Track requester frequency
    if (r.requester) {
      entry.contactCounts[r.requester] = (entry.contactCounts[r.requester] || 0) + 1;
    }

    // In vehicle mode, map directly from craneName (sheet name)
    // In tonnage mode, map from getCraneType(craneName)
    const colKey = isVehicleMode ? r.craneName : getCraneType(r.craneName);
    if (colKey && entry.cranes[colKey]) {
      const delayHrs = (r.delayMin !== null && r.delayMin > 0) ? r.delayMin / 60 : 0;
      entry.cranes[colKey].hours += delayHrs;
      entry.cranes[colKey].amount += (r.amount || 0);
    }
  });

  // Compute totals and contact person
  const result = Object.values(deptMap).map(entry => {
    // Most frequent requester = contact person
    let maxCount = 0, contactPerson = '';
    for (const [name, count] of Object.entries(entry.contactCounts)) {
      if (count > maxCount) { maxCount = count; contactPerson = name; }
    }

    // Dept total amount and delay hours in active columns
    let deptTotal = 0;
    let deptTotalHours = 0;
    cols.forEach(t => {
      deptTotal += (entry.cranes[t]?.amount || 0);
      deptTotalHours += (entry.cranes[t]?.hours || 0);
    });

    return {
      department: entry.department,
      contactPerson,
      cranes: entry.cranes,
      total: deptTotal,
      totalHours: deptTotalHours
    };
  });

  // Sort by total descending
  result.sort((a, b) => b.total - a.total);
  return result;
}

/**
 * Render the department summary table with crane summary row and left-shifted Net Loss.
 */
function renderDeptTable(summary) {
  const tbody = document.getElementById('deptTableBody');
  tbody.innerHTML = '';

  const cols = getActiveDeptColumns();

  // Track crane-wise totals
  const craneTotals = {};
  cols.forEach(type => {
    craneTotals[type] = { hours: 0, amount: 0 };
  });

  let grandTotalHours = 0;
  let netLoss = 0;

  summary.forEach((row, i) => {
    const tr = document.createElement('tr');
    netLoss += row.total;
    grandTotalHours += (row.totalHours || 0);

    let cells = `
      <td>${i + 1}</td>
      <td class="dept-name-cell">${row.department}</td>
      <td>${row.contactPerson}</td>
    `;

    cols.forEach(type => {
      const c = row.cranes[type] || { hours: 0, amount: 0 };
      craneTotals[type].hours += c.hours;
      craneTotals[type].amount += c.amount;

      const hrs = c.hours > 0 ? c.hours.toFixed(2) : '0';
      const amt = c.amount > 0
        ? c.amount.toLocaleString('en-IN', { minimumFractionDigits: 0, maximumFractionDigits: 0 })
        : '0';
      cells += `<td class="hrs-cell">${hrs}</td>`;
      cells += `<td class="amt-cell">${amt}</td>`;
    });

    const deptHrs = row.totalHours > 0 ? row.totalHours.toFixed(2) : '0';
    const deptAmt = row.total > 0
      ? '₹' + row.total.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 })
      : '₹0';

    cells += `<td class="dept-total-cell hrs-cell">${deptHrs}</td>`;
    cells += `<td class="dept-total-cell amt-cell">${deptAmt}</td>`;

    tr.innerHTML = cells;
    tbody.appendChild(tr);
  });

  // Footer: Crane-wise summary row + Net Loss shifted to left
  const deptTable = document.getElementById('deptTable');
  let tfoot = document.getElementById('deptTableFoot') || (deptTable ? deptTable.querySelector('tfoot') : null);
  if (!tfoot && deptTable) {
    tfoot = document.createElement('tfoot');
    tfoot.id = 'deptTableFoot';
    deptTable.appendChild(tfoot);
  }

  if (tfoot) {
    let craneSummaryCells = '';
    cols.forEach(type => {
      const ct = craneTotals[type];
      const hrs = ct.hours > 0 ? ct.hours.toFixed(2) : '0';
      const amt = ct.amount > 0
        ? '₹' + ct.amount.toLocaleString('en-IN', { minimumFractionDigits: 0, maximumFractionDigits: 0 })
        : '0';
      craneSummaryCells += `<td class="hrs-cell crane-summary-hrs">${hrs}</td>`;
      craneSummaryCells += `<td class="amt-cell crane-summary-amt">${amt}</td>`;
    });

    const grandHrsDisplay = grandTotalHours > 0 ? grandTotalHours.toFixed(2) : '0';
    const netLossDisplay = '₹' + netLoss.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    const totalColsCount = 3 + 2 * cols.length + 2;

    tfoot.innerHTML = `
      <tr class="crane-summary-row" id="craneSummaryRow">
        <td colspan="3" class="crane-summary-label">TOTAL (CRANE/VEHICLE-WISE)</td>
        ${craneSummaryCells}
        <td class="dept-total-cell hrs-cell crane-summary-hrs">${grandHrsDisplay}</td>
        <td class="dept-total-cell amt-cell crane-summary-amt">${netLossDisplay}</td>
      </tr>
      <tr class="net-loss-row">
        <td colspan="${totalColsCount}" id="netLossCell">
          <span class="net-loss-label">NET LOSS:</span>
          <span class="net-loss-value">${netLossDisplay}</span>
        </td>
      </tr>
    `;
  }
}

// ─── Header Mapping Mode Controls ─────────────────
function setDeptHeaderMode(mode) {
  if (currentDeptHeaderMode === mode) return;
  currentDeptHeaderMode = mode;
  updateToggleUI();
  updateDeptSummaryView();
}

function toggleDeptHeaderMode() {
  const nextMode = currentDeptHeaderMode === 'tonnage' ? 'vehicle' : 'tonnage';
  setDeptHeaderMode(nextMode);
}

function updateToggleUI() {
  const toggleBtn = document.getElementById('toggleDeptHeadersBtn');
  const optVehicle = document.getElementById('optVehicle');
  if (!toggleBtn) return;

  if (optVehicle && detectedSheetVehicles.length > 0) {
    optVehicle.textContent = `Vehicle Names (${detectedSheetVehicles.length} Sheets)`;
  } else if (optVehicle) {
    optVehicle.textContent = 'Vehicle Names (Sheets)';
  }

  if (currentDeptHeaderMode === 'vehicle') {
    toggleBtn.classList.remove('mode-tonnage');
    toggleBtn.classList.add('mode-vehicle');
    toggleBtn.setAttribute('aria-pressed', 'true');
  } else {
    toggleBtn.classList.remove('mode-vehicle');
    toggleBtn.classList.add('mode-tonnage');
    toggleBtn.setAttribute('aria-pressed', 'false');
  }
}

function updateDeptSummaryView() {
  if (!allRecords || allRecords.length === 0) return;
  renderDeptTableHeaders();
  allDeptSummary = buildDeptSummary(allRecords, currentDeptHeaderMode);
  filterAndRenderDeptTable();
  renderDeptChart(allDeptSummary);
}

function filterAndRenderDeptTable() {
  const term = deptSearchInput ? deptSearchInput.value.toLowerCase().trim() : '';
  const filtered = term
    ? allDeptSummary.filter(d =>
      d.department.toLowerCase().includes(term) ||
      (d.contactPerson && d.contactPerson.toLowerCase().includes(term))
    )
    : allDeptSummary;
  renderDeptTable(filtered);
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

  // Use full department names for chart labels
  const labels = summary.map(d => d.department);

  if (currentChartType === 'bar') {
    renderBarChart(summary, labels);
  } else {
    renderDoughnutChart(summary, labels);
  }
}

function renderBarChart(summary, labels) {
  destroyChart();

  const totalHours = summary.map(d => parseFloat((d.totalHours !== undefined ? d.totalHours : 0).toFixed(2)));

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
            title: function (context) {
              const index = context[0].dataIndex;
              return summary[index].department;
            },
            label: function (ctx) {
              if (ctx.dataset.yAxisID === 'yAmount') {
                return `  Amount: ₹${ctx.parsed.y.toLocaleString('en-IN')}`;
              }
              return `  Hours: ${ctx.parsed.y}h`;
            }
          }
        }
      },
      scales: {
        x: {
          ticks: {
            autoSkip: false,
            color: '#000000',
            font: { size: 9, weight: 700 },
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
            title: function (context) {
              const index = context[0].dataIndex;
              return summary[index].department;
            },
            label: function (ctx) {
              const total = ctx.dataset.data.reduce((a, b) => a + b, 0);
              const pct = total > 0 ? ((ctx.parsed / total) * 100).toFixed(1) : 0;
              return `  ₹${ctx.parsed.toLocaleString('en-IN')}  (${pct}%)`;
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
downloadbtn.addEventListener('click', () => {
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
    'Permit Create Time': r.reportTime,
    'Permit Taken Time': r.permitTime,
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
  const activeCols = getActiveDeptColumns();
  const deptSummary = buildDeptSummary(allRecords, currentDeptHeaderMode);
  const sheet2Data = [];

  // Table Headers (dynamically built from activeCols)
  const headers = ['S.NO', 'DEPARTMENT', 'CONTACT PERSON'];
  activeCols.forEach(type => {
    headers.push(`${type} Delay (hh)`, `${type} Amount (₹)`);
  });
  headers.push('DEPT. TOTAL Delay (hh)', 'DEPT. TOTAL Amount (₹)');
  sheet2Data.push(headers);

  // Track crane-wise totals
  const craneTotals = {};
  activeCols.forEach(type => {
    craneTotals[type] = { hours: 0, amount: 0 };
  });

  let grandTotalHours = 0;
  let netLoss = 0;

  // Table Data
  deptSummary.forEach((row, i) => {
    netLoss += row.total;
    grandTotalHours += (row.totalHours || 0);

    const rowData = [
      i + 1,
      row.department,
      row.contactPerson || ''
    ];

    activeCols.forEach(type => {
      const c = row.cranes[type] || { hours: 0, amount: 0 };
      craneTotals[type].hours += c.hours;
      craneTotals[type].amount += c.amount;

      rowData.push(c.hours > 0 ? parseFloat(c.hours.toFixed(2)) : 0);
      rowData.push(c.amount > 0 ? parseFloat(c.amount.toFixed(2)) : 0);
    });

    rowData.push(row.totalHours > 0 ? parseFloat(row.totalHours.toFixed(2)) : 0);
    rowData.push(parseFloat(row.total.toFixed(2)));
    sheet2Data.push(rowData);
  });

  // Crane-wise Summary Row
  const summaryRow = ['TOTAL', '', ''];
  activeCols.forEach(type => {
    const ct = craneTotals[type];
    summaryRow.push(ct.hours > 0 ? parseFloat(ct.hours.toFixed(2)) : 0);
    summaryRow.push(ct.amount > 0 ? parseFloat(ct.amount.toFixed(2)) : 0);
  });
  summaryRow.push(grandTotalHours > 0 ? parseFloat(grandTotalHours.toFixed(2)) : 0);
  summaryRow.push(parseFloat(netLoss.toFixed(2)));
  sheet2Data.push(summaryRow);

  // Footer / Net Loss (Left-aligned)
  const totalColsCount = 3 + 2 * activeCols.length + 2;
  const footerRow = Array(totalColsCount).fill('');
  footerRow[0] = `NET LOSS: ₹${netLoss.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
  sheet2Data.push(footerRow);

  const ws2 = XLSX.utils.aoa_to_sheet(sheet2Data);

  // Style Sheet 2
  const range2 = XLSX.utils.decode_range(ws2['!ref']);
  const summaryRowIdx = range2.e.r - 1;
  const netLossRowIdx = range2.e.r;

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
      } else if (R === netLossRowIdx) {
        // Net Loss row (shifted to left)
        style.font.bold = true;
        style.font.color = { rgb: 'FFFFFF' };
        style.fill = { fgColor: { rgb: 'C00000' } }; // Red
        style.alignment.horizontal = 'left';
      } else if (R === summaryRowIdx) {
        // Crane Summary row
        style.font.bold = true;
        style.fill = { fgColor: { rgb: 'EAEAEA' } };
        if (C === 0) {
          style.alignment.horizontal = 'center';
        } else if (C >= 3) {
          if (C % 2 === 1) {
            // Hours / Delay
            style.numFmt = '0.00';
            style.alignment.horizontal = 'center';
          } else {
            // Amount
            style.numFmt = '#,##0.00';
            style.alignment.horizontal = 'right';
          }
        }
        if (C === range2.e.c) {
          style.fill = { fgColor: { rgb: 'FFD966' } };
        }
      } else {
        // Data rows
        if (C === 1 || C === 2) {
          style.alignment.horizontal = 'left';
        }

        // Match numbers formatting
        if (C >= 3) {
          if (C % 2 === 1) {
            // Delay (hh) columns: index is odd
            style.numFmt = '0.00';
            style.alignment.horizontal = 'center';
          } else {
            // Amount (₹) columns: index is even
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

        // Gold background for Total amount column
        if (C === range2.e.c) {
          style.fill = { fgColor: { rgb: 'FFD966' } };
          style.numFmt = '#,##0.00';
          style.alignment.horizontal = 'right';
        }

        // Pale yellow applied to Col 4 (first amount col)
        if (C === 4) {
          style.fill = { fgColor: { rgb: 'FFF2CC' } };
        }
      }

      cell.s = style;
    }
  }

  // Merges
  if (!ws2['!merges']) ws2['!merges'] = [];
  // Merge TOTAL across first 3 columns in summary row
  ws2['!merges'].push({
    s: { r: summaryRowIdx, c: 0 },
    e: { r: summaryRowIdx, c: 2 }
  });
  // Merge NET LOSS across all columns
  ws2['!merges'].push({
    s: { r: netLossRowIdx, c: 0 },
    e: { r: netLossRowIdx, c: totalColsCount - 1 }
  });

  // Basic styling (column widths)
  const colsWidth = [
    { wch: 6 },  // SL
    { wch: 25 }, // Dept
    { wch: 15 }  // Contact
  ];
  activeCols.forEach(() => {
    colsWidth.push({ wch: 11 }, { wch: 14 });
  });
  colsWidth.push({ wch: 14 }, { wch: 16 }); // Total Delay, Total Amount
  ws2['!cols'] = colsWidth;

  XLSX.utils.book_append_sheet(wb, ws2, 'Department Permit Delay Summary');

  // --- Write File ---
  const timestamp = new Date().toISOString().split('T')[0];
  XLSX.writeFile(wb, `Permit_Delay_Report_${timestamp}.xlsx`);
}

