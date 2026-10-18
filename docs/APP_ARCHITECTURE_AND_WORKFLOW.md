# Permit Delay System — Application Architecture & Working Details

## 1. System Overview & Executive Summary

The **Permit Delay System (FMM Infra)** is a specialized single-page web application (SPA) created for Tata Steel (TSL) to monitor, quantify, and analyze operational delays in heavy mobile crane permits across diverse industrial plant departments (e.g., Coke Plant, Blast Furnace, LD Shops, Rolling Mills, Sinter Plants).

In industrial construction and maintenance operations, mobile cranes (such as 40T, 55T, 110T, 160T, 350T, 450T) are rented or hired on high hourly billing rates (ranging from ₹1,350/hr up to ₹7,135/hr). When a crane is deployed on-site but the permit-to-work (PTW) is delayed between initial permit creation and final permit acceptance/handover, substantial unproductive rental charges accumulate.

The application automates the ingestion of multi-sheet vendor Excel logbooks, calculates true delay durations (factoring in operational grace periods and shift rollovers), computes department-wise financial liabilities, provides interactive search/filtering and visualizations (Chart.js), and generates formatted executive Excel audit reports.

---

## 2. Technology Stack & Key Dependencies

```
┌─────────────────────────────────────────────────────────────┐
│                    Browser Client (SPA)                     │
├──────────────────────────────┬──────────────────────────────┤
│ Frontend Structure           │ Semantic HTML5               │
│ UI & Design System           │ Vanilla CSS (Swiss Minimal)  │
│ Application & Math Engine    │ Vanilla JavaScript (ES6+)    │
│ Excel Parsing & Styling Engine│ xlsx-js-style v1.2.0 (CDN)   │
│ Charting & Visualization     │ Chart.js v4.4.7 (CDN)        │
│ Static Business Configurations│ persistent_data.js          │
│ Python Diagnostic Environment│ Python 3.12, openpyxl, pandas│
└──────────────────────────────┴──────────────────────────────┘
```

- **Zero Heavy Frameworks**: No React, Next.js, or Vue runtime overhead. Runs as pure client-side static code, capable of being hosted on GitHub Pages or intranet static file servers with zero backend infrastructure.
- **xlsx-js-style (`1.2.0`)**: Enhanced fork of SheetJS supporting cell styling (fills, font colors, borders, custom number formats) required for generating formatted corporate reports.
- **Chart.js (`4.4.7`)**: Interactive responsive dual-axis Bar chart and Doughnut/Pie chart for department delay hours and financial loss comparisons.
- **Design Language**: Swiss Minimalist Design Aesthetic — High-contrast typography (`Inter`), clean monochrome lines, stark borders (`2px solid #000`), intentional red accenting (`#ff0000`), and zero distracting UI fluff.

---

## 3. End-to-End Data Pipeline & Execution Flow

```mermaid
flowchart TD
    A[Excel File(s) Uploaded: Drag & Drop / File Input] --> B[handleFiles: FileReader ArrayBuffer]
    B --> C[parseWorkbook: XLSX.read with cellDates: false]
    C --> D[Iterate Every Sheet: Each Sheet = Crane Asset]
    D --> E[Row Loop: Start at Row 5 / Index 4 to End]
    E --> F[Extract Core Fields: Date, Dept, Requester, Cost, CreateTime, TakenTime]
    F --> G{Validate Row}
    G -- Invalid / Empty --> H[Skip Row]
    G -- Valid --> I[Time Math: getRawTimeFraction & computeDelay]
    I --> J[Subtract 30 min Grace Period]
    J --> K[Compute Amount: Delay Hours x Hourly Crane Rate]
    K --> L[Store Flat Object into allRecords Array]
    L --> M[Sort Chronologically by Date]
    M --> N[Calculate Global KPIs: Total Records, Avg Delay, Total Delay, Total Cost]
    N --> O[Render Records Table with Severity Color Badges]
    N --> P[buildDeptSummary: Aggregate by Dept & Crane Tonnage]
    P --> Q[Determine Contact Person: Most Frequent Requester]
    Q --> R[Render Department Summary Pivot Table & NET LOSS Footer]
    R --> S[Render Chart.js: Dual Axis Bar / Doughnut Chart]
    S --> T[User Action: Search / Filter or Export Excel]
```

---

## 4. Detailed Component & File Breakdown

### 4.1. `index.html` — Layout & DOM Structure
- **Header (`.app-header`)**: Contains the corporate title, sub-title, dynamic logo (`logo2.svg`), and quick navigational buttons (`Records`, `Summary`, `Graph`, `Excel Download`, `Top`).
- **Upload Zone (`.upload-zone`)**: Unified drag-and-drop dropzone with file input triggering and link to format template.
- **Stats Row (`.stats-row`)**: Four executive KPI metric cards:
  1. `Total Records`: Total parsed valid crane job entries.
  2. `Avg. Delay`: Average delay duration (hours and minutes) excluding zero/negative delays.
  3. `Total Delay`: Grand cumulative operational delay time.
- **Data Table (`#tableSection`)**: Granular, searchable table listing every permit record with real-time text filter and severity pills.
- **Department Summary Table (`#deptSection`)**: High-level cross-tabulation matrix breaking down delay hours and costs per crane tonnage category or vehicle name for every department. Features:
  - Dual-mode header toggle (`Tonnage Group` vs `Vehicle Names`).
  - Dual sub-headers (`Delay (hh)` and `Amount (₹)` with nowrap text protection).
  - Dual `DEPT. TOTAL` columns showing both hours and currency.
  - Dedicated `TOTAL (CRANEWISE)` summary row in table footer.
  - Left-aligned high-contrast Swiss Red `NET LOSS` banner.
- **Chart Section (`#chartSection`)**: Dynamic Chart.js canvas with toggle controls between **Bar Graph** (dual-axis: hours vs amount) and **Pie Chart** (percentage contribution per department).
- **Footer (`.app-footer`)**: Displays owner contact (`alqama.hasnain@tatasteel.com`) and organization unit (`FMM-INFRA`).

### 4.2. `assets/persistent_data.js` — Business Constants
Contains the master reference mapping for crane hourly rental rates:
```javascript
const CRANE_RATES = {
  "160T": 2564.1,
  "110T": 2128,
  "55T":  1442.31,
  "40T":  1350,
  "450T": 7135,
  "350T": 6115,
};
```
*Note*: If an input sheet explicitly specifies a rate in Column J (`Hiring cost/ hour`), the application prioritizes the sheet-specific rate; otherwise, it falls back to this dictionary.

### 4.3. `src/script.js` — Core Processing Logic
The script is organized into distinct, modular functional layers:

#### A. Ingestion & Event Handling
- `browseBtn` & `fileInput`: Enables file selection via file picker.
- `dropZone`: Dragover, dragleave, and drop event listeners with visual `.drag-over` state feedback.
- `handleFiles(fileList)`:
  - Supports parallel multi-file ingestion using `Promise.all` and `FileReader.readAsArrayBuffer()`.
  - Merges records across multiple uploaded workbooks into a single master state array `allRecords`.
  - Sorts all records chronologically by date.

#### B. Sheet & Row Parsing (`parseWorkbook`)
- Sets `cellDates: false` in `XLSX.read(data, { type: 'array' })`.
  - *Why this is critical*: When `cellDates` is enabled, SheetJS attempts to instantiate JavaScript `Date` objects, which apply the client browser's local timezone offset to time fractions, corrupting time deltas. Keeping raw numbers allows exact arithmetic on Excel serial day fractions.
- Iterates over `wb.SheetNames`. The sheet name itself serves as the **Crane Identifier** (e.g. `55T-1`, `SANY-110T`, `XCMG-350T-1`).
- Starts row iteration at `R = 4` (0-indexed, corresponding to row 5 in Excel).
- Extracts:
  - Column B (index 1): `Date`
  - Column C (index 2): `Department`
  - Column E (index 4): `Requester Name`
  - Column J (index 9): `Hiring cost/ hour`
  - Column L (index 11): `Permit Create Time` (Step A)
  - Column O (index 14): `Permit Taken time` (Step D)
- Filters out blank rows: Must have a valid date and at least a department or requester name.

#### C. Time Mathematics & Delay Engine
- `getRawTimeFraction(cell)`:
  - If cell value is numeric: Extracts fractional day part (`v % 1`).
  - If cell value is string (e.g., `"10:25:00AM"` or `"1:27:23 PM"`): Parses using regex `/^(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(AM|PM)?$/i` and converts to day fraction `(h*3600 + m*60 + s) / 86400`.
- `computeDelay(createFrac, takenFrac)`:
  - **Overnight Rollover**: If `takenFrac < createFrac`, the shift rolled past midnight; adds 1 full day (`takenFrac += 1`).
  - Converts day fraction difference to minutes: `round((takenFrac - createFrac) * 1440)`.
  - **Grace Period**: Subtracts **30 minutes** of standard operational buffer:
    $$\text{delayMinutes} = \operatorname{round}((\text{takenFrac} - \text{createFrac}) \times 1440) - 30$$
- `amount` calculation:
  $$\text{amount} = \begin{cases} \left(\frac{\text{delayMinutes}}{60}\right) \times \text{craneRate} & \text{if } \text{delayMinutes} > 0 \\ 0 & \text{otherwise} \end{cases}$$

#### D. Department Aggregation Engine (`buildDeptSummary`)
- Aggregates records by `department`.
- **Dual-Mode Header Mapping (Toggle Button)**:
  - **Tonnage Group (Default)**: Detects the crane tonnage category dynamically using regex `getCraneType(craneName)` (`match(/(\d+T)/)`), grouping columns by `CRANE_TYPES` (`450T`, `350T`, `160T`, `110T`, `55T`, `40T`) from `persistent_data.js`.
  - **Vehicle Names (Sheet-Mapped)**: Directly maps headers to the vehicle worksheet names from the input workbook (e.g., `55T-1`, `55T-2`, `SANY-110T`, `XCMG-110T-1112`, etc.), displaying exact per-vehicle delay hours and costs.
  - A toggle button in the `#deptSection` toolbar switches between both modes on demand and syncs with the Chart and Excel Sheet 2 export.
- **Contact Person Determination**: Tracks frequency counts of all requesters in that department (`contactCounts`). The requester with the highest count is designated as the primary department contact person.
- Computes total cost per department and sorts departments by descending financial loss.
- Computes overall plant-wide `NET LOSS`.

#### E. Excel Report Generation (`downloadExcel`)
Generates an audit-ready multi-sheet workbook `Permit_Delay_Report_YYYY-MM-DD.xlsx` containing:
1. `delay-records`: Complete itemized audit table.
2. `Department Permit Delay Summary`: High-level executive matrix with red `#C00000` Net Loss footer and gold `#FFD966` total highlights.
