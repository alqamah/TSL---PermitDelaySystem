# Developer & Agentic AI Maintenance Guide

## 1. Welcome to the Project

This guide provides technical onboarding for human developers and Agentic AI coding assistants working on the **TSL Permit Delay System**.

---

## 2. Repository Structure

```
d:\TSL\TSL---PermitDelaySystem\
├── index.html                  # Root landing page (redirects to src/index.html)
├── README.md                   # Live GitHub Pages link and project overview
├── requirements.txt            # Python dependencies (openpyxl, pandas) for data analysis
├── todo.ignore                 # Development backlog & architecture notes
├── assets\
│   ├── logo1.svg               # Corporate vector logo
│   ├── logo2.svg               # Header logo
│   ├── persistent_data.js      # Business constants (CRANE_RATES master dictionary)
│   └── formats\
│       ├── ip_format.xlsx      # Legacy input template reference
│       └── op_format.xlsx      # Legacy output matrix template reference
├── docs\
│   ├── APP_ARCHITECTURE_AND_WORKFLOW.md   # Complete system architecture & math
│   ├── INPUT_SCHEMA_AND_PARSING_SPEC.md   # Ingestion specification & column mappings
│   ├── OUTPUT_SCHEMA_AND_REPORTS.md       # Export specification & styling rules
│   └── DEVELOPER_AND_AGENTIC_AI_GUIDE.md  # This developer & AI guide
├── IO\
│   ├── rate_chart.xlsx         # Reference rate table for crane tonnages
│   ├── sample\
│   │   ├── inp_out_sample.xlsx # Historical dual-sheet input/output sample
│   │   ├── ip_sample.xlsx      # Sample multi-sheet input workbook (legacy)
│   │   └── op_sample.xlsx      # Reference expected output with Excel formulas
│   └── test\
│       ├── test ip new aug.xlsx# ACTIVE PRODUCTION SAMPLE (Matches screenshot)
│       └── test_ip.xlsx        # Legacy synthetic test workbook
└── src\
    ├── index.html              # Main single-page web application DOM
    ├── script.js               # Ingestion, parsing, calculations, rendering, export
    └── styles.css              # Swiss Minimalist CSS design system
```

---

## 3. How to Run Locally

### Option A: Direct Browser Execution
Because the application is built entirely with client-side HTML5, CSS3, and JavaScript, it can be launched directly:
1. Double click or open `src/index.html` in any modern web browser (Chrome, Edge, Firefox, Safari).
2. Drag and drop `IO/test/test ip new aug.xlsx` into the dropzone to immediately populate test data.

### Option B: Local HTTP Server (Recommended)
Using Python's built-in web server to avoid browser `file://` CORS limitations if referencing external assets:
```bash
python -m http.server 8000
# Open http://localhost:8000/src/index.html in browser
```

---

## 4. In-Memory State Model & Data Structures

The application relies on two global reactive arrays inside `src/script.js`:

```javascript
let allRecords = [];      // Master flat list of all parsed permit jobs
let allDeptSummary = [];  // Aggregated department pivot summary rows
```

### Record Object Schema (`allRecords[i]`)
```typescript
interface PermitRecord {
  craneName: string;      // Worksheet name (e.g. "55T-1", "SANY-110T")
  craneRate: number;      // Hourly cost (e.g. 1442.31) from sheet Col J or persistent_data.js
  date: Date;             // JavaScript Date object (UTC normalized)
  dateStr: string;        // Formatted display string "DD-Mon-YY" (e.g. "01-Jul-26")
  department: string;     // Plant department name (e.g. "BELT GROUP", "F-BF")
  requester: string;      // User / Requester engineer name (e.g. "ANKUR GUPTA")
  reportTime: string;     // Formatted 24h start time "HH:mm" (Step A: Permit Create Time)
  permitTime: string;     // Formatted 24h end time "HH:mm" (Step D: Permit Taken time)
  delayMin: number | null;// Delay in minutes after subtracting 30m grace (can be negative)
  delayStr: string;       // Formatted delay (e.g. "1h 15m" or "-18m")
  amount: number;         // Financial penalty: (delayMin / 60) * craneRate (0 if delayMin <= 0)
}
```

### Department Summary Object Schema (`allDeptSummary[j]`)
```typescript
interface DeptSummary {
  department: string;     // Department identifier
  contactPerson: string;  // Requester name with the highest frequency count
  cranes: {
    [craneType: string]: { // Dynamic keys: "450T", "350T", "160T", "110T", "55T", "40T"
      hours: number;      // Cumulative delay hours
      amount: number;     // Cumulative financial delay amount (₹)
    }
  };
  total: number;          // Department total liability sum across all crane types
  totalHours: number;     // Department cumulative delay hours across all crane types
}
```

---

## 5. Modifying Rates or Adding New Crane Tonnages

Crane rates are maintained in `assets/persistent_data.js`:
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
### When adding a new crane type (e.g. "250T" or "500T"):
1. Add the key and hourly rate to `CRANE_RATES` in `assets/persistent_data.js`.
2. `src/script.js` automatically computes `CRANE_TYPES` dynamically from `CRANE_RATES` sorted in descending tonnage order:
   ```javascript
   const CRANE_TYPES = Object.keys(CRANE_RATES).sort((a, b) => parseInt(b) - parseInt(a));
   ```
3. The table headers in the HTML DOM and in the exported Excel Sheet 2 are dynamically reconstructed on the fly.
4. Ensure the worksheet name in incoming Excel files includes the tonnage pattern (e.g. `XCMG-250T-1`) so that `getCraneType()` regex `(\d+T)` captures it properly.

---

## 6. Technical Debt & Development Backlog (`todo.ignore`)

The project contains a backlog file `todo.ignore` with four high-priority architecture items:

### Item 1: Unified Data Structure Object at Initialization
- **Current State**: Parsing reads cells dynamically row-by-row with hardcoded column indexes (`c: 1`, `c: 2`, `c: 4`, `c: 9`, `c: 11`, `c: 14`).
- **Target Goal**: Build an initialization routine that scans row 3 and row 4 dynamically, finds column headers by semantic label matching, and builds an indexed configuration object. This protects against column shifts when different vendors submit slight template variations.

### Item 2: Multi-Vendor Crane Merging
- **Current State**: Cranes with distinct vendor prefixes (`XCMG-110T-1112`, `SANY-110T`, `XCMG-110T-9206`) are treated as independent sheet entities during flat logging.
- **Target Goal**: In aggregated views, group cranes by vendor or combine common tonnage assets while preserving vendor-specific rate overrides.

### Item 3: Master Asset Classification
- **Current State**: The regex `(\d+T)` extracts the capacity, but special cranes or non-standard asset names may fail extraction.
- **Target Goal**: Maintain a formal master asset registry that maps any sheet name or asset code to standardized tonnage categories and registered vendor IDs.

### Item 4: Department Name Normalization
- **Current State**: Departments are captured as raw strings from the sheet. Differences like `RMBB#1 & 2` vs `RMBB-1` or `LD#1` vs `LD-1` create separate duplicate department rows in the summary table and chart.
- **Target Goal**: Implement a fuzzy or lookup-table normalization map that harmonizes department names across the plant.

---

## 7. Python Data Diagnostic Environment

A dedicated virtual environment `.venv` has been configured in the workspace with `openpyxl` and `pandas` recorded in `requirements.txt`.

### Running Diagnostic Scripts
To inspect input workbooks, verify column indices, or validate formulas:
```bash
# Windows PowerShell
.venv\Scripts\python -c "import openpyxl; wb=openpyxl.load_workbook('IO/test/test ip new aug.xlsx'); print(wb.sheetnames)"
```
Future AI agents should utilize `.venv\Scripts\python` for any heavy offline data processing, regression testing, or synthetic test generation.
