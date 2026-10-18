# Output Schema & Reports Specification

## 1. Overview of Output System

The Permit Delay System provides two output modalities:
1. **Interactive Real-Time Web Dashboard**: Summary KPI cards, filterable raw records table with visual severity badges, department pivot table, and dual-mode Chart.js analytics.
2. **Downloadable Multi-Sheet Excel Audit Report**: Formatted corporate Excel report (`Permit_Delay_Report_YYYY-MM-DD.xlsx`) generated via client-side `xlsx-js-style`.

---

## 2. Generated Excel Report Architecture

When the user clicks the **Excel** button in the top navigation, the application synthesizes all in-memory parsed and computed records into a two-sheet workbook:

```
Permit_Delay_Report_YYYY-MM-DD.xlsx
├── Sheet 1: delay-records
│   └── Granular flat log of every parsed job with calculated delay & cost
└── Sheet 2: Department Permit Delay Summary
    └── Matrix pivot table aggregating delay hours and costs by department & crane tonnage
```

---

## 3. Sheet 1: `delay-records` Schema & Styling

### 3.1. Column Structure

| Col # | Excel Col | Header Name | Data Type | Sample Value | Column Width (`wch`) | Alignment |
|:---:|:---:|:---|:---|:---|:---:|:---:|
| 1 | **A** | `SL` | Integer | `1` | 5 | Center |
| 2 | **B** | `Date` | Formatted Date | `01-Jul-26` | 12 | Center |
| 3 | **C** | `Crane Name` | String | `55T-1` | 15 | Left |
| 4 | **D** | `Dept. Name` | String | `BELT GROUP` | 20 | Left |
| 5 | **E** | `Requester Name` | String | `ANKUR GUPTA` | 20 | Left |
| 6 | **F** | `Permit Create Time` | Time (24h) | `10:25` | 15 | Center |
| 7 | **G** | `Permit Taken Time` | Time (24h) | `10:52` | 15 | Center |
| 8 | **H** | `Delay` | Formatted String | `-3m` or `1h 45m` | 15 | Center |
| 9 | **I** | `Amount (₹)` | Float / Currency | `1442.31` | 15 | Right (`#,##0.00`) |

### 3.2. Cell Styling Specifications
- **Font**: Calibri, 11pt, Regular (`#000000`).
- **Header Row (Row 1)**: Bold text, background fill `#DDDDDD` (Light Gray), centered alignment.
- **Borders**: Thin black border (`rgb: '000000'`) on all 4 sides of every cell.
- **Number Format**: Col I (`Amount (₹)`) formatted with Excel number format string `#,##0.00`.

---

## 4. Sheet 2: `Department Permit Delay Summary` Schema & Styling

### 4.1. Visual Layout of Pivot Grid

```
┌──────┬──────────────┬──────────────┬────────────────────────┬─────┬────────────────────────┐
│ S.NO │ DEPARTMENT   │CONTACT PERSON│ 450T                   │ ... │ DEPT. TOTAL            │
│      │              │              ├────────────┬───────────┼─────┼────────────┬───────────┤
│      │              │              │ Delay (hh) │ Amount (₹)│ ... │ Delay (hh) │ Amount (₹)│
├──────┼──────────────┼──────────────┼────────────┼───────────┼─────┼────────────┼───────────┤
│  1   │ F-BF         │ABHISHEK SHARMA│    0.00    │     -     │ ... │   12.50    │ 42,850.00 │
│  2   │ BELT GROUP   │ANKUR GUPTA   │    0.00    │     -     │ ... │    4.00    │ 18,420.00 │
├──────┴──────────────┴──────────────┼────────────┼───────────┼─────┼────────────┼───────────┤
│ TOTAL (CRANEWISE)                  │    6.50    │ 46,377.50 │ ... │   16.50    │ 61,270.00 │
├────────────────────────────────────┴────────────┴───────────┴─────┴────────────┴───────────┤
│ NET LOSS: ₹61,270.00                                                                       │
└────────────────────────────────────────────────────────────────────────────────────────────┘
```

### 4.2. Column Hierarchy & Sequence
1. `Col A` (Index 0): `S.NO` — Integer rank sorted by descending department total cost.
2. `Col B` (Index 1): `DEPARTMENT` — Department name.
3. `Col C` (Index 2): `CONTACT PERSON` — Most frequent requester in that department.
4. **Dynamic Crane Column Pairs (Dual-Mode Toggle Button)**:
   - **Sub-headers**: `Delay (hh)` and `Amount (₹)`.
   - **Mode A: Tonnage Group (Persistent)**: Derived dynamically from `CRANE_TYPES` (`450T`, `350T`, `160T`, `110T`, `55T`, `40T`):
     - `[Tonnage] Delay (hh)` (Float, 2 decimal places, e.g., `4.50`)
     - `[Tonnage] Amount (₹)` (Float formatted as currency or `'-'` if zero)
   - **Mode B: Vehicle Names (Sheet-wise Direct Mapping)**: Directly mapped from the vehicle sheet names in the uploaded workbook (e.g., `55T-1`, `55T-2`, `SANY-110T`, `XCMG-110T-1112`, `XCMG-350T-1`, etc.):
     - `[VehicleName] Delay (hh)` (Float, 2 decimal places)
     - `[VehicleName] Amount (₹)` (Float formatted as currency or `'-'` if zero)
5. **DEPT. TOTAL Columns**:
   - `DEPT. TOTAL Delay (hh)`: Total delay hours accumulated by that department across all cranes.
   - `DEPT. TOTAL Amount (₹)`: Total financial delay liability for that department.
6. **Footer Rows**:
   - **TOTAL (CRANEWISE) Summary Row**: Spans columns 1–3 (`SL`, `DEPARTMENT`, `CONTACT PERSON`) with label `TOTAL (CRANEWISE)` followed by aggregated delay hours and costs for each crane column, and plant-wide totals under `DEPT. TOTAL`.
   - **NET LOSS Row**: Shifted to the left side, displaying the comprehensive plant-wide liability with high-contrast Swiss Red highlight.

### 4.3. Styling & Conditional Palette
- **Header Row**: Fill color `#DDDDDD`, bold text.
- **Data Rows**:
  - `DEPARTMENT` and `CONTACT PERSON`: Left-aligned.
  - `Hrs` columns: Center-aligned, 2 decimals.
  - `Amount` columns: Right-aligned currency format `#,##0.00`. If value is `0`, displayed as `'-'` and center-aligned.
  - **DEPT. TOTAL column**: Shaded with **Gold background (`#FFD966`)**, bold text, right-aligned.
  - **Col 4 (First crane amount)**: Shaded with **Pale Yellow (`#FFF2CC`)** for visual separation.
- **Footer Row (`NET LOSS:`)**:
  - Merged across all columns from `Col A` up to `DEPT. TOTAL - 1`.
  - **Background Fill**: Intense Red (`#C00000`).
  - **Font**: Bold White (`#FFFFFF`), 11pt.
  - **Value Cell**: Bold White on red background, formatted as `#,##0.00`.

---

## 5. In-App Web Visualizations & Reports

### 5.1. Real-Time KPI Cards
- **Total Records**: Count of all parsed rows across all sheets.
- **Avg. Delay**: Mean delay for delayed jobs ($\text{delayMinutes} > 0$).
- **Total Delay**: Sum of all delay minutes expressed as `Xh YYm`.
- **Total Delay Cost**: Cumulative plant-wide liability formatted in Indian Currency (`₹X,XX,XXX`).

### 5.2. Searchable Records Table with Delay Severity Badges
The raw records table in `#tableSection` applies dynamic color pill badges:
- **Green (`.delay-pill.green`)**: No delay ($\text{delayMinutes} \le 0$).
- **Yellow (`.delay-pill.yellow`)**: Minor delay ($1\text{ to }120\text{ minutes}$).
- **Orange (`.delay-pill.orange`)**: Moderate delay ($121\text{ to }360\text{ minutes}$).
- **Red (`.delay-pill.red`)**: Severe delay ($> 360\text{ minutes}$ / $> 6\text{ hours}$).

### 5.3. Chart.js Department Analytics
- **Dual-Axis Bar Graph**:
  - **Left Y-Axis (Black)**: Cumulative delay hours per department.
  - **Right Y-Axis (Red `#ff0000`)**: Cumulative delay amount (₹) per department.
  - **Tooltip**: Displays full department name, total hours, and formatted rupee total.
- **Doughnut / Pie Chart**:
  - Displays departmental percentage contribution to the total financial loss.
  - Tooltip shows amount and relative share percentage (e.g., `₹2,45,000 (34.2%)`).

---

## 6. Legacy Output Format Comparison (`op_sample.xlsx`)

In legacy system templates (`assets/formats/op_format.xlsx`, `IO/sample/op_sample.xlsx`):
- Headers were: `['S. NO. ', 'DEPARTMENT', 'CONTACT PERSON', '160T', 'Amount', '100T', 'Amount', '80T', 'Amount', '55T', 'Amount', '40T', 'Amount', '160T', 'Amount', '300T', 'Amount', 'DEPT. TOTAL']`.
- Amounts in the legacy Excel files were stored as literal Excel multiplication formulas referencing rate constants:
  - `E2` = `=D2*2564.1` (160T rate)
  - `G2` = `=F2*1602.56` (100T rate)
  - `I2` = `=H2*1362.18` (80T rate)
  - `K2` = `=J2*1442.31` (55T rate)
  - `M2` = `=L2*1350` (40T rate)
  - `O2` = `=N2*2628` (160T alternative rate)
  - `Q2` = `=P2*4247` (300T rate)
  - `R2` = `=Q2+O2+M2+K2+I2+G2+E2` (`DEPT. TOTAL`)
- The modern application computes these values directly in JavaScript to ensure instant display in the web UI, while writing formatted numeric results directly to the generated export workbook.
