# Input Files Schema & Parsing Specification

## 1. Overview of Input Files

The Permit Delay System is designed to consume monthly crane operational permit logs submitted as Microsoft Excel (`.xlsx` / `.xls`) files.

In production operations, vendors (such as IHPL) or plant departments compile monthly logs of crane deployment. Each workbook contains multiple worksheets, where **each worksheet represents one dedicated mobile crane unit** (e.g., `55T-1`, `55T-2`, `SANY-110T`, `XCMG-110T-1112`, `XCMG-350T-1`, `SANY-450T`).

---

## 2. Active Modern Input Format Layout (Production)

The layout below represents the active input format observed in production files (e.g., `test ip new aug.xlsx`, user sample sheets):

```
┌───────┬─────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────┐
│ Row 1 │ Back to Directory (Navigation Hyperlink / Banner)                                                                           │
├───────┼──────────────────────────────────────────────────────┬──────────────────────────────────────────────────────────────────────┤
│ Row 2 │ Vendor Name:- [Vendor Code e.g. IHPL]                │ Month:- [e.g. Month-July-2026]                                       │
├───────┼──────┬──────┬────────────┬──────────┬───────────┬────┴──────────┬──────────┬──────────┬──────┬──────────┬─────────────────────┤
│ Row 3 │S. NO.│ Date │ Department │ Location │ User Name │Jobs nature... │Crane time│Shift time│Type  │(blank)   │Permit No.│PERMIT... │
├───────┼──────┴──────┴────────────┴──────────┴───────────┴───────────────┴──────────┴──────────┴──────┼──────────┼──────────┼──────────┤
│ Row 4 │ (Merged with Row 3 above for Cols A-I, K, S)                                                  │Rate/Hour │(Merged)  │Steps A..D│
├───────┼───────────────────────────────────────────────────────────────────────────────────────────────┴──────────┴──────────┴──────────┤
│ Row 5+│ DATA ROWS (Daily crane permit operational events)                                                                           │
└───────┴─────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────┘
```

### Visual Diagram of Columns (Columns A through S)

```
Col  Letter  Excel Header (Row 3 / Row 4)                       Category / Role
───  ──────  ─────────────────────────────────────────────────  ────────────────────────────
0    A       S. NO.                                             Sequence counter (1, 2, 3...)
1    B       Date                                               Execution Date (DD-MM-YYYY)
2    C       Department                                         Plant Dept Name (e.g. F-BF, PP)
3    D       Location                                           Specific Area (e.g. INBA, MILL)
4    E       User Name                                          Requester / Engineer Name
5    F       Jobs nature SD/N/M SD                              Job Type / Shutdown flag
6    G       Crane reaching time at site                        Physical arrival timestamp
7    H       Crane shift reporting time (A/G/B/C shift)         Designated Shift Code
8    I       Crane Type                                         Crane classification text
9    J       Hiring cost/ hour                                  Hourly rental rate (₹)
10   K       Permit No.                                         Unique permit SAP/EHS number
11   L       A. Permit Create Time (Created by User)            Step A: User creates permit
12   M       B. PTW requested time (User requested to Ops)      Step B: Forwarded to Operation
13   N       C. Permit Issued time (Operation approved permit)  Step C: Operations approves
14   O       D. Permit Taken time (Vendor accepted)             Step D: Vendor receives PTW
15   P       Total permit delay time (D-A)                      Raw time difference (Sheet formula)
16   Q       Total permit delay hours                           Decimal delay hours
17   R       Permit delay cost                                  Vendor-billed delay cost
18   S       Actual reason of permit delay                      Root cause / Delay notes
```

---

## 3. Comprehensive Column Data Dictionary

| Col # | Excel Col | 0-Index | Header Text in Excel | Data Type in Sheet | Parsed by App? | Variable in `script.js` | Business Meaning & Usage |
|:---:|:---:|:---:|:---|:---|:---:|:---|:---|
| **1** | **A** | `0` | `S. NO.` | Integer | No | — | Serial row number within the crane sheet. |
| **2** | **B** | `1` | `Date` | Date serial / Text (`DD-MM-YYYY`) | **YES** | `dateVal`, `dateStr` | Date of crane deployment. Required for row validation. Used for sorting master records chronologically. |
| **3** | **C** | `2` | `Department` | String | **YES** | `department` | Plant department responsible for the crane requisition (e.g., `BELT GROUP`, `SP-1`, `F-BF`, `TSCR`, `RMM`, `PP`, `NBM`, `LD-2 FME`, `BOT PLANT`, `COKE PLANT`). |
| **4** | **D** | `3` | `Location` | String | No | — | Specific shop floor area (e.g., `SP HIGH LINE`, `INBA`, `MILL`, `COAL YARD`, `NEW BPP`, `GAS BOOSTER`). |
| **5** | **E** | `4` | `User Name` | String | **YES** | `requester` | Name of the engineer/requester who ordered the crane. Used to determine the Department Contact Person via frequency analysis. |
| **6** | **F** | `5` | `Jobs nature SD/N/MSD` | String | No | — | Nature of job: Shutdown (SD), Normal (N), Minor Shutdown (MSD). |
| **7** | **G** | `6` | `Crane reaching time at site` | Time serial / String (`HH:mm:ss AM/PM`) | No | — | When the crane physically arrived on the shop floor. |
| **8** | **H** | `7` | `Crane shift reporting time (A/G/B/C shift)` | String (`A`, `B`, `C`, `G`) | No | — | Shift schedule (General Shift `G`, Morning `A`, Evening `B`, Night `C`). |
| **9** | **I** | `8` | `Crane Type` | String | No | — | Sub-classification of crane (often redundant with sheet name). |
| **10** | **J** | `9` | `Hiring cost/ hour` | Float / Currency string | **YES** | `craneRate` | Hourly rental rate (e.g., `1442.31`, `1602.56`, `2128`, `2564.1`). If omitted or 0, fallback to `persistent_data.js`. |
| **11** | **K** | `10` | `Permit No.` | Numeric / String | No | — | The SAP or EHS PTW document tracking number (e.g., `110022250683`). |
| **12** | **L** | `11` | `A. Permit Create Time (Created by User)` | Time serial / String | **YES** | `createFrac`, `createDisplay` | **Start of permit process**: When the requester initiated the permit in the portal. Baseline timestamp for delay math. |
| **13** | **M** | `12` | `B. PTW requested time` | Time serial / String | No | — | Step B: When the permit reached the operations desk. |
| **14** | **N** | `13` | `C. Permit Issued time` | Time serial / String | No | — | Step C: When operations gave safety clearance and issued permit. |
| **15** | **O** | `14` | `D. Permit Taken time (Vendor accepted)` | Time serial / String | **YES** | `takenFrac`, `takenDisplay` | **End of permit process**: When vendor work crew formally accepted the permit and could start crane operation. |
| **16** | **P** | `15` | `Total permit delay time (D-A)` | Time / String / Formula | No (recomputed) | — | Raw sheet calculation. The app calculates its own standardized delay. |
| **17** | **Q** | `16` | `Total permit delay hours` | Float / Formula | No (recomputed) | — | Raw delay hours in sheet. |
| **18** | **R** | `17` | `Permit delay cost` | Float / Formula | No (recomputed) | `amount` (recomputed) | Sheet financial calculation. Recomputed by application using standard grace rules. |
| **19** | **S** | `18` | `Actual reason of permit delay` | String | No | — | Free-text vendor/department explanation for the delay. |

---

## 4. Sheet Naming Convention & Tonnage Extraction

In the uploaded workbooks, each worksheet is named after the crane asset identifier:
- Examples: `55T-1`, `55T-2`, `SANY-110T`, `XCMG-110T-1112`, `XCMG-110T-9206`, `XCMG-110T-2-9245`, `SANY-160T`, `XCMG-350T-1`, `XCMG-350T-2`, `SANY-450T`.

The application parses the tonnage capacity by executing regular expression matching against the sheet name:
```javascript
function getCraneType(craneName) {
  if (!craneName) return null;
  const match = craneName.toUpperCase().match(/(\d+T)/);
  return match ? match[1] : null;
}
```
**Mapping Examples**:
- `55T-1` $\rightarrow$ `55T`
- `SANY-110T` $\rightarrow$ `110T`
- `XCMG-110T-1112` $\rightarrow$ `110T`
- `SANY-160T` $\rightarrow$ `160T`
- `XCMG-350T-1` $\rightarrow$ `350T`
- `SANY-450T` $\rightarrow$ `450T`

---

## 5. Ingestion & Time Parsing Logic

### 5.1. Handling Excel Time Representations
Excel stores times internally as a decimal fraction of a 24-hour day (where `1.0 = 24 hours = 86,400 seconds`).
For example:
- `0.4340277...` $\rightarrow$ `10:25:00 AM`
- `0.4527777...` $\rightarrow$ `10:52:00 AM`

However, in user-edited sheets, times are frequently stored as:
1. Pure fractional numbers (`v < 1.0`).
2. Combined datetime serial numbers (`v > 1.0`, e.g. `45474.434027`).
3. Formatted text strings (e.g., `"10:25:00AM"`, `"11:02:00 AM"`, `"3:05:00PM"`).

`src/script.js` handles all three variations seamlessly:
```javascript
function getRawTimeFraction(cell) {
  if (!cell) return null;
  const v = cell.v;
  if (typeof v === 'number') {
    const frac = v % 1;
    if (frac < 0.00001 && v > 1) return null; // Date-only cell
    return frac;
  }
  if (typeof v === 'string') {
    return parseTimeStringToFraction(v.trim());
  }
  return null;
}
```

### 5.2. Overnight Rollover Calculation
If a crane job started in the evening and the permit was approved in the early hours of the next day, `takenFrac` will be numerically smaller than `createFrac` (e.g., `03:41 AM` < `09:42 PM`).
The app detects this:
```javascript
if (takenFrac < createFrac) {
  takenFrac += 1; // Add 1 full 24-hour cycle
}
```

### 5.3. 30-Minute Grace Period Rule
Permit generation involves necessary administrative and safety steps (safety induction, gas checks, equipment tagging). Therefore, a standard **30-minute grace period** is deducted before applying delay charges:
$$\text{delayMinutes} = \operatorname{round}((\text{takenFrac} - \text{createFrac}) \times 1440) - 30$$
If $\text{delayMinutes} \le 0$, delay is treated as zero or negative (no penalty charge).

---

## 6. Legacy Input Format Comparison

Earlier revisions of the system utilized a simpler, single-header layout (`ip_format.xlsx`, `ip_sample.xlsx`):
- Row 1: `FMM INFRA MOBILE CRANE PERMIT TRACKER`
- Row 2: Metadata: `Form No.`, `CRANE :- [Name]`, `Effective date`, `Revision No.`
- Row 3: Headers:
  `[SL NO, Date, Crane Reg. No., Department, Requestor Name, Requestor Contact, Crane Reporting Time at Site, Permit No., Pemit Hand Over Time, REMARK]`
- Row 4+: Data rows.

In the legacy format:
- Start time was `Crane Reporting Time at Site` (Col G / index 6).
- End time was `Pemit Hand Over Time` (Col I / index 8).
- The modern format replaced this with the 4-step PTW progression (`Steps A, B, C, D`), significantly increasing precision by measuring between user permit creation (Step A) and vendor acceptance (Step D).
