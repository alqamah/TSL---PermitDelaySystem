# AGENTS.md — Agentic AI Coding Guide

This file provides system instructions, architectural constraints, and operational guidelines for AI coding assistants working within the **TSL Permit Delay System** repository.

---

## 1. Project Overview & Tech Stack

- **Purpose**: Production web application for parsing, calculating, tracking, and analyzing Heavy Mobile Crane Permit-To-Work (PTW) delay costs across Tata Steel plant departments.
- **Technology Stack**:
  - **HTML5 & Vanilla JavaScript**: Core reactive application logic (`src/index.html`, `src/script.js`).
  - **Vanilla CSS**: Swiss Minimalist Design System (`src/styles.css`).
  - **Libraries**:
    - SheetJS (`xlsx.full.min.js` v0.18.5) for multi-sheet client-side Excel ingestion and 2-sheet styled workbook generation.
    - Chart.js (v4.4.1) for departmental delay analytics (Dual-Axis Bar Graph & Pie Chart).

---

## 2. Critical UI & Calculation Standards

### A. Number & Currency Formatting
1. **No Spaces After Currency Symbol**:
   - Always format currency as `₹0` or `'₹' + amount.toLocaleString('en-IN')` (e.g., `₹61,270.00`).
   - **Never** place a space after `₹` inside table cells (e.g., do NOT use `'₹ ' + ...`).
   - All amount and delay cells must enforce `white-space: nowrap;` to prevent line breaks between the symbol and the numeric value.
2. **Tabular Numerics**:
   - Numeric cells use `font-variant-numeric: tabular-nums` to guarantee aligned decimal places.

### B. Department Summary Table Structure (`#deptSection`)
1. **Dual-Mode Column Headers**:
   - **Mode A (Tonnage Group)**: Mapped dynamically from `CRANE_TYPES` (`450T`, `350T`, `160T`, `110T`, `55T`, `40T`) sorted descending.
   - **Mode B (Vehicle Names)**: Directly mapped from the vehicle sheet names in the input workbook (e.g. `55T-1`, `SANY-110T`).
   - Header toggle button (`#deptToggleBtn`) allows seamless switching.
2. **Sub-Headers**:
   - Every crane column and the department total column split into two sub-headers:
     - `Delay (hh)`
     - `Amount (₹)`
3. **Dual Department Totals**:
   - `DEPT. TOTAL` spans 2 columns (`colspan="2"`) displaying both cumulative delay hours (`totalHours`) and cumulative financial penalty (`total`).
4. **Table Footer (`tfoot`) Hierarchy**:
   - **`TOTAL (CRANEWISE)` Row** (`.crane-summary-row`):
     - First 3 columns (`SL`, `DEPARTMENT`, `CONTACT PERSON`) are merged with label **`TOTAL (CRANEWISE)`**.
     - Shows cumulative delay hours and total amount per crane column.
     - Displays grand total hours and grand total amount under the Dept Total sub-columns.
     - Dynamically recalculates when the user searches/filters departments.
   - **`NET LOSS` Row** (`.net-loss-row`):
     - Left-aligned banner (`#netLossCell` with `text-align: left`).
     - Labeled as `NET LOSS: ₹X,XX,XXX.XX` with Swiss Red highlight (`#ff0000`).

---

## 3. Excel Export Specification (`downloadExcel`)

When modifying `downloadExcel()` in `src/script.js`:
- Maintain 2 synchronized sheets:
  1. `delay-records`: Complete itemized audit records.
  2. `Department Permit Delay Summary`: High-level pivot matrix matching the in-app view.
- Sheet 2 must include:
  - Header Row 1 (`colspan=2` for each crane and DEPT. TOTAL).
  - Header Row 2: `Delay (hh)` and `Amount (₹)`.
  - Merged `TOTAL (CRANEWISE)` summary row across the first 3 columns (`s: { r: summaryRowIdx, c: 0 }, e: { r: summaryRowIdx, c: 2 }`).
  - Merged left-aligned `NET LOSS` row across all columns with intense red fill (`#C00000`) and bold white text (`#FFFFFF`).

---

## 4. Documentation Index

Detailed specifications are maintained under [`docs/`](docs/):
- **[APP_ARCHITECTURE_AND_WORKFLOW.md](docs/APP_ARCHITECTURE_AND_WORKFLOW.md)**: End-to-end data pipeline, 30m grace math, shift rollover, and component architecture.
- **[INPUT_SCHEMA_AND_PARSING_SPEC.md](docs/INPUT_SCHEMA_AND_PARSING_SPEC.md)**: Ingestion schema for Columns A–S, 4-step PTW workflow, and sample comparisons.
- **[OUTPUT_SCHEMA_AND_REPORTS.md](docs/OUTPUT_SCHEMA_AND_REPORTS.md)**: Output matrix visual layouts, styling tokens, and SheetJS formatting rules.
- **[DEVELOPER_AND_AGENTIC_AI_GUIDE.md](docs/DEVELOPER_AND_AGENTIC_AI_GUIDE.md)**: Local developer setup, in-memory data structures, crane rate updates, and backlog notes.
