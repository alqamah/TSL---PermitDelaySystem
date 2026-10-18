# TSL Permit Delay System — FMM Infra

A high-performance, minimalist web application for tracking, analyzing, and quantifying heavy mobile crane permit-to-work (PTW) delays across Tata Steel plant departments.

## 🚀 Live Application
- **Live Deployment**: [https://alqamah.github.io/TSL---PermitDelaySystem/src/index.html](https://alqamah.github.io/TSL---PermitDelaySystem/src/index.html)

---

## 📚 Complete Project Documentation

Comprehensive documentation has been prepared in the [`docs/`](docs/) directory:

1. **[Application Architecture & Workflow](docs/APP_ARCHITECTURE_AND_WORKFLOW.md)**:
   - System design, Swiss Minimalist UI philosophy, and tech stack.
   - End-to-end data pipeline: file upload, parsing, time math, grace period, and cost computation.
   - Function-by-function implementation reference for `src/script.js`.

2. **[Input Files Schema & Parsing Specification](docs/INPUT_SCHEMA_AND_PARSING_SPEC.md)**:
   - Visual schema diagrams and complete data dictionary for Columns A through S.
   - 4-step PTW workflow (Steps A, B, C, D) and parsing logic.
   - Excel time fraction handling, midnight shift rollover, and 30-minute grace deduction.
   - Comparison between modern production format (`test ip new aug.xlsx`) and legacy templates.

3. **[Output Files Schema & Reports Specification](docs/OUTPUT_SCHEMA_AND_REPORTS.md)**:
   - Specification for exported Excel report (`Permit_Delay_Report_YYYY-MM-DD.xlsx`).
   - Detailed schema for Sheet 1 (`delay-records`) and Sheet 2 (`Department Permit Delay Summary`).
   - Conditional styling rules (Gold total highlights, Red Net Loss footer).
   - In-app interactive visualizations (KPI cards, Table, dual-axis Chart.js analytics).

4. **[Developer & Agentic AI Maintenance Guide](docs/DEVELOPER_AND_AGENTIC_AI_GUIDE.md)**:
   - Local setup and testing workflows.
   - In-memory data structures and state management (`allRecords`, `allDeptSummary`).
   - How to add new crane tonnages and modify hourly billing rates in `assets/persistent_data.js`.
   - Technical debt analysis and implementation roadmap based on `todo.ignore`.

5. **[Agentic AI Operational Instructions](AGENTS.md)**:
   - Behavioral constraints, currency no-break rules (`₹0`), table footer standards (`TOTAL (CRANEWISE)`), and formatting contracts for AI coding assistants.

---

## 🛠️ Quick Local Setup

1. **Direct Browser Execution**:
   Simply open [`src/index.html`](src/index.html) in any modern web browser.
2. **Local HTTP Server (Optional)**:
   ```bash
   python -m http.server 8000
   # Navigate to http://localhost:8000/src/index.html
   ```
3. **Data Analysis Environment**:
   Python dependencies for parsing and offline analysis are specified in [`requirements.txt`](requirements.txt):
   ```bash
   pip install -r requirements.txt
   ```

---

## 📄 Sample Test Files
- **Active Production Sample**: [`IO/test/test ip new aug.xlsx`](IO/test/test%20ip%20new%20aug.xlsx)
- **Hourly Rate Reference**: [`IO/rate_chart.xlsx`](IO/rate_chart.xlsx)
- **Format Templates**: [`assets/formats/ip_format.xlsx`](assets/formats/ip_format.xlsx) and [`assets/formats/op_format.xlsx`](assets/formats/op_format.xlsx)
