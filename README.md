# CityMed Hospital Management System

An automated, relational Patient and Billing Management System built natively in Microsoft Excel 2021 for mid-size facilities (10–50 doctors, 100 beds). This system manages the complete patient lifecycle, operational logs, insurance claims, and financial workflows entirely through formulas—eliminating macro overhead and enterprise ERP dependencies.

## Data Architecture & Core Relationships

The workbook implements a relational database structure using native Excel Tables (`Ctrl+T`) and dynamic `XLOOKUP` chains to pass data seamlessly across 12 sheets divided into 4 functional layers.
### Worksheet Architecture & Schema

* **LOOKUP (Hidden):** Houses 17 named ranges driving strict global data validation dropdowns.
* **INSTRUCTIONS:** Onboarding workspace containing user navigation grids and operational rules.
* **PT_Master (`PT-00000`):** Central patient registry. Records demographics, tracks insurance providers, and automates real-time age calculation via `DATEDIF`.
* **DR_Master (`DR-000`):** Doctor directory tracking availability across 8 medical departments with a 3-tier fee matrix (OPD Consultation, IPD Daily Visit, Emergency).
* **SVC_RateCard (`SVC-0000`):** Centralized price master for 50 services mapping out specific GST tiers (0%, 5%, 12%, 18%) and insurance eligibility flags.
* **BED_Master (`BED-000`):** Live bed inventory across 5 ward variants. Dynamically tracks occupancy, maintenance schedules, and flags ICU availability.
* **ADM_Log (`ADM-00000`):** Encounters log for OPD, IPD, and Emergency admissions. Automates patient/doctor details using nested lookups and tracks ICD-10 codes.
* **SVC_Log (`SL-00000`):** Transactional billing ledger. Logs all services utilized during a visit. Automatically scales room charges based on calculated Length of Stay (LOS).
* **BILLING (`BILL-00000`):** Main invoice engine. Aggregates service totals via `SUMIFS`, handles manual discounts with approval states, and manages the explicit Insurance-vs-Patient split.
* **PAYMENTS (`PAY-00000`):** Accounts Receivable ledger. Keeps a running balance across partial payments and automates a 4-bucket AR aging matrix (0-30, 31-60, 61-90, 90+ days) relative to `TODAY()`.
* **INSURANCE (`INS-00000`):** Claims management engine. Tracks authorization lifecycles, maps TPA references, and surfaces financial variance anomalies.
* **DASHBOARD / ANALYTICS:** Consolidated reporting command centers processing 30+ operational KPIs via `SUMPRODUCT` and `COUNTIFS`. Houses 4 interactive Pivot Charts linked to dynamic global Slicers.

---

## Technical Specifications & Integrity Controls

### Professional Print Optimization
* **Pre-Set Print Areas:** The Executive Dashboard and individual Invoice views are configured with defined boundaries. Core data models are explicitly designed to keep key metrics intact and isolated without text clipping when exported directly to PDF or hardcopy layouts.
* **Visual Formatting:** High-contrast row treatments and formula-driven Conditional Formatting scale clean boundaries across clean, report-ready dimensions.

### Data Governance & Error Prevention
* **Hard Dropdown Enforcement:** All categorical entries are bound to restricted inputs powered by named ranges to block manual transcription or typographical error entry.
* **Formula Isolation Security:** Every calculation row, lookup logic engine, and core data matrix cell is locked under Workbook Sheet Protection to prevent structural corruption by non-technical system operators.
* **Blank Cell Handling:** Nested formulas are cleanly insulated with `IFERROR` strings to suppress native broken formatting expressions such as `#N/A` or `#DIV/0!`.

---

## Installation & Environment

### System Requirements
* Microsoft Excel 2021 or Microsoft 365 (Desktop installation required for full feature compatibility).
* No macro or VBA execution permissions required (100% formulas-based platform).

### Setup & Usage Steps
1. Download `excelProject.xlsx` from this repository.
2. Launch the workbook via your local Microsoft Excel application.
3. Click **Enable Editing** on the standard Microsoft security ribbon if prompted.
4. Open the `INSTRUCTIONS` sheet first to review navigation patterns and structural workflows before populating data.

---
**Developer:** Jyotirmoy Mukherjee
