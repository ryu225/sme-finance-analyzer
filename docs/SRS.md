Title: SME Finance Analyzer – Software Requirements Specification

1. Purpose
   Provide SMEs a fast way to ingest transactions, normalize data, analyze P&L,
   and export executive-ready summaries (CSV/XLSX/PPT).

2. Scope
   - Single-user browser app (Streamlit)
   - CSV/XLSX upload; optional fx_rates.csv and budget.csv in /data
   - Outputs charts, tables, anomalies, PPT

3. Actors
   - Owner/Accountant: uploads files, reviews KPIs, downloads exports

4. Functional Requirements
   FR1 Upload CSV/XLSX with required columns
   FR2 Clean/normalize columns and types
   FR3 Apply FX to base currency (per currency or per date)
   FR4 Compute VAT-inclusive/exclusive metrics
   FR5 P&L by month; Cashflow by month
   FR6 Expense breakdown and drill-down
   FR7 Anomaly detection (Z-score + MAD fallback)
   FR8 Export clean CSV, P&L CSV, XLSX, PPT
   FR9 Budget vs actual variance

5. Non-Functional Requirements
   NFR1 Performance: process ≤100k rows <5s on mid-range laptop
   NFR2 Usability: accessible colors, font ≥14px
   NFR3 Reliability: graceful error handling and schema validation
   NFR4 Portability: Docker + Streamlit Cloud

6. Data Schema (CSV columns)
   date (ISO), type {sale|expense}, category, description, amount (float),
   tax (float, optional), currency (ISO), [derived: rate_to_base, amount_base, net_amount, vat_amount, month]

7. Risks & Mitigations
   - Invalid files → JSON schema + mapping UI
   - FX gaps → default 1.0 + warning banner
   - Tiny datasets → MAD-based anomalies

8. Roadmap
   v0.3 Cashflow + Budget chart
   v0.4 Excel export (formatted)
   v0.5 Dated FX + Alerts + Docker
