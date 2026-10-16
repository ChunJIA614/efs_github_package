# Malaysia Court E-Filing (EFS) Scraper — Final Edition (V9)

Automated extraction tool for **Sistem e-Filing Mahkamah Malaysia (EFS)** (`MyFiling.aspx`) with a built-in Desktop GUI, Date Range Scanner, and Landscape A4 Excel exporter.

## Key Features
- **Production Desktop GUI & CLI Fallback**: Launch Microsoft Edge, monitor live KPI cards, preview captured rows in real time, or stop & save at any time.
- **Date Range Scanning (`From Date` -> `To Date`)**: Scan filings within a specific `DD/MM/YYYY` range with automatic forward seeking, zero-click skipping of out-of-range rows, and automatic early stopping.
- **Reverse Chronological Ordering**: Traverses pages backward (`Page N -> Page 1`) and processes rows on each page from the lowest row (oldest) to the highest row (newest).
- **Master-Detail Multi-Case & Deposit Support**:
  - Links each `rgDetailTable` back to its parent case row in `radGridCaseDetails` to capture every individual party name (preserving leading `N` in names such as `NORFITRI`, `NOOR`, `NURUL`, `NIK`).
  - Captures both standard filing fees (`Fi Pemfailan (RM)`) and deposit amounts (`Amaun Deposit`).
- **Merged & Filterable Excel Output (`.xlsx` Only)**:
  - Exports directly to styled `.xlsx` workbooks in `efs_output/` configured for **A4 Landscape printing**.
  - Vertically merges consecutive `NAME` cells (`Column B`) for the same party within a filing while preserving underlying cell values in every row so Excel/WPS Office filtering by `PRODUCT` or `NAME` always displays all matching records.
- **UNC Network Share Resilient**: Stores the temporary Chromium browser profile in `%LOCALAPPDATA%\EFS_Browser_Profile` so running from shared network drives (e.g. `\\server\Public\...`) never triggers `ProcessSingleton` lock errors.

## Quick Start
1. Double-click **`run_efs_reverse_v9.bat`**.
2. Click **`1. Launch Edge Browser`**, log in to EFS, open **My Filing**, and ensure **Page Size = 10**.
3. (Optional) Enter a **From Date** and **To Date** (`DD/MM/YYYY`), then click **`Scan Date Range`** or **`Reverse Scrape (To Page 1)`**.
