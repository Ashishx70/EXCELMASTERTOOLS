# MASTER EXCEL TOOLS - Web-based Enterprise Excel & Zoho Books Merger

A static, browser-based application designed to merge multiple Excel (`.xlsx`, `.xls`) and `.csv` files into a single consolidated file with high speed and zero server-side file uploads (100% client-side privacy).

---

## Features & Modes

### 1. Standard Excel & CSV Merger (Module 1 & 2 Logic)
- **Target Files**: `.xlsx`, `.xls`, `.csv`
- **Header**: Picked from **Row 1** of the first file.
- **Data**: Starts from **Row 2** to the end of each sheet.
- **Result**: Automatically stacks all rows sequentially.

### 2. Zoho Books Export Merger (Module 3 Logic)
- **Special Formatting**: Handles reports exported from Zoho Books (e.g., Payable Details, Vendor Balances, Sales Invoices).
- **Header**: Picked from **Row 2** (skipping Row 1 which contains report titles/organization headers).
- **Data**: Extracted starting from **Row 3** to the end.
- **3 Sub-Parts & Column Filter Rules**:
  - **Normal**: Full Zoho reports merge. Retains all columns without deleting any (`Merged_Zoho_Normal_[timestamp].xlsx`).
  - **Sale**: Retains ONLY 11 columns and deletes all other columns (`Merged_Zoho_Sale_[timestamp].xlsx`):
    - `transaction_number`, `transaction_date`, `item.CF.Order No`, `quantity_ordered`, `bcy_item_price`, `sgst`, `cgst`, `igst`, `bcy_tax_amount`, `bcy_total_with_tax`, `sku`
  - **Purchase**: Retains ONLY 10 columns and deletes all other columns (`Merged_Zoho_Purchase_[timestamp].xlsx`):
    - `bill_number`, `transaction_date`, `item.CF.Order No`, `quantity_ordered`, `amount_without_tax`, `sgst`, `cgst`, `igst`, `bcy_total`, `sku`
- **Auto-Cleanup**: Filters out empty rows and temporary lock files.

### 3. ERP Merger
- **Target Files**: ERP exported `.xlsx`, `.xls`, `.csv` files.
- **Header**: Picked from **Row 1**.
- **Data**: Starts from **Row 2** to the end of each sheet.
- **3 Sub-Parts & Column Filter Rules**:
  - **Normal**: Full ERP merge. Copies all data without deleting any columns (`Merged_ERP_Normal_[timestamp].xlsx`).
  - **Sale**: Retains ONLY 11 columns and deletes all other columns (`Merged_ERP_Sale_[timestamp].xlsx`):
    - `Invoice`, `Posting Date`, `Order ID`, `Stock Qty`, `Rate`, `Total Tax`, `Output Tax CGST Amount`, `Output Tax SGST Amount`, `IGST Amount`, `Total`, `Item Code`
  - **Purchase**: Retains ONLY 11 columns and deletes all other columns (`Merged_ERP_Purchase_[timestamp].xlsx`):
    - `Bill No`, `Posting Date`, `Order ID`, `Stock Qty`, `Rate`, `Input Tax CGST Amount`, `Input Tax SGST Amount`, `IGST Amount`, `Total Tax`, `Total`, `Item Code`

### 4. Account Center Merger
- **First File**: Starts data extraction from **Row 2** (including Header & initial records).
- **Remaining Files**: Starts data extraction from **Row 3** onwards (skips Row 1 title and duplicate Row 2 header).
- **Default Filename**: `Merged_Account_Center_[timestamp].xlsx`

### 5. ZIP Consolidator (Extract & Repack to Single ZIP)
- **Input**: Upload multiple `.zip` archives.
- **Process**: Auto-extracts all `.xlsx`, `.xls`, and `.csv` files inside all uploaded ZIPs.
- **Output**: Bundles every extracted file into a single consolidated `.zip` archive.

### 6. Direct ZIP Uploads in Merge Workflows
- You can directly drop or upload `.zip` files in Standard, Zoho Books, ERP, Account Center, and Custom modes.
- The files are unpacked client-side and merged seamlessly into one `.xlsx` or `.csv` file.

### 7. Custom Header & Data Row Config
- Define arbitrary header row and data start row numbers for custom ERP exports.

### 8. Zero-RAM GST Register Cleaner (Module 6 Logic)
- **Target Files**: Massive `.xlsx`, `.xlsm`, and `.csv` GST registers (126MB+ files).
- **Zero-RAM Architecture**: Streams disk-to-disk using SAX XML parser and temporary disk-backed SQLite string resolver. **Peak RAM strictly < 35 MB** (100% safe within Render's 512MB RAM limit).
- **Dynamic Header Detection**: Scans rows 1-35 in 2 passes (Exact match, then Substring match on headers < 40 chars).
- **Sales Registers (ERP & Zoho)**:
  - Excludes rows containing "CN", "DN", "SR", or matching regex `R\d`.
  - Retains authorized marketplace series: AJIO (`^AJ27S(J)?\d`), Myntra (`^MY27S(Y)?\d`), Marketplaces (`^(FK27S|CB27S|KT27S|RZ27S|ET27S|KK27S)\d`), CGJ (`^CGJ12627`).
- **Purchase Registers (ERP & Zoho)**:
  - Excludes rows containing "DN", "PDN", "CN", "PR", "SR", or matching regex `R\d`.
  - Retains authorized series: AJIO (`^AJ27P(J)?\d`), Myntra (`^MY27P(Y)?\d`), Marketplaces (`^(FK27P|CB27P|KT27P|KT2P7P|RZ27P|ET27P|KK27P)\d`), PGJ / CGJ (`^(PGJ12627|CGJ12627)`).
- **Credit Notes (Zoho CN)**: Retains only `^(AJ27SCN|MY27SCN|FK27SCN|RCGJ12627|RCGJ126)`.
- **Debit Notes (Zoho DN)**: Retains only `^(AJ27PDN|MY27PDN|FK27PDN|PCGJ12627|PCGJ126)`.
- **Audit & Metrics Report Card**: Displays detected column & row, total original data rows, retained rows, excluded rows, peak memory (MB), and processing time.

---

## Design System & Rules Applied

1. **Full-Size App**: Responsive, edge-to-edge layout utilizing the full viewport.
2. **Light Theme Only**: Crisp, high-contrast, professional light design system. No dark mode.
3. **Professional Vector Icons**: Clean SVG line icons (zero cartoon emojis).
4. **Transparent Glass Buttons**: Frosted glassmorphism (`backdrop-filter: blur`, subtle translucent borders).
5. **Syntax Integrity**: Automated syntax validation script (`validate_syntax.py`) after every edit.

---

## Local Testing

To run the local server with full UI and Zero-RAM Cleaning API:

```bash
# Using python3 (Linux / Mac / Render):
python3 server.py

# Or using python (Windows):
python server.py
```

Then open in your browser:
```
http://localhost:8080
```

---

## Render Deployment Guide

1. Push this folder to a GitHub repository.
2. Log in to [Render](https://render.com).
3. Click **New +** > **Web Service**.
4. Connect your GitHub repository.
5. In settings:
   - **Runtime**: `Python`
   - **Build Command**: `python3 -m pip install -r requirements.txt`
   - **Start Command**: `python3 server.py`
   - **Plan**: `Free` (512 MB RAM) - *Our engine operates at ~35 MB peak RAM!*
6. Click **Deploy Web Service** — Render will build and deploy your site live!
