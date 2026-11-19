# PROJECT RULES & DESIGN SPECIFICATIONS

## 1. App Dimensions & Layout
- **Full Size Layout**: The application must utilize full viewport width and height (`100vw`, `min-height: 100vh`) with edge-to-edge desktop-first layout and responsive scaling.

## 2. Theme & Color Palette
- **Always Light Theme**: Strict light theme only. No dark mode toggle, no dark backgrounds.
- Backgrounds: Pristine light slate / soft gray gradient (`#f8fafc`, `#f1f5f9`).
- Card Surfaces: Pure white (`#ffffff`) or subtle translucent white with refined borders (`#e2e8f0`).
- Text: High contrast professional slate (`#0f172a`, `#334155`).

## 3. Iconography
- **Professional Icons Only**: Strictly use clean SVG vector icons (e.g. Lucide Icons).
- **No Cartoon Emojis**: Never use emoji symbols for UI controls, statuses, or navigation.

## 4. Button Design System
- **Glass Transparent Buttons**: All buttons must adhere to frosted glassmorphism styling:
  - Semi-transparent background (`rgba(255, 255, 255, 0.65)` / `rgba(99, 102, 241, 0.08)`)
  - `backdrop-filter: blur(10px) saturate(180%)`
  - Subtle crisp border (`1px solid rgba(226, 232, 240, 0.8)`)
  - Smooth hover elevation and transitions
  - Clear, readable typography with matching SVG icons

## 5. Merging Logic & Engine
- **Mode 1 - Normal Excel / CSV**:
  - Treats Row 1 as Table Headers.
  - Copies Data from Row 2 to the end of the sheet.
  - Combines rows sequentially into a master sheet.
- **Mode 2 - Zoho Books Reports**:
  - Handles Zoho Books format where Row 1 is a report header/title.
  - Row 2 contains column headers.
  - Copies Data starting from Row 3 to the end.
  - **Sub-Parts & Column Retention**:
    - **Normal**: Standard Zoho reports. Keeps all columns completely without deleting any. Default filename: `Merged_Zoho_Normal_[timestamp].xlsx`.
    - **Sale**: Zoho Sales Invoices. Automatically retains ONLY 11 specified columns and deletes all other columns:
      `transaction_number`, `transaction_date`, `item.CF.Order No`, `quantity_ordered`, `bcy_item_price`, `sgst`, `cgst`, `igst`, `bcy_tax_amount`, `bcy_total_with_tax`, `sku`.
      Default filename: `Merged_Zoho_Sale_[timestamp].xlsx`.
    - **Purchase**: Zoho Purchase Bills. Automatically retains ONLY 10 specified columns and deletes all other columns:
      `transaction_number`, `transaction_date`, `item.CF.Order No`, `quantity_ordered`, `amount_without_tax`, `sgst`, `cgst`, `igst`, `bcy_total`, `sku`.
      Default filename: `Merged_Zoho_Purchase_[timestamp].xlsx`.
- **Mode 3 - ERP Merge**:
  - Treats Row 1 as Table Headers.
  - Copies Data from Row 2 to the end of each sheet.
  - **Sub-Parts & Column Retention**:
    - **Normal**: Standard ERP exports. Copies all transactional data without deleting any columns. Default filename: `Merged_ERP_Normal_[timestamp].xlsx`.
    - **Sale**: ERP Sales Invoices. Automatically retains ONLY 11 specified columns and deletes all other columns:
      `Invoice`, `Posting Date`, `Order ID`, `Stock Qty`, `Rate`, `Total Tax`, `Output Tax CGST Amount`, `Output Tax SGST Amount`, `IGST Amount`, `Total`, `Item Code`.
      Default filename: `Merged_ERP_Sale_[timestamp].xlsx`.
    - **Purchase**: ERP Purchase Bills. Automatically retains ONLY 11 specified columns and deletes all other columns:
      `Bill No`, `Posting Date`, `Order ID`, `Stock Qty`, `Rate`, `Input Tax CGST Amount`, `Input Tax SGST Amount`, `IGST Amount`, `Total Tax`, `Total`, `Item Code`.
      Default filename: `Merged_ERP_Purchase_[timestamp].xlsx`.
- **Mode 4 - Account Center Files**:
  - In the 1st file: extracts from Row 2 (Row 2 Header + Row 3+ Data).
  - In all subsequent files: extracts from Row 3 onwards (skipping Row 1 Title and Row 2 duplicate Header).
- **Mode 5 - ZIP Consolidator (Extract All to Single ZIP)**:
  - Takes multiple uploaded `.zip` archives.
  - Automatically extracts all `.xlsx`, `.xls`, and `.csv` files across all archives.
  - Repackages all extracted files into a single, clean `.zip` archive for bulk download.
- **Mode 6 - Custom Range**:
  - Allows specifying custom Header Row and Data Start Row.
- **Mode 7 - Zero-RAM GST Register Cleaner**:
  - High-performance, disk-streamed cleaning for ERP and Zoho Sales/Purchase/CN/DN registers under 35MB peak RAM.

## 6. Direct ZIP Archive Upload & Auto-Unpack
- In all merging workflows (Standard, Zoho Books, ERP, Account Center, and Custom), direct upload of `.zip` archives is natively supported.
- Archives are unpacked client-side in browser memory without sending data to a backend server.
- Extracted Excel and CSV files are immediately queued into the active merge pipeline.

## 7. Syntax Validation Rule
- **Mandatory Syntax Check**: Every file edited or created must be validated for syntax and structural integrity immediately following edits.
