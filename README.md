# Kytes Adoption Updater

A browser-based utility designed to streamline the process of extracting, aggregating, and pasting project data from Kytes feature exports into a master Adoption Sheet.

## Features
- **Auto-Detection**: Automatically identifies the type of feature export (MRN, PO, WO, GRN, DPR, TDS, Drawing, Vendor Invoice) from the uploaded `.xlsx` files based on column structures and naming conventions.
- **Accurate Row Ordering**: Extracts counts and latest dates per project and outputs them in the *exact same row order* as your master Adoption Sheet, grouped by PGH tabs.
- **One-Click Copy**: Includes a convenient UI to quickly copy specific data columns (either record counts or last updated dates) to your clipboard for easy pasting into SharePoint or Excel.
- **Excel Download**: Export the fully merged dataset as a clean, consolidated Excel workbook.
- **Local & Secure**: Runs entirely in your browser using JavaScript and `SheetJS`. No data is uploaded to any remote server.

## Usage
1. **Step 1: Upload the Adoption Sheet**
   - Drop your master Kytes Adoption Sheet (`.xlsx`) containing the `PGH 1`, `PGH 2`, `PGH 3`, and `PGH 4` tabs.
   - The tool will read the rows and establish the precise order of `Project ID`s for all subsequent outputs.
2. **Step 2: Upload Feature Exports**
   - Drop any number of feature export `.xlsx` files at once. The tool will auto-detect their type.
   - Features supported: MRN, Vendor PO, Vendor WO, GRN, DPR, TDS, Drawing, Vendor Invoice.
3. **Step 3: Copy and Paste**
   - Below, select the desired **PGH tab** and **feature tab**.
   - Click the **Copy column** or **Copy Date** button.
   - Open your master sheet, click the first cell of that specific column for that PGH, and press `Ctrl+V`.

## Technical Details
- **HTML/CSS/JS**: Vanilla frontend stack.
- **[SheetJS (xlsx)](https://sheetjs.com/)**: Used to parse and generate Excel files.
- **Tabler Icons**: Used for the UI icons.

## Setup
Simply open `index.html` in any modern web browser or serve it over a local HTTP server.
