# ASAJ Invoice & OCR Processing Suite

Automated pipeline for extracting tabular data from weighbridge and receipt images via Windows OCR, standardizing dataset records in Excel, and generating export commercial invoices in both HTML and print-ready PDF formats for **ASAJ Company Limited**.

---

## Overview

The `asaj-invoice` workspace provides an end-to-end data processing and document generation system:

```
[Raw Receipt Scans / Images]
           │
           ▼
[Windows Media OCR (WinRT)] ───► system/ocr_windows.ps1
           │
           ▼
[Spatial Table Clustering]  ───► system/table_extract.py
           │
           ▼
[Structured Excel Ledger]   ───► output/table_updated.xlsx
           │
           ▼
[Batch Invoice Engine]      ───► system/generate_invoices.py
     ┌─────┴─────────────────────────┐
     ▼                               ▼
[Dynamic HTML Template]      [Headless Chrome / Edge]
assets/invoice-*.html                │
                                     ▼
                           [Print-Ready PDFs & Manifest]
                           invoices/*.pdf + manifest.json
```

1. **OCR & Tabular Extraction**: Uses native Windows 10/11 WinRT OCR APIs to read text coordinates from photos/scans of weighbridge receipts, clusters bounding boxes into rows and columns, sanitizes domain values (timestamps, ASA identifiers, weight units, status flags), and outputs styled Excel workbooks.
2. **Template-Driven Commercial Invoices**: A responsive, print-optimized HTML invoice template designed for Ghanaian agricultural exports (Cashew Nuts, Dried Ginger Root, Soybeans, Moringa Leaf Powder, Fresh Mango).
3. **Automated Batch PDF Rendering**: Ingests scale weight records from Excel, synthesizes realistic export customer orders, populates the invoice template, compiles PDF files via headless browser automation, and outputs an audit manifest (`manifest.json`).

---

## Project Structure

```text
asaj-invoice/
├── assets/
│   ├── invoice-ORD-1772740748892.html   # Master HTML/CSS/JS commercial invoice template
│   ├── WhatsApp Image 2026-11-12 at 8.42.21 PM.jpeg  # Raw weighbridge receipt photo
│   └── WhatsApp Image 2026-11-12 at 8.42.41 PM.jpeg  # Raw table photo
├── invoices/                            # Generated output directory
│   ├── invoice-ORD-*.pdf                # Compiled commercial invoice PDFs (57 default orders)
│   └── manifest.json                    # Full metadata ledger for all generated invoices
├── output/
│   ├── table.xlsx                       # Initial OCR extracted table
│   └── table_updated.xlsx               # Cleaned & verified lot/weight records
├── system/
│   ├── generate_invoices.py             # Batch invoice synthesis and PDF compilation engine
│   ├── ocr_windows.ps1                  # PowerShell bridge to Windows.Media.Ocr WinRT API
│   └── table_extract.py                 # Table reconstruction & Excel formatting CLI
├── requirements.txt                     # Python dependencies (openpyxl)
└── README.md                            # Project documentation
```

---

## System Requirements & Prerequisites

- **Operating System**: Windows 10 or Windows 11 (required for WinRT `Windows.Media.Ocr`).
- **PowerShell**: PowerShell 5.1+ (standard on Windows).
- **Python**: Python 3.9+ installed.
- **Browser**: Google Chrome or Microsoft Edge installed in standard paths (used for `--headless=new` PDF printing).

### Python Dependencies

Install required dependencies:

```bash
pip install -r requirements.txt
```

*(Only `openpyxl>=3.1` is required for external packages; all other modules use the Python standard library).*

---

## Usage Guide

### 1. Extract Tables from Receipt Images

The `system/table_extract.py` script executes OCR on an image, reconstructs row and column grids based on pixel geometry, normalizes data fields, and exports the data into an Excel spreadsheet.

#### Basic Extraction:

```bash
python system/table_extract.py "assets/WhatsApp Image 2026-11-12 at 8.42.21 PM.jpeg" -o "output/table.xlsx"
```

#### Customized Column Extraction:

You can specify column names, enforce column counts via k-means clustering, reorder columns, or adjust spatial tolerances:

```bash
python system/table_extract.py "assets/WhatsApp Image 2026-11-12 at 8.42.21 PM.jpeg" `
  -o "output/table.xlsx" `
  --columns 10 `
  --names "row_no,lot_code,time,code_936,weight_g,weight_kg,status,unit,asa_id,product_code" `
  --order "row_no,lot_code,time,weight_g,status,asa_id" `
  --row-tolerance 8.0 `
  --column-gap 30.0
```

#### CLI Options:

| Flag | Default | Description |
| :--- | :--- | :--- |
| `image` | *(required)* | Path to the target receipt image. |
| `-o`, `--output` | `<image>.xlsx` | Path for the destination `.xlsx` file. |
| `--columns` | `None` (auto) | Force expected column count using k-means clustering. |
| `--names` | `None` | Comma-separated column names assigned in left-to-right order. |
| `--order` | `None` | Comma-separated list of column names moved to the front. |
| `--row-tolerance` | `8.0` | Vertical tolerance (px) for grouping words into the same row. |
| `--column-gap` | `30.0` | Horizontal separation (px) triggering a new column. |
| `--sheet` | `Data` | Name of the Excel sheet. |
| `--raw` | `False` | Disables field value normalization and keeps raw OCR text. |

---

### 2. Batch Invoice Generation

The `system/generate_invoices.py` script reads the weight data from `output/table_updated.xlsx`, constructs order details using the ASAJ export product catalog and Ghanaian customer records, populates the HTML template, and compiles PDFs via headless Chrome or Edge.

#### Generate Standard Batch (57 Invoices):

```bash
python system/generate_invoices.py
```

#### Custom Options:

```bash
# Generate 10 invoices, keep the generated HTML files, and use a custom random seed:
python system/generate_invoices.py --count 10 --seed 42 --keep-html

# Dry-run mode: Generate manifest.json and HTML without invoking headless browser for PDFs:
python system/generate_invoices.py --no-pdf --keep-html
```

#### CLI Options:

| Flag | Default | Description |
| :--- | :--- | :--- |
| `--table` | `output/table_updated.xlsx` | Source Excel workbook providing weights. |
| `--template` | `assets/invoice-ORD-1772740748892.html` | Master HTML invoice template. |
| `--out-dir` | `invoices/` | Target folder for PDFs and `manifest.json`. |
| `--count` | `57` | Number of invoices to generate. |
| `--seed` | `20261013` | Deterministic random seed for reproducible customer & order assignment. |
| `--keep-html` | `False` | Retains intermediate rendered HTML files in `invoices/html/`. |
| `--no-pdf` | `False` | Skips browser headless PDF rendering (manifest & HTML only). |

---

## Product Catalog & Data Schema

The generation script models ASAJ Company Limited's export products priced in Ghanaian Cedi (GHS):

| Commodity | Grade / Specification | Unit Price (GHS) |
| :--- | :--- | :--- |
| **Cashew nuts** | Premium Raw Cashew Nuts – Grade W240 (Moisture ≤5%) | 159.00 / kg |
| **Ginger** | Dried Ginger Root (High Oil Content, Well Dried) | 326.00 / kg |
| **Soybean** | Non-GMO Soybeans – Export Grade (High Protein) | 340.00 / kg |
| **Moringa** | Moringa Leaf Powder – Certified Organic | 23.80 / kg |
| **Mango** | Fresh Mango – Export Grade | 160.00 / kg |

### Invoice Manifest Schema

The resulting `invoices/manifest.json` contains full audit records for every invoice generated:

```json
{
  "company": {
    "name": "ASAJ Company Limited",
    "subtitle": "ASAJ Commercial Invoice"
  },
  "invoiceNumber": "INV-ORD-1938855491696",
  "transaction": {
    "reference": "ORD-1938855491696",
    "date": "07-Jan-2026",
    "dispatchDate": "12-Jan-2026",
    "deliveryDate": "19-Jan-2026"
  },
  "customer": {
    "name": "Serwaa Addo",
    "email": "serwaaaddo76@gmail.com",
    "phone": "0200398139"
  },
  "delivery": {
    "status": "Processing",
    "type": "Export shipment"
  },
  "currency": "GHS",
  "items": [
    {
      "description": "Moringa Leaf Powder – Certified Organic",
      "qty": 0.29,
      "unit": "kg",
      "unitPrice": 23.8,
      "lineTotal": 6.9,
      "allocatedWeightKg": 5.947
    }
  ],
  "_source": {
    "weight_g": 23786.0,
    "totalWeightKg": 23.786
  }
}
```

---

## Customizing the HTML Template

The HTML template at `assets/invoice-ORD-1772740748892.html` is a standalone single-file document:

- **Print Friendly**: Configured with `@media print` rules, margin sizing, and page break handling. Open the file in any browser and use **Print / Save as PDF** (`Ctrl+P`).
- **Data-Driven**: Driven by a single JavaScript object:
  ```javascript
  const INVOICE = {
    company: { name: "ASAJ Company Limited", subtitle: "ASAJ Commercial Invoice" },
    invoiceNumber: "INV-ORD-1772740748892",
    transaction: { ... },
    customer: { ... },
    items: [ ... ],
    notes: "..."
  };
  ```
- Editing this object and refreshing the page dynamically recalculates sub-totals, discounts, VAT (if configured), and updates delivery badges.

