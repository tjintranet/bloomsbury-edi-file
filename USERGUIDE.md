# User Guide
## BLOUK EDI Order File Generator — Bloomsbury Publishing

---

## What This Tool Does

This tool has two separate functions:

1. **EDI Order Files** — converts a subscription order spreadsheet into the fixed-width `.txt` format required by the TJClays POD Order Importer.
2. **XML Metadata Files** — converts a simple spreadsheet of journal titles into individual XML specification files, ready for use in production workflows.

Both functions are available from the same page. You can use either or both independently.

---

## Before You Start

You will need the two Excel templates. If you don't already have them, download them directly from the app using the buttons in each panel:

- **Download Order Template** → saves `order_file.xlsx`
- **Download User Guide** → saves `USERGUIDE.pdf`
- **Download Metadata Template** → saves `metadata_template.xlsx`

Always use the downloaded files as your starting point. **Do not rename, reorder or add columns** — the app will reject any file that doesn't match the template exactly.

---

## Part 1 — Generating an EDI Order File

### Step 1 — Fill in the order template

Open `order_file.xlsx` and enter your subscription order data in the rows below the column headers. One row per journal issue or line item.

The spreadsheet has three header rows at the top — a title bar, column group labels, and the column headers themselves. **Enter your data from row 4 onwards.** Row 4 contains a pre-filled example row shown in grey italics — you can overwrite it or leave it and start from row 5.

The columns are:

| Column | What to enter |
|---|---|
| Order Ref | Subscription or order reference number. Leading zeros are preserved — do not reformat this column. |
| ISSN (13-digit) | The 13-digit ISSN — numbers only, no spaces or hyphens. Standard ISSNs are 8 digits; pad to 13 as required by your workflow (e.g. `9781472645927`). |
| Journal/ Issue Title | Journal or issue title (for your reference — not written to the EDI file) |
| Volume Number | Volume number (for your reference only) |
| Volume Part | Volume part (for your reference only) |
| Year | Publication year (for your reference only) |
| Quantity | Number of copies ordered |
| Delivery Name | Name of the delivery contact |
| Delivery Company name | Company or institution name |
| Delivery address line 1 | First line of the delivery address |
| Delivery address line 2 | Second line of the delivery address |
| Delivery address line 3 | Third line of the delivery address |
| Delivery Country | ISO country code — see [Country Codes](#country-codes) below |
| Post code | Delivery postcode — enter as text |
| Telephone number | Contact telephone number. Enter as text to preserve leading zeros (e.g. `07700900000`). Include country code for international numbers. |
| Email address | Contact email address |

> **Tip:** Rows that share the same Order Ref, Delivery Company and Delivery Name are automatically combined into a single order with multiple line items. You don't need to do anything special — just enter the rows and the app handles the grouping.

#### Country Codes

The **Delivery Country** column accepts either:

- A **2-letter ISO code** (e.g. `GB` for United Kingdom), or
- A **3-letter ISO alpha-3 code** (e.g. `GBR`)

The column has a dropdown in Excel listing all supported codes. City or region codes such as `LON` or `NYC` are **not valid** and will be rejected — both by the spreadsheet and by the app when you generate the EDI file.

A full list of supported codes is available on the **Country Codes** sheet within `order_file.xlsx`.

#### Formatting notes

All columns in the template are pre-formatted as **text**. This means:

- Leading zeros in Order Ref and Telephone number are preserved exactly as entered
- The ISSN is stored as text so all 13 digits are kept without rounding
- Postcodes with spaces or letters are not altered

Do not change the cell format of any column.

### Step 2 — Upload the file

Drag your completed spreadsheet onto the **Import Excel Data** upload zone, or click the zone to browse for the file.

If the file is accepted, you'll see a green confirmation message and your data will appear in the **Source Data** tab.

If the file is rejected, a red error panel will list exactly which columns don't match the template. Correct the file and try again.

### Step 3 — Check settings (optional)

Click **File Settings** to expand the settings panel. You can change:

- **File ID / Batch Number** — update this for each new batch if required
- **Currency** — default is `GBP`
- **Payment Terms** — choose `FCA` (Free Carrier) or `DAP` (Delivered at Place)
- **Default Qty** — used as a fallback if any Quantity cells are blank

The **Order Number Start** field is set automatically — you don't need to touch it.

### Step 4 — Generate the EDI file

Click **Generate EDI File**.

The app will switch to the **EDI Preview** tab, where you can see the generated file with colour-coded record types. The stats bar at the top shows the number of orders, line items and total records.

If any row contains an unrecognised country code, the app will refuse to generate and display a red error panel listing each affected row with its Order Ref and the invalid value. Correct the country codes in your spreadsheet and re-upload.

### Step 5 — Download

Click **Download .txt** to save the file. The filename follows the format:

```
PO.{fileId}_{time}_({date}).txt
```

For example: `PO.0027816_1423_(24-02-26).txt`

This file is ready to submit to the BLOUK distribution system.

### Clearing and starting again

Click **Clear Data** to reset everything and start a fresh upload.

---

## Part 2 — Generating XML Metadata Files

### Step 1 — Fill in the metadata template

Open `metadata_template.xlsx` and enter your journal data below the header row. One row per title.

The template has three columns:

| Column | What to enter |
|---|---|
| ISSN | The 13-digit ISSN — numbers only, no spaces or hyphens |
| Title | The full journal or issue title |
| Page Extent | The total number of pages |

Everything else — trim size, paper type, spine width, binding and lamination — is calculated automatically by the app.

### Step 2 — Upload the file

Drag your completed spreadsheet onto the **Generate XML Metadata** upload zone, or click to browse.

The app checks two things:

- **Column structure** — the file must match the template exactly
- **ISSN format** — every ISSN must be exactly 13 digits with no spaces or hyphens

If either check fails, the file is rejected and a detailed error report is shown. Correct the issues listed and re-upload.

If the file is accepted, a green confirmation message shows how many rows were loaded.

### Step 3 — Generate

Click **Generate XML & Download ZIP**.

Two files will download automatically:

- **`metadata.zip`** — contains one `.xml` file per row, named by ISSN (e.g. `9771472645051.xml`)
- **`metadata_summary.txt`** — a plain-text report listing every file generated, with all the calculated values shown

### What the app calculates for you

Based on the Page Extent you enter, the app automatically determines:

| Value | Rule |
|---|---|
| Paper Type | `Magno Matt 130 gsm` if Page Extent is 38 pages or fewer; `Magno Matt 90 gsm` if 39 pages or more |
| Spine Size | Calculated from the page extent and paper type, rounded to the nearest whole millimetre |
| Trim Height | Always 245 mm |
| Trim Width | Always 170 mm |
| Binding Style | Always Limp |
| Lamination | Always Matt |

### Clearing and starting again

Click **Clear** to reset the metadata panel and upload a new file.

---

## Troubleshooting

**The file was rejected — "Column mismatch"**
Your spreadsheet doesn't match the expected template. The error panel shows exactly which columns are wrong. The most common causes are:
- Renamed column headers (including changing `ISSN (13-digit)` back to `ISSN`)
- Extra or missing columns
- Columns in the wrong order

Download a fresh copy of the template using the **Download Order Template** button and re-enter your data.

**The file was rejected — "Invalid country code"**
One or more rows contain a Delivery Country value that isn't a recognised ISO code. The error panel lists the row number, the invalid value, and the Order Ref for each affected row. Use a 2-letter code (e.g. `GB`) or 3-letter alpha-3 code (e.g. `GBR`). City or region codes such as `LON` are not accepted. Refer to the **Country Codes** sheet in the template for the full list.

**The file was rejected — "Invalid ISSN"**
One or more ISSNs in your metadata spreadsheet don't meet the format requirements. The error panel lists the row number and the value that failed. ISSNs must be exactly 13 digits — remove any spaces, hyphens or other characters.

**Leading zeros are missing from my Order Ref or phone number**
This happens if the column format has been changed from Text to General or Number. Download a fresh copy of the template — all columns are pre-formatted as text and will preserve leading zeros. Do not reformat any column.

**The download button doesn't appear**
You need to click **Generate EDI File** (or **Generate XML & Download ZIP**) before the download button becomes available. Uploading the file alone is not enough.

**The template download button doesn't work**
The template files (`order_file.xlsx`, `USERGUIDE.pdf`, and `metadata_template.xlsx`) must be present in the same folder as the application. Contact your system administrator if they are missing.

**The page looks wrong or features aren't working**
This app requires a modern browser. Use the latest version of Chrome, Edge or Firefox. Internet Explorer is not supported.

---

*For technical documentation, field position references and development notes, see README.md.*
