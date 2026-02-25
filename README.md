# BLOUK EDI Order File Generator

A browser-based tool for Bloomsbury Publishing that converts journal/book subscription order data from Excel spreadsheets into fixed-width EDI format. The tool also generates XML metadata files from a separate Excel input and packages them as a downloadable ZIP archive.

---

## Table of Contents

1.  [Overview](#overview)
2.  [File Structure](#file-structure)
3.  [Quick Start](#quick-start)
4.  [How to Use](#how-to-use)
5.  [Excel Input Format — EDI Orders](#excel-input-format--edi-orders)
6.  [EDI Output Format](#edi-output-format)
7.  [Record Type Reference](#record-type-reference)
8.  [Order Number Scheme](#order-number-scheme)
9.  [Column Mapping](#column-mapping)
10. [File Settings Reference](#file-settings-reference)
11. [Output Filename Convention](#output-filename-convention)
12. [XML Metadata Generator](#xml-metadata-generator)
13. [Dependencies](#dependencies)
14. [Browser Compatibility](#browser-compatibility)
15. [Known Limitations](#known-limitations)
16. [Development Notes](#development-notes)

---

## Overview

The tool accepts an `.xlsx`, `.xls` or `.csv` file containing subscription order rows and outputs a plain-text EDI file (`.txt`) in the BLOUK T1 fixed-width format. A separate panel accepts a metadata spreadsheet and outputs one XML file per row, bundled into `metadata.zip`. No server, backend, or installation is required — it runs entirely in the browser.

**Key features:**

-   Drag-and-drop or click-to-browse file upload
-   Automatic column mapping with manual override
-   Timestamp-based order number seeding to prevent duplicates across batches
-   Colour-coded EDI preview before download
-   File download action for the generated EDI `.txt`
-   ISO 2-letter → 3-letter country code conversion
-   ISSN/ISBN normalisation (strips hyphens, zero-pads to 13 digits)
-   XML metadata generation from a separate Excel file, downloaded as `metadata.zip`

---

## File Structure

```
blouk_edi/
├── index.html   — Application markup (HTML only, no inline scripts or styles)
├── style.css    — All visual styling and layout
├── script.js    — All application logic
└── README.md    — This document
```

The application is intentionally split into three separate files for maintainability. All logic is in `script.js`; all presentation is in `style.css`; `index.html` contains only semantic structure.

---

## Quick Start

1.  Copy the `blouk_edi/` folder to any location on your machine or web server.
2.  Open `index.html` in a modern web browser (Chrome, Edge or Firefox recommended).
3.  An internet connection is required on first load to fetch CDN libraries (SheetJS, JSZip, Font Awesome, Google Fonts). Subsequent use may work offline if the browser has cached those resources.

> **No installation, Node.js, or build tools are required.**

---

## How to Use

### EDI Order File

#### Step 1 — Upload your Order Excel file

Drag your spreadsheet onto the **Import Excel Data** upload zone, or click it to open the file picker. Accepted formats: `.xlsx`, `.xls`, `.csv`.

On successful load, a green confirmation message shows the number of rows detected, and the **Source Data** tab populates with a preview table.

#### Step 2 — Review File Settings

Adjust the settings in the **File Settings** panel if needed (see [File Settings Reference](#file-settings-reference)). The Order Number Start field is read-only and auto-populated from the current timestamp — no action is required.

#### Step 3 — Check Column Mapping (optional)

The **Column Mapping** panel is collapsed by default. Click the panel header to expand it. The tool attempts to auto-match your spreadsheet's column headers to the expected EDI field names. If your headers differ, use the dropdowns to manually assign each field.

#### Step 4 — Generate

Click **Generate EDI File**. The tool:

-   Computes a fresh timestamp-based order number seed
-   Groups rows into orders (see [Order Number Scheme](#order-number-scheme))
-   Builds all EDI records
-   Switches to the **EDI Preview** tab showing colour-coded output
-   Updates the stats bar with order, line item and record counts
-   Enables the **Download .txt** button

#### Step 5 — Download

Click **Download .txt** to save the file to your downloads folder using the standard filename convention.

---

### XML Metadata

#### Step 1 — Upload your metadata Excel file

Drag your metadata spreadsheet onto the **Generate XML Metadata** upload zone, or click it to browse. Accepted formats: `.xlsx`, `.xls`, `.csv`.

#### Step 2 — Generate and Download

Click **Generate XML & Download ZIP**. The tool generates one XML file per row (named by ISSN) and downloads them bundled as `metadata.zip`. Rows without a valid ISSN are skipped, and a count of any skipped rows is shown in the status message.

---

## Excel Input Format — EDI Orders

The tool expects one row per journal issue / subscription order. The first row must be a header row. Column order does not matter — the tool maps by header name.

### Expected Column Headers

| Column Header | EDI Field | Notes |
|---|---|---|
| `Order Ref` | Order / Item Ref | |
| `ISSN` | ISBN/ISSN (D1) | Normalised to 13 digits |
| `Journal/ Issue Title` | Title | Not written to EDI; for reference only |
| `Quantity` | D1 quantity | Falls back to Default Qty setting if blank |
| `Delivery Name ` | H2 customer name | Trailing space is significant for auto-match |
| `Delivery Company name` | H2 city field | |
| `Delivery address line 1` | H2 addr1 | |
| `Delivery address line 2` | H2 addr2 | |
| `Delivery address line 3` | H2 addr3 | |
| `Delivery Country` | H2 country code | Accepts ISO 2-letter or 3-letter codes |
| `Post code` | H2 postcode | |
| `Telephone number ` | H2 phone | Trailing space is significant for auto-match |
| `Email address` | H2 email | |

> **Tip:** Column headers must match exactly (including trailing spaces) for auto-mapping to work. Use the Column Mapping panel to fix any mismatches.

### Row Grouping

Rows that share the same **Order Ref**, **Delivery Company** and **Delivery Name** are automatically grouped into a single order with multiple D1 line items. Rows with no Order Ref are each treated as a separate single-line order.

---

## EDI Output Format

The generated file is a fixed-width ASCII text file with CRLF (`\r\n`) line endings. All fields occupy exact character positions with no delimiters. Unused space within a field is padded with spaces (text fields) or zeros (numeric fields).

### File Structure

```
$$HDR{sender}  {fileId}   {timestamp}
H1...  (350 chars)
H2...  (358 chars)
H3...  ( 20 chars)
D1...  (266 chars)   ← one per line item
[repeat H1/H2/H3/D1 blocks for each order]
$$EOF{sender}  {fileId}   {timestamp}{recordCount}
```

### Record Counts

The `$$EOF` footer includes a 7-digit record count. This count includes **only** the H1, H2, H3 and D1 lines — the `$$HDR` and `$$EOF` lines themselves are **not** counted. This matches the expectation of the receiving import system.

---

## Record Type Reference

### $$HDR — File Header

| Chars | Content |
|---|---|
| 0–4 | `$$HDR` |
| 5–8 | Sender code (e.g. `BLOO`) |
| 9–10 | Two spaces |
| 11–17 | File ID (7 chars, zero-padded) |
| 18–20 | Three spaces |
| 21–34 | Timestamp `YYYYMMDDHHMMSS` |

---

### H1 — Order Header (350 characters)

| Position | Length | Content |
|---|---|---|
| 0 | 2 | `H1` |
| 2 | 15 | Order number |
| 17 | 8 | Order date `YYYYMMDD` |
| 25 | 30 | Spaces |
| 55 | 1 | `C` (currency flag) |
| 56 | 1 | Space |
| 57 | 28 | Customer/order reference |
| 85 | 7 | Spaces |
| 92 | 8 | Carrier code (`RMA` = Royal Mail) |
| 100 | 2 | `N` |
| 102 | 30 | Spaces |
| 132 | 57 | Zeros |
| 189 | 6 | Space + 5 spaces |
| 195 | 40 | PDF placeholder (`.PDF` padded) |
| 235 | 2 | Spaces |
| 237 | 3 | Currency code (e.g. `GBP`) |
| 240 | 110 | Spaces |

---

### H2 — Customer / Address (358 characters)

| Position | Length | Content |
|---|---|---|
| 0 | 2 | `H2` |
| 2 | 15 | Order number |
| 17 | 27 | Customer code (`ST` + subscription ref) |
| 44 | 50 | Customer / contact name |
| 94 | 50 | Address line 1 |
| 144 | 50 | Address line 2 |
| 194 | 50 | Address line 3 |
| 244 | 50 | Email address |
| 294 | 32 | City / company name |
| 326 | 9 | Post code |
| 335 | 3 | Country code (ISO alpha-3, e.g. `GBR`) |
| 338 | 20 | Telephone number |

---

### H3 — Payment Terms (20 characters)

| Position | Length | Content |
|---|---|---|
| 0 | 2 | `H3` |
| 2 | 15 | Order number |
| 17 | 3 | Incoterms code (`FCA` or `DAP`) |

---

### D1 — Line Item (266 characters)

| Position | Length | Content |
|---|---|---|
| 0 | 2 | `D1` |
| 2 | 15 | Order number |
| 17 | 25 | Item reference (subscription number) |
| 42 | 5 | Line number zero-padded (`00001`) |
| 47 | 3 | Spaces |
| 50 | 18 | Zeros |
| 68 | 78 | Spaces |
| 146 | 28 | Quantity block (`0000001` + 21 zeros) |
| 174 | 12 | Unit price in pence (space + 9 digits + 2 spaces); set to 0 |
| 186 | 40 | Spaces |
| 226 | 13 | ISBN / ISSN (13 digits) |
| 239 | 27 | Spaces |

---

### $$EOF — File Footer

| Chars | Content |
|---|---|
| 0–4 | `$$EOF` |
| 5–8 | Sender code |
| 9–10 | Two spaces |
| 11–17 | File ID |
| 18–20 | Three spaces |
| 21–34 | Timestamp |
| 35–41 | Record count (7 digits, zero-padded) — H1+H2+H3+D1 lines only |

---

## Order Number Scheme

To prevent duplicate order numbers across multiple batch uploads, the starting order number is automatically derived from the current date and time at the moment the **Generate** button is clicked.

**Format:** `3` + `YY` + `MM` + `DD` + `HH` + `MM` (10 digits total)

**Example:** Generated at 14:23 on 24 February 2026 → seed `3260224142`

Orders in the same batch are sequential: `3260224142`, `3260224143`, `3260224144`, etc.

As long as two batches are not generated within the same calendar minute, their order number ranges cannot overlap. The field is displayed as read-only in the UI and refreshes on every Generate click.

---

## Column Mapping

The Column Mapping panel (collapsed by default) shows the relationship between EDI fields and your spreadsheet columns. On file load, the tool attempts to auto-match each EDI field to a spreadsheet column by exact header name.

To override a mapping, expand the panel and use the dropdowns. Select `(none)` to leave a field unmapped — it will be written as spaces/zeros in the output.

Changes to the mapping take effect on the next Generate click.

---

## File Settings Reference

| Setting | Default | Description |
|---|---|---|
| File ID / Batch Number | `0027816` | Appears in $$HDR and $$EOF markers. Update per batch if required. |
| File Prefix | `PO` | Used in the output filename only (e.g. `PO.0027816_...`). |
| Sender Code | `BLOO` | 4–6 character code written to $$HDR / $$EOF. |
| Currency | `GBP` | 3-character currency code written to each H1 record. |
| Payment Terms | `FCA` | Incoterms code written to each H3 record. `FCA` = Free Carrier; `DAP` = Delivered at Place. |
| Order Number Start | Auto | Read-only. Timestamp-derived; regenerated on each Generate click. |
| Default Qty | `1` | Fallback quantity used when the Quantity column is blank or zero. |

---

## Output Filename Convention

Downloaded EDI files use the following naming pattern:

```
{prefix}.{fileId}_{HHMM}_({DD-MM-YY}).txt
```

**Example:** `PO.0027816_1423_(24-02-26).txt`

---

## XML Metadata Generator

The XML Metadata panel accepts a separate Excel file containing book/journal specification data and generates one XML file per row, downloaded as `metadata.zip`.

### Excel Input Format

The tool expects one row per title. The first row must be a header row. Column order does not matter — headers are matched case-insensitively.

| Column Header | XML Element | Notes |
|---|---|---|
| `ISSN` | `<issn>` | Also used as the output filename (`{ISSN}.xml`) |
| `Title` | `<title>` | |
| `Trim Height` | `<trim_height>` | |
| `Trim Width` | `<trim_width>` | |
| `Spine Size` | `<spine_size>` | |
| `Paper Type` | `<paper_type>` | |
| `Binding Style` | `<binding_style>` | |
| `Page Extent` | `<page_extent>` | |
| `Lamination` | `<lamination>` | |

### XML Output Structure

Each generated file follows this structure:

```xml
<?xml version="1.0" encoding="UTF-8"?>
<book>
    <basic_info>
        <issn>9771472645051</issn>
        <title>[2026] 1 FCR 4</title>
    </basic_info>
    <specifications>
        <dimensions>
            <trim_height>245</trim_height>
            <trim_width>170</trim_width>
            <spine_size>5</spine_size>
        </dimensions>
        <materials>
            <paper_type>LetsGo Silk 90 gsm</paper_type>
            <binding_style>Limp</binding_style>
            <lamination>Matt</lamination>
        </materials>
        <page_extent>120</page_extent>
    </specifications>
</book>
```

### Output

All XML files are bundled into a single `metadata.zip` download. Each file is named `{ISSN}.xml`. Rows with no ISSN value are skipped; the status message reports how many were skipped.

---

## Dependencies

All dependencies are loaded from CDN and require an internet connection on first use.

| Library | Version | Purpose | CDN |
|---|---|---|---|
| [SheetJS (xlsx)](https://sheetjs.com/) | 0.18.5 | Excel/CSV parsing | cdnjs.cloudflare.com |
| [JSZip](https://stuk.github.io/jszip/) | 3.10.1 | ZIP file creation for XML metadata | cdnjs.cloudflare.com |
| [Font Awesome](https://fontawesome.com/) | 6.5.1 | UI icons | cdnjs.cloudflare.com |
| [IBM Plex Sans + Mono](https://fonts.google.com/specimen/IBM+Plex+Sans) | — | Typography | fonts.googleapis.com |

No npm, bundler, or build step is required.

---

## Browser Compatibility

| Browser | Support |
|---|---|
| Chrome 90+ | ✅ Full support |
| Edge 90+ | ✅ Full support |
| Firefox 88+ | ✅ Full support |
| Safari 14+ | ✅ Full support |
| Internet Explorer | ❌ Not supported |

---

## Known Limitations

-   **Price field** — Unit price (D1 positions 174–186) is always written as zero. The source Excel does not contain pricing data. If prices are required, a new column mapping would need to be added.
-   **Single H2 record per order** — The tool writes one H2 (ship-to address) per order. Some TFUK files contain a second H2 with a `CS` prefix (carrier address). This is not currently implemented.
-   **No offline support** — CDN libraries must be reachable on first load. A future improvement could bundle the dependencies locally.
-   **Same-minute duplicates** — If two EDI batches are generated within the same minute, their order number ranges will share the same seed. In normal use this is not a practical concern.
-   **XML price data** — The XML metadata output does not include pricing. If price fields are added to the metadata spreadsheet in future, a new column mapping and XML element would need to be implemented.

---

## Development Notes

### EDI Field Positions

All character positions in `script.js` were verified by byte-level comparison against live production files:

-   `T1.M02221600__22-02-26_.txt` (TFUK reference)
-   `PO_TJ_20260223-190749-830.txt` (CUP reference)

The specification document (`hachette_order_file_format_spec.md`) was found to have some inaccuracies versus the actual files. The implementation follows the **actual files**, not the spec document, for all field positions.

### Adding a New EDI Field

1.  Add an entry to the `EDI_FIELDS` array in `script.js` with a unique `key`, display `label`, and `default` header name.
2.  Add a `getCell(row, 'yourKey')` call at the appropriate character position inside `generateEDI()`.
3.  Update this README.

### Adding a New XML Metadata Field

1.  Add an entry to the `XML_FIELDS` array in `script.js` with a unique `key`, `label`, and one or more `aliases` (lowercase strings used for case-insensitive column matching).
2.  Add a `getXMLCell(row, 'yourKey')` call inside `buildXML()` at the appropriate position in the XML template string.
3.  Update this README.

### Changing the Carrier Code

The carrier code at H1 positions `[92:100]` is hardcoded as `'RMA     '` (Royal Mail, 8 chars). To change it, locate this line in `script.js` and update the string, ensuring it remains exactly 8 characters (pad with spaces if shorter).

```js
+ 'RMA     '   // [92:100]  Royal Mail carrier code
```
