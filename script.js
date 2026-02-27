/**
 * script.js — BLOUK EDI Order File Generator
 * Bloomsbury Publishing
 *
 * Responsibilities:
 *   - Parse uploaded Excel / CSV files via SheetJS
 *   - Auto-map spreadsheet columns to EDI field slots
 *   - Build fixed-width EDI records ($$HDR, H1, H2, H3, D1, $$EOF)
 *   - Render the imported data as an HTML table
 *   - Syntax-highlight the generated EDI output in the preview pane
 *   - Provide Copy-to-clipboard and file Download actions
 *
 * EDI Format: T1 fixed-width, BLOUK variant
 * Carrier:    Royal Mail (RMA)
 * Encoding:   ASCII, CRLF line endings
 */

'use strict';

/* ============================================================
   STATE
   Module-level variables shared across functions.
   ============================================================ */

/** Rows from the imported spreadsheet (array of arrays, no header). */
let rawData = [];

/** Column header labels from the first row of the spreadsheet. */
let headers = [];

/**
 * Maps each EDI field key to the 0-based column index in rawData.
 * Value of -1 means the field is unmapped / not present.
 * @type {Object.<string, number>}
 */
let columnMap = {};

/** The last generated EDI string (CRLF-separated). Used by Copy and Download. */
let ediOutput = '';


/* ============================================================
   EDI FIELD DEFINITIONS
   Declares which EDI fields the app expects and their default
   auto-match header names (matching the sample Excel columns).
   ============================================================ */

/**
 * @typedef  {Object} EdiField
 * @property {string} key     - Internal identifier used in columnMap
 * @property {string} label   - Human-readable label shown in the mapping UI
 * @property {string} default - Spreadsheet header name to auto-match against
 */

/** @type {EdiField[]} */
const EDI_FIELDS = [
  { key: 'subscriptionNum', label: 'Order Ref',                 default: 'Order Ref'                  },
  { key: 'isbn',            label: 'ISSN',                      default: 'ISSN'                       },
  { key: 'title',           label: 'Journal/ Issue Title',      default: 'Journal/ Issue  Title'      },
  { key: 'volumeNumber',    label: 'Volume Number',             default: 'Volume Number '             },
  { key: 'volumePart',      label: 'Volume Part',               default: 'Volume Part'                },
  { key: 'year',            label: 'Year',                      default: 'Year'                       },
  { key: 'quantity',        label: 'Quantity',                  default: 'Quantity'                   },
  { key: 'deliveryName',    label: 'Delivery Name',             default: 'Delivery Name '             },
  { key: 'deliveryCompany', label: 'Delivery Company name',     default: 'Delivery Company name'      },
  { key: 'addr1',           label: 'Delivery address line 1',   default: 'Delivery address line 1'    },
  { key: 'addr2',           label: 'Delivery address line 2',   default: 'Delivery address line 2'    },
  { key: 'addr3',           label: 'Delivery address line 3',   default: 'Delivery address line 3'    },
  { key: 'country',         label: 'Delivery Country',          default: 'Delivery Country'           },
  { key: 'postcode',        label: 'Post code',                 default: 'Post code'                  },
  { key: 'phone',           label: 'Telephone number',          default: 'Telephone number '          },
  { key: 'email',           label: 'Email address',             default: 'Email address'              },
];

/**
 * Canonical ordered column list for the Order Excel file.
 * Exact strings as they appear in the template (including trailing spaces).
 * Used to validate uploads — any deviation triggers an error.
 */
const ORDER_TEMPLATE_COLUMNS = [
  'Order Ref',
  'ISSN',
  'Journal/ Issue  Title',
  'Volume Number ',
  'Volume Part',
  'Year',
  'Quantity',
  'Delivery Name ',
  'Delivery Company name',
  'Delivery address line 1',
  'Delivery address line 2',
  'Delivery address line 3',
  'Delivery Country',
  'Post code',
  'Telephone number ',
  'Email address',
];

/**
 * Canonical ordered column list for the Metadata Excel file.
 * Exact strings as they appear in the template.
 */
const METADATA_TEMPLATE_COLUMNS = [
  'ISSN',
  'Title',
  'Trim Height',
  'Trim Width',
  'Spine Size',
  'Paper Type',
  'Binding Style',
  'Page Extent',
  'Lamination',
];

/**
 * Validates that an uploaded file's headers exactly match a canonical template.
 * Checks column count, names, and order.
 * Returns null on success, or an object { type, message, details } on failure.
 *
 * @param {string[]} actual    - Headers from the uploaded file
 * @param {string[]} expected  - Canonical template column list
 * @param {string}   fileLabel - Human-readable label for error messages
 * @return {{ type: string, message: string, details: string[] } | null}
 */
function validateColumns(actual, expected, fileLabel) {
  const errors = [];

  // Check column count
  if (actual.length !== expected.length) {
    errors.push(
      `Expected ${expected.length} column${expected.length !== 1 ? 's' : ''}, ` +
      `found ${actual.length}.`
    );
  }

  // Check each position
  const maxLen = Math.max(actual.length, expected.length);
  for (let i = 0; i < maxLen; i++) {
    const act = actual[i] ?? '(missing)';
    const exp = expected[i] ?? '(unexpected)';
    if (act !== exp) {
      errors.push(
        `Column ${i + 1}: expected "${exp}" — got "${act}".`
      );
    }
  }

  if (errors.length === 0) return null;

  return {
    type: 'column_mismatch',
    message: `The uploaded file does not match the ${fileLabel} template.`,
    details: errors,
  };
}


/* ============================================================
   ISO 2-LETTER → 3-LETTER COUNTRY CODE LOOKUP
   The EDI format requires ISO 3166-1 alpha-3 codes.
   Covers the main territories seen in publishing distribution.
   ============================================================ */

/** @type {Object.<string, string>} */
const ISO_2_TO_3 = {
  AR: 'ARG', AT: 'AUT', AU: 'AUS', BE: 'BEL', BR: 'BRA',
  CA: 'CAN', CH: 'CHE', CL: 'CHL', CN: 'CHN', CO: 'COL',
  CZ: 'CZE', DE: 'DEU', DK: 'DNK', EG: 'EGY', ES: 'ESP',
  FI: 'FIN', FR: 'FRA', GB: 'GBR', GR: 'GRC', HK: 'HKG',
  HU: 'HUN', ID: 'IDN', IE: 'IRL', IL: 'ISR', IN: 'IND',
  IT: 'ITA', JP: 'JPN', KE: 'KEN', KR: 'KOR', MX: 'MEX',
  MY: 'MYS', NG: 'NGA', NL: 'NLD', NO: 'NOR', NZ: 'NZL',
  PH: 'PHL', PL: 'POL', PT: 'PRT', RO: 'ROU', SE: 'SWE',
  SG: 'SGP', TH: 'THA', TR: 'TUR', TW: 'TWN', US: 'USA',
  VN: 'VNM', ZA: 'ZAF',
};

/**
 * Converts an ISO country code to 3-letter alpha-3 format.
 * Accepts 2-letter (ISO 3166-1 alpha-2) or passes through 3-letter codes unchanged.
 *
 * @param  {string} code - Raw country code from the spreadsheet
 * @return {string}        3-character country code, space-padded if needed
 */
function isoToAlpha3(code) {
  if (!code) return '   ';
  code = code.trim().toUpperCase();
  if (code.length === 3) return code;
  if (code.length === 2) return ISO_2_TO_3[code] || (code + ' ');
  return code.substring(0, 3);
}


/* ============================================================
   STRING UTILITY FUNCTIONS
   Fixed-width EDI requires exact field lengths. These helpers
   pad or truncate values to precise character counts.
   ============================================================ */

/**
 * Right-pads a value with spaces to an exact length, or truncates if longer.
 *
 * @param  {*}      val  - Value to format (will be coerced to string)
 * @param  {number} len  - Required field width in characters
 * @param  {string} [fill=' '] - Fill character (default: space)
 * @return {string}        Exactly `len` characters long
 */
function pad(val, len, fill = ' ') {
  return String(val ?? '').substring(0, len).padEnd(len, fill);
}

/**
 * Left-pads a numeric value with zeros to an exact length, or truncates if longer.
 * Used for line numbers, quantities, prices and the record count footer.
 *
 * @param  {*}      val  - Value to format
 * @param  {number} len  - Required field width in characters
 * @param  {string} [fill='0'] - Fill character (default: zero)
 * @return {string}        Exactly `len` characters long
 */
function padLeft(val, len, fill = '0') {
  return String(val ?? '').padStart(len, fill).substring(0, len);
}

/**
 * Normalises an ISSN or ISBN value to a plain 13-digit string.
 * Strips hyphens and spaces, then zero-pads or truncates to 13 chars.
 *
 * @param  {string} val - Raw ISSN/ISBN from the spreadsheet
 * @return {string}       13-character numeric string
 */
function normalizeISSN(val) {
  const digits = String(val ?? '').replace(/[-\s]/g, '').replace(/\D/g, '');
  if (digits.length === 13) return digits;
  if (digits.length > 13)   return digits.substring(0, 13);
  return digits.padStart(13, '0');
}

/**
 * Builds a 14-character timestamp string in YYYYMMDDHHMMSS format.
 * Used in the $$HDR and $$EOF file markers.
 *
 * @param  {Date} date
 * @return {string}
 */
function makeFileTimestamp(date) {
  const y  = date.getFullYear();
  const mo = String(date.getMonth() + 1).padStart(2, '0');
  const d  = String(date.getDate()).padStart(2, '0');
  const h  = String(date.getHours()).padStart(2, '0');
  const mi = String(date.getMinutes()).padStart(2, '0');
  const s  = String(date.getSeconds()).padStart(2, '0');
  return `${y}${mo}${d}${h}${mi}${s}`;
}

/**
 * Converts a string to safe 7-bit ASCII for EDI output.
 * Replaces common Unicode punctuation with ASCII equivalents,
 * then strips any remaining non-ASCII characters.
 * This prevents multi-byte UTF-8 sequences from inflating field lengths
 * and corrupting fixed-width field alignment.
 *
 * @param  {string} str - Input string (may contain Unicode)
 * @return {string}       ASCII-safe string
 */
function toAscii(str) {
  return str
    .replace(/[\u2018\u2019\u201A\u201B]/g, "'")   // curly single quotes -> '
    .replace(/[\u201C\u201D\u201E\u201F]/g, '"')   // curly double quotes -> "
    .replace(/[\u2013\u2014\u2015]/g,       '-')   // en-dash / em-dash   -> -
    .replace(/\u2026/g,                    '...') // ellipsis            -> ...
    .replace(/\u00A0/g,                     ' ')   // non-breaking space  -> space
    .replace(/[^\x00-\x7F]/g,              '');    // strip anything else
}

/**
 * Retrieves and trims a cell value from a data row using an EDI field key.
 * Returns an empty string if the field is unmapped or the cell is null.
 * All values are sanitised to 7-bit ASCII to preserve fixed-width field alignment.
 *
 * @param  {Array}  row - A single data row from rawData
 * @param  {string} key - EDI field key from EDI_FIELDS
 * @return {string}
 */
function getCell(row, key) {
  const idx = columnMap[key];
  if (idx === undefined || idx < 0) return '';
  const val = row[idx];
  if (val === null || val === undefined) return '';
  return toAscii(String(val).trim());
}

/**
 * Reads the quantity for a row, falling back to the Default Qty setting.
 *
 * @param  {Array} row - A single data row from rawData
 * @return {number}      Quantity as a positive integer
 */
function getQuantity(row) {
  const raw = getCell(row, 'quantity');
  const q   = parseInt(raw, 10);
  if (!isNaN(q) && q >= 1) return q;
  return parseInt(document.getElementById('defaultQty').value, 10) || 1;
}


/* ============================================================
   FILE UPLOAD
   Handles both drag-and-drop and click-to-browse interactions.
   Passes the selected File object to processFile().
   ============================================================ */

const dropzone  = document.getElementById('dropzone');
const fileInput = document.getElementById('fileInput');

/** Open the native file picker when the dropzone is clicked. */
dropzone.addEventListener('click', () => fileInput.click());

/** Highlight the dropzone while a file is dragged over it. */
dropzone.addEventListener('dragover', e => {
  e.preventDefault();
  dropzone.classList.add('drag-over');
});

/** Remove highlight when the drag leaves the dropzone. */
dropzone.addEventListener('dragleave', () => {
  dropzone.classList.remove('drag-over');
});

/** Accept a dropped file and pass it for processing. */
dropzone.addEventListener('drop', e => {
  e.preventDefault();
  dropzone.classList.remove('drag-over');
  if (e.dataTransfer.files[0]) processFile(e.dataTransfer.files[0]);
});

/** Accept a file chosen via the native file picker. */
fileInput.addEventListener('change', () => {
  if (fileInput.files[0]) processFile(fileInput.files[0]);
});

/**
 * Reads a File object with SheetJS, extracts headers and data rows,
 * then triggers the mapping UI and table render.
 *
 * @param {File} file - The user-selected spreadsheet file
 */
function processFile(file) {
  const reader = new FileReader();

  reader.onload = e => {
    try {
      const wb   = XLSX.read(e.target.result, { type: 'binary' });
      const ws   = wb.Sheets[wb.SheetNames[0]];
      const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });

      if (data.length < 1) throw new Error('The file appears to be empty.');

      // First row = column headers; remaining rows = data
      const fileHeaders = data[0].map(h => String(h ?? ''));

      // ── Strict column validation against the order template ──────────────
      const validationError = validateColumns(fileHeaders, ORDER_TEMPLATE_COLUMNS, 'Order');
      if (validationError) {
        showColumnError('uploadMsg', validationError, 'order_file.xlsx');
        return;
      }

      if (data.length < 2) throw new Error('No data rows found in the file.');

      headers = fileHeaders;
      rawData = data.slice(1).filter(row => row.some(c => c !== null && c !== ''));

      showMessage(
        'uploadMsg',
        `<i class="fa-solid fa-circle-check"></i> Loaded ${rawData.length} rows from "${file.name}"`,
        'success'
      );

      buildMappingUI();
      renderTable();

    } catch (err) {
      showMessage(
        'uploadMsg',
        `<i class="fa-solid fa-circle-xmark"></i> Error: ${err.message}`,
        'error'
      );
    }
  };

  reader.readAsBinaryString(file);
}


/* ============================================================
   COLUMN MAPPING UI
   Builds a 3-column grid of dropdowns (target → source).
   Each EDI field gets a <select> pre-populated with spreadsheet
   headers, auto-matched where the header name matches exactly.
   ============================================================ */

/**
 * Constructs the column mapping panel from EDI_FIELDS and the
 * loaded spreadsheet headers. Shows the panel and pre-selects
 * any headers that match the default names exactly.
 */
function buildMappingUI() {
  const panel = document.getElementById('mappingPanel');
  const grid  = document.getElementById('mappingGrid');

  grid.innerHTML = '';
  columnMap      = {};

  EDI_FIELDS.forEach(field => {
    // Attempt to auto-match the header by exact name
    const autoIdx = headers.findIndex(h => h.trim() === field.default.trim());
    columnMap[field.key] = autoIdx >= 0 ? autoIdx : -1;

    // Target label (EDI field name)
    const targetDiv = document.createElement('div');
    targetDiv.className   = 'target';
    targetDiv.textContent = field.label;

    // Arrow separator
    const arrow = document.createElement('div');
    arrow.className   = 'arrow';
    arrow.textContent = '→';

    // Source column dropdown
    const sel = document.createElement('select');
    sel.id = `map_${field.key}`;

    // "(none)" option — no mapping for this field
    const noneOpt = document.createElement('option');
    noneOpt.value       = '-1';
    noneOpt.textContent = '(none)';
    sel.appendChild(noneOpt);

    // One option per spreadsheet column
    headers.forEach((h, i) => {
      const opt = document.createElement('option');
      opt.value       = i;
      opt.textContent = h.length > 22 ? h.substring(0, 22) + '…' : h;
      if (i === columnMap[field.key]) opt.selected = true;
      sel.appendChild(opt);
    });

    // Keep columnMap in sync when user changes a dropdown
    sel.addEventListener('change', () => {
      columnMap[field.key] = parseInt(sel.value, 10);
    });

    grid.appendChild(targetDiv);
    grid.appendChild(arrow);
    grid.appendChild(sel);
  });

  // Reveal the mapping panel (collapsed by default)
  panel.style.display = 'block';
}

/**
 * Toggles the File Settings panel open or closed.
 * Called by the panel-title click handler in index.html.
 */
function toggleSettings() {
  const toggle = document.getElementById('settingsToggle');
  const body   = document.getElementById('settingsBody');
  const isOpen = body.classList.contains('open');
  body.classList.toggle('open', !isOpen);
  toggle.classList.toggle('open', !isOpen);
}

/**
 * Toggles the Column Mapping panel open or closed.
 * Called by the panel-title click handler in index.html.
 */
function toggleMapping() {
  const toggle = document.getElementById('mappingToggle');
  const body   = document.getElementById('mappingBody');
  const isOpen = body.classList.contains('open');
  body.classList.toggle('open', !isOpen);
  toggle.classList.toggle('open', !isOpen);
}


/* ============================================================
   TABLE RENDER
   Displays the imported spreadsheet rows in the Source Data tab.
   ============================================================ */

/**
 * Builds and inserts an HTML <table> from rawData and headers.
 * Also updates the stats bar row count.
 */
function renderTable() {
  const wrapper  = document.getElementById('tableWrapper');
  const noMsg    = document.getElementById('noDataMsg');
  const statsRow = document.getElementById('statsRow');

  if (!rawData.length) return;

  noMsg.style.display    = 'none';
  wrapper.style.display  = 'block';
  statsRow.style.display = 'flex';
  document.getElementById('statRows').textContent = rawData.length;

  // Build table
  const table = document.createElement('table');
  table.className = 'data-table';

  // Header row
  const thead = document.createElement('thead');
  const trHead = document.createElement('tr');
  const thNum = document.createElement('th');
  thNum.textContent = '#';
  trHead.appendChild(thNum);
  headers.forEach(h => {
    const th = document.createElement('th');
    th.textContent = h;
    trHead.appendChild(th);
  });
  thead.appendChild(trHead);
  table.appendChild(thead);

  // Data rows
  const tbody = document.createElement('tbody');
  rawData.forEach((row, ri) => {
    const tr = document.createElement('tr');

    const tdNum = document.createElement('td');
    tdNum.className   = 'row-num';
    tdNum.textContent = ri + 1;
    tr.appendChild(tdNum);

    headers.forEach((_, ci) => {
      const td = document.createElement('td');
      td.textContent = row[ci] ?? '';
      td.title       = row[ci] ?? ''; // show full value on hover
      tr.appendChild(td);
    });

    tbody.appendChild(tr);
  });
  table.appendChild(tbody);

  wrapper.innerHTML = '';
  wrapper.appendChild(table);
}


/* ============================================================
   EDI FILE GENERATION
   The core function. Reads settings from the UI, groups rows
   into orders, and constructs each fixed-width record type.

   Record structure (verified against live BLOUK/TFUK files):

     $$HDR  — file header marker (not counted in footer total)
     H1     — order header         (350 chars)
     H2     — customer/address     (358 chars)
     H3     — payment terms        (20 chars)
     D1     — line item (×N)       (266 chars)
     $$EOF  — file footer with record count (excludes $$HDR / $$EOF)

   ============================================================ */

/**
 * Main generation entry point, called by the Generate button.
 * Reads all UI settings, processes rawData into order groups,
 * builds EDI lines, updates stats, renders the preview pane,
 * and enables the Copy/Download buttons.
 */
function generateEDI() {
  if (!rawData.length) {
    alert('Please upload an Excel file first.');
    return;
  }

  // ── Read UI settings ──────────────────────────────────────
  const senderCode = pad(document.getElementById('senderCode').value, 4);
  const currency   = pad(document.getElementById('currency').value, 3);
  const payTerms   = document.getElementById('payTerms').value;

  // ── Timestamp — used for file header/footer and order seed ──
  const now = new Date();
  const ts  = makeFileTimestamp(now);

  // ── Order number seed: 3 + YYMMDDHHММ (10 digits, unique per minute) ──
  const yy = String(now.getFullYear()).slice(2);
  const mo = String(now.getMonth() + 1).padStart(2, '0');
  const dd = String(now.getDate()).padStart(2, '0');
  const hh = String(now.getHours()).padStart(2, '0');
  const mi = String(now.getMinutes()).padStart(2, '0');
  const orderStart = parseInt(`3${yy}${mo}${dd}${hh}${mi}`, 10);

  // Reflect the computed seed in the read-only UI field
  document.getElementById('orderNumStart').value = orderStart;

  // ── File ID string (7 chars, zero-padded left) ─────────────
  const fileIdStr = padLeft(document.getElementById('fileId').value.trim(), 7);

  // Array that will hold all output lines
  const lines = [];

  // Counters for stats bar and footer record count
  let recordCount = 0;
  let orderCount  = 0;
  let lineCount   = 0;

  // ── FILE HEADER ────────────────────────────────────────────
  // Format: $$HDR + senderCode(4) + 2 spaces + fileId(7) + 3 spaces + timestamp(14)
  // Example: $$HDRBLOO  0027816   20260224160055
  lines.push(`$$HDR${senderCode}  ${fileIdStr}   ${ts}`);


  // ── GROUP ROWS INTO ORDERS ─────────────────────────────────
  // Rows sharing the same subscriptionNum + deliveryCompany + deliveryName
  // are grouped as multiple D1 line items under a single order.
  // Rows with no subscriptionNum each become their own separate order.

  /** @type {Array<{key: string, ref: string, rows: Array, orderNum: number}>} */
  const orders   = [];
  const orderMap = {};

  rawData.forEach(row => {
    const subNum  = getCell(row, 'subscriptionNum');
    const company = getCell(row, 'deliveryCompany');
    const name    = getCell(row, 'deliveryName');

    const groupKey = subNum
      ? `${subNum}||${company}||${name}`
      : `__row_${Math.random()}`;    // unique key — each row is its own order

    if (!orderMap[groupKey]) {
      const order = {
        key:      groupKey,
        ref:      subNum,
        rows:     [],
        orderNum: orderStart + orders.length,
      };
      orderMap[groupKey] = order;
      orders.push(order);
    }

    orderMap[groupKey].rows.push(row);
  });


  // ── BUILD RECORDS PER ORDER ────────────────────────────────
  orders.forEach(order => {
    const firstRow    = order.rows[0];
    const orderNumStr = pad(String(order.orderNum), 15);
    const dateStr     = `${now.getFullYear()}${mo}${dd}`;

    // Reference field: subscription number or ISSN as fallback
    const ref     = order.ref || getCell(firstRow, 'isbn') || '';
    const refStr  = pad(ref, 28);
    const pdfName = pad('.PDF', 40);  // no invoice reference required


    // ── H1 RECORD (350 characters) ──────────────────────────
    // Order header: date, customer reference, carrier, currency.
    //
    // Position map (0-based):
    //   [0:2]    'H1'
    //   [2:17]   order number          (15)
    //   [17:25]  date YYYYMMDD          (8)
    //   [25:55]  spaces                (30)
    //   [55:56]  'C' (currency flag)    (1)
    //   [56:57]  space                  (1)
    //   [57:85]  customer reference    (28)
    //   [85:92]  spaces                 (7)
    //   [92:100] carrier code           (8) — 'RMA     ' = Royal Mail
    //   [100:102] ' N'                  (2)
    //   [102:132] spaces               (30)
    //   [132:189] zeros                (57)
    //   [189:190] space                 (1)
    //   [190:195] spaces                (5)
    //   [195:235] PDF placeholder      (40) — '.PDF' padded
    //   [235:237] spaces                (2)
    //   [237:240] currency code         (3) — e.g. 'GBP'
    //   [240:350] spaces              (110)

    let h1 = 'H1'
      + orderNumStr        // [2:17]
      + dateStr            // [17:25]
      + pad('', 30)        // [25:55]
      + 'C '               // [55:57]
      + refStr             // [57:85]
      + pad('', 7)         // [85:92]
      + 'RMA     '         // [92:100]  Royal Mail carrier code
      + ' N'               // [100:102]
      + pad('', 30)        // [102:132]
      + '0'.repeat(57)     // [132:189]
      + ' '                // [189:190]
      + pad('', 5)         // [190:195]
      + pdfName            // [195:235]
      + '  '               // [235:237]
      + currency           // [237:240]
      + pad('', 110);      // [240:350]

    lines.push(h1.substring(0, 350).padEnd(350));
    recordCount++;


    // ── H2 RECORD (358 characters) ──────────────────────────
    // Customer name and delivery address.
    //
    // Position map (0-based):
    //   [0:2]    'H2'
    //   [2:17]   order number          (15)
    //   [17:44]  customer code         (27) — 'ST' + subscription ref
    //   [44:94]  customer name         (50) — delivery contact name
    //   [94:144] address line 1        (50)
    //   [144:194] address line 2       (50)
    //   [194:244] address line 3       (50)
    //   [244:294] email address        (50)
    //   [294:326] city / company       (32)
    //   [326:335] post code             (9)
    //   [335:338] country code (alpha-3)(3)
    //   [338:358] telephone            (20)

    const subCode = order.ref
      ? pad(`ST${order.ref}`, 27)
      : pad('ST', 27);

    const addr2raw = getCell(firstRow, 'addr2');
    const addr3raw = getCell(firstRow, 'addr3');

    let h2 = 'H2'
      + orderNumStr                                                   // [2:17]
      + subCode                                                       // [17:44]
      + pad(getCell(firstRow, 'deliveryName'), 50)                   // [44:94]
      + pad(getCell(firstRow, 'addr1'), 50)                          // [94:144]
      + pad(addr2raw || addr3raw, 50)                                // [144:194]
      + pad(addr2raw && addr3raw ? addr3raw : '', 50)                // [194:244]
      + pad(getCell(firstRow, 'email'), 50)                          // [244:294]
      + pad(getCell(firstRow, 'deliveryCompany'), 32)                // [294:326]
      + pad(getCell(firstRow, 'postcode'), 9)                        // [326:335]
      + pad(isoToAlpha3(getCell(firstRow, 'country')), 3)            // [335:338]
      + pad(getCell(firstRow, 'phone'), 20);                         // [338:358]

    lines.push(h2.substring(0, 358).padEnd(358));
    recordCount++;


    // ── H3 RECORD (20 characters) ───────────────────────────
    // Payment / Incoterms code (e.g. FCA or DAP).
    //
    // Position map (0-based):
    //   [0:2]   'H3'
    //   [2:17]  order number  (15)
    //   [17:20] payment terms  (3)

    lines.push('H3' + orderNumStr + payTerms);
    recordCount++;


    // ── D1 RECORDS (266 characters each) ────────────────────
    // One record per line item (book/journal issue) in the order.
    //
    // Position map (0-based):
    //   [0:2]    'D1'
    //   [2:17]   order number          (15)
    //   [17:42]  item reference        (25) — subscription number
    //   [42:47]  line number '00001'    (5)
    //   [47:50]  spaces                 (3)
    //   [50:68]  zeros                 (18)
    //   [68:146] spaces                (78)
    //   [146:174] quantity block       (28) — '0000001' + 21 zeros
    //   [174:186] price in pence       (12) — space + 9 digits + 2 spaces
    //   [186:226] spaces               (40)
    //   [226:239] ISBN / ISSN          (13)
    //   [239:266] spaces               (27)

    order.rows.forEach((row, lineIdx) => {
      const itemRef  = pad(getCell(row, 'subscriptionNum') || String(lineIdx + 1).padStart(9, '0'), 25);
      const lineNum  = padLeft(lineIdx + 1, 5);
      const qty      = getQuantity(row);
      const qtyBlock = padLeft(qty, 7) + '0'.repeat(21); // 28 chars total
      const isbn     = normalizeISSN(getCell(row, 'isbn'));

      let d1 = 'D1'
        + orderNumStr                      // [2:17]
        + itemRef                          // [17:42]  (25)
        + lineNum                          // [42:47]   (5)
        + '   '                            // [47:50]   (3)
        + '0'.repeat(18)                   // [50:68]  (18)
        + pad('', 78)                      // [68:146] (78)
        + qtyBlock                         // [146:174] (28)
        + ' ' + padLeft(0, 9) + '  '      // [174:186] (12) — price = 0 (not in source)
        + pad('', 40)                      // [186:226] (40)
        + pad(isbn, 13)                    // [226:239] (13)
        + pad('', 27);                     // [239:266] (27)

      lines.push(d1.substring(0, 266).padEnd(266));
      recordCount++;
      lineCount++;
    });

    orderCount++;
  });


  // ── FILE FOOTER ────────────────────────────────────────────
  // recordCount covers H1 + H2 + H3 + D1 lines only.
  // The receiving system does NOT count $$HDR or $$EOF themselves.
  // Format: $$EOF + senderCode + 2 spaces + fileId + 3 spaces + timestamp + recordCount(7)
  const totalRecordsStr = padLeft(recordCount, 7);
  lines.push(`$$EOF${senderCode}  ${fileIdStr}   ${ts}${totalRecordsStr}`);


  // ── ASSEMBLE OUTPUT ────────────────────────────────────────
  // EDI files use CRLF (\r\n) line endings and ASCII encoding.
  ediOutput = lines.join('\r\n') + '\r\n';


  // ── UPDATE STATS BAR ──────────────────────────────────────
  document.getElementById('statOrders').textContent  = orderCount;
  document.getElementById('statLines').textContent   = lineCount;
  document.getElementById('statRecords').textContent = recordCount;
  document.getElementById('statsRow').style.display  = 'flex';

  // ── RENDER PREVIEW & ENABLE ACTIONS ───────────────────────
  renderPreview(lines);
  document.getElementById('clearBtn').style.display    = '';
  document.getElementById('downloadBtn').style.display = '';

  // Switch to the EDI Preview tab automatically
  switchTab('preview', document.querySelectorAll('.tab-btn')[1]);
}


/* ============================================================
   EDI PREVIEW RENDERER
   Injects the generated lines into the preview pane with
   syntax highlighting by record type.
   ============================================================ */

/**
 * CSS class assigned to each EDI record type for colour coding.
 * Keys are line prefixes; values are class names defined in style.css.
 * @type {Object.<string, string>}
 */
const RECORD_COLOUR_MAP = {
  '$$HDR': 'hdr',
  '$$EOF': 'ftr',
  'H1':    'h1',
  'H2':    'h2',
  'H3':    'h3',
  'D1':    'd1',
};

/**
 * Renders the array of EDI lines into the preview pane with
 * per-record-type colour coding.
 *
 * @param {string[]} lines - Array of raw EDI line strings
 */
function renderPreview(lines) {
  const container = document.getElementById('previewOutput');
  const noMsg     = document.getElementById('noPreviewMsg');

  noMsg.style.display      = 'none';
  container.style.display  = 'block';

  const div = document.createElement('div');
  div.className = 'output-preview';

  div.innerHTML = lines.map(line => {
    // Determine which colour class to apply based on line prefix
    let cls = '';
    for (const [prefix, className] of Object.entries(RECORD_COLOUR_MAP)) {
      if (line.startsWith(prefix)) { cls = className; break; }
    }
    // Escape HTML entities to prevent XSS from spreadsheet content
    const escaped = line
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;');
    return cls ? `<span class="${cls}">${escaped}</span>` : escaped;
  }).join('\n');

  container.innerHTML = '';
  container.appendChild(div);
}


/* ============================================================
   COPY & DOWNLOAD ACTIONS
   ============================================================ */

/**
 * Resets the application to its initial state.
 * Clears imported data, EDI output, stats, preview and upload message.
 */
function clearData() {
  // Reset state variables
  rawData   = [];
  headers   = [];
  columnMap = {};
  ediOutput = '';

  // Reset file input so the same file can be re-uploaded if needed
  document.getElementById('fileInput').value = '';

  // Clear upload message
  document.getElementById('uploadMsg').innerHTML = '';

  // Hide and clear the data table
  document.getElementById('tableWrapper').style.display = 'none';
  document.getElementById('tableWrapper').innerHTML     = '';
  document.getElementById('noDataMsg').style.display    = '';

  // Hide and clear the EDI preview
  document.getElementById('previewOutput').style.display = 'none';
  document.getElementById('previewOutput').innerHTML     = '';
  document.getElementById('noPreviewMsg').style.display  = '';

  // Hide the mapping panel
  document.getElementById('mappingPanel').style.display = 'none';
  document.getElementById('mappingGrid').innerHTML      = '';

  // Collapse the mapping toggle in case it was open
  document.getElementById('mappingToggle').classList.remove('open');
  document.getElementById('mappingBody').classList.remove('open');

  // Hide stats bar and action buttons
  document.getElementById('statsRow').style.display    = 'none';
  document.getElementById('clearBtn').style.display    = 'none';
  document.getElementById('downloadBtn').style.display = 'none';

  // Re-initialise the order seed display
  initOrderSeed();

  // Switch back to the Source Data tab
  switchTab('data', document.querySelectorAll('.tab-btn')[0]);
}

/**
 * Triggers a browser download of the EDI file.
 * Filename format: {prefix}.{fileId}_{HHMM}_({DD-MM-YY}).txt
 * Example:         T1.0027816_1423_(24-02-26).txt
 */
function downloadFile() {
  if (!ediOutput) return;

  const fileId = document.getElementById('fileId').value.trim();
  const prefix = document.getElementById('filePrefix').value.trim();
  const now    = new Date();

  const dateStr = [
    String(now.getDate()).padStart(2, '0'),
    String(now.getMonth() + 1).padStart(2, '0'),
    String(now.getFullYear()).slice(2),
  ].join('-');

  const timeStr = [
    String(now.getHours()).padStart(2, '0'),
    String(now.getMinutes()).padStart(2, '0'),
  ].join('');

  const filename = `${prefix}.${fileId}_${timeStr}_(${dateStr}).txt`;

  const blob = new Blob([ediOutput], { type: 'text/plain;charset=ascii' });
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement('a');
  a.href     = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
}


/* ============================================================
   TAB SWITCHING
   ============================================================ */

/**
 * Switches the visible content tab between 'data' and 'preview'.
 *
 * @param {string}      tab - 'data' or 'preview'
 * @param {HTMLElement} btn - The tab button element that was clicked
 */
function switchTab(tab, btn) {
  document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
  btn.classList.add('active');
  document.getElementById('tabData').style.display    = tab === 'data'    ? '' : 'none';
  document.getElementById('tabPreview').style.display = tab === 'preview' ? '' : 'none';
}


/* ============================================================
   MESSAGE HELPER
   ============================================================ */

/**
 * Injects an inline status message into a container element.
 *
 * @param {string} containerId - ID of the target DOM element
 * @param {string} msg         - HTML message string (may include icons)
 * @param {string} type        - 'success' | 'error' | 'info'
 */
function showMessage(containerId, msg, type) {
  const el = document.getElementById(containerId);
  el.innerHTML = `<div class="message ${type}" style="margin-top:10px">${msg}</div>`;
}

/**
 * Renders a structured column-mismatch error with a per-column diff table.
 *
 * @param {string} containerId       - ID of the target DOM element
 * @param {{ message: string, details: string[] }} err - Validation error object
 * @param {string} templateFilename  - Name of the expected template file
 */
function showColumnError(containerId, err, templateFilename) {
  const rows = err.details.map(d => {
    // Highlight the differing parts in each detail line
    return `<li class="col-error-item">${d.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')}</li>`;
  }).join('');

  const html = `
    <div class="message error col-error-box" style="margin-top:10px">
      <div class="col-error-header">
        <i class="fa-solid fa-triangle-exclamation"></i>
        <strong>Column mismatch — file rejected</strong>
      </div>
      <p class="col-error-desc">${err.message} Please use <strong>${templateFilename}</strong> as your template without changing any column names or their order.</p>
      <ul class="col-error-list">${rows}</ul>
    </div>`;

  document.getElementById(containerId).innerHTML = html;
}


/* ============================================================
   INITIALISATION
   Runs once on page load to pre-populate the order number seed
   field so it is visible before the user clicks Generate.
   ============================================================ */

function initOrderSeed() {
  const n  = new Date();
  const yy = String(n.getFullYear()).slice(2);
  const mo = String(n.getMonth() + 1).padStart(2, '0');
  const dd = String(n.getDate()).padStart(2, '0');
  const hh = String(n.getHours()).padStart(2, '0');
  const mi = String(n.getMinutes()).padStart(2, '0');
  document.getElementById('orderNumStart').value = `3${yy}${mo}${dd}${hh}${mi}`;
}

// Run on page load
initOrderSeed();

/* ============================================================
   XML METADATA GENERATOR
   Accepts a separate Excel file containing journal/book metadata
   and generates one XML file per row, bundled as metadata.zip.

   Expected columns (case-insensitive, flexible matching):
     ISSN, Title, Trim Height, Trim Width, Spine Size,
     Paper Type, Binding Style, Page Extent, Lamination

   Output filename per XML: {ISSN}.xml
   Bundle filename:         metadata.zip
   ============================================================ */

/** Rows loaded from the metadata Excel file. */
let xmlRawData = [];

/** Column headers from the metadata Excel file. */
let xmlHeaders = [];

/** Maps each metadata field key to a column index (-1 = unmapped). */
let xmlColumnMap = {};

/**
 * Metadata field definitions.
 * `aliases` lists alternative header names matched case-insensitively.
 */
const XML_FIELDS = [
  { key: 'issn',         label: 'ISSN',          aliases: ['issn'] },
  { key: 'title',        label: 'Title',          aliases: ['title'] },
  { key: 'trimHeight',   label: 'Trim Height',    aliases: ['trim height', 'trimheight', 'height'] },
  { key: 'trimWidth',    label: 'Trim Width',     aliases: ['trim width', 'trimwidth', 'width'] },
  { key: 'spineSize',    label: 'Spine Size',     aliases: ['spine size', 'spinesize', 'spine'] },
  { key: 'paperType',    label: 'Paper Type',     aliases: ['paper type', 'papertype', 'paper'] },
  { key: 'bindingStyle', label: 'Binding Style',  aliases: ['binding style', 'bindingstyle', 'binding'] },
  { key: 'pageExtent',   label: 'Page Extent',    aliases: ['page extent', 'pageextent', 'pages', 'extent'] },
  { key: 'lamination',   label: 'Lamination',     aliases: ['lamination'] },
];

// ── Dropzone wiring ────────────────────────────────────────────────────────
const xmlDropzone  = document.getElementById('xmlDropzone');
const xmlFileInput = document.getElementById('xmlFileInput');

xmlDropzone.addEventListener('click', () => xmlFileInput.click());

xmlDropzone.addEventListener('dragover', e => {
  e.preventDefault();
  xmlDropzone.classList.add('drag-over');
});

xmlDropzone.addEventListener('dragleave', () => {
  xmlDropzone.classList.remove('drag-over');
});

xmlDropzone.addEventListener('drop', e => {
  e.preventDefault();
  xmlDropzone.classList.remove('drag-over');
  if (e.dataTransfer.files[0]) processXMLFile(e.dataTransfer.files[0]);
});

xmlFileInput.addEventListener('change', () => {
  if (xmlFileInput.files[0]) processXMLFile(xmlFileInput.files[0]);
});

/**
 * Parses the metadata Excel file and auto-maps columns.
 * @param {File} file
 */
function processXMLFile(file) {
  const reader = new FileReader();

  reader.onload = e => {
    try {
      const wb   = XLSX.read(e.target.result, { type: 'binary' });
      const ws   = wb.Sheets[wb.SheetNames[0]];
      const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });

      if (data.length < 1) throw new Error('The file appears to be empty.');

      const fileHeaders = data[0].map(h => String(h ?? ''));

      // ── Strict column validation against the metadata template ───────────
      const validationError = validateColumns(fileHeaders, METADATA_TEMPLATE_COLUMNS, 'Metadata');
      if (validationError) {
        showColumnError('xmlUploadMsg', validationError, 'metadata_template.xlsx');
        return;
      }

      if (data.length < 2) throw new Error('No data rows found in the file.');

      xmlHeaders = fileHeaders;
      xmlRawData = data.slice(1).filter(row => row.some(c => c !== null && c !== ''));

      // Columns are validated — map by position directly
      xmlColumnMap = {};
      XML_FIELDS.forEach(field => {
        const idx = xmlHeaders.findIndex(h => {
          const norm = h.trim().toLowerCase();
          return field.aliases.includes(norm);
        });
        xmlColumnMap[field.key] = idx;
      });

      showMessage(
        'xmlUploadMsg',
        `<i class="fa-solid fa-circle-check"></i> Loaded ${xmlRawData.length} row${xmlRawData.length !== 1 ? 's' : ''} from "${file.name}"`,
        'success'
      );

    } catch (err) {
      showMessage(
        'xmlUploadMsg',
        `<i class="fa-solid fa-circle-xmark"></i> Error: ${err.message}`,
        'error'
      );
    }
  };

  reader.readAsBinaryString(file);
}

/**
 * Retrieves a cell value from a metadata row by field key.
 * @param {Array}  row
 * @param {string} key
 * @return {string}
 */
function getXMLCell(row, key) {
  const idx = xmlColumnMap[key];
  if (idx === undefined || idx < 0) return '';
  const val = row[idx];
  if (val === null || val === undefined) return '';
  return String(val).trim();
}

/**
 * Escapes a string for safe embedding in XML text content.
 * @param {string} str
 * @return {string}
 */
function xmlEscape(str) {
  return str
    .replace(/&/g,  '&amp;')
    .replace(/</g,  '&lt;')
    .replace(/>/g,  '&gt;')
    .replace(/"/g,  '&quot;')
    .replace(/'/g,  '&apos;');
}

/**
 * Builds an XML string for a single metadata row.
 * @param {Array} row
 * @return {{ xml: string, issn: string }}
 */
function buildXML(row) {
  const issn         = getXMLCell(row, 'issn');
  const title        = getXMLCell(row, 'title');
  const trimHeight   = getXMLCell(row, 'trimHeight');
  const trimWidth    = getXMLCell(row, 'trimWidth');
  const spineSize    = getXMLCell(row, 'spineSize');
  const paperType    = getXMLCell(row, 'paperType');
  const bindingStyle = getXMLCell(row, 'bindingStyle');
  const pageExtent   = getXMLCell(row, 'pageExtent');
  const lamination   = getXMLCell(row, 'lamination');

  const xml =
`<?xml version="1.0" encoding="UTF-8"?>
<book>
    <basic_info>
        <issn>${xmlEscape(issn)}</issn>
        <title>${xmlEscape(title)}</title>
    </basic_info>
    <specifications>
        <dimensions>
            <trim_height>${xmlEscape(trimHeight)}</trim_height>
            <trim_width>${xmlEscape(trimWidth)}</trim_width>
            <spine_size>${xmlEscape(spineSize)}</spine_size>
        </dimensions>
        <materials>
            <paper_type>${xmlEscape(paperType)}</paper_type>
            <binding_style>${xmlEscape(bindingStyle)}</binding_style>
            <lamination>${xmlEscape(lamination)}</lamination>
        </materials>
        <page_extent>${xmlEscape(pageExtent)}</page_extent>
    </specifications>
</book>`;

  return { xml, issn };
}

/**
 * Main entry point for the XML metadata feature.
 * Generates one XML per row and packages them into metadata.zip.
 */
async function generateXMLMetadata() {
  if (!xmlRawData.length) {
    alert('Please upload a metadata Excel file first.');
    return;
  }

  const btn = document.querySelector('[onclick="generateXMLMetadata()"]');
  const origText = btn.innerHTML;
  btn.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Building ZIP…';
  btn.disabled = true;

  try {
    const zip = new JSZip();
    let count = 0;
    const skipped = [];

    xmlRawData.forEach((row, idx) => {
      const { xml, issn } = buildXML(row);
      if (!issn) {
        skipped.push(idx + 1);
        return;
      }
      // Sanitise ISSN for use as filename (digits only, no punctuation)
      const issnSafe = issn.replace(/[^a-zA-Z0-9_\-]/g, '_');
      zip.file(`${issnSafe}.xml`, xml);
      count++;
    });

    if (count === 0) {
      throw new Error('No rows with a valid ISSN found — no XML files were generated.');
    }

    const blob = await zip.generateAsync({ type: 'blob' });
    const url  = URL.createObjectURL(blob);
    const a    = document.createElement('a');
    a.href     = url;
    a.download = 'metadata.zip';
    a.click();
    URL.revokeObjectURL(url);

    let msg = `<i class="fa-solid fa-circle-check"></i> Generated ${count} XML file${count !== 1 ? 's' : ''} → <strong>metadata.zip</strong>`;
    if (skipped.length) msg += ` <span style="color:var(--warn)">(${skipped.length} row${skipped.length !== 1 ? 's' : ''} skipped — no ISSN)</span>`;
    showMessage('xmlUploadMsg', msg, 'success');

  } catch (err) {
    showMessage('xmlUploadMsg', `<i class="fa-solid fa-circle-xmark"></i> Error: ${err.message}`, 'error');
  } finally {
    btn.innerHTML = origText;
    btn.disabled  = false;
  }
}