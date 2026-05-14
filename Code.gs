const TPSO_CONFIG = {
  apiUrl: 'https://index-api.tpso.go.th/OpenApi/CmiPrice/Month',
  defaultSheetName: 'CmiPrice',
  defaultProvinceType: 14,
  menuName: 'TPSO API',
  sidebarFile: 'Sidebar',
  sidebarTitle: 'TPSO CMI Price',
  timezone: 'Asia/Bangkok',
  buddhistYearOffset: 543,
  metadataStartRow: 1,
  headerRow: 4,
  dataStartRow: 5,
  importSettingsKey: 'TPSO_CMI_PRICE_IMPORT_SETTINGS',
  maxSheetNameLength: 100,
  responseSnippetLength: 300
};

const CMI_PRICE_PREFERRED_HEADERS = [
  'id',
  'type',
  'typeName',
  'commodityCode',
  'commodityNameTH',
  'unitName',
  'curMonth',
  'curYear',
  'priceCur',
  'priceVAT',
  'createdAt'
];

/**
 * Adds a custom menu for opening the TPSO API sidebar in Google Sheets.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu(TPSO_CONFIG.menuName)
    .addItem('เปิด Sidebar ดึงข้อมูล', 'showTpsoSidebar')
    .addSeparator()
    .addItem('ดึงข้อมูลเดือนย้อนหลัง 1 เดือน', 'fetchCmiPricePreviousMonth')
    .addItem('สร้าง Trigger รายเดือน', 'showCreateMonthlyTriggerResult')
    .addToUi();
}

/**
 * Opens the sidebar form for editing API request values before fetching data.
 */
function showTpsoSidebar() {
  const template = HtmlService.createTemplateFromFile(TPSO_CONFIG.sidebarFile);
  template.defaults = getDefaultSidebarValues();

  SpreadsheetApp.getUi().showSidebar(
    template.evaluate().setTitle(TPSO_CONFIG.sidebarTitle)
  );
}

/**
 * Returns default form values for the sidebar.
 *
 * @return {Object} Sidebar defaults.
 */
function getDefaultSidebarValues() {
  const period = getPreviousMonthThaiYear_();
  const savedSettings = loadImportSettings_();

  return {
    thaiYear: period.thaiYear,
    month: period.month,
    type: savedSettings.type,
    sheetName: savedSettings.sheetName,
    overwriteExisting: savedSettings.overwriteExisting
  };
}

/**
 * Fetches CMI price data for the previous month and writes it to the default sheet.
 *
 * @return {Object} Import summary.
 */
function fetchCmiPricePreviousMonth() {
  const period = getPreviousMonthThaiYear_();
  const savedSettings = loadImportSettings_();

  return importCmiPrice_({
    thaiYear: period.thaiYear,
    month: period.month,
    type: savedSettings.type,
    sheetName: savedSettings.sheetName,
    overwriteExisting: savedSettings.overwriteExisting
  });
}

/**
 * Fetches CMI price data using values submitted from the sidebar.
 *
 * @param {Object} form Form values from Sidebar.html.
 * @return {Object} Import summary for rendering in the sidebar.
 */
function fetchCmiPriceFromSidebar(form) {
  return importCmiPrice_(form, { saveSettings: true });
}

/**
 * Creates a monthly trigger and shows the result in the Google Sheets UI.
 */
function showCreateMonthlyTriggerResult() {
  SpreadsheetApp.getUi().alert(createMonthlyTrigger());
}

/**
 * Creates a monthly trigger that imports the previous month on day 5 at 08:00-09:00.
 */
function createMonthlyTrigger() {
  const handlerFunction = 'fetchCmiPricePreviousMonth';
  const existingTrigger = ScriptApp.getProjectTriggers().some(function (trigger) {
    return trigger.getHandlerFunction() === handlerFunction;
  });

  if (existingTrigger) {
    return 'มี Trigger รายเดือนอยู่แล้ว';
  }

  ScriptApp.newTrigger(handlerFunction)
    .timeBased()
    .onMonthDay(5)
    .atHour(8)
    .create();

  return 'สร้าง Trigger รายเดือนสำเร็จ';
}

/**
 * Imports CMI price data from TPSO and writes it to the configured sheet.
 *
 * @param {Object} form User-provided request values.
 * @param {Object=} options Import options.
 * @return {Object} Import summary.
 */
function importCmiPrice_(form, options) {
  const request = normalizeRequest_(form);
  const importOptions = options || {};

  if (importOptions.saveSettings) {
    saveImportSettings_(request);
  }

  const payload = buildCmiPricePayload_(request);
  const data = fetchCmiPriceRows_(payload);

  writeCmiPriceToSheet_(data, request.sheetName, payload, request);

  return buildImportSummary_(request, data.length);
}

/**
 * Builds the TPSO API payload from a normalized request.
 *
 * @param {Object} request Normalized request values.
 * @return {Object} TPSO API payload.
 */
function buildCmiPricePayload_(request) {
  return {
    year: request.thaiYear,
    month: request.month,
    type: request.type
  };
}

/**
 * Calls the TPSO CMI price API and returns validated response rows.
 *
 * @param {Object} payload TPSO API payload.
 * @return {Object[]} API response rows.
 */
function fetchCmiPriceRows_(payload) {
  const response = UrlFetchApp.fetch(TPSO_CONFIG.apiUrl, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const statusCode = response.getResponseCode();
  const text = response.getContentText();

  assertSuccessfulResponse_(statusCode, text);
  return parseCmiPriceRows_(text);
}

/**
 * Throws a readable error when the TPSO API returns a non-successful status.
 *
 * @param {number} statusCode HTTP response code.
 * @param {string} responseText HTTP response body.
 */
function assertSuccessfulResponse_(statusCode, responseText) {
  if (statusCode < 200 || statusCode >= 300) {
    throw new Error('TPSO API error: HTTP ' + statusCode + '\n' + responseText);
  }
}

/**
 * Parses and validates TPSO API response rows.
 *
 * @param {string} responseText Raw JSON response body.
 * @return {Object[]} API response rows.
 */
function parseCmiPriceRows_(responseText) {
  let data;

  try {
    data = JSON.parse(responseText);
  } catch (error) {
    throw new Error(
      'TPSO API response is not valid JSON: ' + truncateText_(responseText)
    );
  }

  if (!Array.isArray(data)) {
    throw new Error('TPSO API response is not an array: ' + truncateText_(responseText));
  }

  return data;
}

/**
 * Truncates long text for readable Sidebar error messages.
 *
 * @param {string} text Text to truncate.
 * @return {string} Truncated text.
 */
function truncateText_(text) {
  const value = String(text || '');

  if (value.length <= TPSO_CONFIG.responseSnippetLength) {
    return value;
  }

  return value.slice(0, TPSO_CONFIG.responseSnippetLength) + '...';
}

/**
 * Builds a compact import summary for sidebar success messages.
 *
 * @param {Object} request Normalized request values.
 * @param {number} rowCount Number of imported rows.
 * @return {Object} Import summary.
 */
function buildImportSummary_(request, rowCount) {
  return {
    rowCount: rowCount,
    sheetName: request.sheetName,
    thaiYear: request.thaiYear,
    month: request.month,
    type: request.type,
    overwriteExisting: request.overwriteExisting
  };
}

/**
 * Saves the latest import settings for future scheduled imports.
 *
 * @param {Object} request Normalized import request.
 */
function saveImportSettings_(request) {
  const settings = {
    type: request.type,
    sheetName: request.sheetName,
    overwriteExisting: request.overwriteExisting
  };

  PropertiesService
    .getDocumentProperties()
    .setProperty(TPSO_CONFIG.importSettingsKey, JSON.stringify(settings));
}

/**
 * Loads saved import settings and falls back to defaults when none exist.
 *
 * @return {Object} Import settings.
 */
function loadImportSettings_() {
  const fallback = {
    type: TPSO_CONFIG.defaultProvinceType,
    sheetName: TPSO_CONFIG.defaultSheetName,
    overwriteExisting: true
  };
  const stored = PropertiesService
    .getDocumentProperties()
    .getProperty(TPSO_CONFIG.importSettingsKey);

  if (!stored) {
    return fallback;
  }

  try {
    const parsed = JSON.parse(stored);
    const settings = {
      type: toInteger_(parsed.type) || fallback.type,
      sheetName: String(parsed.sheetName || fallback.sheetName).trim(),
      overwriteExisting: toBoolean_(parsed.overwriteExisting, fallback.overwriteExisting)
    };

    if (!Number.isInteger(settings.type) || settings.type <= 0) {
      return fallback;
    }

    validateSheetName_(settings.sheetName);
    return settings;
  } catch (error) {
    return fallback;
  }
}

/**
 * Normalizes and validates sidebar request values.
 *
 * @param {Object} form Form values.
 * @return {Object} Validated request values.
 */
function normalizeRequest_(form) {
  const source = form || {};
  const request = {
    thaiYear: toInteger_(source.thaiYear || source.year),
    month: toInteger_(source.month),
    type: toInteger_(source.type),
    sheetName: String(source.sheetName || TPSO_CONFIG.defaultSheetName).trim(),
    overwriteExisting: toBoolean_(source.overwriteExisting, true)
  };

  validateRequest_(request);
  return request;
}

/**
 * Converts a form value to an integer or NaN when invalid.
 *
 * @param {*} value Form value.
 * @return {number} Parsed integer.
 */
function toInteger_(value) {
  const parsed = Number(value);
  return Number.isInteger(parsed) ? parsed : NaN;
}

/**
 * Converts checkbox-like values to a boolean.
 *
 * @param {*} value Form value.
 * @param {boolean} defaultValue Value used when the form value is missing.
 * @return {boolean} Parsed boolean.
 */
function toBoolean_(value, defaultValue) {
  if (value === undefined || value === null || value === '') {
    return defaultValue;
  }

  if (typeof value === 'boolean') {
    return value;
  }

  return ['true', '1', 'yes', 'on'].indexOf(String(value).toLowerCase()) !== -1;
}

/**
 * Validates normalized import request values.
 *
 * @param {Object} request Normalized request values.
 */
function validateRequest_(request) {
  if (!Number.isInteger(request.thaiYear) || request.thaiYear < 2400 || request.thaiYear > 2700) {
    throw new Error('กรุณาระบุปี พ.ศ. ให้ถูกต้อง เช่น 2569');
  }

  if (!Number.isInteger(request.month) || request.month < 1 || request.month > 12) {
    throw new Error('กรุณาระบุเดือนเป็นตัวเลข 1-12');
  }

  if (!Number.isInteger(request.type) || request.type <= 0) {
    throw new Error('กรุณาระบุรหัสจังหวัด/type เป็นตัวเลขที่มากกว่า 0');
  }

  validateSheetName_(request.sheetName);
}

/**
 * Validates the destination Google Sheet name.
 *
 * @param {string} sheetName Destination sheet name.
 */
function validateSheetName_(sheetName) {
  if (!sheetName) {
    throw new Error('กรุณาระบุชื่อชีตปลายทาง');
  }

  if (sheetName.length > TPSO_CONFIG.maxSheetNameLength) {
    throw new Error('ชื่อชีตปลายทางต้องไม่เกิน ' + TPSO_CONFIG.maxSheetNameLength + ' ตัวอักษร');
  }

  if (/[\\\/\?\*\[\]:]/.test(sheetName)) {
    throw new Error('ชื่อชีตปลายทางห้ามมีอักขระ : \\ / ? * [ ]');
  }
}

/**
 * Writes API data to a spreadsheet sheet.
 *
 * @param {Object[]} data API response rows.
 * @param {string} sheetName Destination sheet name.
 * @param {Object} payload API request payload used for the import.
 * @param {Object} request Normalized import request.
 */
function writeCmiPriceToSheet_(data, sheetName, payload, request) {
  const sheet = getOrCreateSheet_(sheetName);
  const headers = collectHeaders_(data);

  if (request.overwriteExisting) {
    sheet.clearContents();
  }

  writeImportMetadata_(sheet, payload);

  if (headers.length === 0) {
    writeNoDataMessage_(sheet);
    return;
  }

  writeTable_(sheet, headers, data);
}

/**
 * Returns an existing destination sheet or creates it when missing.
 *
 * @param {string} sheetName Destination sheet name.
 * @return {GoogleAppsScript.Spreadsheet.Sheet} Destination sheet.
 */
function getOrCreateSheet_(sheetName) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  return spreadsheet.getSheetByName(sheetName) || spreadsheet.insertSheet(sheetName);
}

/**
 * Writes the API payload values used for the import above the data table.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet Destination sheet.
 * @param {Object} payload API payload.
 */
function writeImportMetadata_(sheet, payload) {
  sheet
    .getRange(TPSO_CONFIG.metadataStartRow, 1, 1, 3)
    .setValues([['year', 'month', 'type']]);
  sheet
    .getRange(TPSO_CONFIG.metadataStartRow + 1, 1, 1, 3)
    .setValues([[payload.year, payload.month, payload.type]]);
}

/**
 * Writes a no-data message when the API response is empty.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet Destination sheet.
 */
function writeNoDataMessage_(sheet) {
  sheet
    .getRange(TPSO_CONFIG.headerRow, 1)
    .setValue('ไม่พบข้อมูลจาก API สำหรับเงื่อนไขที่เลือก');
}

/**
 * Writes API rows as a table to the destination sheet.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet Destination sheet.
 * @param {string[]} headers Table headers.
 * @param {Object[]} data API response rows.
 */
function writeTable_(sheet, headers, data) {
  const rows = data.map(function (item) {
    return headers.map(function (key) {
      return getCellValue_(item, key);
    });
  });

  sheet
    .getRange(TPSO_CONFIG.headerRow, 1, 1, headers.length)
    .setValues([headers]);
  sheet
    .getRange(TPSO_CONFIG.dataStartRow, 1, rows.length, headers.length)
    .setValues(rows);
  sheet.autoResizeColumns(1, headers.length);
}

/**
 * Safely returns a value for a sheet cell.
 *
 * @param {Object} item API response row.
 * @param {string} key Header key.
 * @return {*} Cell value.
 */
function getCellValue_(item, key) {
  return item[key] !== undefined && item[key] !== null ? item[key] : '';
}

/**
 * Builds a stable header list from all API response keys.
 *
 * @param {Object[]} data API response rows.
 * @return {string[]} Header names.
 */
function collectHeaders_(data) {
  const keys = collectResponseKeys_(data);
  const ordered = CMI_PRICE_PREFERRED_HEADERS.filter(function (key) {
    return keys[key];
  });
  const extra = Object.keys(keys)
    .filter(function (key) {
      return ordered.indexOf(key) === -1;
    })
    .sort();

  return ordered.concat(extra);
}

/**
 * Collects unique keys from all response rows.
 *
 * @param {Object[]} data API response rows.
 * @return {Object} Key lookup.
 */
function collectResponseKeys_(data) {
  return data.reduce(function (keys, item) {
    Object.keys(item || {}).forEach(function (key) {
      keys[key] = true;
    });
    return keys;
  }, {});
}

/**
 * Gets the previous month based on Thailand timezone and returns Thai Buddhist year.
 *
 * @return {Object} Previous month values.
 */
function getPreviousMonthThaiYear_() {
  const now = new Date();
  const currentYear = Number(Utilities.formatDate(now, TPSO_CONFIG.timezone, 'yyyy'));
  const currentMonth = Number(Utilities.formatDate(now, TPSO_CONFIG.timezone, 'M'));
  const period = subtractOneMonth_(currentYear, currentMonth);

  return {
    year: period.year,
    thaiYear: period.year + TPSO_CONFIG.buddhistYearOffset,
    month: period.month
  };
}

/**
 * Returns a year/month pair for one month before the supplied period.
 *
 * @param {number} year Gregorian year.
 * @param {number} month Month number, 1-12.
 * @return {Object} Previous year/month pair.
 */
function subtractOneMonth_(year, month) {
  if (month === 1) {
    return {
      year: year - 1,
      month: 12
    };
  }

  return {
    year: year,
    month: month - 1
  };
}
