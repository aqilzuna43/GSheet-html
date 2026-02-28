const SHEET_NAME = 'Schedule';
const REQUIRED_HEADERS = [
  'ID',
  'Title',
  'Start Date',
  'End Date',
  'Owner',
  'Department',
  'Status',
  'Description',
  'Tags',
];
const CONFIG_SHEET_NAME = 'Config';
const AUDIT_SHEET_NAME = 'Change_Log';

function doGet() {
  let template;
  try {
    template = HtmlService.createTemplateFromFile('Index');
  } catch (err) {
    return HtmlService.createHtmlOutput(
      '<div style="font-family: sans-serif; padding: 20px;">' +
      '<h3>Setup Error</h3>' +
      '<p>Could not find <b>Index.html</b>. Please add it to your Apps Script project.</p>' +
      '</div>'
    ).setTitle('Configuration Error');
  }

  const title = getDashboardTitle_('standard_title', [SHEET_NAME], 'Visual Schedule Dashboard');
  const source = getGeneralSourceConfig_();
  template.pageTitle = title;
  template.defaultSourceSpreadsheetId = source.spreadsheetId || '';
  template.defaultSourceSheetName = source.sheetName || SHEET_NAME;

  return template.evaluate()
    .setTitle(title)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function onOpen() {
  if (typeof setupScheduleSheet !== 'function') return;
  SpreadsheetApp.getUi()
    .createMenu('Dashboard Tools')
    .addItem('Setup / Repair Standard Sheet', 'setupScheduleSheet')
    .addItem('Validate Standard Data', 'validateScheduleData')
    .addToUi();
}

function onEdit(e) {
  if (typeof onEditScheduleValidation_ === 'function') {
    onEditScheduleValidation_(e);
  }
  logStatusDateChange_(e);
}

function getScheduleItems(forceRefresh, sourcePrefs) {
  const source = resolveGeneralSource_(sourcePrefs);
  const sheet = source.sheet;
  if (!sheet) {
    throw new Error('Source sheet not found. Check spreadsheet ID and sheet name.');
  }

  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) return [];

  const headers = values[0].map((h) => String(h).trim().toLowerCase());
  const headerIndex = REQUIRED_HEADERS.reduce((acc, header) => {
    const idx = headers.indexOf(header.toLowerCase());
    if (idx === -1) {
      throw new Error(`Missing required column: "${header}". Please check row 1 headers.`);
    }
    acc[header] = idx;
    return acc;
  }, {});

  const items = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const title = String(row[headerIndex['Title']] || '').trim();
    if (!title) continue;

    const id = String(row[headerIndex['ID']] || '').trim() || `ROW-${r + 1}`;
    let startDate = normalizeDate_(row[headerIndex['Start Date']]);
    let endDate = normalizeDate_(row[headerIndex['End Date']]);
    if (startDate && !endDate) endDate = startDate;
    if (!startDate && endDate) startDate = endDate;
    if (!startDate || !endDate) continue;
    if (endDate < startDate) endDate = startDate;

    items.push({
      id: id,
      title: title,
      startDate: toIsoDate_(startDate),
      endDate: toIsoDate_(endDate),
      owner: String(row[headerIndex['Owner']] || '').trim(),
      department: String(row[headerIndex['Department']] || '').trim(),
      status: String(row[headerIndex['Status']] || '').trim(),
      description: String(row[headerIndex['Description']] || '').trim(),
      tags: String(row[headerIndex['Tags']] || '').trim(),
    });
  }

  return items;
}

function getGeneralSourceDefaults() {
  const source = getGeneralSourceConfig_();
  return {
    spreadsheetId: source.spreadsheetId || '',
    sheetName: source.sheetName || SHEET_NAME,
  };
}

function resolveGeneralSource_(sourcePrefs) {
  const config = getGeneralSourceConfig_();
  const pref = sourcePrefs && typeof sourcePrefs === 'object' ? sourcePrefs : {};
  const requestedSpreadsheetId = normalizeSpreadsheetIdInput_(String(pref.spreadsheetId || '').trim());
  const requestedSheetName = String(pref.sheetName || '').trim();

  const spreadsheetId = requestedSpreadsheetId || normalizeSpreadsheetIdInput_(config.spreadsheetId || '') || '';
  const sheetName = requestedSheetName || config.sheetName || SHEET_NAME;

  let spreadsheet;
  if (spreadsheetId) {
    try {
      spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    } catch (err) {
      throw new Error(
        'Cannot open source spreadsheet. Use a valid Spreadsheet ID (or Google Sheets URL), ensure access is shared, or click Reset Source.'
      );
    }
  } else {
    spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  }
  if (!spreadsheet) {
    throw new Error('No spreadsheet context found.');
  }

  const sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error(`Sheet "${sheetName}" not found in spreadsheet "${spreadsheet.getName()}".`);
  }

  return {
    spreadsheet: spreadsheet,
    sheet: sheet,
    spreadsheetId: spreadsheet.getId(),
    sheetName: sheetName,
  };
}

function getGeneralSourceConfig_() {
  const values = getConfigKeyValues_();
  const spreadsheetId = normalizeSpreadsheetIdInput_(String(values['standard_source_spreadsheet_id'] || '').trim());
  const sheetName = String(values['standard_source_sheet_name'] || '').trim();
  return {
    spreadsheetId: spreadsheetId,
    sheetName: sheetName || SHEET_NAME,
  };
}

function normalizeSpreadsheetIdInput_(input) {
  const raw = String(input || '').trim();
  if (!raw) return '';
  const match = raw.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/i);
  const candidate = match && match[1] ? match[1] : raw;
  return /^[a-zA-Z0-9-_]{20,}$/.test(candidate) ? candidate : raw;
}

function getConfigKeyValues_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = ss.getSheetByName(CONFIG_SHEET_NAME);
  if (!config) return {};
  const values = config.getDataRange().getDisplayValues();
  const out = {};
  for (let r = 0; r < values.length; r++) {
    const key = String(values[r][0] || '').trim().toLowerCase();
    const value = String(values[r][1] || '').trim();
    if (key) out[key] = value;
  }
  return out;
}

function normalizeDate_(value) {
  if (!value) return null;
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return value;
  }
  if (typeof value === 'number') {
    return new Date(Math.round((value - 25569) * 86400 * 1000));
  }
  if (typeof value === 'string' && value.trim() !== '') {
    const str = value.trim();
    let parsed = new Date(str);
    if (!isNaN(parsed.getTime())) return parsed;
    const parts = str.split(/[-/]/);
    if (parts.length === 3) {
      let d = parseInt(parts[0], 10);
      let m = parseInt(parts[1], 10) - 1;
      let y = parseInt(parts[2], 10);
      if (y < 100) y += 2000;
      parsed = new Date(y, m, d);
      if (!isNaN(parsed.getTime())) return parsed;
    }
  }
  return null;
}

function toIsoDate_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function getDashboardTitle_(titleKey, sheetNames, fallbackTitle) {
  const configTitle = getDashboardTitleFromConfig_(titleKey);
  if (configTitle) return configTitle;
  return getDashboardTitleFromLegacySheetCells_(sheetNames, fallbackTitle);
}

function getDashboardTitleFromConfig_(titleKey) {
  const values = getConfigKeyValues_();
  return String(values[String(titleKey || '').toLowerCase()] || '').trim();
}

function getDashboardTitleFromLegacySheetCells_(sheetNames, fallbackTitle) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  for (let i = 0; i < sheetNames.length; i++) {
    const sheet = ss.getSheetByName(sheetNames[i]);
    if (!sheet) continue;
    const marker = String(sheet.getRange('A1').getDisplayValue() || '').trim().toLowerCase();
    const value = String(sheet.getRange('B1').getDisplayValue() || '').trim();
    if (marker === 'title:' && value) return value;
  }
  return fallbackTitle;
}

function logStatusDateChange_(e) {
  const range = e && e.range ? e.range : null;
  if (!range) return;
  if (range.getNumRows() !== 1 || range.getNumColumns() !== 1) return;

  const sheet = range.getSheet();
  if (sheet.getName() !== SHEET_NAME) return;

  const row = range.getRow();
  const col = range.getColumn();
  const headerRow = 1;
  if (row <= headerRow) return;

  const headerRaw = String(sheet.getRange(headerRow, col).getDisplayValue() || '').trim();
  const header = headerRaw.toLowerCase();
  const watched = ['status', 'start date', 'end date'];
  if (watched.indexOf(header) === -1) return;

  const oldValue = String((e && Object.prototype.hasOwnProperty.call(e, 'oldValue')) ? e.oldValue : '').trim();
  const newValue = String(range.getDisplayValue() || '').trim();
  if (oldValue === newValue) return;

  const idCol = findHeaderColumn_(sheet, headerRow, ['id']);
  const titleCol = findHeaderColumn_(sheet, headerRow, ['title']);
  const rowId = idCol > 0 ? String(sheet.getRange(row, idCol).getDisplayValue() || '').trim() : '';
  const rowTitle = titleCol > 0 ? String(sheet.getRange(row, titleCol).getDisplayValue() || '').trim() : '';

  const audit = getOrCreateAuditSheet_();
  audit.appendRow([
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'),
    Session.getEffectiveUser().getEmail() || '',
    sheet.getName(),
    row,
    rowId,
    rowTitle,
    headerRaw,
    oldValue,
    newValue,
  ]);
}

function findHeaderColumn_(sheet, headerRow, namesLower) {
  const headers = sheet.getRange(headerRow, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  for (let i = 0; i < headers.length; i++) {
    const label = String(headers[i] || '').trim().toLowerCase();
    if (namesLower.indexOf(label) !== -1) return i + 1;
  }
  return -1;
}

function getOrCreateAuditSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(AUDIT_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(AUDIT_SHEET_NAME);
  }
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, 9).setValues([[
      'Timestamp',
      'User',
      'Sheet',
      'Row',
      'Record ID',
      'Record Title',
      'Field',
      'Old Value',
      'New Value',
    ]]);
    sheet.getRange(1, 1, 1, 9).setFontWeight('bold').setBackground('#f3f4f6');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 165);
    sheet.setColumnWidth(2, 190);
    sheet.setColumnWidth(3, 140);
    sheet.setColumnWidth(4, 50);
    sheet.setColumnWidth(5, 110);
    sheet.setColumnWidth(6, 280);
    sheet.setColumnWidth(7, 120);
    sheet.setColumnWidth(8, 200);
    sheet.setColumnWidth(9, 200);
  }
  return sheet;
}
