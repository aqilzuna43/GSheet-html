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
const MECOE_SHEET_CANDIDATES = ['ME_COE_Schedule', 'ME_COE', 'Schedule_ME_COE'];
const MECOE_DATE_HEADERS = [
  'ID',
  'Phase',
  'Deliverable / Document',
  'Site',
  'Docs',
  'Priority',
  'Status',
  'Owner',
  'Start Date',
  'End Date',
  'Milestone Date',
  'Notes / Actions',
];

function doGet(e) {
  let app = e && e.parameter && e.parameter.app ? String(e.parameter.app).toLowerCase() : '';

  // Auto-detect the app if no URL parameter is provided
  if (!app) {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      if (ss) {
        const sheets = ss.getSheets().map(s => s.getName());
        const hasMECOE = MECOE_SHEET_CANDIDATES.some(name => sheets.indexOf(name) !== -1);
        const hasStandard = sheets.indexOf(SHEET_NAME) !== -1;
        
        // If ME COE sheet exists but standard doesn't, automatically load ME COE
        if (hasMECOE && !hasStandard) {
          app = 'mecoe';
        }
      }
    } catch (err) {
      // Ignore if not container-bound
    }
  }

  let isMECOE = app === 'mecoe';
  let templateName = isMECOE ? 'IndexMECOE' : 'Index';
  let template;

  try {
    template = HtmlService.createTemplateFromFile(templateName);
  } catch (err) {
    // Fallback: If the requested template is missing, try the other one
    isMECOE = !isMECOE;
    templateName = isMECOE ? 'IndexMECOE' : 'Index';
    try {
      template = HtmlService.createTemplateFromFile(templateName);
    } catch (fallbackErr) {
      // If neither HTML file is found, return a friendly user message
      return HtmlService.createHtmlOutput(
        '<div style="font-family: sans-serif; padding: 20px;">' +
        '<h3>Setup Error</h3>' +
        '<p>Could not find the required HTML files. Please make sure you have added either <b>Index.html</b> or <b>IndexMECOE.html</b> to your Apps Script project.</p>' +
        '</div>'
      ).setTitle('Configuration Error');
    }
  }

  const fallbackTitle = isMECOE ? 'ME COE Timeline Dashboard' : 'Visual Schedule Dashboard';
  const sheetNames = isMECOE ? MECOE_SHEET_CANDIDATES : [SHEET_NAME];
  const titleKey = isMECOE ? 'mecoe_title' : 'standard_title';
  const title = getDashboardTitle_(titleKey, sheetNames, fallbackTitle);

  template.pageTitle = title;

  return template.evaluate()
    .setTitle(title)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('Dashboard Tools');
  let hasItems = false;

  if (typeof setupScheduleSheet === 'function') {
    menu
      .addItem('Setup / Repair Standard Sheet', 'setupScheduleSheet')
      .addItem('Validate Standard Data', 'validateScheduleData');
    hasItems = true;
  }

  if (typeof setupMECOEScheduleSheet === 'function') {
    if (hasItems) menu.addSeparator();
    menu
      .addItem('Setup / Repair ME COE Sheet', 'setupMECOEScheduleSheet')
      .addItem('Migrate Legacy ME COE (ME_COE)', 'runMECOELegacyMigrationDefault');
    hasItems = true;
  }

  if (hasItems) menu.addToUi();
}

function onEdit(e) {
  if (typeof onEditScheduleValidation_ === 'function') {
    onEditScheduleValidation_(e);
  }
  logStatusDateChange_(e);
}

function runMECOELegacyMigrationDefault() {
  if (typeof migrateLegacyMECOEToDateModel !== 'function') {
    SpreadsheetApp.getUi().alert('mecoe_setup.gs is not loaded.');
    return;
  }
  migrateLegacyMECOEToDateModel('ME_COE');
}

function getScheduleItems(forceRefresh) {
  // Cache removed to ensure browser refreshes load live sheet edits instantly.
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) {
    throw new Error(`Sheet "${SHEET_NAME}" not found.`);
  }

  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) return [];

  // Match headers case-insensitively to prevent typo failures
  const headers = values[0].map((h) => String(h).trim().toLowerCase());
  const headerIndex = REQUIRED_HEADERS.reduce((acc, header) => {
    const search = header.toLowerCase();
    const idx = headers.indexOf(search);
    if (idx === -1) {
      throw new Error(`Missing required column: "${header}". Please check your row 1 headers.`);
    }
    acc[header] = idx;
    return acc;
  }, {});

  const items = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const title = String(row[headerIndex['Title']] || '').trim();
    if (!title) continue; // Skip strictly if title is missing entirely

    // Fallback: Generate an ID if the user left it blank
    const id = String(row[headerIndex['ID']] || '').trim() || `ROW-${r + 1}`;
    
    let startDate = normalizeDate_(row[headerIndex['Start Date']]);
    let endDate = normalizeDate_(row[headerIndex['End Date']]);

    // Auto-heal missing or inverted dates to prevent frontend crash
    if (startDate && !endDate) endDate = startDate;
    if (!startDate && endDate) startDate = endDate;
    if (!startDate || !endDate) continue; // Still skip if both dates are completely missing

    // Auto-fix inverted dates instead of skipping the row
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

function normalizeDate_(value) {
  if (!value) return null;
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return value;
  }
  // Handle Excel/Sheets internal serial dates (if format gets stripped)
  if (typeof value === 'number') {
    return new Date(Math.round((value - 25569) * 86400 * 1000));
  }
  // Handle string dates robustly
  if (typeof value === 'string' && value.trim() !== '') {
    const str = value.trim();
    let parsed = new Date(str);
    if (!isNaN(parsed.getTime())) return parsed;
    
    // Try recovering DD/MM/YYYY or DD-MM-YYYY formats
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

function getMECOESchedule(forceRefresh) {
  // Cache removed to ensure live data loads properly
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = MECOE_SHEET_CANDIDATES
    .map((name) => ss.getSheetByName(name))
    .filter((s) => !!s)[0];
  if (!sheet) {
    throw new Error('ME COE sheet not found. Try one of: ' + MECOE_SHEET_CANDIDATES.join(', '));
  }

  const rawValues = sheet.getDataRange().getValues();
  const displayValues = sheet.getDataRange().getDisplayValues();
  if (displayValues.length <= 1) {
    return { timelineHeaders: [], timelineBuckets: [], items: [], sheetName: sheet.getName(), model: 'date' };
  }

  const headerRowIdx = findMECOEHeaderRow_(displayValues);
  if (headerRowIdx === -1) {
    throw new Error('Could not locate ME COE header row. Expected first columns: Phase, Deliverable / Document.');
  }

  const headers = displayValues[headerRowIdx].map((h) => String(h || '').trim());
  const isDateModel = isMECOEDateModel_(headers);
  let payload;
  
  if (isDateModel) {
    payload = parseMECOEDateModel_(sheet, rawValues, displayValues, headerRowIdx, headers);
  } else {
    payload = parseMECOELegacyModel_(sheet, displayValues, headerRowIdx, headers);
  }

  return payload;
}

function parseMECOEDateModel_(sheet, rawValues, displayValues, headerRowIdx, headers) {
  const tz = Session.getScriptTimeZone();
  const lowerHeaders = headers.map(h => h.toLowerCase());
  
  const idx = {};
  MECOE_DATE_HEADERS.forEach(expected => {
    idx[expected] = lowerHeaders.indexOf(expected.toLowerCase());
  });

  const items = [];
  const datePoints = [];
  for (let r = headerRowIdx + 1; r < displayValues.length; r++) {
    const rawRow = rawValues[r] || [];
    const row = displayValues[r] || [];
    if (!rowHasContent_(row)) continue;

    const deliverable = String(row[idx['Deliverable / Document']] || '').trim();
    const phase = String(row[idx['Phase']] || '').trim();
    if (!deliverable && !phase) continue;

    let start = normalizeDate_(rawRow[idx['Start Date']]);
    let end = normalizeDate_(rawRow[idx['End Date']]);
    const milestone = normalizeDate_(rawRow[idx['Milestone Date']]);

    // Auto-fix inverted dates
    if (start && end && end < start) end = start;

    const item = {
      id: String(row[idx['ID']] || '').trim() || `ROW-${r + 1}`,
      phase: phase,
      deliverable: deliverable,
      site: String(row[idx['Site']] || '').trim(),
      docs: String(row[idx['Docs']] || '').trim(),
      priority: String(row[idx['Priority']] || '').trim(),
      status: String(row[idx['Status']] || '').trim(),
      owner: String(row[idx['Owner']] || '').trim(),
      notes: String(row[idx['Notes / Actions']] || '').trim(),
      startDate: start ? toIsoDate_(start) : '',
      endDate: end ? toIsoDate_(end) : '',
      milestoneDate: milestone ? toIsoDate_(milestone) : '',
    };
    items.push(item);
    
    if (start) datePoints.push(start);
    if (end) datePoints.push(end);
    if (milestone) datePoints.push(milestone);
  }

  const today = new Date();
  if (!datePoints.length) {
    datePoints.push(today, today);
  } else {
    datePoints.push(today);
  }
  const timelineBuckets = buildHalfMonthBuckets_(minDate_(datePoints), maxDate_(datePoints), tz);

  return {
    model: 'date',
    sheetName: sheet.getName(),
    timelineHeaders: timelineBuckets.map((b) => b.label),
    timelineBuckets: timelineBuckets,
    items: items,
    todayIso: toIsoDate_(today),
    updatedAt: Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss z'),
  };
}

function parseMECOELegacyModel_(sheet, values, headerRowIdx, headers) {
  const timelineStartIdx = findTimelineStart_(headers);
  if (timelineStartIdx < 0) {
    throw new Error('Could not locate timeline columns for legacy ME COE sheet.');
  }

  const timelineHeaders = headers.slice(timelineStartIdx).map(cleanLabel_);
  const items = [];
  for (let r = headerRowIdx + 1; r < values.length; r++) {
    const row = values[r];
    if (!rowHasContent_(row)) continue;

    const item = {
      id: `ROW-${r + 1}`,
      phase: String(row[0] || '').trim(),
      deliverable: String(row[1] || '').trim(),
      site: String(row[2] || '').trim(),
      docs: String(row[3] || '').trim(),
      priority: String(row[4] || '').trim(),
      status: String(row[5] || '').trim(),
      owner: '',
      notes: String(row[6] || '').trim(),
      startDate: '',
      endDate: '',
      milestoneDate: '',
      timeline: [],
    };

    for (let c = timelineStartIdx; c < headers.length; c++) {
      item.timeline.push(String(row[c] || '').trim());
    }
    if (!item.phase && !item.deliverable) continue;
    items.push(item);
  }

  return {
    model: 'legacy',
    sheetName: sheet.getName(),
    timelineHeaders: timelineHeaders,
    timelineBuckets: [],
    items: items,
    todayIso: toIsoDate_(new Date()),
    updatedAt: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss z'),
  };
}

function isMECOEDateModel_(headers) {
  const lowerHeaders = headers.map(h => h.toLowerCase());
  return MECOE_DATE_HEADERS.every((h) => lowerHeaders.indexOf(h.toLowerCase()) !== -1);
}

function findMECOEHeaderRow_(values) {
  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const c0 = String(row[0] || '').trim().toLowerCase();
    const c1 = String(row[1] || '').trim().toLowerCase();
    const c2 = String(row[2] || '').trim().toLowerCase();
    const isLegacy = c0 === 'phase' && c1.indexOf('deliverable') !== -1;
    const isDateModel = c0 === 'id' && c1 === 'phase' && c2.indexOf('deliverable') !== -1;
    if (isLegacy || isDateModel) return i;
  }
  return -1;
}

function findTimelineStart_(headers) {
  const months = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'];
  for (let i = 7; i < headers.length; i++) {
    const lower = String(headers[i] || '').toLowerCase();
    const hasMonth = months.some((m) => lower.indexOf(m) !== -1);
    if (hasMonth || lower.indexOf('now') !== -1) return i;
  }
  return -1;
}

function cleanLabel_(label) {
  return String(label || '').replace(/\s+/g, ' ').trim();
}

function rowHasContent_(row) {
  for (let i = 0; i < row.length; i++) {
    if (String(row[i] || '').trim()) return true;
  }
  return false;
}

function minDate_(dates) {
  let min = dates[0];
  for (let i = 1; i < dates.length; i++) {
    if (dates[i].getTime() < min.getTime()) min = dates[i];
  }
  return min;
}

function maxDate_(dates) {
  let max = dates[0];
  for (let i = 1; i < dates.length; i++) {
    if (dates[i].getTime() > max.getTime()) max = dates[i];
  }
  return max;
}

function buildHalfMonthBuckets_(minDate, maxDate, tz) {
  const start = new Date(minDate.getFullYear(), minDate.getMonth(), minDate.getDate());
  const end = new Date(maxDate.getFullYear(), maxDate.getMonth(), maxDate.getDate());
  const cursor = new Date(start.getFullYear(), start.getMonth(), start.getDate() <= 15 ? 1 : 16);
  const lastBoundary = new Date(end.getFullYear(), end.getMonth(), end.getDate() <= 15 ? 15 : daysInMonth_(end.getFullYear(), end.getMonth()));

  const buckets = [];
  while (cursor.getTime() <= lastBoundary.getTime()) {
    const year = cursor.getFullYear();
    const month = cursor.getMonth();
    const day = cursor.getDate();
    const startDay = day === 1 ? 1 : 16;
    const endDay = day === 1 ? 15 : daysInMonth_(year, month);
    const bucketStart = new Date(year, month, startDay);
    const bucketEnd = new Date(year, month, endDay);
    buckets.push({
      key: Utilities.formatDate(bucketStart, tz, 'yyyy-MM-dd'),
      label: Utilities.formatDate(bucketStart, tz, 'MMM') + ' ' + startDay + '-' + endDay,
      start: Utilities.formatDate(bucketStart, tz, 'yyyy-MM-dd'),
      end: Utilities.formatDate(bucketEnd, tz, 'yyyy-MM-dd'),
    });

    if (day === 1) {
      cursor.setDate(16);
    } else {
      cursor.setMonth(cursor.getMonth() + 1);
      cursor.setDate(1);
    }
  }
  return buckets;
}

function daysInMonth_(year, month) {
  return new Date(year, month + 1, 0).getDate();
}

function getDashboardTitle_(titleKey, sheetNames, fallbackTitle) {
  const configTitle = getDashboardTitleFromConfig_(titleKey);
  if (configTitle) return configTitle;
  return getDashboardTitleFromLegacySheetCells_(sheetNames, fallbackTitle);
}

function getDashboardTitleFromConfig_(titleKey) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = ss.getSheetByName(CONFIG_SHEET_NAME);
  if (!config) return '';
  const values = config.getDataRange().getDisplayValues();
  for (let r = 0; r < values.length; r++) {
    const key = String(values[r][0] || '').trim().toLowerCase();
    const value = String(values[r][1] || '').trim();
    if (key === String(titleKey || '').toLowerCase() && value) return value;
  }
  return '';
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
  const sheetName = sheet.getName();
  const isStandard = sheetName === SHEET_NAME;
  const isMECOE = MECOE_SHEET_CANDIDATES.indexOf(sheetName) !== -1;
  if (!isStandard && !isMECOE) return;

  const row = range.getRow();
  const col = range.getColumn();
  let headerRow = 1;
  if (isMECOE) {
    const values = sheet.getDataRange().getDisplayValues();
    headerRow = findMECOEHeaderRow_(values);
    if (headerRow === -1) return;
    headerRow += 1; // convert 0-based index to sheet row number
  }
  if (row <= headerRow) return;

  const headerRaw = String(sheet.getRange(headerRow, col).getDisplayValue() || '').trim();
  const header = headerRaw.toLowerCase();
  const watched = ['status', 'start date', 'end date', 'milestone date'];
  if (watched.indexOf(header) === -1) return;

  const oldValue = String((e && Object.prototype.hasOwnProperty.call(e, 'oldValue')) ? e.oldValue : '').trim();
  const newValue = String(range.getDisplayValue() || '').trim();
  if (oldValue === newValue) return;

  const idCol = findHeaderColumn_(sheet, headerRow, ['id']);
  const titleCol = isStandard
    ? findHeaderColumn_(sheet, headerRow, ['title'])
    : findHeaderColumn_(sheet, headerRow, ['deliverable / document']);
  const rowId = idCol > 0 ? String(sheet.getRange(row, idCol).getDisplayValue() || '').trim() : '';
  const rowTitle = titleCol > 0 ? String(sheet.getRange(row, titleCol).getDisplayValue() || '').trim() : '';

  const audit = getOrCreateAuditSheet_();
  audit.appendRow([
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'),
    Session.getEffectiveUser().getEmail() || '',
    sheetName,
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