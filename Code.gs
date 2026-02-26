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
  const app = e && e.parameter && e.parameter.app ? String(e.parameter.app).toLowerCase() : '';
  const isMECOE = app === 'mecoe';
  const templateName = isMECOE ? 'IndexMECOE' : 'Index';
  const fallbackTitle = isMECOE ? 'ME COE Timeline Dashboard' : 'Visual Schedule Dashboard';
  const sheetNames = isMECOE ? MECOE_SHEET_CANDIDATES : [SHEET_NAME];
  const titleKey = isMECOE ? 'mecoe_title' : 'standard_title';
  const title = getDashboardTitle_(titleKey, sheetNames, fallbackTitle);

  const template = createTemplateStrict_(templateName);
  template.pageTitle = title;

  return template.evaluate()
    .setTitle(title)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function createTemplateStrict_(name) {
  try {
    return HtmlService.createTemplateFromFile(name);
  } catch (err) {
    throw new Error('Missing HTML template: ' + name);
  }
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
  const cache = CacheService.getScriptCache();
  const cacheKey = 'schedule_items_v1';
  if (!forceRefresh) {
    const cached = cache.get(cacheKey);
    if (cached) return JSON.parse(cached);
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) {
    throw new Error('Sheet "Schedule" not found.');
  }

  const values = sheet.getDataRange().getValues();
  if (!values.length) return [];

  const headers = values[0].map((h) => String(h).trim());
  const headerIndex = REQUIRED_HEADERS.reduce((acc, header) => {
    const idx = headers.indexOf(header);
    if (idx === -1) {
      throw new Error('Missing required column: ' + header);
    }
    acc[header] = idx;
    return acc;
  }, {});

  const items = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const id = String(row[headerIndex['ID']] || '').trim();
    const title = String(row[headerIndex['Title']] || '').trim();
    const startRaw = row[headerIndex['Start Date']];
    const endRaw = row[headerIndex['End Date']];
    if (!id || !title || !startRaw || !endRaw) continue;

    const startDate = normalizeDate_(startRaw);
    const endDate = normalizeDate_(endRaw);
    if (!startDate || !endDate || endDate < startDate) continue;

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

  cache.put(cacheKey, JSON.stringify(items), 300);
  return items;
}

function normalizeDate_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return value;
  }
  if (typeof value === 'string') {
    const parsed = new Date(value);
    if (!isNaN(parsed.getTime())) return parsed;
  }
  return null;
}

function toIsoDate_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function getMECOESchedule(forceRefresh) {
  const cache = CacheService.getScriptCache();
  const cacheKey = 'mecoe_schedule_v1';
  if (!forceRefresh) {
    const cached = cache.get(cacheKey);
    if (cached) return JSON.parse(cached);
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = MECOE_SHEET_CANDIDATES
    .map((name) => ss.getSheetByName(name))
    .filter((s) => !!s)[0];
  if (!sheet) {
    throw new Error('ME COE sheet not found. Try one of: ' + MECOE_SHEET_CANDIDATES.join(', '));
  }

  const rawValues = sheet.getDataRange().getValues();
  const displayValues = sheet.getDataRange().getDisplayValues();
  if (!displayValues.length) {
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

  cache.put(cacheKey, JSON.stringify(payload), 300);
  return payload;
}

function parseMECOEDateModel_(sheet, rawValues, displayValues, headerRowIdx, headers) {
  const tz = Session.getScriptTimeZone();
  const idx = {};
  for (let i = 0; i < headers.length; i++) {
    idx[headers[i]] = i;
  }

  const items = [];
  const datePoints = [];
  for (let r = headerRowIdx + 1; r < displayValues.length; r++) {
    const rawRow = rawValues[r] || [];
    const row = displayValues[r] || [];
    if (!rowHasContent_(row)) continue;

    const deliverable = String(row[idx['Deliverable / Document']] || '').trim();
    const phase = String(row[idx['Phase']] || '').trim();
    if (!deliverable && !phase) continue;

    const start = normalizeDate_(rawRow[idx['Start Date']]);
    const end = normalizeDate_(rawRow[idx['End Date']]);
    if (!start || !end || end < start) continue;
    const milestone = normalizeDate_(rawRow[idx['Milestone Date']]);

    const item = {
      id: String(row[idx['ID']] || '').trim(),
      phase: phase,
      deliverable: deliverable,
      site: String(row[idx['Site']] || '').trim(),
      docs: String(row[idx['Docs']] || '').trim(),
      priority: String(row[idx['Priority']] || '').trim(),
      status: String(row[idx['Status']] || '').trim(),
      owner: String(row[idx['Owner']] || '').trim(),
      notes: String(row[idx['Notes / Actions']] || '').trim(),
      startDate: toIsoDate_(start),
      endDate: toIsoDate_(end),
      milestoneDate: milestone ? toIsoDate_(milestone) : '',
    };
    items.push(item);
    datePoints.push(start, end);
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
      id: '',
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
  return MECOE_DATE_HEADERS.every((h) => headers.indexOf(h) !== -1);
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

  const header = String(sheet.getRange(headerRow, col).getDisplayValue() || '').trim();
  const watched = ['Status', 'Start Date', 'End Date', 'Milestone Date'];
  if (watched.indexOf(header) === -1) return;

  const oldValue = String((e && Object.prototype.hasOwnProperty.call(e, 'oldValue')) ? e.oldValue : '').trim();
  const newValue = String(range.getDisplayValue() || '').trim();
  if (oldValue === newValue) return;

  const idCol = findHeaderColumn_(sheet, headerRow, ['ID']);
  const titleCol = isStandard
    ? findHeaderColumn_(sheet, headerRow, ['Title'])
    : findHeaderColumn_(sheet, headerRow, ['Deliverable / Document']);
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
    header,
    oldValue,
    newValue,
  ]);
}

function findHeaderColumn_(sheet, headerRow, names) {
  const headers = sheet.getRange(headerRow, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  for (let i = 0; i < headers.length; i++) {
    const label = String(headers[i] || '').trim();
    if (names.indexOf(label) !== -1) return i + 1;
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
