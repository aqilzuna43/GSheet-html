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
const MECOE_CONFIG_SHEET_NAME = 'Config';

function doGet() {
  let template;
  try {
    template = HtmlService.createTemplateFromFile('IndexMECOE');
  } catch (err) {
    return HtmlService.createHtmlOutput(
      '<div style="font-family: sans-serif; padding: 20px;">' +
      '<h3>Setup Error</h3>' +
      '<p>Could not find <b>IndexMECOE.html</b>. Please add it to your Apps Script project.</p>' +
      '</div>'
    ).setTitle('Configuration Error');
  }

  const title = getMECOEDashboardTitle_(
    'mecoe_title',
    MECOE_SHEET_CANDIDATES,
    'ME COE Timeline Dashboard'
  );
  template.pageTitle = title;

  return template.evaluate()
    .setTitle(title)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function onOpen() {
  if (typeof setupMECOEScheduleSheet !== 'function') return;
  SpreadsheetApp.getUi()
    .createMenu('ME COE Tools')
    .addItem('Setup / Repair ME COE Sheet', 'setupMECOEScheduleSheet')
    .addItem('Migrate Legacy ME COE (ME_COE)', 'runMECOELegacyMigrationDefault')
    .addToUi();
}

function runMECOELegacyMigrationDefault() {
  if (typeof migrateLegacyMECOEToDateModel !== 'function') {
    SpreadsheetApp.getUi().alert('mecoe_setup.gs is not loaded.');
    return;
  }
  migrateLegacyMECOEToDateModel('ME_COE');
}

function getMECOESchedule(forceRefresh) {
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
  const lowerHeaders = headers.map((h) => h.toLowerCase());

  const idx = {};
  MECOE_DATE_HEADERS.forEach((expected) => {
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
  const lowerHeaders = headers.map((h) => h.toLowerCase());
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
  const lastBoundary = new Date(
    end.getFullYear(),
    end.getMonth(),
    end.getDate() <= 15 ? 15 : daysInMonth_(end.getFullYear(), end.getMonth())
  );

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

function getMECOEDashboardTitle_(titleKey, sheetNames, fallbackTitle) {
  const configTitle = getMECOEDashboardTitleFromConfig_(titleKey);
  if (configTitle) return configTitle;
  return getMECOEDashboardTitleFromLegacySheetCells_(sheetNames, fallbackTitle);
}

function getMECOEDashboardTitleFromConfig_(titleKey) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = ss.getSheetByName(MECOE_CONFIG_SHEET_NAME);
  if (!config) return '';
  const values = config.getDataRange().getDisplayValues();
  for (let r = 0; r < values.length; r++) {
    const key = String(values[r][0] || '').trim().toLowerCase();
    const value = String(values[r][1] || '').trim();
    if (key === String(titleKey || '').toLowerCase() && value) return value;
  }
  return '';
}

function getMECOEDashboardTitleFromLegacySheetCells_(sheetNames, fallbackTitle) {
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
