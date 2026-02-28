/**
 * Off Day Tracker for Google Sheets
 *
 * Functions to bind to buttons:
 * - buildLayout
 * - addOffDay
 * - useOffDay
 * - addPersonnel
 * - deletePersonnel
 * - manageOffGranted
 * - manageOffUsed
 */

const SHEET_DASHBOARD = 'Dashboard';
const SHEET_GRANTED = 'Offs (Granted)';
const SHEET_GRANTED_LEGACY = 'Offs (Unused)';
const SHEET_USED = 'Offs (Used)';
const SHEET_CALENDAR = 'Calendar';
const SHEET_PERSONNEL = 'Personnel';
const SHEET_EDIT_LOGS = 'Edit Logs';
const SHEET_BKP_GRANTED = '__BKP_GRANTED__';
const SHEET_BKP_USED = '__BKP_USED__';
const SHEET_BKP_CALENDAR = '__BKP_CALENDAR__';
const SHEET_BKP_LOGS = '__BKP_LOGS__';

const DEFAULT_PERSONNEL = 'Default';
const DASHBOARD_PERSONNEL_CELL = 'B2';

const UNUSED_HEADERS = [
  'ID',
  'Date Off Granted',
  'Duration Type',
  'Duration Value',
  'Reason Type',
  'Weekend Ops Duty Date',
  'Reason Details',
  'Provided By',
  'Used Value',
  'Remaining Value',
  'Status',
  'Created At',
  'Personnel'
];

const USED_HEADERS = [
  'Use ID',
  'Date Intended',
  'Session',
  'Duration Used',
  'Off IDs Used',
  'Comments',
  'Created At',
  'Personnel'
];

const EDIT_LOG_HEADERS = [
  'Log ID',
  'Timestamp',
  'Action',
  'Personnel',
  'Record Type',
  'Record ID',
  'Summary',
  'Before',
  'After',
  'Edited By'
];

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Off Day Tracker')
    .addItem('Reset All & Rebuild', 'resetAllAndRebuild')
    .addSeparator()
    .addItem('Build Layout', 'buildLayout')
    .addItem('Add Personnel', 'addPersonnel')
    .addItem('Delete Personnel', 'deletePersonnel')
    .addSeparator()
    .addItem('Add Off Day', 'addOffDay')
    .addItem('Use Off Day', 'useOffDay')
    .addItem('Edit Off Granted', 'manageOffGranted')
    .addItem('Delete Off Granted', 'deleteOffGranted')
    .addItem('Edit/Undo Off Used', 'manageOffUsed')
    .addSeparator()
    .addItem('Refresh Dashboard', 'refreshDashboard')
    .addItem('Refresh Calendar', 'refreshCalendar')
    .addToUi();

  try {
    protectManagedSheets_(SpreadsheetApp.getActiveSpreadsheet());
  } catch (err) {
    Logger.log(err && err.stack ? err.stack : String(err));
  }

  try {
    ensureOnChangeTrigger_();
  } catch (err) {
    Logger.log(err && err.stack ? err.stack : String(err));
  }
}

function onEdit(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();
  const a1 = e.range.getA1Notation();
  const name = sheet.getName();

  if (
    name === SHEET_GRANTED ||
    name === SHEET_USED ||
    name === SHEET_EDIT_LOGS ||
    (name === SHEET_CALENDAR && a1 !== 'B2')
  ) {
    revertProtectedEdit_(e);
    if (name === SHEET_CALENDAR) {
      refreshCalendar();
    } else if (name === SHEET_GRANTED || name === SHEET_USED) {
      refreshDashboard();
      refreshCalendar();
    } else if (name === SHEET_EDIT_LOGS) {
      const ss = sheet.getParent();
      refreshPersonnelSheetViews_(ss, getSelectedPersonnel_(ss));
      syncManagedBackupForSheet_(ss, SHEET_EDIT_LOGS, SHEET_BKP_LOGS);
    }
    return;
  }

  if (name === SHEET_CALENDAR && a1 === 'B2') {
    refreshCalendar();
    return;
  }

  if (name === SHEET_DASHBOARD && a1 === DASHBOARD_PERSONNEL_CELL) {
    refreshDashboard();
    refreshCalendar();
  }
}

function onChange(e) {
  if (!e || !e.changeType) return;

  // Protect against row/column/sheet structure changes in managed sheets.
  const type = String(e.changeType || '').toUpperCase();
  if (
    type === 'REMOVE_ROW' ||
    type === 'REMOVE_COLUMN' ||
    type === 'INSERT_ROW' ||
    type === 'INSERT_COLUMN' ||
    type === 'REMOVE_GRID' ||
    type === 'INSERT_GRID'
  ) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    restoreAllManagedSheetsFromBackup_(ss);
    protectManagedSheets_(ss);
    refreshDashboard();
    refreshCalendar();
  }
}

function ensureOnChangeTrigger_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const triggers = ScriptApp.getProjectTriggers();
  const exists = triggers.some((t) => t.getHandlerFunction() === 'onChange');
  if (!exists) {
    ScriptApp.newTrigger('onChange')
      .forSpreadsheet(ss)
      .onChange()
      .create();
  }
}

// Backward-compatible alias.
function setupOffDayTracker() {
  buildLayout();
}

function buildLayout() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dash = getOrCreateSheet_(ss, SHEET_DASHBOARD);
    const granted = getOrCreateGrantedSheet_(ss);
    const used = getOrCreateSheet_(ss, SHEET_USED);
    const calendar = getOrCreateSheet_(ss, SHEET_CALENDAR);
    const personnel = getOrCreateSheet_(ss, SHEET_PERSONNEL);
    const logs = getOrCreateSheet_(ss, SHEET_EDIT_LOGS);

    ensureSheetDimensions_(dash, 30, 6);
    ensureSheetDimensions_(granted, 1000, UNUSED_HEADERS.length);
    ensureSheetDimensions_(used, 1000, USED_HEADERS.length);
    ensureSheetDimensions_(calendar, 80, 24);
    ensureSheetDimensions_(personnel, 200, 1);
    ensureSheetDimensions_(logs, 2000, EDIT_LOG_HEADERS.length);

    granted.clear();
    used.clear();
    dash.clear();
    calendar.clear();
    personnel.clear();
    logs.clear();

    granted.getRange(1, 1, 1, UNUSED_HEADERS.length).setValues([UNUSED_HEADERS]).setFontWeight('bold');
    used.getRange(1, 1, 1, USED_HEADERS.length).setValues([USED_HEADERS]).setFontWeight('bold');
    personnel.getRange('A1').setValue('Name').setFontWeight('bold');
    personnel.getRange('A2').setValue(DEFAULT_PERSONNEL);
    logs.getRange(1, 1, 1, EDIT_LOG_HEADERS.length).setValues([EDIT_LOG_HEADERS]).setFontWeight('bold');
    personnel.setFrozenRows(1);
    granted.setFrozenRows(1);
    used.setFrozenRows(1);
    logs.setFrozenRows(1);

    granted.setColumnWidths(1, UNUSED_HEADERS.length, 150);
    used.setColumnWidths(1, USED_HEADERS.length, 170);

    granted.getRange('B:B').setNumberFormat('yyyy-mm-dd');
    granted.getRange('F:F').setNumberFormat('yyyy-mm-dd');
    granted.getRange('D:D').setNumberFormat('0.0');
    granted.getRange('I:I').setNumberFormat('0.0');
    granted.getRange('J:J').setNumberFormat('0.0');
    granted.getRange('L:L').setNumberFormat('yyyy-mm-dd hh:mm:ss');
    applyGrantedStatusHighlighting_(granted);

    used.getRange('B:B').setNumberFormat('yyyy-mm-dd');
    used.getRange('D:D').setNumberFormat('0.0');
    used.getRange('G:G').setNumberFormat('yyyy-mm-dd hh:mm:ss');
    logs.getRange('B:B').setNumberFormat('yyyy-mm-dd hh:mm:ss');

    dash.getRange('A1').setValue('Off Day Tracker').setFontSize(16).setFontWeight('bold');
    dash.getRange('A2').setValue('Personnel');
    dash.getRange('A3').setValue('Total Offs Granted');
    dash.getRange('A4').setValue('Total Offs Used');
    dash.getRange('A5').setValue('Off Balance (Remaining)');
    dash.getRange('A2:A5').setFontWeight('bold');

    refreshPersonnelDropdown_(ss, DEFAULT_PERSONNEL);

    refreshDashboard();

    dash.getRange('A7').setValue('How to add buttons').setFontWeight('bold');
    dash.getRange('A8').setValue('1) Insert > Drawing > shape "Reset", then Assign script: resetAllAndRebuild');
    dash.getRange('A9').setValue('2) Insert > Drawing > shape "Add Personnel", then Assign script: addPersonnel');
    dash.getRange('A10').setValue('3) Insert > Drawing > shape "Delete Personnel", then Assign script: deletePersonnel');
    dash.getRange('A11').setValue('4) Insert > Drawing > shape "Build Layout", then Assign script: buildLayout');
    dash.getRange('A12').setValue('5) Insert > Drawing > shape "Add Off Day", then Assign script: addOffDay');
    dash.getRange('A13').setValue('6) Insert > Drawing > shape "Use Off Day", then Assign script: useOffDay');
    dash.getRange('A14').setValue('7) Insert > Drawing > shape "Edit Granted", then Assign script: manageOffGranted');
    dash.getRange('A15').setValue('8) Insert > Drawing > shape "Delete Granted", then Assign script: deleteOffGranted');
    dash.getRange('A16').setValue('9) Insert > Drawing > shape "Edit/Undo Used", then Assign script: manageOffUsed');
    dash.getRange('A17').setValue('10) Or use the "Off Day Tracker" custom menu.');
    dash.autoResizeColumn(1);
    dash.autoResizeColumn(2);

    renderCalendar_(ss, null, DEFAULT_PERSONNEL);
    protectManagedSheets_(ss);
    syncManagedBackups_(ss);
    ss.setActiveSheet(dash);
  } catch (err) {
    const message = err && err.message ? err.message : String(err);
    Logger.log(err && err.stack ? err.stack : message);
    SpreadsheetApp.getUi().alert(`Build layout failed: ${message}`);
    throw err;
  }
}

function resetAllAndRebuild() {
  const ui = SpreadsheetApp.getUi();
  const confirm = ui.alert(
    'Confirm Reset',
    'This will delete existing tracker sheets and rebuild from scratch. Continue?',
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) {
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tmpName = '__TMP_RESET__';
  const targetSheets = [
    SHEET_GRANTED,
    SHEET_GRANTED_LEGACY,
    SHEET_USED,
    SHEET_CALENDAR,
    SHEET_PERSONNEL,
    SHEET_EDIT_LOGS,
    SHEET_BKP_GRANTED,
    SHEET_BKP_USED,
    SHEET_BKP_CALENDAR,
    SHEET_BKP_LOGS
  ];

  let tmp = ss.getSheetByName(tmpName);
  if (!tmp) tmp = ss.insertSheet(tmpName);
  ss.setActiveSheet(tmp);

  for (let i = 0; i < targetSheets.length; i += 1) {
    const sh = ss.getSheetByName(targetSheets[i]);
    if (sh) ss.deleteSheet(sh);
  }

  buildLayout();

  const tmpAfter = ss.getSheetByName(tmpName);
  if (tmpAfter) ss.deleteSheet(tmpAfter);
}

function addPersonnel() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'Add Personnel',
    'Enter personnel name:',
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;

  const name = String(response.getResponseText() || '').trim();
  if (!name) {
    ui.alert('Name is required.');
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const names = getPersonnelNames_(ss);
  const exists = names.some((n) => n.toLowerCase() === name.toLowerCase());
  if (exists) {
    ui.alert(`Personnel "${name}" already exists.`);
    return;
  }

  names.push(name);
  setPersonnelNames_(ss, names);
  refreshPersonnelDropdown_(ss, name);
  refreshDashboard();
  refreshCalendar();
  ui.alert(`Added personnel: ${name}`);
}

function deletePersonnel() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const selectedPersonnel = getSelectedPersonnel_(ss);
  const names = getPersonnelNames_(ss);

  if (names.length <= 1) {
    ui.alert('At least one personnel must remain.');
    return;
  }

  const namePrompt = ui.prompt(
    'Delete Personnel',
    `Current personnel: ${names.join(', ')}\n\nCurrently selected: ${selectedPersonnel}\n\nEnter exact name to delete:`,
    ui.ButtonSet.OK_CANCEL
  );
  if (namePrompt.getSelectedButton() !== ui.Button.OK) return;

  const name = String(namePrompt.getResponseText() || '').trim();
  if (!name) {
    ui.alert('Name is required.');
    return;
  }

  const deleteDataPrompt = ui.prompt(
    'Delete Related Records?',
    'Type YES to also delete all Offs (Granted) and Offs (Used) rows for this personnel. Leave blank to keep rows.',
    ui.ButtonSet.OK_CANCEL
  );
  if (deleteDataPrompt.getSelectedButton() !== ui.Button.OK) return;

  const result = submitDeletePersonnel({
    name,
    deleteData: String(deleteDataPrompt.getResponseText() || '').trim().toUpperCase() === 'YES'
  });
  ui.alert(result.message || (result.ok ? 'Done.' : 'Failed.'));
}

function submitDeletePersonnel(form) {
  const nameInput = normalizePersonnel_(form && form.name);
  const deleteData = !!(form && form.deleteData);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const names = getPersonnelNames_(ss);
  const name = names.find((n) => n.toLowerCase() === nameInput.toLowerCase());

  if (names.length <= 1) {
    return { ok: false, message: 'At least one personnel must remain.' };
  }
  if (!name) {
    return { ok: false, message: `Personnel "${nameInput}" not found.` };
  }

  const granted = getOrCreateGrantedSheet_(ss);
  const used = getOrCreateSheet_(ss, SHEET_USED);
  const logs = getOrCreateEditLogsSheet_(ss);
  const grantedRows = findRowsByPersonnel_(granted, 13, name);
  const usedRows = findRowsByPersonnel_(used, 8, name);
  const logRows = findRowsByPersonnel_(logs, 4, name);

  if (!deleteData && (grantedRows.length > 0 || usedRows.length > 0 || logRows.length > 0)) {
    return {
      ok: false,
      message: `Personnel "${name}" has existing records. Tick "Delete all related records" to proceed.`
    };
  }

  if (deleteData) {
    deleteRowsDescending_(used, usedRows);
    deleteRowsDescending_(granted, grantedRows);
    deleteRowsDescending_(logs, logRows);
  }

  const nextNames = names.filter((n) => n !== name);
  setPersonnelNames_(ss, nextNames);
  refreshPersonnelDropdown_(ss, nextNames[0]);
  refreshDashboard();
  refreshCalendar();

  return {
    ok: true,
    message: deleteData
      ? `Deleted personnel "${name}" and all related records.`
      : `Deleted personnel "${name}".`
  };
}

function addOffDay() {
  const selectedPersonnel = getSelectedPersonnel_(SpreadsheetApp.getActiveSpreadsheet());
  const html = HtmlService
    .createHtmlOutput(getAddOffDayDialogHtml_(selectedPersonnel))
    .setWidth(520)
    .setHeight(560);

  SpreadsheetApp.getUi().showModalDialog(html, 'Add Off Day');
}

function submitAddOffDay(form) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const granted = getOrCreateGrantedSheet_(ss);
  const selectedPersonnel = getSelectedPersonnel_(ss);

  const grantedDate = parseDateInput_(String(form.grantedDate || '').trim());
  if (!grantedDate) {
    return { ok: false, message: 'Invalid date. Use YYYY-MM-DD.' };
  }

  const durationTypeRaw = String(form.durationType || '').trim().toUpperCase();
  const reasonTypeRaw = String(form.reasonType || '').trim().toUpperCase();
  const weekendOpsDate = parseDateInput_(String(form.weekendOpsDate || '').trim());
  const otherDetails = String(form.otherDetails || '').trim();

  let providedBy = String(form.providedBy || '').trim();

  let durationType;
  let durationValue;
  if (durationTypeRaw === 'FULL') {
    durationType = 'Full Day';
    durationValue = 1;
  } else if (durationTypeRaw === 'HALF') {
    durationType = 'Half Day';
    durationValue = 0.5;
  } else {
    return { ok: false, message: 'Duration must be FULL or HALF.' };
  }

  let reasonType;
  let reasonDetails;
  let weekendOpsDutyDate = '';

  if (reasonTypeRaw === 'OPS') {
    reasonType = 'Ops';
    if (!weekendOpsDate) {
      return { ok: false, message: 'Please provide the Weekend Ops duty date.' };
    }
    if (!isWeekendDate_(weekendOpsDate)) {
      return { ok: false, message: 'Weekend Ops duty date must be Saturday or Sunday.' };
    }

    weekendOpsDutyDate = weekendOpsDate;
    reasonDetails = `Weekend Ops on ${formatDateYmd_(weekendOpsDate)}`;

    if (!providedBy) {
      providedBy = 'Yourself';
    }
  } else if (reasonTypeRaw === 'OTHERS') {
    reasonType = 'Others';
    if (!otherDetails) {
      return { ok: false, message: 'Please provide comments/details for Others.' };
    }
    if (!providedBy) {
      return { ok: false, message: 'Please fill in "Provided by who".' };
    }

    reasonDetails = otherDetails;
  } else {
    return { ok: false, message: 'Reason must be Ops or Others.' };
  }

  const nextRow = granted.getLastRow() + 1;
  const id = `G-${String(nextRow - 1).padStart(4, '0')}`;
  const now = new Date();

  granted.getRange(nextRow, 1, 1, UNUSED_HEADERS.length).setValues([[
    id,
    grantedDate,
    durationType,
    durationValue,
    reasonType,
    weekendOpsDutyDate,
    reasonDetails,
    providedBy,
    0,
    durationValue,
    'Unused',
    now,
    selectedPersonnel
  ]]);

  granted.getRange(nextRow, 2).setNumberFormat('yyyy-mm-dd');
  granted.getRange(nextRow, 4).setNumberFormat('0.0');
  granted.getRange(nextRow, 6).setNumberFormat('yyyy-mm-dd');
  granted.getRange(nextRow, 9, 1, 2).setNumberFormat('0.0');
  granted.getRange(nextRow, 12).setNumberFormat('yyyy-mm-dd hh:mm:ss');
  applyGrantedStatusHighlighting_(granted);

  refreshDashboard();
  refreshCalendar();

  return { ok: true, message: `Added ${durationType} (${durationValue}) off day as ${id} for ${selectedPersonnel}.` };
}

function useOffDay() {
  const selectedPersonnel = getSelectedPersonnel_(SpreadsheetApp.getActiveSpreadsheet());
  const initialOptions = getUseOffOptions();
  const html = HtmlService
    .createHtmlOutput(getUseOffDayDialogHtml_(initialOptions, selectedPersonnel))
    .setWidth(560)
    .setHeight(620);

  SpreadsheetApp.getUi().showModalDialog(html, 'Use Off Day');
}

function getUseOffOptions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const granted = getOrCreateGrantedSheet_(ss);
  const selectedPersonnel = getSelectedPersonnel_(ss);
  const lastRow = granted.getLastRow();

  if (lastRow < 2) return [];

  const values = granted.getRange(2, 1, lastRow - 1, UNUSED_HEADERS.length).getValues();
  const options = [];

  for (let i = 0; i < values.length; i += 1) {
    const row = values[i];
    const id = String(row[0] || '').trim();
    if (!id) continue;
    if (!isPersonnelMatch_(row[12], selectedPersonnel)) continue;

    const durationValue = toNumber_(row[3]);
    const usedValue = toNumber_(row[8]);
    let remainingValue = toNumber_(row[9]);

    if (Number.isNaN(remainingValue)) {
      remainingValue = Math.max(durationValue - usedValue, 0);
    }

    if (remainingValue <= 0) continue;

    const reasonType = String(row[4] || '').trim().toLowerCase();
    const weekendOpsDate = row[5];
    const reasonDetails = String(row[6] || '').trim();
    const providedBy = String(row[7] || '').trim();

    let label;
    if (reasonType === 'ops') {
      const weekendText = formatDateYmdSafe_(weekendOpsDate);
      label = `${id}, ${formatDuration_(remainingValue)} day, Weekend Ops on (${weekendText})`;
    } else {
      const providerText = providedBy || 'Unknown';
      const detailsText = reasonDetails || 'No details';
      label = `${id}, ${formatDuration_(remainingValue)} day, Off provided by (${providerText}) For (${detailsText})`;
    }

    options.push({
      id,
      label,
      remaining: remainingValue
    });
  }

  options.sort((a, b) => a.id.localeCompare(b.id));
  return options;
}

function submitUseOffDay(form) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const granted = getOrCreateGrantedSheet_(ss);
  const used = getOrCreateSheet_(ss, SHEET_USED);
  const selectedPersonnel = getSelectedPersonnel_(ss);

  const intendedDate = parseDateInput_(String(form.intendedDate || '').trim());
  if (!intendedDate) {
    return { ok: false, message: 'Invalid intended date. Use YYYY-MM-DD.' };
  }

  const sessionRaw = String(form.session || '').trim().toUpperCase();

  let session;
  let durationNeeded;
  if (sessionRaw === 'FULL') {
    session = 'Full Day';
    durationNeeded = 1;
  } else if (sessionRaw === 'AM') {
    session = 'AM';
    durationNeeded = 0.5;
  } else if (sessionRaw === 'PM') {
    session = 'PM';
    durationNeeded = 0.5;
  } else {
    return { ok: false, message: 'Session must be AM, PM, or FULL.' };
  }

  const rawIds = Array.isArray(form.selectedIds) ? form.selectedIds : [];
  const selectedIds = rawIds
    .map((id) => String(id || '').trim())
    .filter((id) => id);

  if (selectedIds.length === 0) {
    return { ok: false, message: 'Please choose at least one OFF ID.' };
  }

  const comments = String(form.comments || '').trim();
  const map = getUnusedRecordMap_(granted, selectedPersonnel);

  let selectedTotal = 0;
  for (let i = 0; i < selectedIds.length; i += 1) {
    const id = selectedIds[i];
    const rec = map[id];
    if (!rec) {
      return { ok: false, message: `OFF ID ${id} does not exist or has no remaining balance.` };
    }
    selectedTotal += rec.remaining;
  }

  if (selectedTotal + 1e-9 < durationNeeded) {
    if (durationNeeded === 1 && Math.abs(selectedTotal - 0.5) < 1e-9) {
      return {
        ok: false,
        message: 'You selected only 0.5 day. For Full Day OFF, choose another ID to make a total of 1 day.'
      };
    }

    return {
      ok: false,
      message: `Selected IDs total ${formatDuration_(selectedTotal)} day, but ${formatDuration_(durationNeeded)} day is required.`
    };
  }

  let remainingNeed = durationNeeded;
  const allocations = [];

  for (let i = 0; i < selectedIds.length; i += 1) {
    if (remainingNeed <= 0) break;

    const id = selectedIds[i];
    const rec = map[id];
    if (!rec || rec.remaining <= 0) continue;

    const useAmount = Math.min(rec.remaining, remainingNeed);
    if (useAmount <= 0) continue;

    rec.used += useAmount;
    rec.remaining -= useAmount;
    rec.status = rec.remaining <= 0 ? 'Used' : 'Partial';

    allocations.push({
      id,
      amount: useAmount,
      rowIndex: rec.rowIndex,
      used: rec.used,
      remaining: rec.remaining,
      status: rec.status
    });

    remainingNeed -= useAmount;
  }

  if (remainingNeed > 1e-9) {
    return {
      ok: false,
      message: 'Unable to allocate enough off balance from selected IDs. Please try again.'
    };
  }

  for (let i = 0; i < allocations.length; i += 1) {
    const a = allocations[i];
    granted.getRange(a.rowIndex, 9, 1, 3).setValues([[a.used, a.remaining, a.status]]);
  }
  applyGrantedStatusHighlighting_(granted);

  const nextRow = used.getLastRow() + 1;
  const useId = `U-${String(nextRow - 1).padStart(4, '0')}`;
  const now = new Date();
  const offIdsUsed = allocations
    .map((a) => `${a.id} (${formatDuration_(a.amount)})`)
    .join(' + ');

  used.getRange(nextRow, 1, 1, USED_HEADERS.length).setValues([[
    useId,
    intendedDate,
    session,
    durationNeeded,
    offIdsUsed,
    comments,
    now,
    selectedPersonnel
  ]]);

  used.getRange(nextRow, 2).setNumberFormat('yyyy-mm-dd');
  used.getRange(nextRow, 4).setNumberFormat('0.0');
  used.getRange(nextRow, 7).setNumberFormat('yyyy-mm-dd hh:mm:ss');

  refreshDashboard();
  refreshCalendar();

  return {
    ok: true,
    message: `Recorded ${session} usage (${formatDuration_(durationNeeded)} day) for ${selectedPersonnel} using ${offIdsUsed}.`
  };
}

function refreshDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dash = getOrCreateSheet_(ss, SHEET_DASHBOARD);
  const selectedPersonnel = getSelectedPersonnel_(ss);
  const granted = getOrCreateGrantedSheet_(ss);
  const used = getOrCreateSheet_(ss, SHEET_USED);

  const totalGranted = sumColumnByPersonnel_(granted, 4, 13, selectedPersonnel);
  const totalUsed = sumColumnByPersonnel_(used, 4, 8, selectedPersonnel);
  const balance = sumColumnByPersonnel_(granted, 10, 13, selectedPersonnel);

  dash.getRange(DASHBOARD_PERSONNEL_CELL).setValue(selectedPersonnel);
  dash.getRange('B3:B5').setValues([[totalGranted], [totalUsed], [balance]]);
  dash.getRange('B3:B5').setNumberFormat('0.0');
  ensureManagedColumnFormats_(ss);
  refreshPersonnelSheetViews_(ss, selectedPersonnel);
  syncManagedBackups_(ss);
}

function refreshCalendar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const calendar = getOrCreateSheet_(ss, SHEET_CALENDAR);
  const selectedPersonnel = getSelectedPersonnel_(ss);
  const selectedMonth = normalizeMonthStart_(calendar.getRange('B2').getValue());
  renderCalendar_(ss, selectedMonth, selectedPersonnel);
  syncManagedBackupForSheet_(ss, SHEET_CALENDAR, SHEET_BKP_CALENDAR);
}

function getCurrentBalance_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const granted = getOrCreateGrantedSheet_(ss);
  const selectedPersonnel = getSelectedPersonnel_(ss);
  return sumColumnByPersonnel_(granted, 10, 13, selectedPersonnel);
}

function refreshPersonnelSheetViews_(ss, selectedPersonnel) {
  applyPersonnelFilter_(getOrCreateGrantedSheet_(ss), 13, selectedPersonnel);
  applyPersonnelFilter_(getOrCreateSheet_(ss, SHEET_USED), 8, selectedPersonnel);
  applyPersonnelFilter_(getOrCreateEditLogsSheet_(ss), 4, selectedPersonnel);
}

function protectManagedSheets_(ss) {
  if (!ss) return;

  const granted = ss.getSheetByName(SHEET_GRANTED) || ss.getSheetByName(SHEET_GRANTED_LEGACY);
  const used = ss.getSheetByName(SHEET_USED);
  const calendar = ss.getSheetByName(SHEET_CALENDAR);
  const logs = ss.getSheetByName(SHEET_EDIT_LOGS);

  if (granted) {
    protectSheetForTracker_(granted, []);
  }
  if (used) {
    protectSheetForTracker_(used, []);
  }
  if (calendar) {
    protectSheetForTracker_(calendar, ['B2']);
  }
  if (logs) {
    protectSheetForTracker_(logs, []);
  }
}

function protectSheetForTracker_(sheet, unprotectedA1Ranges) {
  const existing = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  for (let i = 0; i < existing.length; i += 1) {
    if (String(existing[i].getDescription() || '') === 'OFF_TRACKER_LOCK') {
      existing[i].remove();
    }
  }

  const protection = sheet.protect();
  protection.setDescription('OFF_TRACKER_LOCK');
  protection.setWarningOnly(false);

  const me = Session.getEffectiveUser();
  const meEmail = me && me.getEmail ? me.getEmail() : '';

  if (meEmail) {
    protection.addEditor(meEmail);
  }

  const editors = protection.getEditors();
  const removableEditors = [];
  for (let i = 0; i < editors.length; i += 1) {
    if (!meEmail || editors[i].getEmail() !== meEmail) {
      removableEditors.push(editors[i]);
    }
  }
  if (removableEditors.length > 0) {
    protection.removeEditors(removableEditors);
  }
  if (protection.canDomainEdit()) {
    protection.setDomainEdit(false);
  }

  const unprotected = Array.isArray(unprotectedA1Ranges)
    ? unprotectedA1Ranges.map((a1) => sheet.getRange(a1))
    : [];
  protection.setUnprotectedRanges(unprotected);
}

function revertProtectedEdit_(e) {
  const range = e.range;
  const ss = range.getSheet().getParent();
  const sheetName = range.getSheet().getName();

  const isSingleCell = range.getNumRows() === 1 && range.getNumColumns() === 1;
  const canUseOldValue = isSingleCell && typeof e.oldValue !== 'undefined';

  if (canUseOldValue) {
    setProtectedCellToOldValue_(range, e.oldValue);
  } else {
    restoreManagedSheetFromBackup_(ss, sheetName);
  }

  ensureManagedColumnFormats_(ss);
}

function setProtectedCellToOldValue_(range, oldValue) {
  const sheet = range.getSheet();
  const sheetName = sheet.getName();
  const col = range.getColumn();
  const dateColFormat = getManagedDateColumnFormat_(sheetName, col);

  if (dateColFormat) {
    const dateValue = toDateValueForCell_(oldValue);
    if (dateValue) {
      range.setValue(dateValue);
    } else if (oldValue === null || String(oldValue).trim() === '') {
      range.clearContent();
    } else {
      range.setValue(oldValue);
    }
    range.setNumberFormat(dateColFormat);
    return;
  }

  if (isManagedNumericColumn_(sheetName, col) && typeof oldValue === 'string') {
    const num = Number(oldValue);
    if (!Number.isNaN(num)) {
      range.setValue(num);
      range.setNumberFormat('0.0');
      return;
    }
  }

  range.setValue(oldValue);
}

function toDateValueForCell_(raw) {
  if (raw instanceof Date && !Number.isNaN(raw.getTime())) return raw;
  if (raw === null || typeof raw === 'undefined') return null;

  if (typeof raw === 'number' && !Number.isNaN(raw)) {
    return sheetSerialToDate_(raw);
  }

  const text = String(raw).trim();
  if (!text) return null;

  if (/^-?\d+(\.\d+)?$/.test(text)) {
    const num = Number(text);
    if (!Number.isNaN(num)) {
      return sheetSerialToDate_(num);
    }
  }

  const iso = parseDateInput_(text);
  if (iso) return iso;

  const parsed = new Date(text);
  if (!Number.isNaN(parsed.getTime())) return parsed;

  return null;
}

function sheetSerialToDate_(serial) {
  const whole = Math.floor(serial);
  const frac = serial - whole;
  const base = new Date(1899, 11, 30);
  const millis = (whole * 86400000) + Math.round(frac * 86400000);
  return new Date(base.getTime() + millis);
}

function getManagedDateColumnFormat_(sheetName, column) {
  if (sheetName === SHEET_GRANTED) {
    if (column === 2 || column === 6) return 'yyyy-mm-dd';
    if (column === 12) return 'yyyy-mm-dd hh:mm:ss';
  }
  if (sheetName === SHEET_USED) {
    if (column === 2) return 'yyyy-mm-dd';
    if (column === 7) return 'yyyy-mm-dd hh:mm:ss';
  }
  if (sheetName === SHEET_CALENDAR && column === 2) {
    return 'mmmm yyyy';
  }
  return '';
}

function isManagedNumericColumn_(sheetName, column) {
  if (sheetName === SHEET_GRANTED) {
    return column === 4 || column === 9 || column === 10;
  }
  if (sheetName === SHEET_USED) {
    return column === 4;
  }
  return false;
}

function ensureManagedColumnFormats_(ss) {
  const granted = ss.getSheetByName(SHEET_GRANTED) || ss.getSheetByName(SHEET_GRANTED_LEGACY);
  const used = ss.getSheetByName(SHEET_USED);
  const calendar = ss.getSheetByName(SHEET_CALENDAR);

  if (granted) {
    granted.getRange('B:B').setNumberFormat('yyyy-mm-dd');
    granted.getRange('F:F').setNumberFormat('yyyy-mm-dd');
    granted.getRange('D:D').setNumberFormat('0.0');
    granted.getRange('I:I').setNumberFormat('0.0');
    granted.getRange('J:J').setNumberFormat('0.0');
    granted.getRange('L:L').setNumberFormat('yyyy-mm-dd hh:mm:ss');
  }

  if (used) {
    used.getRange('B:B').setNumberFormat('yyyy-mm-dd');
    used.getRange('D:D').setNumberFormat('0.0');
    used.getRange('G:G').setNumberFormat('yyyy-mm-dd hh:mm:ss');
  }

  if (calendar) {
    calendar.getRange('B2').setNumberFormat('mmmm yyyy');
  }
}

function applyPersonnelFilter_(sheet, personnelColumnIndex, selectedPersonnel) {
  if (!sheet || sheet.getLastColumn() < personnelColumnIndex) return;

  const existingFilter = sheet.getFilter();
  if (existingFilter) {
    existingFilter.remove();
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 1) return;

  const range = sheet.getRange(1, 1, lastRow, sheet.getLastColumn());
  range.createFilter();

  if (lastRow < 2) return;

  sheet.getFilter().setColumnFilterCriteria(
    personnelColumnIndex,
    SpreadsheetApp.newFilterCriteria()
      .whenTextEqualTo(normalizePersonnel_(selectedPersonnel))
      .build()
  );
}

function sumColumn_(sheet, columnIndex) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;

  const values = sheet.getRange(2, columnIndex, lastRow - 1, 1).getValues();
  return values.reduce((acc, row) => acc + (Number(row[0]) || 0), 0);
}

function renderCalendar_(ss, requestedMonthStart, selectedPersonnel) {
  const calendar = getOrCreateSheet_(ss, SHEET_CALENDAR);
  ensureSheetDimensions_(calendar, 80, 24);
  const monthOptions = getNext24Months_();
  const monthStart = pickMonthOption_(requestedMonthStart, monthOptions);
  const personnel = selectedPersonnel || getSelectedPersonnel_(ss);

  calendar.clear();

  // Hidden helper range for month dropdown options.
  calendar.getRange(2, 24, monthOptions.length, 1)
    .setValues(monthOptions.map((d) => [d]))
    .setNumberFormat('mmmm yyyy');
  calendar.hideColumns(24, 1);

  calendar.getRange('A1').setValue('Calendar').setFontSize(16).setFontWeight('bold');
  calendar.getRange('A2').setValue('Month / Year').setFontWeight('bold');

  const selector = calendar.getRange('B2');
  selector
    .setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInRange(calendar.getRange(2, 24, monthOptions.length, 1), true)
        .setAllowInvalid(false)
        .build()
    )
    .setNumberFormat('mmmm yyyy')
    .setValue(monthStart);

  const dayShift = (monthStart.getDay() + 6) % 7; // Monday-first calendar.
  const firstGridDate = new Date(monthStart);
  firstGridDate.setDate(monthStart.getDate() - dayShift);

  const givenByDate = getDurationByDate_(getOrCreateGrantedSheet_(ss), 2, 4, 13, personnel);
  const usedSummaryByDate = getUsedSummaryByDate_(getOrCreateSheet_(ss, SHEET_USED), personnel);

  const headers = [['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun']];
  calendar.getRange(4, 1, 1, 7)
    .setValues(headers)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  const values = [];
  const backgrounds = [];
  const fontColors = [];

  for (let r = 0; r < 6; r += 1) {
    const rowValues = [];
    const rowBackgrounds = [];
    const rowFontColors = [];

    for (let c = 0; c < 7; c += 1) {
      const d = new Date(firstGridDate);
      d.setDate(firstGridDate.getDate() + r * 7 + c);

      const key = dateKey_(d);
      const given = givenByDate[key] || 0;
      const usedSummary = usedSummaryByDate[key] || null;
      const used = usedSummary ? usedSummary.total : 0;
      const inMonth = d.getMonth() === monthStart.getMonth();

      let label = String(d.getDate());
      const chips = [];
      if (given > 0) chips.push(`+${formatDuration_(given)}`);
      if (used > 0) chips.push(formatUsedCalendarChip_(usedSummary));
      if (chips.length > 0) label += `\n${chips.join('  ')}`;
      rowValues.push(label);

      if (!inMonth) {
        rowBackgrounds.push('#f3f4f6');
        rowFontColors.push('#9ca3af');
      } else if (given > 0 && used > 0) {
        rowBackgrounds.push('#fff7cc');
        rowFontColors.push('#111827');
      } else if (given > 0) {
        rowBackgrounds.push('#e7f7ec');
        rowFontColors.push('#111827');
      } else if (used > 0) {
        rowBackgrounds.push('#fdecec');
        rowFontColors.push('#111827');
      } else {
        rowBackgrounds.push('#ffffff');
        rowFontColors.push('#111827');
      }
    }

    values.push(rowValues);
    backgrounds.push(rowBackgrounds);
    fontColors.push(rowFontColors);
  }

  calendar.getRange(5, 1, 6, 7)
    .setValues(values)
    .setBackgrounds(backgrounds)
    .setFontColors(fontColors)
    .setWrap(true)
    .setVerticalAlignment('top')
    .setHorizontalAlignment('left')
    .setFontSize(11)
    .setBorder(true, true, true, true, true, true);

  calendar.setColumnWidths(1, 7, 110);
  for (let row = 5; row <= 10; row += 1) {
    calendar.setRowHeight(row, 72);
  }

  calendar.getRange('A12').setValue('Legend').setFontWeight('bold');
  calendar.getRange('B12').setValue('+ Granted').setBackground('#e7f7ec');
  calendar.getRange('C12').setValue('- Used').setBackground('#fdecec');
  calendar.getRange('D12').setValue('+/- Same Day').setBackground('#fff7cc');
  calendar.autoResizeColumn(1);
}

function getDurationByDate_(sheet, dateColumnIndex, durationColumnIndex, personnelColumnIndex, selectedPersonnel) {
  const out = {};
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return out;

  const values = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();

  for (let i = 0; i < values.length; i += 1) {
    const row = values[i];
    if (
      selectedPersonnel &&
      personnelColumnIndex &&
      !isPersonnelMatch_(row[personnelColumnIndex - 1], selectedPersonnel)
    ) {
      continue;
    }

    const rawDate = row[dateColumnIndex - 1];
    if (!rawDate) continue;

    const d = rawDate instanceof Date ? rawDate : new Date(rawDate);
    if (Number.isNaN(d.getTime())) continue;

    const key = dateKey_(d);
    const duration = Number(row[durationColumnIndex - 1]) || 0;
    out[key] = (out[key] || 0) + duration;
  }

  return out;
}

function getUsedSummaryByDate_(sheet, selectedPersonnel) {
  const out = {};
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return out;

  const values = sheet.getRange(2, 1, lastRow - 1, USED_HEADERS.length).getValues();
  for (let i = 0; i < values.length; i += 1) {
    const row = values[i];
    if (selectedPersonnel && !isPersonnelMatch_(row[7], selectedPersonnel)) continue;

    const rawDate = row[1];
    if (!rawDate) continue;

    const d = rawDate instanceof Date ? rawDate : new Date(rawDate);
    if (Number.isNaN(d.getTime())) continue;

    const key = dateKey_(d);
    const duration = Number(row[3]) || 0;
    const session = String(row[2] || '').trim().toUpperCase();

    if (!out[key]) {
      out[key] = {
        total: 0,
        sessions: []
      };
    }

    out[key].total += duration;
    if (session && out[key].sessions.indexOf(session) === -1) {
      out[key].sessions.push(session);
    }
  }

  return out;
}

function formatUsedCalendarChip_(usedSummary) {
  if (!usedSummary) return '';

  const used = Number(usedSummary.total) || 0;
  if (used <= 0) return '';

  let suffix = '';
  const isHalfDay = Math.abs(used - 0.5) < 1e-9;
  if (isHalfDay && usedSummary.sessions.length === 1) {
    const session = usedSummary.sessions[0];
    if (session === 'AM' || session === 'PM') {
      suffix = ` (${session})`;
    }
  }

  return `-${formatDuration_(used)}${suffix}`;
}

function getUnusedRecordMap_(unusedSheet, selectedPersonnel) {
  const map = {};
  const lastRow = unusedSheet.getLastRow();
  if (lastRow < 2) return map;

  const values = unusedSheet.getRange(2, 1, lastRow - 1, UNUSED_HEADERS.length).getValues();
  for (let i = 0; i < values.length; i += 1) {
    const row = values[i];
    const id = String(row[0] || '').trim();
    if (!id) continue;
    if (selectedPersonnel && !isPersonnelMatch_(row[12], selectedPersonnel)) continue;

    const duration = Number(row[3]) || 0;
    let used = Number(row[8]);
    let remaining = Number(row[9]);

    if (Number.isNaN(used)) used = 0;
    if (Number.isNaN(remaining)) remaining = Math.max(duration - used, 0);

    map[id] = {
      rowIndex: i + 2,
      used,
      remaining,
      status: String(row[10] || 'Unused')
    };
  }

  return map;
}

function formatDuration_(value) {
  const rounded = Math.round(value * 10) / 10;
  return Number.isInteger(rounded) ? String(rounded) : rounded.toFixed(1);
}

function dateKey_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function formatDateYmd_(value) {
  const d = value instanceof Date ? value : new Date(value);
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function formatDateYmdSafe_(value) {
  if (!value) return 'Unknown date';

  const d = value instanceof Date ? value : new Date(value);
  if (Number.isNaN(d.getTime())) return String(value);

  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function toNumber_(value) {
  if (typeof value === 'number') return value;
  if (value === null || value === undefined || value === '') return NaN;

  // Accept "0,5" and "0.5 day" style legacy values.
  const cleaned = String(value).replace(',', '.').match(/-?\d+(\.\d+)?/);
  if (!cleaned) return NaN;
  return Number(cleaned[0]);
}

function isWeekendDate_(date) {
  const day = date.getDay();
  return day === 0 || day === 6;
}

function getNext24Months_() {
  const now = new Date();
  const start = new Date(now.getFullYear(), now.getMonth(), 1);

  const out = [];
  for (let i = 0; i < 24; i += 1) {
    out.push(new Date(start.getFullYear(), start.getMonth() + i, 1));
  }
  return out;
}

function normalizeMonthStart_(value) {
  if (!value) return null;

  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return new Date(value.getFullYear(), value.getMonth(), 1);
  }

  const text = String(value).trim();
  if (!text) return null;

  const parsed = new Date(text);
  if (Number.isNaN(parsed.getTime())) return null;

  return new Date(parsed.getFullYear(), parsed.getMonth(), 1);
}

function pickMonthOption_(requestedMonthStart, monthOptions) {
  if (!requestedMonthStart) return monthOptions[0];

  for (let i = 0; i < monthOptions.length; i += 1) {
    const opt = monthOptions[i];
    if (
      opt.getFullYear() === requestedMonthStart.getFullYear() &&
      opt.getMonth() === requestedMonthStart.getMonth()
    ) {
      return opt;
    }
  }

  return monthOptions[0];
}

function quoteSheetName_(name) {
  return `'${name.replace(/'/g, "''")}'`;
}

function parseDateInput_(input) {
  if (!input) {
    const now = new Date();
    return new Date(now.getFullYear(), now.getMonth(), now.getDate());
  }

  const match = input.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) return null;

  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);
  const d = new Date(year, month - 1, day);

  if (
    d.getFullYear() !== year ||
    d.getMonth() !== month - 1 ||
    d.getDate() !== day
  ) {
    return null;
  }
  return d;
}

function normalizeDateInputString_(value) {
  if (!value) return '';

  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return formatDateYmd_(value);
  }

  const text = String(value).trim();
  if (!text) return '';

  const strict = parseDateInput_(text);
  if (strict) return formatDateYmd_(strict);

  const parsed = new Date(text);
  if (!Number.isNaN(parsed.getTime())) {
    return formatDateYmd_(parsed);
  }

  return '';
}

function normalizeSessionRaw_(value) {
  const text = String(value || '').trim().toUpperCase();
  if (text === 'FULL DAY' || text === 'FULL') return 'FULL';
  if (text === 'AM') return 'AM';
  if (text === 'PM') return 'PM';
  return 'FULL';
}

function getOrCreateSheet_(ss, name) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  return sheet;
}

function ensureSheetDimensions_(sheet, minRows, minColumns) {
  const currentRows = sheet.getMaxRows();
  const currentColumns = sheet.getMaxColumns();

  if (currentRows < minRows) {
    sheet.insertRowsAfter(currentRows, minRows - currentRows);
  }
  if (currentColumns < minColumns) {
    sheet.insertColumnsAfter(currentColumns, minColumns - currentColumns);
  }
}

function getPersonnelNames_(ss) {
  const personnelSheet = getOrCreateSheet_(ss, SHEET_PERSONNEL);
  if (personnelSheet.getLastRow() < 1) {
    personnelSheet.getRange('A1').setValue('Name').setFontWeight('bold');
  }

  const lastRow = personnelSheet.getLastRow();
  const rawNames = lastRow >= 2
    ? personnelSheet.getRange(2, 1, lastRow - 1, 1).getValues().map((r) => normalizePersonnel_(r[0]))
    : [];

  const uniqueNames = [];
  for (let i = 0; i < rawNames.length; i += 1) {
    const name = rawNames[i];
    if (!name) continue;
    if (uniqueNames.indexOf(name) === -1) uniqueNames.push(name);
  }

  return setPersonnelNames_(ss, uniqueNames);
}

function setPersonnelNames_(ss, rawNames) {
  const personnelSheet = getOrCreateSheet_(ss, SHEET_PERSONNEL);
  const normalized = [];
  const incoming = Array.isArray(rawNames) ? rawNames : [];

  for (let i = 0; i < incoming.length; i += 1) {
    const name = normalizePersonnel_(incoming[i]);
    if (!name) continue;
    if (normalized.indexOf(name) === -1) normalized.push(name);
  }

  const nonDefault = normalized.filter((name) => name.toLowerCase() !== DEFAULT_PERSONNEL.toLowerCase());
  const finalNames = nonDefault.length > 0
    ? nonDefault
    : (normalized.length > 0 ? normalized : [DEFAULT_PERSONNEL]);

  personnelSheet.getRange('A1').setValue('Name').setFontWeight('bold');
  if (personnelSheet.getMaxRows() > 1) {
    personnelSheet.getRange(2, 1, personnelSheet.getMaxRows() - 1, 1).clearContent();
  }
  personnelSheet.getRange(2, 1, finalNames.length, 1).setValues(finalNames.map((name) => [name]));

  return finalNames;
}

function refreshPersonnelDropdown_(ss, preferredName) {
  const dash = getOrCreateSheet_(ss, SHEET_DASHBOARD);
  const personnelSheet = getOrCreateSheet_(ss, SHEET_PERSONNEL);
  const names = getPersonnelNames_(ss);

  const listRange = personnelSheet.getRange(2, 1, names.length, 1);
  const dropdown = dash.getRange(DASHBOARD_PERSONNEL_CELL);
  dropdown.setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInRange(listRange, true)
      .setAllowInvalid(false)
      .build()
  );

  const current = normalizePersonnel_(dropdown.getValue());
  const chosen = preferredName && names.indexOf(preferredName) !== -1
    ? preferredName
    : (names.indexOf(current) !== -1 ? current : names[0]);

  dropdown.setValue(chosen);
  dropdown.setHorizontalAlignment('left');
}

function getSelectedPersonnel_(ss) {
  const dash = getOrCreateSheet_(ss, SHEET_DASHBOARD);
  refreshPersonnelDropdown_(ss);
  return normalizePersonnel_(dash.getRange(DASHBOARD_PERSONNEL_CELL).getValue());
}

function normalizePersonnel_(value) {
  const text = String(value || '').trim();
  return text || DEFAULT_PERSONNEL;
}

function isPersonnelMatch_(rawValue, selectedPersonnel) {
  return normalizePersonnel_(rawValue) === normalizePersonnel_(selectedPersonnel);
}

function sumColumnByPersonnel_(sheet, amountColumnIndex, personnelColumnIndex, selectedPersonnel) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;

  const values = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  let total = 0;
  for (let i = 0; i < values.length; i += 1) {
    const row = values[i];
    if (!isPersonnelMatch_(row[personnelColumnIndex - 1], selectedPersonnel)) continue;
    total += Number(row[amountColumnIndex - 1]) || 0;
  }
  return total;
}

function manageOffGranted() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const selectedPersonnel = getSelectedPersonnel_(ss);
  const options = getManageOffGrantedOptions_();

  if (options.length === 0) {
    SpreadsheetApp.getUi().alert(`No Offs (Granted) records found for ${selectedPersonnel}.`);
    return;
  }

  const html = HtmlService
    .createHtmlOutput(getManageOffGrantedDialogHtml_(options, selectedPersonnel))
    .setWidth(760)
    .setHeight(720);

  SpreadsheetApp.getUi().showModalDialog(html, 'Edit Off Granted');
}

function getManageOffGrantedOptions_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const selectedPersonnel = getSelectedPersonnel_(ss);
  const records = getGrantedRecordsForSelectedPersonnel_(ss, selectedPersonnel);

  return records.map((r) => {
    const reasonRaw = String(r.reasonType || '').trim().toUpperCase();
    const isOps = reasonRaw === 'OPS';
    const durationType = Number(r.durationValue) >= 1 ? 'FULL' : 'HALF';

    return {
      id: r.id,
      label: `${r.id}, ${formatDuration_(Number(r.remaining) || 0)} day remaining, ${r.reasonDetails || '-'}`,
      grantedDate: normalizeDateInputString_(r.grantedDate),
      durationType,
      reasonType: isOps ? 'OPS' : 'OTHERS',
      weekendOpsDate: isOps ? normalizeDateInputString_(r.weekendOpsDutyDate) : '',
      otherDetails: isOps ? '' : String(r.reasonDetails || ''),
      providedBy: String(r.providedBy || '')
    };
  });
}

function submitManageOffGranted(form) {
  const selectedIds = Array.isArray(form && form.selectedIds)
    ? form.selectedIds.map((id) => String(id || '').trim()).filter((id) => id)
    : [];

  if (selectedIds.length !== 1) {
    return { ok: false, message: 'Tick exactly one OFF ID to edit.' };
  }

  return submitEditOffGranted({
    id: selectedIds[0],
    grantedDate: String(form.grantedDate || '').trim(),
    durationType: String(form.durationType || '').trim(),
    reasonType: String(form.reasonType || '').trim(),
    weekendOpsDate: String(form.weekendOpsDate || '').trim(),
    otherDetails: String(form.otherDetails || '').trim(),
    providedBy: String(form.providedBy || '').trim()
  });
}

function deleteOffGranted() {
  const selectedPersonnel = getSelectedPersonnel_(SpreadsheetApp.getActiveSpreadsheet());
  const initialOptions = getDeleteOffGrantedOptions();
  const html = HtmlService
    .createHtmlOutput(getDeleteOffGrantedDialogHtml_(initialOptions, selectedPersonnel))
    .setWidth(560)
    .setHeight(620);

  SpreadsheetApp.getUi().showModalDialog(html, 'Delete Off Granted');
}

function getDeleteOffGrantedOptions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const granted = getOrCreateGrantedSheet_(ss);
  const selectedPersonnel = getSelectedPersonnel_(ss);
  const lastRow = granted.getLastRow();
  if (lastRow < 2) return [];

  const values = granted.getRange(2, 1, lastRow - 1, UNUSED_HEADERS.length).getValues();
  const options = [];
  for (let i = 0; i < values.length; i += 1) {
    const row = values[i];
    const id = String(row[0] || '').trim();
    if (!id) continue;
    if (!isPersonnelMatch_(row[12], selectedPersonnel)) continue;

    const durationValue = roundDuration_(toNumber_(row[3]) || 0);
    const usedValue = roundDuration_(toNumber_(row[8]) || 0);
    const remainingValue = roundDuration_(toNumber_(row[9]));
    const effectiveRemaining = Number.isNaN(remainingValue)
      ? roundDuration_(Math.max(durationValue - usedValue, 0))
      : remainingValue;

    if (usedValue > 1e-9) continue;

    const reasonType = String(row[4] || '').trim().toLowerCase();
    let detail = '';
    if (reasonType === 'ops') {
      detail = `Weekend Ops on (${formatDateYmdSafe_(row[5])})`;
    } else {
      const provider = String(row[7] || '').trim() || 'Unknown';
      const reason = String(row[6] || '').trim() || 'No details';
      detail = `Off provided by (${provider}) For (${reason})`;
    }

    options.push({
      id,
      label: `${id}, ${formatDuration_(effectiveRemaining)} day, ${detail}`
    });
  }

  options.sort((a, b) => a.id.localeCompare(b.id));
  return options;
}

function submitDeleteOffGrantedBatch(form) {
  const selectedIds = Array.isArray(form && form.selectedIds)
    ? form.selectedIds.map((id) => String(id || '').trim()).filter((id) => id)
    : [];

  if (selectedIds.length === 0) {
    return { ok: false, message: 'Please tick at least one OFF ID to delete.' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const granted = getOrCreateGrantedSheet_(ss);
  const selectedPersonnel = getSelectedPersonnel_(ss);
  const lastRow = granted.getLastRow();
  if (lastRow < 2) return { ok: false, message: 'No Offs (Granted) records found.' };

  const values = granted.getRange(2, 1, lastRow - 1, UNUSED_HEADERS.length).getValues();
  const toDeleteRows = [];
  const foundMap = {};
  const seen = {};
  const deletedSnapshots = [];

  for (let i = 0; i < selectedIds.length; i += 1) {
    if (!seen[selectedIds[i]]) seen[selectedIds[i]] = true;
  }

  for (let i = 0; i < values.length; i += 1) {
    const row = values[i];
    const id = String(row[0] || '').trim();
    if (!id || !seen[id]) continue;
    if (!isPersonnelMatch_(row[12], selectedPersonnel)) continue;

    const usedValue = roundDuration_(toNumber_(row[8]) || 0);
    if (usedValue > 1e-9) {
      return { ok: false, message: `Cannot delete ${id}. It already has used amount ${formatDuration_(usedValue)}.` };
    }

    foundMap[id] = true;
    toDeleteRows.push(i + 2);
    deletedSnapshots.push(grantedRowSnapshotFromValues_(row));
  }

  for (let i = 0; i < selectedIds.length; i += 1) {
    if (!foundMap[selectedIds[i]]) {
      return { ok: false, message: `OFF ID ${selectedIds[i]} not found for selected personnel.` };
    }
  }

  deleteRowsDescending_(granted, toDeleteRows);
  appendEditLog_(ss, {
    action: 'DELETE_GRANTED',
    personnel: selectedPersonnel,
    recordType: 'Off Granted',
    recordId: selectedIds.join(', '),
    summary: `Deleted ${toDeleteRows.length} Offs (Granted): ${selectedIds.join(', ')}.`,
    before: deletedSnapshots,
    after: { deleted: true }
  });
  applyGrantedStatusHighlighting_(granted);
  refreshDashboard();
  refreshCalendar();

  return { ok: true, message: `Deleted ${toDeleteRows.length} Off Granted record(s).` };
}

function manageOffUsed() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const selectedPersonnel = getSelectedPersonnel_(ss);
  const usedOptions = getManageOffUsedOptions_();

  if (usedOptions.length === 0) {
    SpreadsheetApp.getUi().alert(`No Offs (Used) records found for ${selectedPersonnel}.`);
    return;
  }

  const additionalOptions = getUseOffOptions();
  const html = HtmlService
    .createHtmlOutput(getManageOffUsedDialogHtml_(usedOptions, additionalOptions, selectedPersonnel))
    .setWidth(780)
    .setHeight(760);

  SpreadsheetApp.getUi().showModalDialog(html, 'Edit/Undo Off Used');
}

function getManageOffUsedOptions_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const selectedPersonnel = getSelectedPersonnel_(ss);
  const records = getUsedRecordsForSelectedPersonnel_(ss, selectedPersonnel);

  return records.map((r) => {
    const sessionRaw = normalizeSessionRaw_(r.session);
    return {
      useId: r.useId,
      label: `${r.useId}, ${r.intendedDate}, ${r.session}, ${formatDuration_(Number(r.duration) || 0)} day`,
      intendedDate: normalizeDateInputString_(r.intendedDate),
      session: sessionRaw,
      duration: roundDuration_(Number(r.duration) || 0),
      comments: String(r.comments || '')
    };
  });
}

function submitManageOffUsed(form) {
  const selectedIds = Array.isArray(form && form.selectedIds)
    ? form.selectedIds.map((id) => String(id || '').trim()).filter((id) => id)
    : [];
  if (selectedIds.length !== 1) {
    return { ok: false, message: 'Tick exactly one Use ID.' };
  }

  const action = String(form && form.action || 'EDIT').trim().toUpperCase();
  const useId = selectedIds[0];
  if (action === 'UNDO') {
    return submitUndoOffUsed({ useId });
  }

  return submitEditOffUsed({
    useId,
    intendedDate: String(form.intendedDate || '').trim(),
    session: String(form.session || '').trim(),
    comments: String(form.comments || '').trim(),
    additionalIds: Array.isArray(form.additionalIds) ? form.additionalIds : []
  });
}

function getGrantedRecordsForSelectedPersonnel_(ss, selectedPersonnel) {
  const sheet = getOrCreateGrantedSheet_(ss);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const values = sheet.getRange(2, 1, lastRow - 1, UNUSED_HEADERS.length).getValues();
  const out = [];

  for (let i = 0; i < values.length; i += 1) {
    const row = values[i];
    const id = String(row[0] || '').trim();
    if (!id) continue;
    if (!isPersonnelMatch_(row[12], selectedPersonnel)) continue;

    out.push({
      rowIndex: i + 2,
      id,
      grantedDate: formatDateYmdSafe_(row[1]),
      durationType: String(row[2] || ''),
      durationValue: toNumber_(row[3]),
      reasonType: String(row[4] || ''),
      weekendOpsDutyDate: formatDateYmdSafe_(row[5]),
      reasonDetails: String(row[6] || ''),
      providedBy: String(row[7] || ''),
      used: toNumber_(row[8]),
      remaining: toNumber_(row[9]),
      status: String(row[10] || '')
    });
  }

  out.sort((a, b) => a.id.localeCompare(b.id));
  return out;
}

function submitEditOffGranted(form) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const granted = getOrCreateGrantedSheet_(ss);
  const selectedPersonnel = getSelectedPersonnel_(ss);
  const id = String(form && form.id || '').trim();
  const rec = findGrantedRowById_(granted, id, selectedPersonnel);

  if (!rec) {
    return { ok: false, message: `OFF ID ${id} not found for selected personnel.` };
  }
  const beforeSnapshot = grantedRowSnapshotFromValues_(rec.rawRow);

  const grantedDate = parseDateInput_(String(form.grantedDate || '').trim());
  if (!grantedDate) {
    return { ok: false, message: 'Invalid date. Use YYYY-MM-DD.' };
  }

  const durationTypeRaw = String(form.durationType || '').trim().toUpperCase();
  const reasonTypeRaw = String(form.reasonType || '').trim().toUpperCase();
  const weekendOpsDate = parseDateInput_(String(form.weekendOpsDate || '').trim());
  const otherDetails = String(form.otherDetails || '').trim();
  let providedBy = String(form.providedBy || '').trim();

  let durationType;
  let durationValue;
  if (durationTypeRaw === 'FULL') {
    durationType = 'Full Day';
    durationValue = 1;
  } else if (durationTypeRaw === 'HALF') {
    durationType = 'Half Day';
    durationValue = 0.5;
  } else {
    return { ok: false, message: 'Duration must be FULL or HALF.' };
  }

  const usedValue = roundDuration_(toNumber_(rec.rawRow[8]) || 0);
  if (durationValue + 1e-9 < usedValue) {
    return {
      ok: false,
      message: `Cannot reduce duration below already used amount (${formatDuration_(usedValue)}).`
    };
  }

  let reasonType;
  let reasonDetails;
  let weekendOpsDutyDate = '';
  if (reasonTypeRaw === 'OPS') {
    reasonType = 'Ops';
    if (!weekendOpsDate) {
      return { ok: false, message: 'Please provide the Weekend Ops duty date.' };
    }
    if (!isWeekendDate_(weekendOpsDate)) {
      return { ok: false, message: 'Weekend Ops duty date must be Saturday or Sunday.' };
    }
    weekendOpsDutyDate = weekendOpsDate;
    reasonDetails = `Weekend Ops on ${formatDateYmd_(weekendOpsDate)}`;
    if (!providedBy) providedBy = 'Yourself';
  } else if (reasonTypeRaw === 'OTHERS') {
    reasonType = 'Others';
    if (!otherDetails) {
      return { ok: false, message: 'Please provide comments/details for Others.' };
    }
    if (!providedBy) {
      return { ok: false, message: 'Please fill in "Provided by who".' };
    }
    reasonDetails = otherDetails;
  } else {
    return { ok: false, message: 'Reason must be OPS or OTHERS.' };
  }

  const remainingValue = roundDuration_(durationValue - usedValue);
  const status = computeGrantedStatus_(usedValue, remainingValue);

  granted.getRange(rec.rowIndex, 2, 1, 10).setValues([[
    grantedDate,
    durationType,
    durationValue,
    reasonType,
    weekendOpsDutyDate,
    reasonDetails,
    providedBy,
    usedValue,
    remainingValue,
    status
  ]]);

  granted.getRange(rec.rowIndex, 2).setNumberFormat('yyyy-mm-dd');
  granted.getRange(rec.rowIndex, 4).setNumberFormat('0.0');
  granted.getRange(rec.rowIndex, 6).setNumberFormat('yyyy-mm-dd');
  granted.getRange(rec.rowIndex, 9, 1, 2).setNumberFormat('0.0');
  applyGrantedStatusHighlighting_(granted);

  const afterRow = [
    id,
    grantedDate,
    durationType,
    durationValue,
    reasonType,
    weekendOpsDutyDate,
    reasonDetails,
    providedBy,
    usedValue,
    remainingValue,
    status,
    rec.rawRow[11],
    rec.rawRow[12]
  ];
  appendEditLog_(ss, {
    action: 'EDIT_GRANTED',
    personnel: selectedPersonnel,
    recordType: 'Off Granted',
    recordId: id,
    summary: buildGrantedEditSummary_(beforeSnapshot, grantedRowSnapshotFromValues_(afterRow)),
    before: beforeSnapshot,
    after: grantedRowSnapshotFromValues_(afterRow)
  });

  refreshDashboard();
  refreshCalendar();

  return { ok: true, message: `Updated ${id}.` };
}

function submitDeleteOffGranted(form) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const granted = getOrCreateGrantedSheet_(ss);
  const selectedPersonnel = getSelectedPersonnel_(ss);
  const id = String(form && form.id || '').trim();
  const rec = findGrantedRowById_(granted, id, selectedPersonnel);

  if (!rec) {
    return { ok: false, message: `OFF ID ${id} not found for selected personnel.` };
  }

  const usedValue = roundDuration_(toNumber_(rec.rawRow[8]) || 0);
  if (usedValue > 1e-9) {
    return { ok: false, message: `Cannot delete ${id}. It already has used amount ${formatDuration_(usedValue)}.` };
  }

  const beforeSnapshot = grantedRowSnapshotFromValues_(rec.rawRow);
  granted.deleteRow(rec.rowIndex);
  appendEditLog_(ss, {
    action: 'DELETE_GRANTED',
    personnel: selectedPersonnel,
    recordType: 'Off Granted',
    recordId: id,
    summary: `Deleted ${id} (Date ${beforeSnapshot.dateOffGranted}, Duration ${beforeSnapshot.durationType} ${beforeSnapshot.durationValue}, Reason ${beforeSnapshot.reasonType}).`,
    before: beforeSnapshot,
    after: { deleted: true }
  });
  applyGrantedStatusHighlighting_(granted);
  refreshDashboard();
  refreshCalendar();

  return { ok: true, message: `Deleted ${id}.` };
}

function getUsedRecordsForSelectedPersonnel_(ss, selectedPersonnel) {
  const sheet = getOrCreateSheet_(ss, SHEET_USED);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const values = sheet.getRange(2, 1, lastRow - 1, USED_HEADERS.length).getValues();
  const out = [];
  for (let i = 0; i < values.length; i += 1) {
    const row = values[i];
    const useId = String(row[0] || '').trim();
    if (!useId) continue;
    if (!isPersonnelMatch_(row[7], selectedPersonnel)) continue;

    out.push({
      rowIndex: i + 2,
      useId,
      intendedDate: formatDateYmdSafe_(row[1]),
      session: String(row[2] || ''),
      duration: toNumber_(row[3]),
      offIdsUsed: String(row[4] || ''),
      comments: String(row[5] || ''),
      rawRow: row
    });
  }

  out.sort((a, b) => a.useId.localeCompare(b.useId));
  return out;
}

function submitEditOffUsed(form) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const selectedPersonnel = getSelectedPersonnel_(ss);
  const usedSheet = getOrCreateSheet_(ss, SHEET_USED);
  const grantedSheet = getOrCreateGrantedSheet_(ss);

  const useId = String(form && form.useId || '').trim();
  const usedRec = findUsedRowByUseId_(usedSheet, useId, selectedPersonnel);
  if (!usedRec) {
    return { ok: false, message: `Use ID ${useId} not found for selected personnel.` };
  }
  const beforeSnapshot = usedRowSnapshotFromValues_(usedRec.rawRow);

  const intendedDate = parseDateInput_(String(form.intendedDate || '').trim());
  if (!intendedDate) {
    return { ok: false, message: 'Invalid intended date. Use YYYY-MM-DD.' };
  }

  const sessionRaw = String(form.session || '').trim().toUpperCase();
  let session;
  let targetDuration;
  if (sessionRaw === 'FULL') {
    session = 'Full Day';
    targetDuration = 1;
  } else if (sessionRaw === 'AM') {
    session = 'AM';
    targetDuration = 0.5;
  } else if (sessionRaw === 'PM') {
    session = 'PM';
    targetDuration = 0.5;
  } else {
    return { ok: false, message: 'Session must be FULL, AM, or PM.' };
  }

  const currentDuration = roundDuration_(toNumber_(usedRec.rawRow[3]) || 0);
  const comments = String(form.comments || '').trim();
  const allocations = parseOffAllocations_(String(usedRec.rawRow[4] || ''));
  if (allocations.length === 0) {
    return { ok: false, message: `Unable to parse Off IDs Used for ${useId}.` };
  }

  const grantedMap = getUnusedRecordMap_(grantedSheet, selectedPersonnel);
  const touchedIds = {};
  const delta = roundDuration_(targetDuration - currentDuration);

  if (delta > 1e-9) {
    const additionalIds = Array.isArray(form.additionalIds)
      ? form.additionalIds.map((x) => String(x || '').trim()).filter((x) => x)
      : [];
    if (additionalIds.length === 0) {
      return {
        ok: false,
        message: `Need additional ${formatDuration_(delta)} day. Please provide more OFF ID(s).`
      };
    }

    let remainingNeed = delta;
    for (let i = 0; i < additionalIds.length; i += 1) {
      if (remainingNeed <= 1e-9) break;
      const id = additionalIds[i];
      const rec = grantedMap[id];
      if (!rec || rec.remaining <= 1e-9) continue;

      const useAmount = roundDuration_(Math.min(rec.remaining, remainingNeed));
      rec.used = roundDuration_(rec.used + useAmount);
      rec.remaining = roundDuration_(rec.remaining - useAmount);
      rec.status = computeGrantedStatus_(rec.used, rec.remaining);
      touchedIds[id] = true;
      mergeAllocation_(allocations, id, useAmount);
      remainingNeed = roundDuration_(remainingNeed - useAmount);
    }

    if (remainingNeed > 1e-9) {
      return {
        ok: false,
        message: `Additional OFF IDs are insufficient. Still need ${formatDuration_(remainingNeed)} day.`
      };
    }
  } else if (delta < -1e-9) {
    let releaseNeed = roundDuration_(-delta);
    for (let i = allocations.length - 1; i >= 0 && releaseNeed > 1e-9; i -= 1) {
      const alloc = allocations[i];
      if (alloc.amount <= 1e-9) continue;

      const releaseAmount = roundDuration_(Math.min(alloc.amount, releaseNeed));
      alloc.amount = roundDuration_(alloc.amount - releaseAmount);

      const rec = grantedMap[alloc.id];
      if (!rec) {
        return { ok: false, message: `Granted row not found for OFF ID ${alloc.id}.` };
      }

      rec.used = roundDuration_(Math.max(rec.used - releaseAmount, 0));
      rec.remaining = roundDuration_(rec.remaining + releaseAmount);
      rec.status = computeGrantedStatus_(rec.used, rec.remaining);
      touchedIds[alloc.id] = true;
      releaseNeed = roundDuration_(releaseNeed - releaseAmount);
    }

    if (releaseNeed > 1e-9) {
      return { ok: false, message: `Unable to release enough allocation for ${useId}.` };
    }
  }

  for (const id in touchedIds) {
    if (!Object.prototype.hasOwnProperty.call(touchedIds, id)) continue;
    const rec = grantedMap[id];
    if (!rec) continue;
    grantedSheet.getRange(rec.rowIndex, 9, 1, 3).setValues([[
      roundDuration_(rec.used),
      roundDuration_(rec.remaining),
      rec.status
    ]]);
  }

  const normalizedAllocations = allocations.filter((a) => a.amount > 1e-9);
  const offIdsUsed = formatOffAllocations_(normalizedAllocations);
  if (!offIdsUsed) {
    return { ok: false, message: 'No OFF IDs remain allocated after edit.' };
  }

  usedSheet.getRange(usedRec.rowIndex, 2, 1, 5).setValues([[
    intendedDate,
    session,
    targetDuration,
    offIdsUsed,
    comments
  ]]);
  usedSheet.getRange(usedRec.rowIndex, 2).setNumberFormat('yyyy-mm-dd');
  usedSheet.getRange(usedRec.rowIndex, 4).setNumberFormat('0.0');

  const afterRow = [
    useId,
    intendedDate,
    session,
    targetDuration,
    offIdsUsed,
    comments,
    usedRec.rawRow[6],
    usedRec.rawRow[7]
  ];
  appendEditLog_(ss, {
    action: 'EDIT_USED',
    personnel: selectedPersonnel,
    recordType: 'Off Used',
    recordId: useId,
    summary: buildUsedEditSummary_(beforeSnapshot, usedRowSnapshotFromValues_(afterRow)),
    before: beforeSnapshot,
    after: usedRowSnapshotFromValues_(afterRow)
  });

  applyGrantedStatusHighlighting_(grantedSheet);
  refreshDashboard();
  refreshCalendar();

  return {
    ok: true,
    message: `Updated ${useId} to ${session} (${formatDuration_(targetDuration)} day).`
  };
}

function submitUndoOffUsed(form) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const selectedPersonnel = getSelectedPersonnel_(ss);
  const usedSheet = getOrCreateSheet_(ss, SHEET_USED);
  const grantedSheet = getOrCreateGrantedSheet_(ss);

  const useId = String(form && form.useId || '').trim();
  const usedRec = findUsedRowByUseId_(usedSheet, useId, selectedPersonnel);
  if (!usedRec) {
    return { ok: false, message: `Use ID ${useId} not found for selected personnel.` };
  }
  const beforeSnapshot = usedRowSnapshotFromValues_(usedRec.rawRow);

  const allocations = parseOffAllocations_(String(usedRec.rawRow[4] || ''));
  if (allocations.length === 0) {
    return { ok: false, message: `Unable to parse Off IDs Used for ${useId}.` };
  }

  const grantedMap = getUnusedRecordMap_(grantedSheet, selectedPersonnel);
  const touchedIds = {};
  for (let i = 0; i < allocations.length; i += 1) {
    const alloc = allocations[i];
    const rec = grantedMap[alloc.id];
    if (!rec) {
      return { ok: false, message: `Granted row not found for OFF ID ${alloc.id}.` };
    }

    rec.used = roundDuration_(Math.max(rec.used - alloc.amount, 0));
    rec.remaining = roundDuration_(rec.remaining + alloc.amount);
    rec.status = computeGrantedStatus_(rec.used, rec.remaining);
    touchedIds[alloc.id] = true;
  }

  for (const id in touchedIds) {
    if (!Object.prototype.hasOwnProperty.call(touchedIds, id)) continue;
    const rec = grantedMap[id];
    if (!rec) continue;
    grantedSheet.getRange(rec.rowIndex, 9, 1, 3).setValues([[
      roundDuration_(rec.used),
      roundDuration_(rec.remaining),
      rec.status
    ]]);
  }

  usedSheet.deleteRow(usedRec.rowIndex);
  appendEditLog_(ss, {
    action: 'UNDO_USED',
    personnel: selectedPersonnel,
    recordType: 'Off Used',
    recordId: useId,
    summary: `Undid ${useId}. Restored ${beforeSnapshot.durationUsed} day from: ${beforeSnapshot.offIdsUsed}. Intended date was ${beforeSnapshot.dateIntended} (${beforeSnapshot.session}).`,
    before: beforeSnapshot,
    after: { deleted: true }
  });
  applyGrantedStatusHighlighting_(grantedSheet);
  refreshDashboard();
  refreshCalendar();

  return { ok: true, message: `Undid ${useId}. Allocated Off balance has been restored.` };
}

function findGrantedRowById_(sheet, id, selectedPersonnel) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  const values = sheet.getRange(2, 1, lastRow - 1, UNUSED_HEADERS.length).getValues();
  for (let i = 0; i < values.length; i += 1) {
    const row = values[i];
    if (String(row[0] || '').trim() !== id) continue;
    if (!isPersonnelMatch_(row[12], selectedPersonnel)) continue;
    return { rowIndex: i + 2, rawRow: row };
  }
  return null;
}

function findUsedRowByUseId_(sheet, useId, selectedPersonnel) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  const values = sheet.getRange(2, 1, lastRow - 1, USED_HEADERS.length).getValues();
  for (let i = 0; i < values.length; i += 1) {
    const row = values[i];
    if (String(row[0] || '').trim() !== useId) continue;
    if (!isPersonnelMatch_(row[7], selectedPersonnel)) continue;
    return { rowIndex: i + 2, rawRow: row };
  }
  return null;
}

function parseOffAllocations_(text) {
  const out = [];
  const source = String(text || '');
  const re = /([A-Za-z]-\d+)\s*\((\d+(?:\.\d+)?)\)/g;
  let m;
  while ((m = re.exec(source)) !== null) {
    out.push({
      id: m[1],
      amount: roundDuration_(Number(m[2]) || 0)
    });
  }
  return out.filter((a) => a.id && a.amount > 0);
}

function mergeAllocation_(allocations, id, amount) {
  if (!id || amount <= 0) return;
  for (let i = 0; i < allocations.length; i += 1) {
    if (allocations[i].id === id) {
      allocations[i].amount = roundDuration_(allocations[i].amount + amount);
      return;
    }
  }
  allocations.push({ id, amount: roundDuration_(amount) });
}

function formatOffAllocations_(allocations) {
  if (!Array.isArray(allocations) || allocations.length === 0) return '';
  return allocations
    .filter((a) => a.id && a.amount > 0)
    .map((a) => `${a.id} (${formatDuration_(roundDuration_(a.amount))})`)
    .join(' + ');
}

function computeGrantedStatus_(usedValue, remainingValue) {
  const used = roundDuration_(usedValue || 0);
  const remaining = roundDuration_(remainingValue || 0);
  if (used <= 1e-9) return 'Unused';
  if (remaining <= 1e-9) return 'Used';
  return 'Partial';
}

function roundDuration_(value) {
  return Math.round((Number(value) || 0) * 10) / 10;
}

function findRowsByPersonnel_(sheet, personnelColumnIndex, selectedPersonnel) {
  const rows = [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2 || sheet.getLastColumn() < personnelColumnIndex) return rows;

  const values = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  for (let i = 0; i < values.length; i += 1) {
    if (isPersonnelMatch_(values[i][personnelColumnIndex - 1], selectedPersonnel)) {
      rows.push(i + 2);
    }
  }
  return rows;
}

function deleteRowsDescending_(sheet, rows) {
  if (!Array.isArray(rows) || rows.length === 0) return;
  const sorted = rows.slice().sort((a, b) => b - a);
  for (let i = 0; i < sorted.length; i += 1) {
    sheet.deleteRow(sorted[i]);
  }
}

function getManagedBackupPairs_() {
  return [
    [SHEET_GRANTED, SHEET_BKP_GRANTED],
    [SHEET_USED, SHEET_BKP_USED],
    [SHEET_CALENDAR, SHEET_BKP_CALENDAR],
    [SHEET_EDIT_LOGS, SHEET_BKP_LOGS]
  ];
}

function resolveManagedSheet_(ss, sheetName) {
  if (sheetName === SHEET_GRANTED) {
    return getOrCreateGrantedSheet_(ss);
  }
  if (sheetName === SHEET_EDIT_LOGS) {
    return getOrCreateEditLogsSheet_(ss);
  }
  return getOrCreateSheet_(ss, sheetName);
}

function getOrCreateBackupSheet_(ss, backupName) {
  const sh = getOrCreateSheet_(ss, backupName);
  if (!sh.isSheetHidden()) sh.hideSheet();
  return sh;
}

function syncManagedBackups_(ss) {
  const pairs = getManagedBackupPairs_();
  for (let i = 0; i < pairs.length; i += 1) {
    const sourceName = pairs[i][0];
    const backupName = pairs[i][1];
    syncManagedBackupForSheet_(ss, sourceName, backupName);
  }
}

function syncManagedBackupForSheet_(ss, sourceName, backupName) {
  const source = resolveManagedSheet_(ss, sourceName);
  const backup = getOrCreateBackupSheet_(ss, backupName);

  const rows = source.getMaxRows();
  const cols = source.getMaxColumns();
  ensureSheetDimensions_(backup, rows, cols);
  backup.clear();
  source.getRange(1, 1, rows, cols).copyTo(backup.getRange(1, 1, rows, cols), { contentsOnly: false });
  if (!backup.isSheetHidden()) backup.hideSheet();
}

function restoreManagedSheetFromBackup_(ss, sourceName) {
  const pairs = getManagedBackupPairs_();
  let backupName = '';
  for (let i = 0; i < pairs.length; i += 1) {
    if (pairs[i][0] === sourceName) {
      backupName = pairs[i][1];
      break;
    }
  }
  if (!backupName) return false;

  const backup = ss.getSheetByName(backupName);
  if (!backup) return false;

  const source = resolveManagedSheet_(ss, sourceName);
  const rows = backup.getMaxRows();
  const cols = backup.getMaxColumns();
  ensureSheetDimensions_(source, rows, cols);
  source.clear();
  backup.getRange(1, 1, rows, cols).copyTo(source.getRange(1, 1, rows, cols), { contentsOnly: false });

  if (sourceName === SHEET_GRANTED) {
    applyGrantedStatusHighlighting_(source);
  }
  return true;
}

function restoreAllManagedSheetsFromBackup_(ss) {
  const pairs = getManagedBackupPairs_();
  for (let i = 0; i < pairs.length; i += 1) {
    restoreManagedSheetFromBackup_(ss, pairs[i][0]);
  }
}

function getOrCreateEditLogsSheet_(ss) {
  const sh = getOrCreateSheet_(ss, SHEET_EDIT_LOGS);
  if (sh.getMaxColumns() < EDIT_LOG_HEADERS.length) {
    sh.insertColumnsAfter(sh.getMaxColumns(), EDIT_LOG_HEADERS.length - sh.getMaxColumns());
  }
  if (sh.getMaxRows() < 2) {
    sh.insertRowsAfter(sh.getMaxRows(), 2 - sh.getMaxRows());
  }
  sh.getRange(1, 1, 1, EDIT_LOG_HEADERS.length).setValues([EDIT_LOG_HEADERS]).setFontWeight('bold');
  sh.setFrozenRows(1);
  sh.getRange('B:B').setNumberFormat('yyyy-mm-dd hh:mm:ss');
  return sh;
}

function appendEditLog_(ss, payload) {
  const sh = getOrCreateEditLogsSheet_(ss);
  const nextRow = sh.getLastRow() + 1;
  const logId = `L-${String(nextRow - 1).padStart(5, '0')}`;
  const beforeText = truncateLogText_(safeJson_(payload.before || {}), 45000);
  const afterText = truncateLogText_(safeJson_(payload.after || {}), 45000);

  sh.getRange(nextRow, 1, 1, EDIT_LOG_HEADERS.length).setValues([[
    logId,
    new Date(),
    String(payload.action || ''),
    String(payload.personnel || ''),
    String(payload.recordType || ''),
    String(payload.recordId || ''),
    String(payload.summary || ''),
    beforeText,
    afterText,
    getEditorEmail_()
  ]]);
  syncManagedBackupForSheet_(ss, SHEET_EDIT_LOGS, SHEET_BKP_LOGS);
}

function buildChangeSummary_(before, after, fieldLabels) {
  const labels = fieldLabels || {};
  const parts = [];
  const keys = Object.keys(labels);

  for (let i = 0; i < keys.length; i += 1) {
    const key = keys[i];
    const beforeText = normalizeSummaryValue_(before ? before[key] : '');
    const afterText = normalizeSummaryValue_(after ? after[key] : '');
    if (beforeText === afterText) continue;
    parts.push(`${labels[key]}: ${beforeText} -> ${afterText}`);
  }

  if (parts.length === 0) return 'No field changes detected.';
  return truncateLogText_(parts.join(' | '), 1200);
}

function normalizeSummaryValue_(value) {
  if (value === null || typeof value === 'undefined') return '-';
  const text = String(value).trim();
  return text || '-';
}

function buildGrantedEditSummary_(before, after) {
  return buildChangeSummary_(before, after, {
    dateOffGranted: 'Date Off Granted',
    durationType: 'Duration Type',
    durationValue: 'Duration Value',
    reasonType: 'Reason Type',
    weekendOpsDutyDate: 'Weekend Ops Duty Date',
    reasonDetails: 'Reason Details',
    providedBy: 'Provided By',
    usedValue: 'Used Value',
    remainingValue: 'Remaining Value',
    status: 'Status'
  });
}

function buildUsedEditSummary_(before, after) {
  return buildChangeSummary_(before, after, {
    dateIntended: 'Date Intended',
    session: 'Session',
    durationUsed: 'Duration Used',
    offIdsUsed: 'Off IDs Used',
    comments: 'Comments'
  });
}

function safeJson_(value) {
  try {
    return JSON.stringify(value);
  } catch (err) {
    return String(value);
  }
}

function truncateLogText_(text, limit) {
  const value = String(text || '');
  if (value.length <= limit) return value;
  return `${value.slice(0, limit - 3)}...`;
}

function getEditorEmail_() {
  const active = Session.getActiveUser();
  const effective = Session.getEffectiveUser();
  return (active && active.getEmail && active.getEmail()) ||
    (effective && effective.getEmail && effective.getEmail()) ||
    '';
}

function grantedRowSnapshotFromValues_(row) {
  return {
    id: String(row[0] || ''),
    dateOffGranted: normalizeDateInputString_(row[1]),
    durationType: String(row[2] || ''),
    durationValue: roundDuration_(toNumber_(row[3]) || 0),
    reasonType: String(row[4] || ''),
    weekendOpsDutyDate: normalizeDateInputString_(row[5]),
    reasonDetails: String(row[6] || ''),
    providedBy: String(row[7] || ''),
    usedValue: roundDuration_(toNumber_(row[8]) || 0),
    remainingValue: roundDuration_(toNumber_(row[9]) || 0),
    status: String(row[10] || ''),
    createdAt: normalizeDateInputString_(row[11]),
    personnel: String(row[12] || '')
  };
}

function usedRowSnapshotFromValues_(row) {
  return {
    useId: String(row[0] || ''),
    dateIntended: normalizeDateInputString_(row[1]),
    session: String(row[2] || ''),
    durationUsed: roundDuration_(toNumber_(row[3]) || 0),
    offIdsUsed: String(row[4] || ''),
    comments: String(row[5] || ''),
    createdAt: normalizeDateInputString_(row[6]),
    personnel: String(row[7] || '')
  };
}

function escapeHtml_(value) {
  return String(value === null || value === undefined ? '' : value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function getOrCreateGrantedSheet_(ss) {
  let sheet = ss.getSheetByName(SHEET_GRANTED);
  if (sheet) return sheet;

  const legacy = ss.getSheetByName(SHEET_GRANTED_LEGACY);
  if (legacy) {
    legacy.setName(SHEET_GRANTED);
    return legacy;
  }

  sheet = ss.insertSheet(SHEET_GRANTED);
  return sheet;
}

function applyGrantedStatusHighlighting_(sheet) {
  const range = sheet.getRange('A2:M');

  const greenRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($A2<>"",$J2>0)')
    .setBackground('#e7f7ec')
    .setRanges([range])
    .build();

  const redRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($A2<>"",$J2=0,$I2>0)')
    .setBackground('#fdecec')
    .setRanges([range])
    .build();

  sheet.setConditionalFormatRules([greenRule, redRule]);
}

function getAddOffDayDialogHtml_(selectedPersonnel) {
  const personnelLabel = escapeHtml_(selectedPersonnel || DEFAULT_PERSONNEL);
  return `
<!DOCTYPE html>
<html>
  <head>
    <base target="_top">
    <style>
      body { font-family: Arial, sans-serif; margin: 16px; color: #1f2937; }
      h3 { margin: 0 0 12px 0; }
      label { display: block; font-weight: 600; margin-top: 10px; margin-bottom: 4px; }
      input, select, textarea { width: 100%; box-sizing: border-box; padding: 8px; }
      textarea { min-height: 70px; resize: vertical; }
      .hidden { display: none; }
      .actions { margin-top: 14px; display: flex; gap: 8px; }
      button { padding: 8px 12px; cursor: pointer; }
      button:disabled { background: #d1d5db; color: #6b7280; cursor: not-allowed; }
      .error { margin-top: 10px; color: #b91c1c; white-space: pre-wrap; }
      .hint { color: #6b7280; font-size: 12px; margin-top: 4px; }
    </style>
  </head>
  <body>
    <h3>Add Off Day</h3>
    <div class="hint">Tracking Personnel: ${personnelLabel}</div>

    <label for="grantedDate">Enter date Off was granted</label>
    <input id="grantedDate" type="date" />

    <label for="durationType">Duration type</label>
    <select id="durationType">
      <option value="FULL">Full Day (1)</option>
      <option value="HALF">Half Day (0.5)</option>
    </select>

    <label for="reasonType">Reason</label>
    <select id="reasonType" onchange="toggleReasonInputs()">
      <option value="OPS">Ops</option>
      <option value="OTHERS">Others</option>
    </select>

    <div id="opsFields">
      <label for="weekendOpsDate">When was your Weekend Ops duty?</label>
      <input id="weekendOpsDate" type="date" />
      <div class="hint">Must be a weekend date (Saturday/Sunday).</div>
    </div>

    <div id="othersFields" class="hidden">
      <label for="otherDetails">Provide comments/details (what it was for)</label>
      <textarea id="otherDetails" placeholder="Enter details..."></textarea>
    </div>

    <label for="providedBy">Provided by who? (Yourself if Ops related)</label>
    <input id="providedBy" type="text" placeholder="Name" />

    <div class="actions">
      <button id="submitBtn" type="button" onclick="submitForm()">Submit</button>
      <button onclick="google.script.host.close()">Cancel</button>
    </div>

    <div id="error" class="error"></div>

    <script>
      function todayYmd() {
        const d = new Date();
        const yyyy = d.getFullYear();
        const mm = String(d.getMonth() + 1).padStart(2, '0');
        const dd = String(d.getDate()).padStart(2, '0');
        return yyyy + '-' + mm + '-' + dd;
      }

      function toggleReasonInputs() {
        const reason = document.getElementById('reasonType').value;
        document.getElementById('opsFields').classList.toggle('hidden', reason !== 'OPS');
        document.getElementById('othersFields').classList.toggle('hidden', reason !== 'OTHERS');
      }

      function setError(message) {
        document.getElementById('error').textContent = message || '';
      }

      function setSubmitting(isSubmitting) {
        const btn = document.getElementById('submitBtn');
        if (!btn) return;
        btn.disabled = isSubmitting;
        btn.textContent = isSubmitting ? 'Submitting...' : 'Submit';
      }

      function submitForm() {
        if (window.__submitInFlight) return;
        window.__submitInFlight = true;
        setError('');
        setSubmitting(true);

        const payload = {
          grantedDate: document.getElementById('grantedDate').value,
          durationType: document.getElementById('durationType').value,
          reasonType: document.getElementById('reasonType').value,
          weekendOpsDate: document.getElementById('weekendOpsDate').value,
          otherDetails: document.getElementById('otherDetails').value,
          providedBy: document.getElementById('providedBy').value
        };

        if (!(window.google && google.script && google.script.run)) {
          setError('Apps Script runtime unavailable in dialog.');
          window.__submitInFlight = false;
          setSubmitting(false);
          return;
        }

        google.script.run
          .withFailureHandler((err) => {
            setError(err && err.message ? err.message : String(err));
            window.__submitInFlight = false;
            setSubmitting(false);
          })
          .withSuccessHandler((result) => {
            if (!result || !result.ok) {
              setError(result && result.message ? result.message : 'Failed to add off day.');
              window.__submitInFlight = false;
              setSubmitting(false);
              return;
            }
            alert(result.message || 'Added.');
            window.__submitInFlight = false;
            setSubmitting(false);
            google.script.host.close();
          })
          .submitAddOffDay(payload);
      }

      document.getElementById('grantedDate').value = todayYmd();
      toggleReasonInputs();
    </script>
  </body>
</html>`;
}

function getUseOffDayDialogHtml_(initialOptions, selectedPersonnel) {
  const options = Array.isArray(initialOptions) ? initialOptions : [];
  const personnelLabel = escapeHtml_(selectedPersonnel || DEFAULT_PERSONNEL);
  const checkboxTags = options.length > 0
    ? options
      .map((opt) => {
        const id = escapeHtml_(opt.id);
        const remaining = escapeHtml_(String(opt.remaining || 0));
        const label = escapeHtml_(opt.label);
        return `<label class="off-item"><input class="off-id-checkbox" type="checkbox" name="offIds" value="${id}" data-remaining="${remaining}"> ${label}</label>`;
      })
      .join('\n')
    : '<div class="off-empty">No available OFF IDs</div>';

  return `
<!DOCTYPE html>
<html>
  <head>
    <base target="_top">
    <style>
      body { font-family: Arial, sans-serif; margin: 16px; color: #1f2937; }
      h3 { margin: 0 0 12px 0; }
      label { display: block; font-weight: 600; margin-top: 10px; margin-bottom: 4px; }
      input, select, textarea { width: 100%; box-sizing: border-box; padding: 8px; }
      .off-list { min-height: 220px; max-height: 260px; overflow: auto; border: 1px solid #9ca3af; padding: 8px; border-radius: 4px; background: #fff; }
      .off-item { display: block; margin: 4px 0; line-height: 1.35; }
      .off-item input { width: auto; margin-right: 8px; vertical-align: middle; }
      .off-empty { color: #6b7280; }
      textarea { min-height: 80px; resize: vertical; }
      .actions { margin-top: 14px; display: flex; gap: 8px; }
      button { padding: 8px 12px; cursor: pointer; }
      .error { margin-top: 10px; color: #b91c1c; white-space: pre-wrap; }
      .hint { margin-top: 8px; color: #374151; font-size: 12px; white-space: pre-wrap; }
    </style>
  </head>
  <body>
    <h3>Use Off Day</h3>
    <div class="hint">Tracking Personnel: ${personnelLabel}</div>

    <label for="intendedDate">Enter date intended to use Off</label>
    <input id="intendedDate" type="date" />

    <label for="session">Session</label>
    <select id="session">
      <option value="FULL">Full Day (1)</option>
      <option value="AM">AM (0.5)</option>
      <option value="PM">PM (0.5)</option>
    </select>

    <label for="offIds">Which OFF ID number do you want to use?</label>
    <div id="offIdsList" class="off-list">
      ${checkboxTags}
    </div>
    <div class="hint">Tick one or more OFF IDs.</div>

    <div id="selectionHint" class="hint"></div>

    <label for="comments">Comments (optional)</label>
    <textarea id="comments" placeholder="Add comments..."></textarea>

    <div class="actions">
      <button id="submitBtn" type="button" onclick="submitForm()">Submit</button>
      <button id="cancelBtn" type="button" onclick="if(window.google&&google.script&&google.script.host){google.script.host.close();}">Cancel</button>
    </div>

    <div id="error" class="error"></div>

    <script>
      function todayYmd() {
        const d = new Date();
        const yyyy = d.getFullYear();
        const mm = String(d.getMonth() + 1).padStart(2, '0');
        const dd = String(d.getDate()).padStart(2, '0');
        return yyyy + '-' + mm + '-' + dd;
      }

      function setError(message) {
        document.getElementById('error').textContent = message || '';
      }

      function setSubmitting(isSubmitting) {
        const btn = document.getElementById('submitBtn');
        if (!btn) return;
        btn.disabled = isSubmitting;
        btn.textContent = isSubmitting ? 'Submitting...' : 'Submit';
      }

      function getSelectedOptions() {
        return Array.from(document.querySelectorAll('.off-id-checkbox:checked'));
      }

      function updateSelectionHint() {
        const session = document.getElementById('session').value;
        const needed = session === 'FULL' ? 1 : 0.5;

        const selected = getSelectedOptions();
        let total = 0;
        for (const opt of selected) {
          total += Number(opt.dataset.remaining || 0);
        }

        const hint = document.getElementById('selectionHint');
        const totalText = Number.isInteger(total) ? String(total) : total.toFixed(1);
        const neededText = Number.isInteger(needed) ? String(needed) : needed.toFixed(1);

        let message = 'Selected total: ' + totalText + ' day\\nRequired: ' + neededText + ' day';
        if (total < needed) {
          message += '\\nChoose another ID to meet required total.';
        }

        hint.textContent = message;
      }

      function submitForm() {
        if (window.__submitInFlight) {
          return;
        }
        window.__submitInFlight = true;
        setError('');
        setSubmitting(true);

        const selectedIds = getSelectedOptions().map((opt) => opt.value);
        if (selectedIds.length === 0) {
          setError('Please choose at least one OFF ID.');
          window.__submitInFlight = false;
          setSubmitting(false);
          return;
        }

        const payload = {
          intendedDate: document.getElementById('intendedDate').value,
          session: document.getElementById('session').value,
          selectedIds,
          comments: document.getElementById('comments').value
        };

        if (!(window.google && google.script && google.script.run)) {
          setError('Apps Script runtime unavailable in dialog.');
          window.__submitInFlight = false;
          setSubmitting(false);
          return;
        }

        google.script.run
          .withFailureHandler((err) => {
            const msg = err && err.message ? err.message : String(err);
            setError(msg);
            window.__submitInFlight = false;
            setSubmitting(false);
          })
          .withSuccessHandler((result) => {
            if (!result || !result.ok) {
              setError(result && result.message ? result.message : 'Failed to use off day.');
              window.__submitInFlight = false;
              setSubmitting(false);
              return;
            }
            alert(result.message || 'Recorded.');
            window.__submitInFlight = false;
            setSubmitting(false);
            google.script.host.close();
          })
          .submitUseOffDay(payload);
      }

      document.getElementById('intendedDate').value = todayYmd();
      const hasNoOptions = document.querySelectorAll('.off-id-checkbox').length === 0;
      if (hasNoOptions) {
        setError('No available OFF IDs in Offs (Granted).');
      } else {
        setError('');
      }
      document.getElementById('session').addEventListener('change', updateSelectionHint);
      Array.from(document.querySelectorAll('.off-id-checkbox')).forEach(function (cb) {
        cb.addEventListener('change', updateSelectionHint);
      });
      updateSelectionHint();
    </script>
  </body>
</html>`;
}

function getDeleteOffGrantedDialogHtml_(initialOptions, selectedPersonnel) {
  const options = Array.isArray(initialOptions) ? initialOptions : [];
  const personnelLabel = escapeHtml_(selectedPersonnel || DEFAULT_PERSONNEL);
  const checkboxTags = options.length > 0
    ? options
      .map((opt) => {
        const id = escapeHtml_(opt.id);
        const label = escapeHtml_(opt.label);
        return `<label class="off-item"><input class="off-id-checkbox" type="checkbox" name="offIds" value="${id}"> ${label}</label>`;
      })
      .join('\n')
    : '<div class="off-empty">No deletable OFF IDs (only completely unused IDs can be deleted).</div>';

  return `
<!DOCTYPE html>
<html>
  <head>
    <base target="_top">
    <style>
      body { font-family: Arial, sans-serif; margin: 16px; color: #1f2937; }
      h3 { margin: 0 0 12px 0; }
      .off-list { min-height: 260px; max-height: 320px; overflow: auto; border: 1px solid #9ca3af; padding: 8px; border-radius: 4px; background: #fff; }
      .off-item { display: block; margin: 4px 0; line-height: 1.35; font-weight: 600; }
      .off-item input { width: auto; margin-right: 8px; vertical-align: middle; }
      .off-empty { color: #6b7280; }
      .actions { margin-top: 14px; display: flex; gap: 8px; }
      button { padding: 8px 12px; cursor: pointer; }
      .error { margin-top: 10px; color: #b91c1c; white-space: pre-wrap; }
      .hint { margin-top: 8px; color: #374151; font-size: 12px; white-space: pre-wrap; }
    </style>
  </head>
  <body>
    <h3>Delete Off Granted</h3>
    <div class="hint">Tracking Personnel: ${personnelLabel}</div>
    <div class="hint">Tick one or more OFF IDs to delete.</div>

    <div id="offIdsList" class="off-list">
      ${checkboxTags}
    </div>

    <div id="selectionHint" class="hint"></div>

    <div class="actions">
      <button id="submitBtn" type="button" onclick="submitForm()">Delete Selected</button>
      <button id="cancelBtn" type="button" onclick="if(window.google&&google.script&&google.script.host){google.script.host.close();}">Cancel</button>
    </div>

    <div id="error" class="error"></div>

    <script>
      function setError(message) {
        document.getElementById('error').textContent = message || '';
      }

      function setSubmitting(isSubmitting) {
        const btn = document.getElementById('submitBtn');
        if (!btn) return;
        btn.disabled = isSubmitting;
        btn.textContent = isSubmitting ? 'Deleting...' : 'Delete Selected';
      }

      function getSelectedOptions() {
        return Array.from(document.querySelectorAll('.off-id-checkbox:checked'));
      }

      function updateSelectionHint() {
        const selected = getSelectedOptions();
        document.getElementById('selectionHint').textContent = 'Selected IDs: ' + selected.length;
      }

      function submitForm() {
        if (window.__submitInFlight) return;
        setError('');

        const selectedIds = getSelectedOptions().map((opt) => opt.value);
        if (selectedIds.length === 0) {
          setError('Please tick at least one OFF ID.');
          return;
        }

        if (!confirm('Delete selected OFF ID(s)? This cannot be undone.')) {
          return;
        }

        window.__submitInFlight = true;
        setSubmitting(true);

        google.script.run
          .withFailureHandler((err) => {
            const msg = err && err.message ? err.message : String(err);
            setError(msg);
            window.__submitInFlight = false;
            setSubmitting(false);
          })
          .withSuccessHandler((result) => {
            if (!result || !result.ok) {
              setError(result && result.message ? result.message : 'Failed to delete.');
              window.__submitInFlight = false;
              setSubmitting(false);
              return;
            }
            alert(result.message || 'Deleted.');
            window.__submitInFlight = false;
            setSubmitting(false);
            google.script.host.close();
          })
          .submitDeleteOffGrantedBatch({ selectedIds });
      }

      Array.from(document.querySelectorAll('.off-id-checkbox')).forEach(function (cb) {
        cb.addEventListener('change', updateSelectionHint);
      });
      updateSelectionHint();
    </script>
  </body>
</html>`;
}

function getManageOffGrantedDialogHtml_(initialOptions, selectedPersonnel) {
  const options = Array.isArray(initialOptions) ? initialOptions : [];
  const personnelLabel = escapeHtml_(selectedPersonnel || DEFAULT_PERSONNEL);
  const optionsJson = JSON.stringify(options).replace(/</g, '\\u003c');

  return `
<!DOCTYPE html>
<html>
  <head>
    <base target="_top">
    <style>
      body { font-family: Arial, sans-serif; margin: 16px; color: #1f2937; }
      h3 { margin: 0 0 12px 0; }
      label { display: block; font-weight: 600; margin-top: 10px; margin-bottom: 4px; }
      input, select, textarea { width: 100%; box-sizing: border-box; padding: 8px; }
      textarea { min-height: 70px; resize: vertical; }
      .off-list { min-height: 160px; max-height: 210px; overflow: auto; border: 1px solid #9ca3af; padding: 8px; border-radius: 4px; background: #fff; }
      .off-item { display: block; margin: 4px 0; line-height: 1.35; }
      .off-item input { width: auto; margin-right: 8px; vertical-align: middle; }
      .off-empty { color: #6b7280; }
      .hint { margin-top: 6px; color: #374151; font-size: 12px; white-space: pre-wrap; }
      .hidden { display: none; }
      .actions { margin-top: 14px; display: flex; gap: 8px; }
      button { padding: 8px 12px; cursor: pointer; }
      .error { margin-top: 10px; color: #b91c1c; white-space: pre-wrap; }
    </style>
  </head>
  <body>
    <h3>Edit Off Granted</h3>
    <div class="hint">Tracking Personnel: ${personnelLabel}</div>

    <label>Which OFF ID do you want to edit?</label>
    <div id="offList" class="off-list"></div>
    <div id="selectionHint" class="hint">Tick exactly one OFF ID.</div>

    <label for="grantedDate">Date Off Granted</label>
    <input id="grantedDate" type="date" />

    <label for="durationType">Duration type</label>
    <select id="durationType">
      <option value="FULL">Full Day (1)</option>
      <option value="HALF">Half Day (0.5)</option>
    </select>

    <label for="reasonType">Reason</label>
    <select id="reasonType">
      <option value="OPS">Ops</option>
      <option value="OTHERS">Others</option>
    </select>

    <div id="opsFields">
      <label for="weekendOpsDate">When was your Weekend Ops duty?</label>
      <input id="weekendOpsDate" type="date" />
    </div>

    <div id="othersFields" class="hidden">
      <label for="otherDetails">Reason details</label>
      <textarea id="otherDetails" placeholder="Enter details..."></textarea>
    </div>

    <label for="providedBy">Provided by who? (Yourself if Ops related)</label>
    <input id="providedBy" type="text" />

    <div class="actions">
      <button id="submitBtn" type="button" onclick="submitForm()">Save Changes</button>
      <button type="button" onclick="google.script.host.close()">Cancel</button>
    </div>

    <div id="error" class="error"></div>

    <script>
      const options = ${optionsJson};

      function esc(value) {
        return String(value == null ? '' : value)
          .replace(/&/g, '&amp;')
          .replace(/</g, '&lt;')
          .replace(/>/g, '&gt;')
          .replace(/"/g, '&quot;')
          .replace(/'/g, '&#39;');
      }

      function setError(message) {
        document.getElementById('error').textContent = message || '';
      }

      function toggleReasonInputs() {
        const reason = document.getElementById('reasonType').value;
        document.getElementById('opsFields').classList.toggle('hidden', reason !== 'OPS');
        document.getElementById('othersFields').classList.toggle('hidden', reason !== 'OTHERS');
      }

      function renderOptions() {
        const container = document.getElementById('offList');
        if (!options.length) {
          container.innerHTML = '<div class="off-empty">No OFF IDs found.</div>';
          return;
        }
        container.innerHTML = options.map((opt) => (
          '<label class="off-item">' +
          '<input class="off-id-checkbox" type="checkbox" value="' + esc(opt.id) + '">' +
          esc(opt.label) +
          '</label>'
        )).join('');

        Array.from(document.querySelectorAll('.off-id-checkbox')).forEach((cb) => {
          cb.addEventListener('change', () => {
            if (cb.checked) {
              Array.from(document.querySelectorAll('.off-id-checkbox')).forEach((other) => {
                if (other !== cb) other.checked = false;
              });
            }
            applySelectedValues();
          });
        });
      }

      function getSelectedIds() {
        return Array.from(document.querySelectorAll('.off-id-checkbox:checked')).map((el) => el.value);
      }

      function applySelectedValues() {
        const selected = getSelectedIds();
        if (selected.length !== 1) {
          document.getElementById('selectionHint').textContent = 'Tick exactly one OFF ID.';
          return;
        }
        const opt = options.find((o) => o.id === selected[0]);
        if (!opt) return;

        document.getElementById('grantedDate').value = opt.grantedDate || '';
        document.getElementById('durationType').value = opt.durationType || 'FULL';
        document.getElementById('reasonType').value = opt.reasonType || 'OPS';
        document.getElementById('weekendOpsDate').value = opt.weekendOpsDate || '';
        document.getElementById('otherDetails').value = opt.otherDetails || '';
        document.getElementById('providedBy').value = opt.providedBy || '';
        toggleReasonInputs();
        document.getElementById('selectionHint').textContent = 'Editing ' + opt.id;
      }

      function setSubmitting(isSubmitting) {
        const btn = document.getElementById('submitBtn');
        btn.disabled = isSubmitting;
        btn.textContent = isSubmitting ? 'Saving...' : 'Save Changes';
      }

      function submitForm() {
        if (window.__submitInFlight) return;
        setError('');

        const selectedIds = getSelectedIds();
        if (selectedIds.length !== 1) {
          setError('Please tick exactly one OFF ID.');
          return;
        }

        const payload = {
          selectedIds,
          grantedDate: document.getElementById('grantedDate').value,
          durationType: document.getElementById('durationType').value,
          reasonType: document.getElementById('reasonType').value,
          weekendOpsDate: document.getElementById('weekendOpsDate').value,
          otherDetails: document.getElementById('otherDetails').value,
          providedBy: document.getElementById('providedBy').value
        };

        window.__submitInFlight = true;
        setSubmitting(true);

        google.script.run
          .withFailureHandler((err) => {
            setError(err && err.message ? err.message : String(err));
            window.__submitInFlight = false;
            setSubmitting(false);
          })
          .withSuccessHandler((result) => {
            if (!result || !result.ok) {
              setError(result && result.message ? result.message : 'Failed to edit Off Granted.');
              window.__submitInFlight = false;
              setSubmitting(false);
              return;
            }
            alert(result.message || 'Updated.');
            window.__submitInFlight = false;
            setSubmitting(false);
            google.script.host.close();
          })
          .submitManageOffGranted(payload);
      }

      document.getElementById('reasonType').addEventListener('change', toggleReasonInputs);
      renderOptions();
      toggleReasonInputs();
    </script>
  </body>
</html>`;
}

function getManageOffUsedDialogHtml_(initialUsedOptions, additionalOffOptions, selectedPersonnel) {
  const usedOptions = Array.isArray(initialUsedOptions) ? initialUsedOptions : [];
  const additionalOptions = Array.isArray(additionalOffOptions) ? additionalOffOptions : [];
  const personnelLabel = escapeHtml_(selectedPersonnel || DEFAULT_PERSONNEL);
  const usedJson = JSON.stringify(usedOptions).replace(/</g, '\\u003c');
  const additionalJson = JSON.stringify(additionalOptions).replace(/</g, '\\u003c');

  return `
<!DOCTYPE html>
<html>
  <head>
    <base target="_top">
    <style>
      body { font-family: Arial, sans-serif; margin: 16px; color: #1f2937; }
      h3 { margin: 0 0 12px 0; }
      label { display: block; font-weight: 600; margin-top: 10px; margin-bottom: 4px; }
      input, select, textarea { width: 100%; box-sizing: border-box; padding: 8px; }
      textarea { min-height: 70px; resize: vertical; }
      .off-list { min-height: 140px; max-height: 190px; overflow: auto; border: 1px solid #9ca3af; padding: 8px; border-radius: 4px; background: #fff; }
      .off-item { display: block; margin: 4px 0; line-height: 1.35; }
      .off-item input { width: auto; margin-right: 8px; vertical-align: middle; }
      .off-empty { color: #6b7280; }
      .hint { margin-top: 6px; color: #374151; font-size: 12px; white-space: pre-wrap; }
      .hidden { display: none; }
      .actions { margin-top: 14px; display: flex; gap: 8px; }
      button { padding: 8px 12px; cursor: pointer; }
      .error { margin-top: 10px; color: #b91c1c; white-space: pre-wrap; }
    </style>
  </head>
  <body>
    <h3>Edit/Undo Off Used</h3>
    <div class="hint">Tracking Personnel: ${personnelLabel}</div>

    <label>Which Use ID do you want to manage?</label>
    <div id="usedList" class="off-list"></div>
    <div id="selectionHint" class="hint">Tick exactly one Use ID.</div>

    <label for="actionType">Action</label>
    <select id="actionType">
      <option value="EDIT">Edit</option>
      <option value="UNDO">Undo</option>
    </select>

    <div id="editFields">
      <label for="intendedDate">Date Intended</label>
      <input id="intendedDate" type="date" />

      <label for="session">Session</label>
      <select id="session">
        <option value="FULL">Full Day (1)</option>
        <option value="AM">AM (0.5)</option>
        <option value="PM">PM (0.5)</option>
      </select>

      <label for="comments">Comments</label>
      <textarea id="comments" placeholder="Add comments..."></textarea>

      <label>Additional OFF IDs (used only when session increases)</label>
      <div id="additionalList" class="off-list"></div>
      <div id="additionalHint" class="hint"></div>
    </div>

    <div class="actions">
      <button id="submitBtn" type="button" onclick="submitForm()">Submit</button>
      <button type="button" onclick="google.script.host.close()">Cancel</button>
    </div>

    <div id="error" class="error"></div>

    <script>
      const usedOptions = ${usedJson};
      const additionalOptions = ${additionalJson};

      function esc(value) {
        return String(value == null ? '' : value)
          .replace(/&/g, '&amp;')
          .replace(/</g, '&lt;')
          .replace(/>/g, '&gt;')
          .replace(/"/g, '&quot;')
          .replace(/'/g, '&#39;');
      }

      function setError(message) {
        document.getElementById('error').textContent = message || '';
      }

      function renderUsedOptions() {
        const container = document.getElementById('usedList');
        if (!usedOptions.length) {
          container.innerHTML = '<div class="off-empty">No Use IDs found.</div>';
          return;
        }
        container.innerHTML = usedOptions.map((opt) => (
          '<label class="off-item">' +
          '<input class="use-id-checkbox" type="checkbox" value="' + esc(opt.useId) + '">' +
          esc(opt.label) +
          '</label>'
        )).join('');

        Array.from(document.querySelectorAll('.use-id-checkbox')).forEach((cb) => {
          cb.addEventListener('change', () => {
            if (cb.checked) {
              Array.from(document.querySelectorAll('.use-id-checkbox')).forEach((other) => {
                if (other !== cb) other.checked = false;
              });
            }
            applySelectedUsedValues();
            updateAdditionalHint();
          });
        });
      }

      function renderAdditionalOptions() {
        const container = document.getElementById('additionalList');
        if (!additionalOptions.length) {
          container.innerHTML = '<div class="off-empty">No additional OFF IDs available.</div>';
          return;
        }
        container.innerHTML = additionalOptions.map((opt) => (
          '<label class="off-item">' +
          '<input class="additional-id-checkbox" type="checkbox" value="' + esc(opt.id) + '" data-remaining="' + esc(opt.remaining || 0) + '">' +
          esc(opt.label) +
          '</label>'
        )).join('');
        Array.from(document.querySelectorAll('.additional-id-checkbox')).forEach((cb) => {
          cb.addEventListener('change', updateAdditionalHint);
        });
      }

      function getSelectedUseIds() {
        return Array.from(document.querySelectorAll('.use-id-checkbox:checked')).map((el) => el.value);
      }

      function getSelectedAdditionalIds() {
        return Array.from(document.querySelectorAll('.additional-id-checkbox:checked')).map((el) => el.value);
      }

      function getSelectedAdditionalTotal() {
        let total = 0;
        Array.from(document.querySelectorAll('.additional-id-checkbox:checked')).forEach((el) => {
          total += Number(el.dataset.remaining || 0);
        });
        return total;
      }

      function getSelectedUsedRecord() {
        const ids = getSelectedUseIds();
        if (ids.length !== 1) return null;
        return usedOptions.find((x) => x.useId === ids[0]) || null;
      }

      function applySelectedUsedValues() {
        const rec = getSelectedUsedRecord();
        if (!rec) {
          document.getElementById('selectionHint').textContent = 'Tick exactly one Use ID.';
          return;
        }
        document.getElementById('selectionHint').textContent = 'Managing ' + rec.useId;
        document.getElementById('intendedDate').value = rec.intendedDate || '';
        document.getElementById('session').value = rec.session || 'FULL';
        document.getElementById('comments').value = rec.comments || '';
      }

      function toggleActionFields() {
        const action = document.getElementById('actionType').value;
        document.getElementById('editFields').classList.toggle('hidden', action === 'UNDO');
      }

      function updateAdditionalHint() {
        const rec = getSelectedUsedRecord();
        const action = document.getElementById('actionType').value;
        const session = document.getElementById('session').value;
        const hint = document.getElementById('additionalHint');

        if (!rec || action !== 'EDIT') {
          hint.textContent = '';
          return;
        }

        const target = session === 'FULL' ? 1 : 0.5;
        const current = Number(rec.duration || 0);
        const delta = target - current;
        if (delta <= 0) {
          hint.textContent = 'No additional OFF ID needed.';
          return;
        }

        const selectedTotal = getSelectedAdditionalTotal();
        const selectedText = Number.isInteger(selectedTotal) ? String(selectedTotal) : selectedTotal.toFixed(1);
        const needText = Number.isInteger(delta) ? String(delta) : delta.toFixed(1);
        hint.textContent = 'Need additional ' + needText + ' day. Selected additional total: ' + selectedText + ' day.';
      }

      function setSubmitting(isSubmitting) {
        const btn = document.getElementById('submitBtn');
        btn.disabled = isSubmitting;
        btn.textContent = isSubmitting ? 'Submitting...' : 'Submit';
      }

      function submitForm() {
        if (window.__submitInFlight) return;
        setError('');

        const selectedIds = getSelectedUseIds();
        if (selectedIds.length !== 1) {
          setError('Please tick exactly one Use ID.');
          return;
        }

        const action = document.getElementById('actionType').value;
        if (action === 'UNDO') {
          if (!confirm('Undo selected Use ID? This will restore OFF allocations.')) return;
        }

        const payload = {
          selectedIds,
          action,
          intendedDate: document.getElementById('intendedDate').value,
          session: document.getElementById('session').value,
          comments: document.getElementById('comments').value,
          additionalIds: getSelectedAdditionalIds()
        };

        window.__submitInFlight = true;
        setSubmitting(true);

        google.script.run
          .withFailureHandler((err) => {
            setError(err && err.message ? err.message : String(err));
            window.__submitInFlight = false;
            setSubmitting(false);
          })
          .withSuccessHandler((result) => {
            if (!result || !result.ok) {
              setError(result && result.message ? result.message : 'Failed to process request.');
              window.__submitInFlight = false;
              setSubmitting(false);
              return;
            }
            alert(result.message || 'Done.');
            window.__submitInFlight = false;
            setSubmitting(false);
            google.script.host.close();
          })
          .submitManageOffUsed(payload);
      }

      document.getElementById('actionType').addEventListener('change', () => {
        toggleActionFields();
        updateAdditionalHint();
      });
      document.getElementById('session').addEventListener('change', updateAdditionalHint);

      renderUsedOptions();
      renderAdditionalOptions();
      toggleActionFields();
      updateAdditionalHint();
    </script>
  </body>
</html>`;
}
