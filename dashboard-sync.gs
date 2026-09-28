/**
 * Goals app -> Daily Dashboard sync
 *
 * Paste this into the Apps Script editor of the "Daily Dashboard - To Do" sheet
 * (Extensions > Apps Script), then Deploy > New deployment > Web app.
 * Adds one row to the "To Do Database" tab each time the Goals app sends a task.
 */
const SHEET_NAME = 'To Do Database';
const VALID_CATEGORIES = ['Household', 'Work', 'Personal', 'Hobby', 'Travel'];

function doGet() {
  return out_({ ok: true, message: 'Goals to Dashboard sync is live' });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const data = JSON.parse(e.postData.contents);
    const task = String(data.task || '').trim();
    if (!task) return out_({ ok: false, error: 'missing task' });

    // Safe to retry: a goal that was already sent is never added twice.
    const props = PropertiesService.getScriptProperties();
    const sentKey = data.goalId ? 'sent_' + data.goalId : null;
    if (sentKey && props.getProperty(sentKey)) return out_({ ok: true, duplicate: true });

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sheet) return out_({ ok: false, error: 'sheet "' + SHEET_NAME + '" not found' });

    // Last row that has a task in column B
    const colB = sheet.getRange(1, 2, sheet.getMaxRows(), 1).getValues();
    let last = 1;
    for (let i = colB.length - 1; i >= 0; i--) {
      if (String(colB[i][0]).trim() !== '') { last = i + 1; break; }
    }
    const row = last + 1;
    if (row > sheet.getMaxRows()) sheet.insertRowsAfter(sheet.getMaxRows(), 1);

    // Keep checkboxes / dropdowns consistent with the row above
    if (last > 1) {
      sheet.getRange(last, 1, 1, 12).copyTo(
        sheet.getRange(row, 1, 1, 12),
        SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
    }

    const category = VALID_CATEGORIES.indexOf(data.category) >= 0 ? data.category : 'Personal';
    // Columns: Complete | Task | Status | Due Date | Priority | Category | Sub Category | Repeat | Notes | Assignee | Family Agenda | GP Agenda
    sheet.getRange(row, 1, 1, 12).setValues([[
      false, task, 'Not started', '', 'Moderate', category, 'None', '', '', '', false, false
    ]]);

    if (/^\d{4}-\d{2}-\d{2}$/.test(String(data.dueDate || ''))) {
      writeDueDate_(sheet, row, last, data.dueDate);
    }

    if (sentKey) props.setProperty(sentKey, new Date().toISOString());
    return out_({ ok: true, row: row });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

// Matches however the existing Due Date cells are stored (real dates or text like "Sep-27").
function writeDueDate_(sheet, row, lastRow, iso) {
  const p = iso.split('-').map(Number);
  const d = new Date(p[0], p[1] - 1, p[2], 12, 0, 0);
  let ref = null;
  if (lastRow >= 2) {
    const vals = sheet.getRange(2, 4, lastRow - 1, 1).getValues();
    for (let i = vals.length - 1; i >= 0; i--) {
      const v = vals[i][0];
      if (v !== '' && v !== null && v !== 'Due Date') { ref = { v: v, row: i + 2 }; break; }
    }
  }
  const cell = sheet.getRange(row, 4);
  if (ref && Object.prototype.toString.call(ref.v) !== '[object Date]') {
    cell.setNumberFormat('@');
    cell.setValue(Utilities.formatDate(d, Session.getScriptTimeZone(), 'MMM-dd'));
  } else {
    cell.setValue(d);
    cell.setNumberFormat(ref ? sheet.getRange(ref.row, 4).getNumberFormat() : 'MMM-dd');
  }
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
