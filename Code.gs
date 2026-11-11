/**
 * 24/7 roster generator for Google Sheets (Apps Script).
 *
 * Reads Settings, Names and Unavailability (including soft preferred days off), solves the roster with Apps Script's
 * built-in Linear Optimization Service (mixed-integer), and writes period-named Roster, Points and Log sheets.
 *
 * Use: Extensions > Apps Script > paste this file > save > reload the sheet > menu Roster.
 */

var BLOCKS = ['AM', 'PM', 'Graveyard'];
var BLOCK_TITLES = ['AM week', 'PM week', 'Graveyard week', 'Admin week (fill in manually)'];
var BLOCK_LEN = [7, 7, 7, 5];
var START_COL = [3, 11, 19, 27];   // C, K, S, AA (1-based columns)
var COUNT_COL = 33;                // AG: per-person counters
var FIRST_ROW = 3;
var BIG = 1000000;                 // stands in for "no limit" in constraints
var MAX_SOLVE_SECONDS = 270;       // an Apps Script run is limited to 6 minutes
var MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
var DAYS = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'];

// label on the Settings sheet -> [key, default]  (default null = required unless listed in OPTIONAL)
var SETTINGS = {
  'AM week start date': ['start', null],
  'Plan Graveyard week (Yes/No)': ['planGraveyard', 'Yes'],
  'Mode (QUOTA or EXACT)': ['mode', 'QUOTA'],
  'AM/PM minimum on duty': ['ampmMin', 15],
  'AM/PM minimum # names': ['ampmHash', 3],
  'AM/PM maximum # names (blank = no limit)': ['ampmHmax', null],
  'Graveyard minimum on duty': ['gMin', 5],
  'Graveyard minimum # names': ['gHash', 1],
  'AM days per person per week': ['qAm', 5],
  'PM days per person per week': ['qPm', 5],
  'Graveyard days per person per week': ['qG', 2],
  'Perm-AM people: AM days in PM week': ['permPm', 5],
  'Perm-AM people: AM days in Graveyard week': ['permG', 3],
  'Solver time limit (seconds)': ['tlimit', 200],
  'Show absence codes on roster (Yes/No)': ['showAbs', 'Yes'],
  'Weekend days (e.g. Sat,Sun)': ['wkDays', 'Sat,Sun'],
  'Max weekend shifts per person per week (blank = no limit)': ['wkMax', null],
  'Weekend limit type (HARD or SOFT)': ['wkType', 'SOFT'],
  'Balance points when scheduling (Yes/No)': ['balance', 'No'],
  'Points: evening hour weight (1800-2200)': ['wEve', 1.5],
  'Points: night hour weight (2200-0600)': ['wNight', 2.0],
  'Points: weekend add-on per hour (Fri 1500 to Sun 2400)': ['wWk', 0.5],
  'Points: recovery debit per night shift': ['pRecov', 8],
  'Points: back-to-back night penalty': ['pB2b', 4],
  'Points: bonus per day worked': ['pDay', 2],
  'Admin shift hours (for points)': ['adminHrs', 7],
  'Preferred days off penalty': ['preferOffPenalty', 20]
};
var OPTIONAL = { ampmHmax: true, wkMax: true };

var TZ = null;   // spreadsheet time zone, set at run time

function PlanError(message) {
  this.name = 'PlanError';
  this.message = message;
}
PlanError.prototype = Object.create(Error.prototype);

// ------------------------------------------------------------------ menu and entry point
function onOpen() {
  var ss = SpreadsheetApp.getActive();
  try {
    if (ss && ss.getSheetByName('Settings')) ensureGraveyardPlanningSetting_(ss);
    if (ss && ss.getSheetByName('Names')) ensureNamesWorkdayColumns_(ss);
    if (ss) {
      var visibleRosters = ss.getSheets().filter(function (sh) {
        return /^Roster(?: |$)/i.test(sh.getName()) && !/^Archived Roster/i.test(sh.getName()) && !sh.isSheetHidden();
      });
      if (visibleRosters.length === 1) updateCurrentRosterImportTab_(ss, visibleRosters[0].getName());
    }
  } catch (setupError) {
    console.error('Could not add optional Settings/Names controls: ' + String(setupError && setupError.stack ? setupError.stack : setupError));
  }
  SpreadsheetApp.getUi().createMenu('Roster')
    .addItem('Generate schedule', 'generateSchedule')
    .addItem('Enhanced Graveyard', 'generateEnhancedGraveyard')
    .addItem('Fix roster colours', 'fixRosterColors')
    .addItem('Refresh Points Log', 'refreshPointsLog')
    .addToUi();
}

function ensureNamesWorkdayColumns_(ss) {
  var sh = ss.getSheetByName('Names');
  if (!sh) return;
  var headers = ['AM workdays/week', 'PM workdays/week', 'Graveyard workdays/week'];
  var headerRange = sh.getRange(1, 5, 1, 3), current = headerRange.getValues()[0];
  var blank = current.every(function (v) { return v === '' || v === null || v === undefined; });
  var matches = headers.every(function (v, i) { return String(current[i] || '').trim() === v; });
  if (blank) {
    sh.getRange(1, 4).copyTo(headerRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    headerRange.setValues([headers]);
    matches = true;
  }
  if (!matches) return; // Preserve any existing custom columns in E:G.
  sh.setColumnWidth(5, 115);
  sh.setColumnWidth(6, 115);
  sh.setColumnWidth(7, 155);
  headerRange.setWrap(true);
  var rule = SpreadsheetApp.newDataValidation()
    .requireFormulaSatisfied('=OR(E2="",AND(ISNUMBER(E2),E2=INT(E2),E2>=0,E2<=7))')
    .setAllowInvalid(false).build();
  sh.getRange(2, 5, sh.getMaxRows() - 1, 3).setDataValidation(rule);
}

function ensureGraveyardPlanningSetting_(ss) {
  var sh = ss.getSheetByName('Settings');
  if (!sh) return;
  var last = Math.max(sh.getLastRow(), 1), labels = sh.getRange(1, 1, last, 1).getValues(), row = 0;
  for (var i = 0; i < labels.length; i++) {
    if (String(labels[i][0] || '').trim() === 'Plan Graveyard week (Yes/No)') { row = i + 1; break; }
  }
  if (!row) {
    row = last + 1;
    sh.getRange(row, 1, 1, 3).setValues([['Plan Graveyard week (Yes/No)', 'Yes', 'Yes generates Graveyard assignments. No leaves the week pending for Roster > Enhanced Graveyard.']]);
  }
  var rule = SpreadsheetApp.newDataValidation().requireValueInList(['Yes', 'No'], true).setAllowInvalid(false).build();
  sh.getRange(row, 2).setDataValidation(rule);
  if (!sh.getRange(row, 2).getValue()) sh.getRange(row, 2).setValue('Yes');
}

function refreshPointsLog() {
  var ss = SpreadsheetApp.getActive();
  SpreadsheetApp.flush();
  var message = buildPointsLog(ss);
  SpreadsheetApp.getUi().alert('Points Log refreshed', message, SpreadsheetApp.getUi().ButtonSet.OK);
}

// Simple edit trigger keeps the log in sync with member-list, roster and Points edits.
function onEdit(e) {
  if (!e || !e.range) return;
  var sh = e.range.getSheet(), name = sh.getName(), row = e.range.getRow(), col = e.range.getColumn();
  var relevant = false;
  if (name === 'Names') relevant = row >= 2 && col <= 4;
  else if (/^(?:Archived )?Points(?: |$)/i.test(name)) relevant = row >= 2;
  else if (isRosterOutputSheet(sh)) relevant = row >= FIRST_ROW && col >= 3 && col <= 31;
  if (!relevant || name === 'Points Log') return;
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(5000)) return;
  try {
    SpreadsheetApp.flush();
    buildPointsLog(e.source || SpreadsheetApp.getActive());
  } catch (err) {
    console.error('Points Log refresh failed: ' + String(err && err.stack ? err.stack : err));
  } finally {
    lock.releaseLock();
  }
}

// Build a cycle-based, newcomer-fair summary. Each generated Points sheet is a
// cycle; duplicate archived versions of the same date range are ignored in favor
// of the newest sheet. The live Names list contributes rows even with no history.
function buildPointsLog(ss) {
  var member = {}, memberOrder = [];
  var namesSheet = ss.getSheetByName('Names');
  if (namesSheet) {
    var nv = namesSheet.getDataRange().getValues();
    for (var n = 1; n < nv.length; n++) {
      var name = String(nv[n][0] || '').trim();
      if (!name) continue;
      var key = norm(name);
      if (!(key in member)) {
        member[key] = { key: key, name: name, section: String(nv[n][1] || ''), cycles: 0, pointSum: 0, targetSum: 0, shareSum: 0, latest: '' };
        memberOrder.push(key);
      } else {
        member[key].section = String(nv[n][1] || member[key].section);
      }
    }
  }

  var processedCycles = {}, completedCycles = 0;
  ss.getSheets().forEach(function (pointsSheet) {
    var pointsName = pointsSheet.getName();
    if (/^Points Log$/i.test(pointsName) || !/^(?:Archived )?Points(?: |$)/i.test(pointsName)) return;
    var cycleKey = pointsName.replace(/^Archived /i, '').replace(/\s+\(\d+\)(?:\s+\d+)?$/, '');
    if (processedCycles[cycleKey]) return;
    processedCycles[cycleKey] = true;
    if (/^Graveyard planning:\s*PENDING$/i.test(String(pointsSheet.getRange('Q1').getValue() || '').trim())) return;
    completedCycles++;
    var values = pointsSheet.getDataRange().getValues();
    var rosterName = pointsName.replace(/^Archived Points/i, 'Archived Roster').replace(/^Points/i, 'Roster')
      .replace(/\s+\(\d+\)(?:\s+\d+)?$/, '');
    var roster = ss.getSheetByName(rosterName);
    var rosterSections = {};
    if (roster) {
      var rosterValues = roster.getDataRange().getValues();
      for (var rr = FIRST_ROW - 1; rr < rosterValues.length; rr++) {
        var rName = String(rosterValues[rr][0] || '').trim();
        if (rName && rosterValues[rr][1]) rosterSections[norm(rName)] = String(rosterValues[rr][1]);
      }
    }
    var cycleLabel = cycleKey.replace(/^Points\s*/i, '');
    for (var r = 5; r < values.length; r++) {
      var memberName = String(values[r][1] || '').trim();
      if (!memberName) continue;
      var key = norm(memberName);
      if (!(key in member)) {
        member[key] = { key: key, name: memberName, section: rosterSections[key] || '', cycles: 0, pointSum: 0, targetSum: 0, shareSum: 0, latest: '' };
        memberOrder.push(key);
      } else if (!member[key].section && rosterSections[key]) {
        member[key].section = rosterSections[key];
      }
      var actual = Number(values[r][2]);
      var target = Number(values[r][8]);
      var share = Number(values[r][9]);
      if (!isFinite(actual) || !isFinite(target) || target <= 0 || !isFinite(share)) continue;
      var rec = member[key];
      rec.cycles++;
      rec.pointSum += actual;
      rec.targetSum += target;
      rec.shareSum += share;
      if (!rec.latest) rec.latest = cycleLabel;
    }
  });

  var ranked = memberOrder.map(function (key) { return member[key]; });
  ranked.forEach(function (m) {
    m.meanPoints = m.cycles ? m.pointSum / m.cycles : null;
    m.meanTarget = m.cycles ? m.targetSum / m.cycles : null;
    m.meanShare = m.cycles ? m.shareSum / m.cycles : null;
  });
  var withHistory = ranked.filter(function (m) { return m.cycles > 0; });
  withHistory.sort(function (a, b) { return a.meanShare - b.meanShare || a.name.localeCompare(b.name); });
  var lastShare = null, lastRank = 0;
  withHistory.forEach(function (m, i) {
    if (lastShare === null || Math.abs(m.meanShare - lastShare) > 1e-9) lastRank = i + 1;
    m.rank = lastRank;
    lastShare = m.meanShare;
  });
  var noHistory = ranked.filter(function (m) { return !m.cycles; }).sort(function (a, b) { return a.name.localeCompare(b.name); });
  var rows = withHistory.concat(noHistory).map(function (m) {
    return [m.rank || '', m.name, m.section, m.cycles, m.meanPoints, m.meanTarget, m.meanShare, m.latest || 'No completed cycle'];
  });

  // Keep the previous month-based reports as hidden history when upgrading an existing workbook.
  ['Monthly Points', 'Monthly Rankings'].forEach(function (legacy) {
    var old = ss.getSheetByName(legacy);
    if (old && !old.isSheetHidden()) old.hideSheet();
  });
  var log = ss.getSheetByName('Points Log');
  if (!log) log = ss.insertSheet('Points Log', 0);
  log.showSheet();
  log.getRange('A1:H2').breakApart();
  log.clearContents();
  log.clearFormats();
  log.getRange('A1:H1').merge();
  log.getRange('A1').setValue('Points Log — Fair-share contribution across cycles');
  log.getRange('A2:H2').merge();
  log.getRange('A2').setValue('Mean fair-share points is the average of each cycle’s Total points ÷ Fair-share target. Rank 1 is the lowest average share. Only cycles where a member appears and has a target count, so new members are not penalized for earlier cycles. Members in Names with no completed cycle appear at the bottom without a rank.');
  log.getRange('A4:H4').setValues([['Rank (1 = lowest share)', 'Member', 'Section', 'Cycles counted', 'Mean total points', 'Mean fair-share target', 'Mean fair-share points (% of target)', 'Most recent cycle']]);
  if (rows.length) log.getRange(5, 1, rows.length, 8).setValues(rows);
  log.getRange('A1:H1').setBackground(C_HEAD).setFontColor('#ffffff').setFontWeight('bold').setFontSize(14);
  log.getRange('A2:H2').setWrap(true).setFontStyle('italic');
  log.getRange('A4:H4').setBackground(C_HEAD).setFontColor('#ffffff').setFontWeight('bold').setWrap(true);
  if (rows.length) {
    log.getRange(5, 5, rows.length, 2).setNumberFormat('0.0');
    log.getRange(5, 7, rows.length, 1).setNumberFormat('0.0%');
  }
  log.setFrozenRows(4);
  [150, 230, 90, 105, 130, 150, 220, 180].forEach(function (w, i) { log.setColumnWidth(i + 1, w); });
  log.setRowHeight(2, 48);
  return 'Points Log updated for ' + ranked.length + ' members across ' + completedCycles + ' completed cycle(s).';
}

function isRosterOutputSheet(sh) {
  return /^(Roster|Archived Roster)(?: |$)/i.test(sh.getName());
}

function findRosterSectionRows(sh) {
  var lastRow = sh.getLastRow();
  if (lastRow < FIRST_ROW) return null;
  var labels = sh.getRange(FIRST_ROW, 1, lastRow - FIRST_ROW + 1, 1).getValues();
  for (var i = 0; i < labels.length; i++) {
    if (String(labels[i][0] || '').trim().toLowerCase() === 'total on duty') {
      var summaryRow = FIRST_ROW + i;
      return { personLastRow: summaryRow - 2, summaryRow: summaryRow };
    }
  }
  return { personLastRow: lastRow, summaryRow: null };
}

// Schedule-cell fills follow their values: X/AM is yellow and other entries are red.
function addAssignmentColorRules(sh, personLastRow) {
  if (personLastRow < FIRST_ROW) return;
  var rules = sh.getConditionalFormatRules();
  var ownedFormulas = {};
  for (var kk = 0; kk < 4; kk++) {
    var ownedCol = colLetter(START_COL[kk]);
    ownedFormulas['=OR(' + ownedCol + FIRST_ROW + '="X",' + ownedCol + FIRST_ROW + '="AM")'] = true;
    ownedFormulas['=AND(' + ownedCol + FIRST_ROW + '<>"",' + ownedCol + FIRST_ROW + '<>"X",' + ownedCol + FIRST_ROW + '<>"AM")'] = true;
  }
  // Replace our prior rules so older versions' whole-sheet ranges are narrowed too.
  rules = rules.filter(function (rule) {
    var condition = rule.getBooleanCondition();
    if (!condition || condition.getCriteriaType() !== SpreadsheetApp.BooleanCriteria.CUSTOM_FORMULA) return true;
    var criteria = condition.getCriteriaValues();
    return !criteria.length || !ownedFormulas[String(criteria[0])];
  });
  for (var k = 0; k < 4; k++) {
    var col = colLetter(START_COL[k]);
    var yellowFormula = '=OR(' + col + FIRST_ROW + '="X",' + col + FIRST_ROW + '="AM")';
    var redFormula = '=AND(' + col + FIRST_ROW + '<>"",' + col + FIRST_ROW + '<>"X",' + col + FIRST_ROW + '<>"AM")';
    var range = sh.getRange(FIRST_ROW, START_COL[k], personLastRow - FIRST_ROW + 1, BLOCK_LEN[k]);
    [
      { formula: yellowFormula, background: C_YEL },
      { formula: redFormula, background: C_RED }
    ].forEach(function (spec) {
      rules.push(SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(spec.formula)
        .setBackground(spec.background)
        .setFontColor('#000000')
        .setBold(true)
        .setRanges([range])
        .build());
    });
  }
  sh.setConditionalFormatRules(rules);
}

// One-time repair for rosters generated before the conditional-format rule existed.
function fixRosterColors() {
  var ss = SpreadsheetApp.getActive();
  var fixed = 0;
  ss.getSheets().forEach(function (sh) {
    if (!isRosterOutputSheet(sh)) return;
    var lastRow = sh.getLastRow();
    var sectionRows = findRosterSectionRows(sh);
    var personLastRow = sectionRows ? sectionRows.personLastRow : lastRow;
    if (lastRow >= FIRST_ROW) {
      var range = sh.getRange(1, 1, lastRow, COUNT_COL + 7);
      var values = range.getValues(), backgrounds = range.getBackgrounds(), changed = false;
      for (var r = FIRST_ROW - 1; r < personLastRow; r++) {
        if (!values[r][0] || !values[r][1]) continue;
        for (var k = 0; k < 4; k++) for (var d = 0; d < BLOCK_LEN[k]; d++) {
          var c = START_COL[k] - 1 + d;
          var existingBg = String(backgrounds[r][c]).toLowerCase();
          if (existingBg === C_YEL || existingBg === C_GREY) {
            var defaultBg = k === 3 ? C_ADMIN : C_WHITE;
            if (existingBg !== defaultBg) {
              backgrounds[r][c] = defaultBg;
              changed = true;
            }
          }
        }
      }
      if (changed) range.setBackgrounds(backgrounds);
      if (sectionRows && sectionRows.summaryRow !== null) {
        var summary = sh.getRange(sectionRows.summaryRow, 1, lastRow - sectionRows.summaryRow + 1, COUNT_COL + 7);
        summary.setBackground(C_WHITE).setFontColor('#000000').setFontWeight('bold').setFontStyle('normal');
      }
    }
    addAssignmentColorRules(sh, personLastRow);
    fixed++;
  });
  SpreadsheetApp.getUi().alert('Roster colours fixed', 'Updated ' + fixed + ' roster sheet(s). X/AM cells are yellow; other nonblank schedule cells are red with bold black text.', SpreadsheetApp.getUi().ButtonSet.OK);
}

function generateSchedule() { generateScheduleMode_(false); }

function generateEnhancedGraveyard() { updateEnhancedGraveyardInPlace_(); }

function currentRosterSheet_(ss) {
  var active = ss.getActiveSheet();
  if (/^Roster(?: |$)/i.test(active.getName()) && !active.isSheetHidden()) return active;
  var visible = ss.getSheets().filter(function (sh) {
    return /^Roster(?: |$)/i.test(sh.getName()) && !sh.isSheetHidden();
  });
  if (visible.length === 1) return visible[0];
  if (!visible.length) throw new PlanError('There is no visible generated Roster sheet. Generate a regular roster first.');
  throw new PlanError('More than one Roster sheet is visible. Open the roster you want to update, then press Enhanced Graveyard again.');
}

function datesFromRoster_(sh) {
  var headers = sh.getRange(2, 1, 1, START_COL[3] + BLOCK_LEN[3] - 1).getValues()[0];
  var dates = [];
  for (var k = 0; k < 4; k++) {
    var block = [];
    for (var d = 0; d < BLOCK_LEN[k]; d++) {
      var dn = toDate(headers[START_COL[k] - 1 + d], 'Roster date header');
      if (dn === null || dn === 'INDEF') throw new PlanError('Could not read every date from the open Roster sheet.');
      block.push(dn);
    }
    dates.push(block);
  }
  return dates;
}

function graveyardOnlyMap_(source) {
  var out = {};
  Object.keys(source || {}).forEach(function (key) {
    if (/_2_\d+$/.test(key)) out[key] = source[key];
  });
  return out;
}

function currentSharePriorities_(ss, pointsSheet, people) {
  SpreadsheetApp.flush();
  buildPointsLog(ss);
  var graveyardPending = /^Graveyard planning:\s*PENDING$/i.test(String(pointsSheet.getRange('Q1').getValue() || '').trim());
  var log = ss.getSheetByName('Points Log');
  var logValues = log ? log.getDataRange().getValues() : [];
  var currentValues = pointsSheet.getDataRange().getValues();
  var currentByKey = {}, pointsNames = {}, logByKey = {};
  for (var r = 5; r < currentValues.length; r++) {
    var currentName = String(currentValues[r][1] || '').trim();
    var actual = Number(currentValues[r][2]), target = Number(currentValues[r][8]);
    if (currentName) pointsNames[norm(currentName)] = true;
    if (currentName && isFinite(actual) && isFinite(target) && target > 0) currentByKey[norm(currentName)] = {
      currentShare: actual / target, amReq: Number(currentValues[r][12]), pmReq: Number(currentValues[r][13]),
      adminDays: Number(currentValues[r][15]) || 0
    };
  }
  people.forEach(function (p) {
    if (!pointsNames[p.key]) throw new PlanError("The matching Points sheet is missing '" + p.name + "'. Update the regular roster first.");
  });
  for (var l = 4; l < logValues.length; l++) {
    var name = String(logValues[l][1] || '').trim(), cycles = Number(logValues[l][3]), mean = Number(logValues[l][6]);
    if (name && isFinite(cycles) && cycles > 0 && isFinite(mean)) logByKey[norm(name)] = { cycles: cycles, meanShare: mean };
  }
  var result = {};
  people.forEach(function (p) {
    var hist = logByKey[p.key], current = currentByKey[p.key];
    if (!hist || !current) return;
    // Points Log includes the current cycle. Remove its old ratio before finding
    // the ratio this cycle needs to reach a 100% average after the update.
    var targetRatio = hist.cycles * (1 - hist.meanShare) + (graveyardPending ? 1 : current.currentShare);
    result[p.key] = { cycles: hist.cycles, meanShare: hist.meanShare, currentShare: current.currentShare,
      amReq: current.amReq, pmReq: current.pmReq, adminDays: current.adminDays,
      targetRatio: Math.max(0, targetRatio) };
  });
  return result;
}

function readCurrentRosterAssignments_(sh, people) {
  var values = sh.getDataRange().getValues();
  var sections = findRosterSectionRows(sh);
  if (!sections) throw new PlanError('The selected Roster sheet does not have the expected schedule layout.');
  var byKey = {}, rowByKey = {}, fixed = [];
  people.forEach(function (p, i) {
    byKey[p.key] = i;
    fixed.push([[0, 0, 0, 0, 0, 0, 0], [0, 0, 0, 0, 0, 0, 0], [0, 0, 0, 0, 0, 0, 0]]);
  });
  for (var r = FIRST_ROW - 1; r < sections.personLastRow; r++) {
    var name = String(values[r][0] || '').trim();
    if (!name || !values[r][1]) continue;
    var key = norm(name);
    if (!(key in byKey)) throw new PlanError("Roster member '" + name + "' is missing from Names. Update the regular roster first.");
    if (key in rowByKey) throw new PlanError("Roster contains duplicate member '" + name + "'.");
    rowByKey[key] = r + 1;
    var p = byKey[key];
    for (var k = 0; k < 3; k++) for (var d = 0; d < 7; d++) {
      var v = String(values[r][START_COL[k] - 1 + d] || '').trim().toUpperCase();
      fixed[p][k][d] = (v === 'X' || v === 'AM') ? 1 : 0;
    }
  }
  people.forEach(function (p) {
    if (!(p.key in rowByKey)) throw new PlanError("Member '" + p.name + "' is not on the selected Roster. Update the regular roster first.");
  });
  return { fixed: fixed, rowByKey: rowByKey, personLastRow: sections.personLastRow };
}

function updateCurrentRosterGraveyard_(roster, m, X, rowByKey) {
  var cfg = m.cfg, summary = findRosterSectionRows(roster);
  var pointValues = [], labels = roster.getRange(1, 1, roster.getLastRow(), 1).getValues();
  var pointsRow = null, reqRow = null, meetsRow = null;
  labels.forEach(function (row, i) {
    var label = String(row[0] || '').trim();
    if (label === 'Points per shift (X)') pointsRow = i + 1;
    if (label === 'Requirement (from Settings)') reqRow = i + 1;
    if (label === 'Meets requirement?') meetsRow = i + 1;
  });
  if (!summary || !pointsRow || !reqRow || !meetsRow) throw new PlanError('The selected Roster sheet is missing summary or requirement rows.');
  for (var p = 0; p < m.P; p++) {
    var row = rowByKey[m.people[p].key], values = [];
    for (var d = 0; d < 7; d++) {
      var key = p + '_2_' + d;
      if (X[p][2][d]) values.push(m.counted(p, 2) ? 'X' : 'AM');
      else if (key in m.absent && cfg.showAbs) values.push(m.absent[key]);
      else values.push('');
    }
    roster.getRange(row, START_COL[2], 1, 7).setValues([values]);
  }
  for (var d2 = 0; d2 < 7; d2++) pointValues.push(Math.round(shiftBurden('G', m.dates[2][d2], cfg) * 10000) / 10000);
  roster.getRange(pointsRow, START_COL[2], 1, 7).setValues([pointValues]);

  roster.getRange(reqRow + 3, 1, 3, 2).setValues([
    ['Graveyard minimum on duty', 10],
    ['Graveyard minimum # names', 2],
    ['Graveyard maximum # names', 3]
  ]);
  var totalRow = summary.summaryRow, hashRow = totalRow + 1;
  var graveyardChecks = [], adminChecks = [];
  for (var d3 = 0; d3 < 7; d3++) {
    var col = colLetter(START_COL[2] + d3);
    graveyardChecks.push('=IF(AND(' + col + totalRow + '=10,' + col + hashRow + '>=2,' + col + hashRow + '<=3),"OK","CHECK")');
  }
  for (var d4 = 0; d4 < 5; d4++) {
    var adminCol = colLetter(START_COL[3] + d4);
    adminChecks.push('=IF(AND(' + adminCol + totalRow + '>=' + cfg.adminMin + ',' + adminCol + hashRow + '>=' + cfg.adminHash + '),"OK","CHECK")');
  }
  roster.getRange(meetsRow, START_COL[2], 1, 7).setFormulas([graveyardChecks]);
  roster.getRange(meetsRow, START_COL[3], 1, 5).setFormulas([adminChecks]);
}

function updateCurrentPointsSheet_(points, m, X) {
  var values = points.getDataRange().getValues(), rowsByKey = {};
  for (var r = 5; r < values.length; r++) {
    var name = String(values[r][1] || '').trim();
    if (name) rowsByKey[norm(name)] = r + 1;
  }
  var tab = m.points(X);
  points.getRange('B2').setValue(m.cfg.pDay);
  points.getRange('B3').setValue(m.cfg.pB2b);
  points.getRange('E2').setValue(m.avg[0]);
  points.getRange('E3').setValue(m.avg[1]);
  points.getRange('H2').setValue(m.avg[2]);
  points.getRange('H3').setValue(m.cfg.adminHrs);
  points.getRange('Q1').setValue('Graveyard planning: COMPLETE');
  for (var p = 0; p < m.P; p++) {
    var row = rowsByKey[m.people[p].key];
    if (!row) throw new PlanError("The matching Points sheet is missing '" + m.people[p].name + "'. Update the regular roster first.");
    points.getRange(row, 15).setValue(tab[p].req[2]);
  }
  return 'Updated existing Points sheet: ' + points.getName() + '.';
}

function writeLogInPlace_(ss, text) {
  var sh = ss.getSheetByName('Log');
  if (!sh) return;
  sh.clearContents();
  var lines = text.split('\n').map(function (line) { return [line]; });
  sh.getRange(1, 1, lines.length, 1).setValues(lines);
  sh.setColumnWidth(1, 900);
}

function updateEnhancedGraveyardInPlace_() {
  var ss = SpreadsheetApp.getActive();
  TZ = ss.getSpreadsheetTimeZone();
  var log = ['Mode: Enhanced Graveyard — update current roster in place'];
  var ui = SpreadsheetApp.getUi();
  try {
    ss.toast('Updating only the Graveyard week and Points sheet.', 'Roster', 30);
    var roster = currentRosterSheet_(ss);
    var rosterName = roster.getName();
    var pointsName = rosterName.replace(/^Roster/i, 'Points');
    var points = ss.getSheetByName(pointsName);
    if (!points) throw new PlanError("The current roster has no matching Points sheet called '" + pointsName + "'. Generate the regular roster once, then retry.");
    var dates = datesFromRoster_(roster);
    var inputs = readInputs(sheetValues(ss, 'Settings'), sheetValues(ss, 'Names'), sheetValues(ss, 'Unavailability'), dates);
    inputs.cfg.adminMin = inputs.cfg.gMin;
    inputs.cfg.adminHash = inputs.cfg.gHash;
    inputs.cfg.enhancedGraveyard = true;
    inputs.cfg.planGraveyard = true;
    inputs.cfg.gMin = 10;
    inputs.cfg.gHash = 2;
    inputs.cfg.gHashMax = 3;
    inputs.absent = graveyardOnlyMap_(inputs.absent);
    inputs.preferred = graveyardOnlyMap_(inputs.preferred);

    var current = readCurrentRosterAssignments_(roster, inputs.people);
    var history = currentSharePriorities_(ss, points, inputs.people);
    var m = new Model(inputs.cfg, inputs.people, inputs.dates, inputs.absent, inputs.preferred, null, current.fixed, history);
    m.precheck();
    var res = m.solve(LinearOptimizationService);
    var v = m.verify(res.X);
    log.push('Roster updated: ' + rosterName);
    log.push('Graveyard rules: exactly 10 on duty each day, 2–3 # names each day, no weekly or weekend cap.');
    log.push('All existing Graveyard assignments were retained; additional assignments fill the remaining spaces.');
    log.push('AM and PM assignments were kept as they were.');
    log.push('The solver favors higher point assignments for members below their 100% fair-share target in Points Log.');
    log.push('Solver: ' + res.status + ' (' + res.solveSeconds + ' seconds)');
    log = log.concat(v.lines);
    if (v.warnings.length) log.push('WARNINGS:\n- ' + v.warnings.join('\n- '));
    updateCurrentRosterGraveyard_(roster, m, res.X, current.rowByKey);
    SpreadsheetApp.flush();
    try {
      updateCurrentRosterImportTab_(ss, rosterName);
      log.push('Current Roster import tab refreshed.');
    } catch (mirrorError) {
      log.push('Current Roster import tab could not be refreshed: ' + String(mirrorError && mirrorError.message ? mirrorError.message : mirrorError));
    }
    log.push(updateCurrentPointsSheet_(points, m, res.X));
    SpreadsheetApp.flush();
    log.push(buildPointsLog(ss));
    writeLogInPlace_(ss, log.join('\n'));
    roster.activate();
    ui.alert('Enhanced Graveyard updated', log.join('\n').substring(0, 1500), ui.ButtonSet.OK);
  } catch (e) {
    log.push(e instanceof PlanError ? 'CANNOT UPDATE GRAVEYARD' : 'UNEXPECTED ERROR');
    log.push(String(e && e.stack ? e.stack : e));
    writeLogInPlace_(ss, log.join('\n'));
    ui.alert('Enhanced Graveyard was not updated', String(e && e.message ? e.message : e), ui.ButtonSet.OK);
  }
}

function generateScheduleMode_(enhancedGraveyard) {
  var ss = SpreadsheetApp.getActive();
  TZ = ss.getSpreadsheetTimeZone();
  var log = [];
  var ui = SpreadsheetApp.getUi();
  try {
    ss.toast('Building the roster. This can take a few minutes.', 'Roster', 30);
    ensureGraveyardPlanningSetting_(ss);
    ensureNamesWorkdayColumns_(ss);
    var inputs = readInputs(sheetValues(ss, 'Settings'), sheetValues(ss, 'Names'), sheetValues(ss, 'Unavailability'));
    // Admin checks keep using the standard Graveyard thresholds; the enhanced
    // override below applies only to the generated Graveyard week.
    inputs.cfg.adminMin = inputs.cfg.gMin;
    inputs.cfg.adminHash = inputs.cfg.gHash;
    if (enhancedGraveyard) {
      inputs.cfg.enhancedGraveyard = true;
      inputs.cfg.planGraveyard = true;
      inputs.cfg.gMin = 10;
      inputs.cfg.gHash = 2;
      inputs.cfg.gHashMax = 3;
    }
    var outputNames = outputSheetNames(inputs.dates);
    var previous = readPreviousSchedule(ss, inputs.people);
    var m = new Model(inputs.cfg, inputs.people, inputs.dates, inputs.absent, inputs.preferred, previous);
    if (!m.cfg.planGraveyard) log.push('Graveyard week: NOT PLANNED. It is marked pending and can be completed in place with Enhanced Graveyard.');
    log.push('Mode: ' + (m.cfg.enhancedGraveyard ? 'Enhanced Graveyard' : m.cfg.mode) + '   People: ' + m.P + '   Plan: ' + fmtDate(m.dates[0][0]) + ' to ' + fmtDate(m.dates[2][6]) +
      '   Admin week: ' + fmtDate(m.dates[3][0]) + ' to ' + fmtDate(m.dates[3][4]));
    log.push('Preferred days off: soft penalty ' + m.cfg.preferOffPenalty + ' per matching assignment.');
    m.precheck();
    var res = m.solve(LinearOptimizationService);
    var v = m.verify(res.X);
    log.push('Solver: ' + res.status);
    if (res.solveSeconds !== undefined) log.push('Solver time: ' + res.solveSeconds + ' seconds');
    if (previous) {
      var diversity = scheduleDifference(previous, res.X);
      log.push('Schedule variation: ' + diversity.changed + ' of ' + diversity.previousAssignments +
        ' previous shift assignments changed (at least one change is required when a prior roster exists).');
    }
    log = log.concat(v.lines);
    if (v.warnings.length) {
      log.push('WARNINGS:');
      v.warnings.forEach(function (w) { log.push('- ' + w); });
    } else {
      log.push('All hard rules and goals are met.');
    }
    archiveOutputSheets(ss, outputNames.roster, outputNames.points);
    writeRoster(ss, m, res.X, outputNames.roster, outputNames.points);
    try {
      updateCurrentRosterImportTab_(ss, outputNames.roster);
      log.push('Current Roster import tab now points to ' + outputNames.roster + '.');
    } catch (mirrorError) {
      log.push('Current Roster import tab could not be refreshed: ' + String(mirrorError && mirrorError.message ? mirrorError.message : mirrorError));
    }
    try {
      SpreadsheetApp.flush();
      log.push(buildPointsLog(ss));
    } catch (reportError) {
      log.push('Points Log could not be refreshed: ' + String(reportError && reportError.message ? reportError.message : reportError));
    }
    writeLog(ss, log.join('\n'));
    ss.getSheetByName(outputNames.roster).activate();
    ui.alert('Roster generated', log.join('\n').substring(0, 1500), ui.ButtonSet.OK);
  } catch (e) {
    if (e instanceof PlanError) {
      log.push('CANNOT BUILD ROSTER');
      log.push(e.message);
      writeLog(ss, log.join('\n'));
      ui.alert('No roster was created', log.join('\n').substring(0, 1500), ui.ButtonSet.OK);
    } else {
      log.push('UNEXPECTED ERROR');
      log.push(String(e && e.stack ? e.stack : e));
      writeLog(ss, log.join('\n'));
      ui.alert('Something went wrong', String(e && e.message ? e.message : e), ui.ButtonSet.OK);
    }
  }
}

function sheetValues(ss, name) {
  var sh = ss.getSheetByName(name);
  if (!sh) throw new PlanError("The spreadsheet has no sheet called '" + name + "'.");
  var lastRow = Math.max(sh.getLastRow(), 1);
  var lastCol = Math.max(sh.getLastColumn(), 1);
  return sh.getRange(1, 1, lastRow, lastCol).getValues();
}

function outputSheetNames(dates) {
  var first = fmtDateLong(dates[0][0]);
  var last = fmtDateLong(dates[3][dates[3].length - 1]);
  return { roster: 'Roster ' + first + ' to ' + last, points: 'Points ' + first + ' to ' + last };
}

// Keep a stable source range for destination workbooks using IMPORTRANGE.
// The formula follows the newest generated Roster and also reflects manual
// edits to that roster without changing the destination formula each cycle.
function updateCurrentRosterImportTab_(ss, rosterName) {
  var marker = 'Managed by the roster scheduler for IMPORTRANGE.';
  var source = ss.getSheetByName(rosterName);
  if (!source) throw new PlanError("The generated Roster sheet '" + rosterName + "' could not be found.");
  var mirror = ss.getSheetByName('Current Roster');
  if (mirror && mirror.getRange('A1').getNote() !== marker) {
    throw new PlanError("A sheet named 'Current Roster' already exists and is not managed by the scheduler. Rename it, then generate again.");
  }
  if (!mirror) mirror = ss.insertSheet('Current Roster');
  mirror.clearContents();
  var lastRow = Math.max(source.getLastRow(), 1), lastCol = Math.max(source.getLastColumn(), 1);
  mirror.getRange('A1').setFormula('=ARRAYFORMULA(' + quotedSheet(rosterName) + '!A1:' + colLetter(lastCol) + lastRow + ')');
  mirror.getRange('A1').setNote(marker);
  SpreadsheetApp.flush();
  return mirror;
}

function readPreviousSchedule(ss, people) {
  var byKey = {};
  people.forEach(function (p, i) { byKey[p.key] = i; });
  var old = ss.getSheets().filter(function (sh) { return /^Roster(?: |$)/i.test(sh.getName()); })[0];
  if (!old) return null;
  var vals = old.getDataRange().getValues(), out = [];
  for (var p = 0; p < people.length; p++) out.push([[], [], []]);
  for (var r = FIRST_ROW - 1; r < vals.length; r++) {
    var name = vals[r][0];
    if (name === '' || name === null || name === undefined) continue;
    var key = norm(name);
    if (!(key in byKey)) continue;
    var pi = byKey[key];
    for (var k = 0; k < 3; k++) {
      if (!out[pi][k].length) for (var z = 0; z < 7; z++) out[pi][k].push(0);
      for (var d = 0; d < 7; d++) {
        var value = String(vals[r][START_COL[k] - 1 + d] || '').trim().toUpperCase();
        if (value === 'X' || value === 'AM') out[pi][k][d] = 1;
      }
    }
  }
  var count = out.reduce(function (s, row) {
    return s + row.reduce(function (a, shifts) { return a + shifts.reduce(function (b, x) { return b + x; }, 0); }, 0);
  }, 0);
  return count ? out : null;
}

function scheduleDifference(previous, X) {
  var total = 0, same = 0;
  for (var p = 0; p < X.length; p++) for (var k = 0; k < 3; k++) for (var d = 0; d < 7; d++) {
    if (previous[p][k][d]) { total++; if (X[p][k][d]) same++; }
  }
  return { previousAssignments: total, changed: total - same };
}

function archiveOutputSheets(ss, rosterName, pointsName) {
  // A numeric timestamp avoids a locale/time-zone dependency while keeping archive names unique.
  var now = new Date().getTime();
  ss.getSheets().forEach(function (sh) {
    var name = sh.getName();
    if (/^Points Log$/i.test(name)) return;
    if (!/^(Roster|Points)( |$)/i.test(name)) return;
    sh.hideSheet();
    if (name === rosterName || name === pointsName) {
      var base = 'Archived ' + name + ' (' + now + ')', archiveName = base, n = 2;
      while (ss.getSheetByName(archiveName)) archiveName = base + ' ' + n++;
      sh.setName(archiveName);
    }
  });
}

function quotedSheet(name) { return "'" + String(name).replace(/'/g, "''") + "'"; }

// ------------------------------------------------------------------ dates (day numbers since 1970-01-01)
function dnFromYMD(y, mo, d) { return Math.round(Date.UTC(y, mo - 1, d) / 86400000); }
function ymdFromDn(dn) { var d = new Date(dn * 86400000); return [d.getUTCFullYear(), d.getUTCMonth() + 1, d.getUTCDate()]; }
function weekdayOf(dn) { return (((dn + 3) % 7) + 7) % 7; }   // Mon = 0
function pad2(n) { return (n < 10 ? '0' : '') + n; }
function fmtDate(dn) { var t = ymdFromDn(dn); return DAYS[weekdayOf(dn)] + ' ' + pad2(t[2]) + ' ' + MONTHS[t[1] - 1]; }
function fmtDateLong(dn) { var t = ymdFromDn(dn); return pad2(t[2]) + ' ' + MONTHS[t[1] - 1] + ' ' + t[0]; }

function toDate(v, what) {
  if (v === null || v === undefined || (typeof v === 'string' && v.trim() === '')) return null;
  if (v instanceof Date) {
    var s = (typeof Utilities !== 'undefined' && TZ) ? Utilities.formatDate(v, TZ, 'yyyy-MM-dd')
      : (v.getFullYear() + '-' + pad2(v.getMonth() + 1) + '-' + pad2(v.getDate()));
    var p = s.split('-');
    return dnFromYMD(+p[0], +p[1], +p[2]);
  }
  var str = String(v).trim();
  if (/^(indefinite|ind|ongoing)$/i.test(str)) return 'INDEF';
  var m = str.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (m) return dnFromYMD(+m[1], +m[2], +m[3]);
  m = str.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
  if (m) return dnFromYMD(+m[3], +m[2], +m[1]);
  throw new PlanError(what + ": cannot read '" + v + "' as a date (use a real date or YYYY-MM-DD).");
}

function norm(s) { return String(s).replace(/\u00a0/g, ' ').replace(/#/g, '').replace(/\s+/g, '').toLowerCase(); }
function yes(v) { return (v !== null && v !== undefined && String(v).trim() !== '') ? /^y/i.test(String(v).trim()) : false; }
function num(v, label) {
  var n = Number(v);
  if (v === '' || v === null || isNaN(n)) throw new PlanError("Settings: '" + label + "' must be a number (found '" + v + "').");
  return n;
}
function optionalWorkDays(v, label, name) {
  if (v === '' || v === null || v === undefined || (typeof v === 'string' && v.trim() === '')) return null;
  var n = Number(v);
  if (!isFinite(n) || Math.floor(n) !== n || n < 0 || n > 7) {
    throw new PlanError("Names: '" + label + "' for " + name + " must be a whole number from 0 to 7, or left blank.");
  }
  return n;
}

// ------------------------------------------------------------------ input
function readInputs(settingsVals, namesVals, unavailVals, scheduleDates) {
  var raw = {};
  settingsVals.forEach(function (row) {
    if (row[0] !== '' && row[0] !== null && row[0] !== undefined) raw[String(row[0]).trim()] = row[1];
  });
  var cfg = {};
  Object.keys(SETTINGS).forEach(function (label) {
    var key = SETTINGS[label][0], def = SETTINGS[label][1];
    var v = Object.prototype.hasOwnProperty.call(raw, label) ? raw[label] : def;
    if ((v === '' || v === null || v === undefined) && !OPTIONAL[key] && def === null) throw new PlanError("Settings: missing '" + label + "'.");
    cfg[key] = (v === undefined) ? null : v;
  });
  var start = scheduleDates ? scheduleDates[0][0] : toDate(cfg.start, 'Settings: AM week start date');
  if (start === null || start === 'INDEF') throw new PlanError("Settings: 'AM week start date' is empty.");
  cfg.startDn = start;
  cfg.mode = String(cfg.mode).trim().toUpperCase();
  if (cfg.mode !== 'QUOTA' && cfg.mode !== 'EXACT') throw new PlanError('Settings: Mode must be QUOTA or EXACT.');
  if (cfg.preferOffPenalty < 0) throw new PlanError("Settings: 'Preferred days off penalty' must be zero or greater.");
    ['ampmMin', 'ampmHash', 'gMin', 'gHash', 'qAm', 'qPm', 'qG', 'permPm', 'permG', 'tlimit', 'pRecov', 'pB2b', 'pDay', 'adminHrs', 'wEve', 'wNight', 'wWk', 'preferOffPenalty']
    .forEach(function (k) {
      var label = Object.keys(SETTINGS).filter(function (l) { return SETTINGS[l][0] === k; })[0];
      cfg[k] = num(cfg[k], label);
    });
  cfg.ampmHmax = (cfg.ampmHmax === '' || cfg.ampmHmax === null) ? null : num(cfg.ampmHmax, 'AM/PM maximum # names');
  cfg.wkMax = (cfg.wkMax === '' || cfg.wkMax === null) ? null : num(cfg.wkMax, 'Max weekend shifts per person per week');
  cfg.showAbs = yes(cfg.showAbs);
  cfg.balance = yes(cfg.balance);
  cfg.planGraveyard = yes(cfg.planGraveyard);
  cfg.wkType = String(cfg.wkType).trim().toUpperCase();
  if (cfg.wkType !== 'HARD' && cfg.wkType !== 'SOFT') throw new PlanError('Settings: Weekend limit type must be HARD or SOFT.');
  if (cfg.tlimit > MAX_SOLVE_SECONDS) cfg.tlimit = MAX_SOLVE_SECONDS;
  cfg.wkSet = {};
  String(cfg.wkDays).replace(/;/g, ',').split(',').forEach(function (tok) {
    var t = tok.trim().toLowerCase().substring(0, 3);
    if (!t) return;
    var i = DAYS.map(function (x) { return x.toLowerCase(); }).indexOf(t);
    if (i < 0) throw new PlanError("Settings: '" + tok.trim() + "' is not a day name (use Mon, Tue, ... Sun).");
    cfg.wkSet[i] = true;
  });

  var people = [], seen = {};
  var workdayHeaders = ['AM workdays/week', 'PM workdays/week', 'Graveyard workdays/week'];
  var hasWorkdayOverrideColumns = workdayHeaders.every(function (label, i) {
    return String(namesVals[0] && namesVals[0][4 + i] || '').trim() === label;
  });
  for (var r = 1; r < namesVals.length; r++) {
    var row = namesVals[r];
    if (row[0] === '' || row[0] === null || row[0] === undefined || String(row[0]).trim() === '') continue;
    var name = String(row[0]).replace(/\u00a0/g, ' ').trim();
    var key = norm(name);
    if (seen[key]) throw new PlanError("Names: '" + name + "' appears twice.");
    seen[key] = true;
    var sec = (row[1] === '' || row[1] === null || row[1] === undefined) ? null : String(row[1]).trim();
    if (sec === null) throw new PlanError("Names: '" + name + "' has no section.");
    people.push({
      name: name, key: key, sec: sec, hash: name.charAt(0) === '#', perm: yes(row[2]), nobonus: yes(row[3]),
      amDays: hasWorkdayOverrideColumns ? optionalWorkDays(row[4], 'AM days/week', name) : null,
      pmDays: hasWorkdayOverrideColumns ? optionalWorkDays(row[5], 'PM days/week', name) : null,
      graveyardDays: hasWorkdayOverrideColumns ? optionalWorkDays(row[6], 'Graveyard days/week', name) : null
    });
  }
  if (!people.length) throw new PlanError('Names: the sheet is empty.');

  var dates = scheduleDates || [];
  if (!scheduleDates) {
    for (var k = 0; k < 4; k++) {
      var arr = [];
      for (var d = 0; d < BLOCK_LEN[k]; d++) arr.push(start + 7 * k + d);
      dates.push(arr);
    }
  }
  var planStart = dates[0][0], planEnd = dates[3][4];
  var byKey = {};
  people.forEach(function (p, i) { byKey[p.key] = i; });
  var absent = {}, preferred = {}, problems = [];
  for (var u = 1; u < unavailVals.length; u++) {
    var ur = unavailVals[u], rn = u + 1;
    if (ur[0] === '' || ur[0] === null || ur[0] === undefined || String(ur[0]).trim() === '') continue;
    var ukey = norm(ur[0]);
    if (!(ukey in byKey)) { problems.push('Unavailability row ' + rn + ": '" + String(ur[0]).trim() + "' is not on the Names sheet."); continue; }
    var dFrom = toDate(ur[1], 'Unavailability row ' + rn + " 'From'");
    var dTo = toDate(ur[2], 'Unavailability row ' + rn + " 'To'");
    if (dFrom === 'INDEF') throw new PlanError('Unavailability row ' + rn + ": 'From' cannot be 'Indefinite'.");
    if (dFrom === null) dFrom = planStart;
    if (dTo === null) dTo = dFrom;
    if (dTo === 'INDEF') dTo = planEnd;
    var reason = (ur[3] === '' || ur[3] === null || ur[3] === undefined) ? 'N/A' : String(ur[3]).trim();
    if (reason.toUpperCase() === 'X' || reason.toUpperCase() === 'AM') reason += '*';
    if (dTo < dFrom) { problems.push('Unavailability row ' + rn + ": 'To' is before 'From'."); continue; }
    for (var kk = 0; kk < 4; kk++) {
      for (var dd = 0; dd < BLOCK_LEN[kk]; dd++) {
        if (dates[kk][dd] >= dFrom && dates[kk][dd] <= dTo) absent[byKey[ukey] + '_' + kk + '_' + dd] = reason;
      }
    }
  }
  // Soft preferences are entered in columns F:G of Unavailability, with headers
  // on row 1 and one member plus one or more dates per row starting on row 2.
  for (var f = 1; f < unavailVals.length; f++) {
    var fr = unavailVals[f], fn = String(fr[5] || '').trim();
    var prefDateValue = fr[6];
    if (!fn && (prefDateValue === '' || prefDateValue === null || prefDateValue === undefined)) continue;
    if (!fn || String(fn).toLowerCase() === 'name') continue;
    var fkey = norm(fn);
    if (!(fkey in byKey)) { problems.push('Preferred days off row ' + (f + 1) + ": '" + fn + "' is not on the Names sheet."); continue; }
    var prefDns;
    try {
      prefDns = parsePreferredDates(prefDateValue, dates, 'Preferred days off row ' + (f + 1));
    } catch (prefError) {
      problems.push(String(prefError.message || prefError));
      continue;
    }
    prefDns.forEach(function (prefDn) {
      for (var pk = 0; pk < 3; pk++) {
        var pd = dates[pk].indexOf(prefDn);
        if (pd >= 0) preferred[byKey[fkey] + '_' + pk + '_' + pd] = true;
      }
    });
    // The manually completed Admin week is not part of the solver.
  }
  if (problems.length) throw new PlanError(problems.join('\n'));
  return { cfg: cfg, people: people, dates: dates, absent: absent, preferred: preferred };
}

// Accept real dates, full dates, or comma-separated DD/MM dates. Dates without a
// year are matched against the current schedule so year-boundary cycles resolve correctly.
function parsePreferredDates(value, dates, label) {
  if (value === null || value === undefined || String(value).trim() === '') {
    throw new PlanError(label + ': enter at least one date.');
  }
  var parts = value instanceof Date ? [value] : String(value).split(',');
  var planned = [];
  dates.forEach(function (block) { block.forEach(function (dn) { planned.push(dn); }); });
  var results = [];
  parts.forEach(function (part) {
    if (part instanceof Date) {
      var dateNumber = toDate(part, label + " 'Date'");
      if (dateNumber === null || dateNumber === 'INDEF') throw new PlanError(label + ': enter specific dates.');
      results.push(dateNumber);
      return;
    }
    var text = String(part).trim();
    if (!text) throw new PlanError(label + ': remove empty items after commas.');
    var shortDate = text.match(/^(\d{1,2})\/(\d{1,2})$/);
    if (shortDate) {
      var day = Number(shortDate[1]), month = Number(shortDate[2]);
      var match = null;
      for (var i = 0; i < planned.length; i++) {
        var ymd = ymdFromDn(planned[i]);
        if (ymd[1] === month && ymd[2] === day) { match = planned[i]; break; }
      }
      if (match === null) throw new PlanError(label + ": '" + text + "' does not fall within this schedule period (use DD/MM).");
      results.push(match);
      return;
    }
    var fullDate = toDate(text, label + " 'Date' (use DD/MM or DD/MM/YYYY)");
    if (fullDate === null || fullDate === 'INDEF') throw new PlanError(label + ': enter specific dates.');
    results.push(fullDate);
  });
  return results;
}

// ------------------------------------------------------------------ points
function shiftBurden(kind, dn, cfg) {
  var range = { AM: [8, 15], PM: [15, 22], G: [22, 32] }[kind];
  var w = weekdayOf(dn);
  var tot = 0;
  for (var h = range[0]; h < range[1]; h++) {
    var hr = h % 24;
    var wt = (hr >= 22 || hr < 6) ? cfg.wNight : ((hr >= 18 && hr < 22) ? cfg.wEve : 1.0);
    var offset = w * 24 + h;                       // hours since Monday 00:00 of that week
    tot += wt + ((offset >= 4 * 24 + 15 && offset < 7 * 24) ? cfg.wWk : 0);
  }
  if (kind === 'G') tot += cfg.pRecov;
  return tot;
}

// ------------------------------------------------------------------ model
function Model(cfg, people, dates, absent, preferred, previous, fixedAssignments, historicShare) {
  var self = this;
  this.cfg = cfg; this.people = people; this.dates = dates; this.absent = absent; this.preferred = preferred || {};
  this.previous = previous;
  this.fixedAssignments = fixedAssignments || null;
  this.historicShare = historicShare || {};
  this.P = people.length;
  this.avail = [];
  for (var p = 0; p < this.P; p++) {
    this.avail.push([]);
    for (var k = 0; k < 3; k++) {
      var row = [];
      for (var d = 0; d < 7; d++) row.push(((p + '_' + k + '_' + d) in absent) ? 0 : 1);
      this.avail[p].push(row);
    }
  }
  this.quota = [cfg.qAm, cfg.qPm, cfg.qG];
  this.reqTotal = [cfg.ampmMin, cfg.ampmMin, cfg.planGraveyard || cfg.enhancedGraveyard ? cfg.gMin : 0];
  this.reqHash = [cfg.ampmHash, cfg.ampmHash, cfg.planGraveyard || cfg.enhancedGraveyard ? cfg.gHash : 0];
  this.maxHash = [cfg.ampmHmax, cfg.ampmHmax, cfg.enhancedGraveyard ? cfg.gHashMax : null];
  this.secs = [];
  people.forEach(function (pr) { if (self.secs.indexOf(pr.sec) < 0) self.secs.push(pr.sec); });
  var kinds = ['AM', 'PM', 'G'];
  this.burden = []; this.amBurden = []; this.avg = []; this.avgAm = []; this.wkdays = [];
  for (var k2 = 0; k2 < 3; k2++) {
    var b = [], a = [], wk = [];
    for (var d2 = 0; d2 < 7; d2++) {
      b.push(shiftBurden(kinds[k2], dates[k2][d2], cfg));
      a.push(shiftBurden('AM', dates[k2][d2], cfg));
      if (cfg.wkSet[weekdayOf(dates[k2][d2])]) wk.push(d2);
    }
    this.burden.push(b); this.amBurden.push(a); this.wkdays.push(wk);
    this.avg.push(b.reduce(function (s, x) { return s + x; }, 0) / 7);
    this.avgAm.push(a.reduce(function (s, x) { return s + x; }, 0) / 7);
  }
}

Model.prototype.availSum = function (p, k) { return this.avail[p][k].reduce(function (s, x) { return s + x; }, 0); };
Model.prototype.counted = function (p, k) { return !(this.people[p].perm && k > 0); };
Model.prototype.absenceReduction = function (p, k) {
  var count = 0;
  for (var d = 0; d < BLOCK_LEN[k]; d++) {
    var reason = this.absent[p + '_' + k + '_' + d];
    if (/^(OFF|OSL|LL)$/i.test(String(reason || '').trim())) count++;
  }
  return count;
};
Model.prototype.personQuota = function (p, k) {
  var person = this.people[p], override = [person.amDays, person.pmDays, person.graveyardDays][k];
  var usual = override !== null && override !== undefined ? override :
    (person.perm ? [this.cfg.qAm, this.cfg.permPm, this.cfg.permG][k] : this.quota[k]);
  return Math.max(0, usual - this.absenceReduction(p, k));
};
Model.prototype.fixedDays = function (p, k) {
  if (this.people[p].perm) return Math.min(this.personQuota(p, k), this.availSum(p, k));
  return null;
};
Model.prototype.graveyardFairShareDays = function (p) {
  if (!this.cfg.enhancedGraveyard || !this.counted(p, 2)) return 0;
  var available = 0;
  for (var q = 0; q < this.P; q++) if (this.counted(q, 2)) available += this.availSum(q, 2);
  return available ? this.reqTotal[2] * 7 * this.availSum(p, 2) / available : 0;
};
Model.prototype.bonus = function (p) { return this.people[p].nobonus ? 0 : this.cfg.pDay; };
Model.prototype.shiftValue = function (p, k, d) { return this.counted(p, k) ? this.burden[k][d] : this.amBurden[k][d]; };
Model.prototype.xName = function (p, k, d) { return 'x' + p + '_' + k + '_' + d; };

Model.prototype.precheck = function () {
  var issues = [], cfg = this.cfg, self = this;
  if (cfg.enhancedGraveyard && this.fixedAssignments) {
    for (var fp = 0; fp < this.P; fp++) for (var fd = 0; fd < 7; fd++) {
      if (this.fixedAssignments[fp][2][fd] && !this.avail[fp][2][fd]) {
        issues.push('Graveyard ' + fmtDate(this.dates[2][fd]) + ': existing assignment for ' + this.people[fp].name +
          ' conflicts with a current hard unavailability entry. Clear that absence or remove the old assignment before running Enhanced Graveyard.');
      }
    }
  }
  for (var k = 0; k < 3; k++) {
    if (cfg.enhancedGraveyard && k < 2) continue;
    if (!cfg.planGraveyard && !cfg.enhancedGraveyard && k === 2) continue;
    for (var d = 0; d < 7; d++) {
      var cnt = [];
      for (var p = 0; p < this.P; p++) if (this.counted(p, k) && this.avail[p][k][d]) cnt.push(p);
      var nh = cnt.filter(function (q) { return self.people[q].hash; }).length;
      var day = fmtDate(this.dates[k][d]);
      if (cnt.length < this.reqTotal[k]) issues.push(BLOCKS[k] + ' ' + day + ': only ' + cnt.length + ' people are available but ' + this.reqTotal[k] + ' are needed.');
      if (nh < this.reqHash[k]) issues.push(BLOCKS[k] + ' ' + day + ": only " + nh + " '#' names are available but " + this.reqHash[k] + ' are needed.');
    }
  }
  if (cfg.mode === 'QUOTA') {
    for (var k2 = 0; k2 < 3; k2++) {
      if (cfg.enhancedGraveyard) continue;
      if (!cfg.planGraveyard && k2 === 2) continue;
      var supply = 0;
      for (var p2 = 0; p2 < this.P; p2++) {
        if (!this.counted(p2, k2)) continue;
        var fd = this.fixedDays(p2, k2);
        supply += (fd !== null) ? fd : Math.min(this.personQuota(p2, k2), this.availSum(p2, k2));
      }
      var need = this.reqTotal[k2] * 7;
      if (supply < need) issues.push(BLOCKS[k2] + ' week: the days per person supply ' + supply + ' shifts but ' + need + ' are needed (' +
        this.reqTotal[k2] + ' people x 7 days). Raise the days per person or add people.');
    }
  }
  if (cfg.wkMax !== null && cfg.wkType === 'HARD') {
    var lim = cfg.wkMax;
    for (var k3 = 0; k3 < 3; k3++) {
      if (cfg.enhancedGraveyard) continue;
      if (!cfg.planGraveyard && k3 === 2) continue;
      var W = this.wkdays[k3];
      if (!W.length) continue;
      var demand = this.reqTotal[k3] * W.length;
      var peopleOk = [];
      for (var p3 = 0; p3 < this.P; p3++) {
        if (this.counted(p3, k3) && W.some(function (dd) { return self.avail[p3][k3][dd]; })) peopleOk.push(p3);
      }
      if (demand > lim * peopleOk.length) issues.push(BLOCKS[k3] + ' week: ' + W.length + ' weekend days need ' + demand + ' shifts, but only ' +
        peopleOk.length + ' people are free on a weekend day and each may work ' + lim + '. Use SOFT, raise the limit, or add people.');
      peopleOk.forEach(function (q) {
        var aNon = 0;
        for (var dd = 0; dd < 7; dd++) if (W.indexOf(dd) < 0 && self.avail[q][k3][dd]) aNon++;
        if (cfg.mode === 'QUOTA') {
          var f = self.fixedDays(q, k3);
          var n = (f !== null) ? f : Math.min(self.personQuota(q, k3), self.availSum(q, k3));
          if (n - aNon > lim) issues.push(BLOCKS[k3] + ': ' + self.people[q].name + ' must work ' + n + ' days but is only free on ' + aNon +
            ' non-weekend days, so cannot stay within ' + lim + ' weekend shift(s).');
        }
      });
    }
  }
  if (issues.length) throw new PlanError('The rules cannot be met with these inputs:\n- ' + issues.join('\n- '));
};

Model.prototype.exactBounds = function () {
  var fixed = {}, bounds = {}, self = this;
  for (var k = 0; k < 3; k++) {
    if (this.cfg.enhancedGraveyard && k === 2) continue;
    if (!this.cfg.planGraveyard && k === 2) continue;
    [true, false].forEach(function (cls) {
      var members = [];
      for (var p = 0; p < self.P; p++) if (self.people[p].hash === cls && self.counted(p, k)) members.push(p);
      var free = [], fx = 0;
      members.forEach(function (q) {
        var fd = self.fixedDays(q, k), a = self.availSum(q, k);
        if (fd !== null) { fixed[q + '_' + k] = fd; fx += fd; }
        else {
          var personCap = Math.min(self.personQuota(q, k), a);
          if (personCap === 0) fixed[q + '_' + k] = 0;
          else free.push(q);
        }
      });
      var want = (cls ? self.reqHash[k] : self.reqTotal[k] - self.reqHash[k]) * 7;
      var label = BLOCKS[k] + ' ' + (cls ? '# names' : 'non-# names');
      var total, lo = 0, hi = 0;
      while (true) {
        total = want - fx;
        if (total < 0) throw new PlanError(label + ': fixed shifts already exceed what is needed.');
        if (!free.length) {
          if (total !== 0) throw new PlanError(label + ': nobody is available to cover ' + total + ' shifts.');
          break;
        }
        lo = Math.floor(total / free.length); hi = Math.ceil(total / free.length);
        var short = free.filter(function (q) { return Math.min(self.personQuota(q, k), self.availSum(q, k)) < lo; });
        if (!short.length) break;
        short.forEach(function (q) {
          var a = Math.min(self.personQuota(q, k), self.availSum(q, k));
          fixed[q + '_' + k] = a; fx += a;
          free.splice(free.indexOf(q), 1);
        });
      }
      if (free.length) {
        var totalCapacity = free.reduce(function (sum, q) { return sum + Math.min(self.personQuota(q, k), self.availSum(q, k)); }, 0);
        if (total > totalCapacity) throw new PlanError(label + ': remaining members can cover only ' + totalCapacity + ' of the ' + total +
          ' required shifts after OFF/OSL/LL reductions. Add eligible members or change the staffing requirements.');
        free.forEach(function (q) { bounds[q + '_' + k] = [lo, Math.min(hi, self.personQuota(q, k), self.availSum(q, k))]; });
      }
    });
  }
  return { fixed: fixed, bounds: bounds };
};

Model.prototype.requiredDays = function (X) {
  var req = [];
  for (var p = 0; p < this.P; p++) {
    var row = [];
    for (var k = 0; k < 3; k++) {
      if (!this.cfg.planGraveyard && !this.cfg.enhancedGraveyard && k === 2) { row.push(0); continue; }
      if (this.cfg.enhancedGraveyard && k === 2) row.push(this.graveyardFairShareDays(p));
      else if (this.cfg.mode === 'EXACT') row.push(X[p][k].reduce(function (s, x) { return s + x; }, 0));
      else {
        var fd = this.fixedDays(p, k);
        row.push(fd !== null ? fd : Math.min(this.personQuota(p, k), this.availSum(p, k)));
      }
    }
    req.push(row);
  }
  return req;
};

/** Build and solve the model. `service` is LinearOptimizationService (or a stand-in when testing). */
Model.prototype.solve = function (service) {
  var cfg = this.cfg, P = this.P, self = this;
  var exact = cfg.mode === 'EXACT';
  var balance = cfg.balance && (!exact || cfg.enhancedGraveyard);
  var names = [], lbs = [], ubs = [], types = [], objs = [];
  var INT = service.VariableType.INTEGER, CONT = service.VariableType.CONTINUOUS;
  function addVar(name, lb, ub, type, obj) { names.push(name); lbs.push(lb); ubs.push(ub); types.push(type); objs.push(obj || 0); }

  for (var p = 0; p < P; p++)
    for (var k = 0; k < 3; k++)
      for (var d = 0; d < 7; d++) {
        var prefCost = ((p + '_' + k + '_' + d) in this.preferred) ? cfg.preferOffPenalty : 0;
        var fixedValue = null;
        if (cfg.enhancedGraveyard && this.fixedAssignments) {
          if (k < 2) fixedValue = this.fixedAssignments[p][k][d];
          else if (this.fixedAssignments[p][k][d]) fixedValue = 1;
        }
        var disabledGraveyard = !cfg.planGraveyard && !cfg.enhancedGraveyard && k === 2;
        addVar(this.xName(p, k, d), fixedValue === null ? 0 : fixedValue,
          disabledGraveyard ? 0 : (fixedValue === null ? (this.avail[p][k][d] ? 1 : 0) : fixedValue), INT, prefCost);
      }
  for (var k1 = 0; k1 < 3; k1++) {
    this.secs.forEach(function (s, i) {
      var sectionCount = 0;
      for (var q1 = 0; q1 < P; q1++) if (self.counted(q1, k1) && self.people[q1].sec === s) sectionCount++;
      addVar('smax_' + k1 + '_' + i, 0, sectionCount, CONT, 0);
      addVar('smin_' + k1 + '_' + i, 0, sectionCount, CONT, 0);
      addVar('sex_' + k1 + '_' + i, 0, sectionCount, CONT, 10);  // section even across days
    });
    var countedCount = 0, hashCount = 0;
    for (var q2 = 0; q2 < P; q2++) if (this.counted(q2, k1)) {
      countedCount++;
      if (this.people[q2].hash) hashCount++;
    }
    addVar('tmax_' + k1, 0, countedCount, CONT, 0); addVar('tmin_' + k1, 0, countedCount, CONT, 0); addVar('tex_' + k1, 0, countedCount, CONT, 1000);  // headcount even
    addVar('hmax_' + k1, 0, hashCount, CONT, 0); addVar('hmin_' + k1, 0, hashCount, CONT, 0); addVar('hex_' + k1, 0, hashCount, CONT, 100);  // '#' names even
  }
  for (var p1 = 0; p1 < P; p1++)
    for (var k2 = 0; k2 < 3; k2++) addVar('wex_' + p1 + '_' + k2, 0, this.wkdays[k2].length, CONT, 50);  // extra weekend shifts
  if (balance) {
    addVar('pmax', 0, 100, CONT, 200);
    addVar('pmin', 0, 100, CONT, -200);
  }
  if (balance || cfg.enhancedGraveyard) {
    for (var p2 = 0; p2 < P; p2++) {
      if (this.people[p2].perm) continue;
      for (var d2 = 0; d2 < 6; d2++) addVar('y_' + p2 + '_' + d2, 0, 1, CONT, 0);
    }
  }
  if (cfg.enhancedGraveyard) {
    for (var hp = 0; hp < P; hp++) if (this.historicShare[this.people[hp].key]) {
      addVar('histBelow_' + hp, 0, 100, CONT, 250);
      addVar('histAbove_' + hp, 0, 100, CONT, 250);
    }
  }

  var cLo = [], cHi = [], cNames = [], cCoefs = [];
  function add(co, lo, hi) {
    var ns = [], cs = [];
    co.forEach(function (t) { ns.push(t[0]); cs.push(t[1]); });
    cNames.push(ns); cCoefs.push(cs); cLo.push(lo); cHi.push(hi);
  }
  function neg(v) { return v === -Infinity ? -BIG : v; }
  function pos(v) { return v === Infinity ? BIG : v; }
  function addC(co, lo, hi) { add(co, neg(lo), pos(hi)); }

  var fixed = {}, bounds = {};
  if (exact && !cfg.enhancedGraveyard) { var eb = this.exactBounds(); fixed = eb.fixed; bounds = eb.bounds; }
  var need = {};
  for (var p3 = 0; p3 < P; p3++) {
    for (var k3 = 0; k3 < 3; k3++) {
      var co = [];
      for (var d3 = 0; d3 < 7; d3++) co.push([this.xName(p3, k3, d3), 1]);
      var key = p3 + '_' + k3;
      if (!cfg.planGraveyard && !cfg.enhancedGraveyard && k3 === 2) { need[key] = 0; continue; }
      var a = this.availSum(p3, k3), fd = this.fixedDays(p3, k3);
      if (exact) {
        if (this.cfg.enhancedGraveyard && k3 < 2) {
          var fixedCount = 0;
          for (var fdn = 0; fdn < 7; fdn++) fixedCount += this.fixedAssignments[p3][k3][fdn];
          need[key] = fixedCount;
        } else if (this.cfg.enhancedGraveyard && k3 === 2) {
          need[key] = this.graveyardFairShareDays(p3);
        } else if (key in fixed) { addC(co, fixed[key], fixed[key]); need[key] = fixed[key]; }
        else if (key in bounds) { addC(co, bounds[key][0], bounds[key][1]); need[key] = (bounds[key][0] + bounds[key][1]) / 2; }
      } else {
        if (this.cfg.enhancedGraveyard && k3 < 2) {
          var fixedQuotaCount = 0;
          for (var fqn = 0; fqn < 7; fqn++) fixedQuotaCount += this.fixedAssignments[p3][k3][fqn];
          need[key] = fixedQuotaCount;
        } else if (this.cfg.enhancedGraveyard && k3 === 2) {
          need[key] = this.graveyardFairShareDays(p3);
        } else {
          var n = (fd !== null) ? fd : Math.min(this.personQuota(p3, k3), a);
          need[key] = n;
          addC(co, n, n);
        }
      }
    }
  }

  for (var k4 = 0; k4 < 3; k4++) {
    if (cfg.enhancedGraveyard && k4 < 2) continue;
    if (!cfg.planGraveyard && !cfg.enhancedGraveyard && k4 === 2) continue;
    for (var d4 = 0; d4 < 7; d4++) {
      var tot = [], hs = [];
      for (var q = 0; q < P; q++) {
        if (!this.counted(q, k4)) continue;
        tot.push([this.xName(q, k4, d4), 1]);
        if (this.people[q].hash) hs.push([this.xName(q, k4, d4), 1]);
      }
      if (exact || (cfg.enhancedGraveyard && k4 === 2)) {
        addC(tot, this.reqTotal[k4], this.reqTotal[k4]);
        addC(hs, this.reqHash[k4], this.maxHash[k4] === null ? this.reqHash[k4] : this.maxHash[k4]);
      }
      else {
        addC(tot, this.reqTotal[k4], Infinity);
        addC(hs, this.reqHash[k4], this.maxHash[k4] === null ? Infinity : this.maxHash[k4]);
      }
      addC(tot.concat([['tmax_' + k4, -1]]), -Infinity, 0);
      addC(tot.concat([['tmin_' + k4, -1]]), 0, Infinity);
      addC(hs.concat([['hmax_' + k4, -1]]), -Infinity, 0);
      addC(hs.concat([['hmin_' + k4, -1]]), 0, Infinity);
      this.secs.forEach(function (s, i) {
        var ss = [];
        for (var r = 0; r < P; r++) if (self.counted(r, k4) && self.people[r].sec === s) ss.push([self.xName(r, k4, d4), 1]);
        addC(ss.concat([['smax_' + k4 + '_' + i, -1]]), -Infinity, 0);
        addC(ss.concat([['smin_' + k4 + '_' + i, -1]]), 0, Infinity);
      });
    }
    addC([['tmax_' + k4, 1], ['tmin_' + k4, -1], ['tex_' + k4, -1]], -Infinity, 1);
    addC([['hmax_' + k4, 1], ['hmin_' + k4, -1], ['hex_' + k4, -1]], -Infinity, 1);
    this.secs.forEach(function (s, i) {
      addC([['smax_' + k4 + '_' + i, 1], ['smin_' + k4 + '_' + i, -1], ['sex_' + k4 + '_' + i, -1]], -Infinity, 1);
    });
  }

  if (cfg.wkMax !== null) {
    for (var p4 = 0; p4 < P; p4++) {
      for (var k5 = 0; k5 < 3; k5++) {
        if (!this.counted(p4, k5) || !this.wkdays[k5].length || cfg.enhancedGraveyard) continue;
        if (!cfg.planGraveyard && k5 === 2) continue;
        var wco = this.wkdays[k5].map(function (dd) { return [self.xName(p4, k5, dd), 1]; });
        if (cfg.wkType === 'HARD') addC(wco, -Infinity, cfg.wkMax);
        else addC(wco.concat([['wex_' + p4 + '_' + k5, -1]]), -Infinity, cfg.wkMax);
      }
    }
  }

  if (balance || cfg.enhancedGraveyard) {
    for (var p5 = 0; p5 < P; p5++) {
      var history = this.historicShare[this.people[p5].key];
      var t = 0;
      if (cfg.enhancedGraveyard && history) {
        var gDays = this.graveyardFairShareDays(p5);
        var targetDays = history.amReq + history.pmReq + gDays + history.adminDays;
        if (this.people[p5].perm) t = (history.amReq + history.pmReq + gDays) * this.avgAm[0] + history.adminDays * cfg.adminHrs;
        else t = history.amReq * this.avg[0] + history.pmReq * this.avg[1] + gDays * this.avg[2] + history.adminDays * cfg.adminHrs;
        t += this.bonus(p5) * targetDays;
      } else {
        for (var k6 = 0; k6 < 3; k6++) {
          var nn = need[p5 + '_' + k6] || 0;
          t += nn * (this.counted(p5, k6) ? this.avg[k6] : this.avgAm[k6]) + nn * this.bonus(p5);
        }
      }
      if (t <= 0) continue;
      var rco = [];
      for (var k7 = 0; k7 < 3; k7++)
        for (var d7 = 0; d7 < 7; d7++) rco.push([this.xName(p5, k7, d7), (this.shiftValue(p5, k7, d7) + this.bonus(p5)) / t]);
      var fixedRatio = (cfg.enhancedGraveyard && history) ? history.adminDays * (cfg.adminHrs + this.bonus(p5)) / t : 0;
      if (!this.people[p5].perm) {
        for (var d8 = 0; d8 < 6; d8++) {
          var yv = 'y_' + p5 + '_' + d8;
          rco.push([yv, cfg.pB2b / t]);
          var xa = this.xName(p5, 2, d8), xb = this.xName(p5, 2, d8 + 1);
          addC([[yv, 1], [xa, -1]], -Infinity, 0);
          addC([[yv, 1], [xb, -1]], -Infinity, 0);
          addC([[yv, 1], [xa, -1], [xb, -1]], -1, Infinity);
        }
      }
      if (balance) {
        addC(rco.concat([['pmax', -1]]), -Infinity, -fixedRatio);
        addC(rco.concat([['pmin', -1]]), -fixedRatio, Infinity);
      }
      if (cfg.enhancedGraveyard && history) {
        var negativeRatio = rco.map(function (term) { return [term[0], -term[1]]; });
        addC(rco.concat([['histAbove_' + p5, -1]]), -Infinity, history.targetRatio - fixedRatio);
        addC(negativeRatio.concat([['histBelow_' + p5, -1]]), -Infinity, fixedRatio - history.targetRatio);
      }
    }
  }
  if (balance) {
    addC([['pmax', 1], ['pmin', -1]], 0, Infinity);
  }

  // When a previous roster exists, require at least one prior assignment to move.
  // This guarantees consecutive generated rosters cannot be identical without adding
  // a costly secondary objective for every prior assignment.
  if (this.previous) {
    var previousAssignments = [];
    for (var p6 = 0; p6 < P; p6++) for (var k8 = 0; k8 < 3; k8++) for (var d9 = 0; d9 < 7; d9++) {
      if (this.previous[p6][k8][d9]) {
        previousAssignments.push([this.xName(p6, k8, d9), 1]);
      }
    }
    if (previousAssignments.length) addC(previousAssignments, -Infinity, previousAssignments.length - 1);
  }

  var engine = service.createEngine();
  engine.addVariables(names, lbs, ubs, types, objs);
  engine.addConstraints(cLo, cHi, cNames, cCoefs);
  engine.setMinimization();
  var solveStarted = new Date().getTime();
  var sol = engine.solve(Math.min(cfg.tlimit, MAX_SOLVE_SECONDS));
  var solveSeconds = Math.round((new Date().getTime() - solveStarted) / 1000);
  if (!sol.isValid()) {
    throw new PlanError('The optimiser found no roster that meets every hard rule.\n' +
      "Likely causes: too many people unavailable on the same days, the '#' names are short on some day, or the minimums " +
      'are too high for the days per person.');
  }
  var X = [];
  for (var pp = 0; pp < P; pp++) {
    X.push([]);
    for (var kk = 0; kk < 3; kk++) {
      var row = [];
      for (var dd = 0; dd < 7; dd++) row.push(Math.round(sol.getVariableValue(this.xName(pp, kk, dd))));
      X[pp].push(row);
    }
  }
  var status = (sol.getStatus() === service.Status.OPTIMAL) ? 'optimal' : 'best found within the time limit';
  return { X: X, status: status, solveSeconds: solveSeconds };
};

Model.prototype.points = function (X) {
  var cfg = this.cfg, req = this.requiredDays(X), out = [];
  for (var p = 0; p < this.P; p++) {
    var pts = 0, days = 0, hrs = 0, nights = 0, pairs = 0;
    for (var k = 0; k < 3; k++) {
      for (var d = 0; d < 7; d++) {
        if (X[p][k][d]) {
          pts += this.shiftValue(p, k, d); days++;
          var night = (k === 2 && this.counted(p, k));
          hrs += night ? 10 : 7;
          if (night) nights++;
        }
      }
    }
    if (this.counted(p, 2)) for (var d2 = 0; d2 < 6; d2++) if (X[p][2][d2] && X[p][2][d2 + 1]) pairs++;
    var bonus = this.bonus(p) * days;
    pts += cfg.pB2b * pairs + bonus;
    var tgt = this.bonus(p) * (req[p][0] + req[p][1] + req[p][2]);
    for (var k2 = 0; k2 < 3; k2++) tgt += req[p][k2] * (this.counted(p, k2) ? this.avg[k2] : this.avgAm[k2]);
    out.push({ pts: pts, days: days, hrs: hrs, nights: nights, pairs: pairs, bonus: bonus, tgt: tgt, pct: tgt ? pts / tgt : 0, req: req[p] });
  }
  return out;
};

Model.prototype.verify = function (X) {
  var lines = [], warnings = [], P = this.P, cfg = this.cfg, self = this;
  var viol = 0;
  for (var p = 0; p < P; p++) for (var k = 0; k < 3; k++) for (var d = 0; d < 7; d++) {
    if (cfg.enhancedGraveyard && k < 2) continue;
    if (X[p][k][d] && !this.avail[p][k][d]) viol++;
  }
  lines.push('Shifts on unavailable dates: ' + viol);
  if (viol) warnings.push('A shift was placed on an unavailable date.');
  function sumOver(k, d, filter) {
    var s = 0;
    for (var q = 0; q < P; q++) if (self.counted(q, k) && filter(q)) s += X[q][k][d];
    return s;
  }
  var range = function (a) { return Math.min.apply(null, a) + '-' + Math.max.apply(null, a); };
  for (var k2 = 0; k2 < 3; k2++) {
    if (cfg.enhancedGraveyard && k2 < 2) continue;
    if (!cfg.planGraveyard && !cfg.enhancedGraveyard && k2 === 2) { lines.push('Graveyard: not planned; pending completion.'); continue; }
    var cnt = [], hs = [];
    for (var d2 = 0; d2 < 7; d2++) {
      cnt.push(sumOver(k2, d2, function () { return true; }));
      hs.push(sumOver(k2, d2, function (q) { return self.people[q].hash; }));
    }
    lines.push(BLOCKS[k2] + ': on duty ' + range(cnt) + " per day, '#' names " + range(hs) + ' per day');
    if (Math.min.apply(null, cnt) < this.reqTotal[k2]) warnings.push(BLOCKS[k2] + ': a day is below the minimum headcount.');
    if (this.cfg.enhancedGraveyard && k2 === 2 && Math.max.apply(null, cnt) > this.reqTotal[k2]) warnings.push(BLOCKS[k2] + ': a day exceeds the target of ' + this.reqTotal[k2] + ' people.');
    if (Math.min.apply(null, hs) < this.reqHash[k2]) warnings.push(BLOCKS[k2] + ": a day is below the minimum '#' count.");
    if (this.cfg.enhancedGraveyard && k2 === 2 && Math.max.apply(null, hs) > this.maxHash[k2]) warnings.push(BLOCKS[k2] + ': a day exceeds the maximum of ' + this.maxHash[k2] + " '#' names.");
    if (Math.max.apply(null, cnt) - Math.min.apply(null, cnt) > 1) warnings.push(BLOCKS[k2] + ': headcount differs by more than 1 between days (goal not met).');
    if (Math.max.apply(null, hs) - Math.min.apply(null, hs) > 1) warnings.push(BLOCKS[k2] + ": '#' names differ by more than 1 between days (goal not met).");
    var worst = 0;
    this.secs.forEach(function (s) {
      var sc = [];
      for (var dd = 0; dd < 7; dd++) sc.push(sumOver(k2, dd, function (q) { return self.people[q].sec === s; }));
      worst = Math.max(worst, Math.max.apply(null, sc) - Math.min.apply(null, sc));
    });
    lines.push(BLOCKS[k2] + ': largest section difference across days = ' + worst);
    if (worst > 1) warnings.push(BLOCKS[k2] + ': a section differs by more than 1 between days (goal not met).');
    var lim = cfg.wkMax;
    if (lim !== null && this.wkdays[k2].length && !(cfg.enhancedGraveyard && k2 === 2)) {
      var over = [], wkShifts = 0, peopleOk = 0;
      for (var q2 = 0; q2 < P; q2++) {
        if (!this.counted(q2, k2)) continue;
        var n = 0;
        this.wkdays[k2].forEach(function (dd) { n += X[q2][k2][dd]; });
        wkShifts += n;
        if (n > lim) over.push(n - lim);
        if (this.wkdays[k2].some(function (dd) { return self.avail[q2][k2][dd]; })) peopleOk++;
      }
      var unavoidable = Math.max(0, wkShifts - lim * peopleOk);
      var extra = over.reduce(function (s, x) { return s + x; }, 0);
      lines.push(BLOCKS[k2] + ': ' + over.length + ' people work more than ' + lim + ' weekend shift(s) (' + extra + ' extra shifts; at least ' +
        unavoidable + ' are unavoidable for ' + wkShifts + ' weekend shifts shared by ' + peopleOk + ' people)');
      if (extra > unavoidable) warnings.push(BLOCKS[k2] + ': ' + (extra - unavoidable) + ' more extra weekend shift(s) than the minimum possible for this weekend headcount (solver time limit; raise it for a better result).');
    }
  }
  var tab = this.points(X).filter(function (t) { return t.tgt > 0; });
  if (tab.length) {
    var pc = tab.map(function (t) { return t.pct * 100; });
    var lo = Math.min.apply(null, pc), hi = Math.max.apply(null, pc);
    lines.push('Points: everyone is between ' + lo.toFixed(1) + '% and ' + hi.toFixed(1) + '% of their fair share (spread ' + (hi - lo).toFixed(1) + ' points)');
    if ((cfg.mode === 'EXACT' && !cfg.enhancedGraveyard) || !cfg.balance) lines.push('Points were not balanced by the solver in this configuration; they are reported only.');
  }
  return { lines: lines, warnings: warnings };
};

// ------------------------------------------------------------------ output
function colLetter(c) {
  var s = '';
  while (c > 0) { var r = (c - 1) % 26; s = String.fromCharCode(65 + r) + s; c = Math.floor((c - 1) / 26); }
  return s;
}
function grid(rows, cols, fill) {
  var g = [];
  for (var r = 0; r < rows; r++) { var row = []; for (var c = 0; c < cols; c++) row.push(fill); g.push(row); }
  return g;
}
function dateFormula(dn) { var t = ymdFromDn(dn); return '=DATE(' + t[0] + ',' + t[1] + ',' + t[2] + ')'; }

var C_HEAD = '#1f4e79', C_SEC = '#dce6f1', C_YEL = '#ffff00', C_GREY = '#d9d9d9', C_ADMIN = '#f2f2f2', C_WHITE = '#ffffff';
var C_RED = '#f4cccc';

function freshSheet(ss, name, index) {
  var old = ss.getSheetByName(name);
  if (old) ss.deleteSheet(old);
  return ss.insertSheet(name, index);
}

function newOutputSheet(ss, name, index) {
  if (ss.getSheetByName(name)) throw new PlanError("An output sheet named '" + name + "' already exists.");
  return ss.insertSheet(name, index);
}

function writeRoster(ss, m, X, rosterName, pointsName) {
  var cfg = m.cfg, people = m.people, secs = m.secs;
  var lastCol = COUNT_COL + 7;
  // ----- row plan
  var rowPlan = [];                       // {type:'sec'|'person', sec, p}
  var secRows = {};
  secs.forEach(function (s) {
    rowPlan.push({ type: 'sec', sec: s });
    var first = rowPlan.length + FIRST_ROW;      // 1-based sheet row of the first person
    people.forEach(function (pr, p) { if (pr.sec === s) rowPlan.push({ type: 'person', sec: s, p: p }); });
    secRows[s] = [first, rowPlan.length + FIRST_ROW - 1];
  });
  var lastPerson = FIRST_ROW + rowPlan.length - 1;
  var checkLabels = ['Total on duty', '# names on duty'].concat(secs.map(function (s) { return s + ' on duty'; })).concat(['Meets requirement?']);
  var check0 = lastPerson + 2;
  var ptsR = check0 + checkLabels.length;
  var wkR = ptsR + 2;
  var reqR = wkR + 2;
  var noteR = reqR + 9;
  var totalRows = noteR;

  var V = grid(totalRows, lastCol, ''), BG = grid(totalRows, lastCol, C_WHITE), FW = grid(totalRows, lastCol, 'normal'), FC = grid(totalRows, lastCol, '#000000');
  function put(r, c, v, o) {            // r, c are 1-based
    V[r - 1][c - 1] = v;
    if (o) {
      if (o.bg) BG[r - 1][c - 1] = o.bg;
      if (o.bold) FW[r - 1][c - 1] = 'bold';
      if (o.color) FC[r - 1][c - 1] = o.color;
    }
  }
  put(1, 1, 'Roster', { bg: C_HEAD, bold: true, color: '#ffffff' });
  put(2, 1, 'Name', { bg: C_HEAD, bold: true, color: '#ffffff' });
  put(2, 2, 'Section', { bg: C_HEAD, bold: true, color: '#ffffff' });
  put(1, 2, '', { bg: C_HEAD });
  for (var k = 0; k < 4; k++) {
    for (var d = 0; d < BLOCK_LEN[k]; d++) {
      var col = START_COL[k] + d;
      put(1, col, d === 0 ? (BLOCK_TITLES[k] + (k === 2 && !cfg.planGraveyard ? ' — NOT PLANNED' : '') + ' (' + fmtDateLong(m.dates[k][0]).substring(0, 6) + ' to ' + fmtDateLong(m.dates[k][BLOCK_LEN[k] - 1]) + ')') : '', { bg: C_HEAD, bold: true, color: '#ffffff' });
      put(2, col, dateFormula(m.dates[k][d]), { bg: C_HEAD, bold: true, color: '#ffffff' });
    }
  }
  var heads = ['AM days', 'PM days', 'Graveyard days', 'Admin days', 'Total days', 'AM wknd shifts', 'PM wknd shifts', 'G wknd shifts'];
  heads.forEach(function (t, j) { put(2, COUNT_COL + j, t, { bg: C_HEAD, bold: true, color: '#ffffff' }); });

  var personRow = {};
  rowPlan.forEach(function (rp, i) {
    var r = FIRST_ROW + i;
    if (rp.type === 'sec') {
      for (var c = 1; c <= lastCol; c++) put(r, c, c === 1 ? rp.sec : '', { bg: C_SEC, bold: c === 1 });
      return;
    }
    var p = rp.p, pr = people[p];
    personRow[p] = r;
    put(r, 1, pr.name); put(r, 2, rp.sec);
    for (var kk = 0; kk < 4; kk++) {
      for (var dd = 0; dd < BLOCK_LEN[kk]; dd++) {
        var cc = START_COL[kk] + dd;
        if (kk < 3 && X[p][kk][dd]) put(r, cc, m.counted(p, kk) ? 'X' : 'AM', { bold: true });
        else if (((p + '_' + kk + '_' + dd) in m.absent) && cfg.showAbs) put(r, cc, m.absent[p + '_' + kk + '_' + dd]);
        else if (kk === 3) put(r, cc, '', { bg: C_ADMIN });
      }
    }
    for (var b = 0; b < 4; b++) {
      var rg = colLetter(START_COL[b]) + r + ':' + colLetter(START_COL[b] + BLOCK_LEN[b] - 1) + r;
      put(r, COUNT_COL + b, '=COUNTIF(' + rg + ',"X")+COUNTIF(' + rg + ',"AM")');
    }
    put(r, COUNT_COL + 4, '=SUM(' + colLetter(COUNT_COL) + r + ':' + colLetter(COUNT_COL + 3) + r + ')');
    for (var b2 = 0; b2 < 3; b2++) {
      var c0 = START_COL[b2], c1 = START_COL[b2] + 6;
      var flags = colLetter(c0) + '$' + wkR + ':' + colLetter(c1) + '$' + wkR;
      put(r, COUNT_COL + 5 + b2, '=SUMPRODUCT((' + colLetter(c0) + r + ':' + colLetter(c1) + r + '="X")*' + flags + ')');
    }
  });

  checkLabels.forEach(function (t, i) { put(check0 + i, 1, t, { bold: true }); });
  put(ptsR, 1, 'Points per shift (X)'); put(ptsR + 1, 1, 'Points if AM shift'); put(wkR, 1, 'Weekend day (1 = counts toward weekend limit)');
  put(reqR, 1, 'Requirement (from Settings)', { bold: true });
  var reqs = [['AM/PM minimum on duty', cfg.ampmMin], ['AM/PM minimum # names', cfg.ampmHash],
    ['Graveyard minimum on duty', cfg.planGraveyard ? cfg.gMin : 'Pending'], ['Graveyard minimum # names', cfg.planGraveyard ? cfg.gHash : 'Pending'],
    ['Graveyard maximum # names', cfg.enhancedGraveyard ? cfg.gHashMax : 'No limit'],
    ['Admin minimum on duty', cfg.adminMin], ['Admin minimum # names', cfg.adminHash]];
  reqs.forEach(function (rq, i) { put(reqR + 1 + i, 1, rq[0]); put(reqR + 1 + i, 2, rq[1]); });
  var R_AMPM = '$B$' + (reqR + 1), R_AMPMH = '$B$' + (reqR + 2), R_G = '$B$' + (reqR + 3),
    R_GH = '$B$' + (reqR + 4), R_GHMAX = '$B$' + (reqR + 5), R_ADMIN = '$B$' + (reqR + 6), R_ADMINH = '$B$' + (reqR + 7);
  var kinds = ['AM', 'PM', 'G'];
  for (var k2 = 0; k2 < 4; k2++) {
    for (var d2 = 0; d2 < BLOCK_LEN[k2]; d2++) {
      var col2 = START_COL[k2] + d2, L = colLetter(col2);
      put(check0, col2, '=COUNTIF(' + L + FIRST_ROW + ':' + L + lastPerson + ',"X")');
      put(check0 + 1, col2, '=SUMPRODUCT((LEFT($A$' + FIRST_ROW + ':$A$' + lastPerson + ',1)="#")*(' + L + FIRST_ROW + ':' + L + lastPerson + '="X"))');
      secs.forEach(function (s, i) {
        put(check0 + 2 + i, col2, '=COUNTIF(' + L + secRows[s][0] + ':' + L + secRows[s][1] + ',"X")');
      });
      var tot, hsh, op;
      var checkFormula;
      if (k2 < 2) { tot = R_AMPM; hsh = R_AMPMH; op = cfg.mode === 'EXACT' ? '=' : '>='; }
      else if (k2 === 2) {
        tot = R_G; hsh = R_GH; op = (cfg.mode === 'EXACT' || cfg.enhancedGraveyard) ? '=' : '>=';
      } else { tot = R_ADMIN; hsh = R_ADMINH; op = '>='; }
      if (k2 === 2 && !cfg.planGraveyard) {
        checkFormula = '="PENDING"';
      } else if (k2 === 2 && cfg.enhancedGraveyard) {
        checkFormula = '=IF(AND(' + L + check0 + '=' + tot + ',' + L + (check0 + 1) + '>=' + hsh + ',' + L + (check0 + 1) + '<=' + R_GHMAX + '),"OK","CHECK")';
      } else {
        checkFormula = '=IF(AND(' + L + check0 + op + tot + ',' + L + (check0 + 1) + op + hsh + '),"OK","CHECK")';
      }
      put(check0 + checkLabels.length - 1, col2, checkFormula);
      if (k2 < 3) {
        put(ptsR, col2, Math.round(shiftBurden(kinds[k2], m.dates[k2][d2], cfg) * 10000) / 10000);
        put(ptsR + 1, col2, Math.round(shiftBurden('AM', m.dates[k2][d2], cfg) * 10000) / 10000);
        put(wkR, col2, cfg.wkSet[weekdayOf(m.dates[k2][d2])] ? 1 : 0);
      } else {
        put(ptsR, col2, cfg.adminHrs);
      }
    }
  }
  put(noteR, 1, 'X = on duty (bold on yellow). AM = perm-AM person working an AM shift in the PM/Graveyard week (not counted in those totals). Grey = unavailable (reason shown). Admin cells are for manual entry: type X.');

  var sh = newOutputSheet(ss, rosterName, 0);
  var rng = sh.getRange(1, 1, totalRows, lastCol);
  rng.setValues(V);
  rng.setBackgrounds(BG);
  rng.setFontWeights(FW);
  rng.setFontColors(FC);
  sh.getRange(1, 3, totalRows, lastCol - 2).setHorizontalAlignment('center');
  sh.getRange(noteR, 1, 1, 1).setHorizontalAlignment('left');
  for (var kb = 0; kb < 4; kb++) {
    sh.getRange(1, START_COL[kb], 1, BLOCK_LEN[kb]).merge();
    sh.getRange(2, START_COL[kb], 1, BLOCK_LEN[kb]).setNumberFormat('ddd d mmm');
  }
  sh.getRange(ptsR, 3, 3, lastCol - 2).setFontStyle('italic');
  sh.setFrozenRows(2);
  sh.setFrozenColumns(2);
  sh.setColumnWidth(1, 230);
  sh.setColumnWidth(2, 60);
  for (var c = 3; c <= lastCol; c++) sh.setColumnWidth(c, 82);
  [10, 18, 26, 32].forEach(function (c) { sh.setColumnWidth(c, 18); });
  for (var cn = COUNT_COL; cn <= lastCol; cn++) sh.setColumnWidth(cn, 100);
  addAssignmentColorRules(sh, lastPerson);
  sh.getRange(check0, 1, totalRows - check0 + 1, lastCol)
    .setBackground(C_WHITE).setFontColor('#000000').setFontWeight('bold').setFontStyle('normal');
  writePoints(ss, m, X, personRow, ptsR, rosterName, pointsName);
}

function writePoints(ss, m, X, personRow, ptsR, rosterName, pointsName) {
  var cfg = m.cfg, people = m.people, n = people.length;
  var tab = m.points(X);
  var HR = 5, r0 = HR + 1;
  var totalRows = r0 + n + 2, lastCol = 17;
  var V = grid(totalRows, lastCol, ''), BG = grid(totalRows, lastCol, C_WHITE), FW = grid(totalRows, lastCol, 'normal'), FC = grid(totalRows, lastCol, '#000000');
  function put(r, c, v, o) {
    V[r - 1][c - 1] = v;
    if (o) { if (o.bg) BG[r - 1][c - 1] = o.bg; if (o.bold) FW[r - 1][c - 1] = 'bold'; if (o.color) FC[r - 1][c - 1] = o.color; }
  }
  put(1, 1, 'Points parameters (blue cells can be edited and the table updates)', { bold: true });
  put(1, 17, 'Graveyard planning: ' + (cfg.planGraveyard ? 'COMPLETE' : 'PENDING'));
  put(2, 1, 'Day bonus per day worked'); put(2, 2, cfg.pDay, { color: '#0000ff' });
  put(3, 1, 'Back-to-back night penalty'); put(3, 2, cfg.pB2b, { color: '#0000ff' });
  put(2, 4, 'Average points per AM shift'); put(2, 5, m.avg[0]);
  put(3, 4, 'Average points per PM shift'); put(3, 5, m.avg[1]);
  put(2, 7, 'Average points per Graveyard shift'); put(2, 8, m.avg[2]);
  put(3, 7, 'Admin shift points'); put(3, 8, cfg.adminHrs);
  var heads = ['Rank', 'Name', 'Total points', 'Days worked', 'Hours worked', 'Night shifts', 'Back-to-back nights', 'Days affected',
    'Fair-share target', '% of fair share', 'Day bonus eligible', 'Day bonus points', 'Required AM', 'Required PM', 'Required Graveyard', 'Admin days'];
  heads.forEach(function (t, j) { put(HR, j + 1, t, { bg: C_HEAD, bold: true, color: '#ffffff' }); });
  for (var p = 0; p < n; p++) {
    var r = r0 + p, rr = personRow[p];
    var rosterRef = quotedSheet(rosterName) + '!';
    var rowAll = rosterRef + '$C' + rr + ':$AE' + rr;
    var wX = rosterRef + '$C$' + ptsR + ':$AE$' + ptsR, wAm = rosterRef + '$C$' + (ptsR + 1) + ':$AE$' + (ptsR + 1);
    put(r, 2, people[p].name);
    put(r, 11, people[p].nobonus ? 'No' : 'Yes', { color: '#0000ff' });
    for (var k = 0; k < 3; k++) put(r, 13 + k, tab[p].req[k]);
    put(r, 16, '=COUNTIF(' + rosterRef + '$AA' + rr + ':$AE' + rr + ',"X")');
    put(r, 3, '=SUMPRODUCT((' + rowAll + '="X")*' + wX + ')+SUMPRODUCT((' + rowAll + '="AM")*' + wAm + ')+$B$3*G' + r + '+L' + r);
    put(r, 4, '=COUNTIF(' + rowAll + ',"X")+COUNTIF(' + rowAll + ',"AM")');
    put(r, 5, '=7*(COUNTIF(' + rosterRef + '$C' + rr + ':$I' + rr + ',"X")+COUNTIF(' + rosterRef + '$K' + rr + ':$Q' + rr + ',"X")+COUNTIF(' + rowAll + ',"AM"))+10*F' + r + '+P' + r + '*' + cfg.adminHrs);
    put(r, 6, '=COUNTIF(' + rosterRef + '$S' + rr + ':$Y' + rr + ',"X")');
    put(r, 7, '=SUMPRODUCT((' + rosterRef + '$S' + rr + ':$X' + rr + '="X")*(' + rosterRef + '$T' + rr + ':$Y' + rr + '="X"))');
    put(r, 8, '=D' + r + '+2*F' + r + '-2*G' + r);
    var base = people[p].perm ? '(M' + r + '+N' + r + '+O' + r + ')*$E$2' : 'M' + r + '*$E$2+N' + r + '*$E$3+O' + r + '*$H$2';
    put(r, 9, '=' + base + '+P' + r + '*$H$3+IF(K' + r + '="Yes",$B$2,0)*(M' + r + '+N' + r + '+O' + r + '+P' + r + ')');
    put(r, 10, '=IF(I' + r + '=0,0,C' + r + '/I' + r + ')');
    put(r, 12, '=IF(K' + r + '="Yes",$B$2*D' + r + ',0)');
    put(r, 1, '=RANK(J' + r + ',$J$' + r0 + ':$J$' + (r0 + n - 1) + ',0)');
  }
  put(r0 + n + 1, 1, 'Total points = shift values (hour weights, weekend add-on, recovery) + back-to-back penalty + day bonus. Admin days entered by hand add points and also raise the fair-share target, so they do not skew the % column.');
  var sh = newOutputSheet(ss, pointsName, 1);
  var rng = sh.getRange(1, 1, totalRows, lastCol);
  rng.setValues(V); rng.setBackgrounds(BG); rng.setFontWeights(FW); rng.setFontColors(FC);
  sh.getRange(r0, 3, n, 1).setNumberFormat('0.0');
  sh.getRange(r0, 9, n, 1).setNumberFormat('0.0');
  sh.getRange(r0, 10, n, 1).setNumberFormat('0.0%');
  sh.getRange(HR, 1, n + 1, lastCol).setHorizontalAlignment('center');
  sh.getRange(r0, 2, n, 1).setHorizontalAlignment('left');
  sh.setFrozenRows(HR);
  sh.setColumnWidth(1, 90);
  sh.setColumnWidth(2, 220);
  for (var c = 3; c <= lastCol; c++) sh.setColumnWidth(c, 105);
}

function writeLog(ss, text) {
  var sh = freshSheet(ss, 'Log', ss.getSheets().length);
  var lines = text.split('\n').map(function (l) { return [l]; });
  sh.getRange(1, 1, lines.length, 1).setValues(lines);
  sh.setColumnWidth(1, 900);
}
