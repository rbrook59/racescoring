
/** Stamford Yacht Club Scoring Template **
 * Version 3.1 June 2026
 *
 * Functions:
 *   scoreSailors    - main entry point; reads data, calls scoring functions, writes results, sorts
 *   cleanScores     - converts letter codes (DNF/DSQ/etc) to numbers, strips un-sailed races
 *   csScore         - Cox-Sprague high-point percentage score (thin wrapper over csScoreDetail)
 *   csScoreDetail   - Cox-Sprague score with full per-race breakdown and discard flags
 *   hpScore         - High Point of Perfection percentage score
 *   avgScore        - low-point average score
 *   lowScore        - low-point total score
 *   buildCSTable    - builds the hard-coded Cox-Sprague lookup table
 *
 *   Helpers:
 *   loadScoreRange  - reads a named range, separating header from data rows
 *   writeScoreRange - writes header + data rows back to a named range
 *   applyDiscards   - sorts a score array ascending and removes the worst `discards` entries
 *   writeCSDetail   - builds the 'CS Detail' sheet with per-race num/den/ratio breakdown
 *
 * ___________________________________________________________________________
 * Developed by Rick Brook (rick.brook@gmail.com) for Stamford Yacht Club.
 * Version 3.1 - refactored June 2026
 *
 * The Cox-Sprague algorithm was originally developed in VBA by Witold Gesing <wgesing@gmail.com>
 *
 * This software may be copied and re-distributed freely.
 * If you make changes, please share them with the above authors.
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Race Scoring')
    .addItem('Calculate Scores', 'scoreSailors')
    .addSeparator()
    .addToUi();
}

// ─── HELPERS ─────────────────────────────────────────────────────────────────

function loadScoreRange(sheet, rangeName) {
  const data = sheet.getRange(rangeName).getValues();
  return { header: data[0], rows: data.slice(1) };
}

function writeScoreRange(sheet, rangeName, rows, header) {
  sheet.getRange(rangeName).setValues([header, ...rows]);
}

// Sorts cleanArray ascending by column 0 and removes the `discards` worst (highest) scores.
// Mutates in place — always call on a deep copy, not the original referenceArray.
function applyDiscards(cleanArray, discards) {
  cleanArray.sort((a, b) => a[0] - b[0]);
  if (discards > 0) cleanArray.splice(-discards);
}

// ─── MAIN ────────────────────────────────────────────────────────────────────

function scoreSailors() {
  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();

  const controlSheet = spreadSheet.getSheetByName('ControlPanel');
  if (!controlSheet) throw new Error("Sheet 'ControlPanel' not found.");
  const resultsSheet = spreadSheet.getSheetByName('RaceResults');
  if (!resultsSheet) throw new Error("Sheet 'RaceResults' not found.");

  // Column positions are stored as 1-based physical column numbers; subtract 1 for 0-based JS arrays.
  const arrColumns     = controlSheet.getRange('rangeColumns').getValues();
  const colSkipperName = Number(arrColumns[0]) - 1;
  const colQualified   = Number(arrColumns[1]) - 1;
  const colDiscards    = Number(arrColumns[2]) - 1;
  const colCoxSprague  = Number(arrColumns[3]) - 1;
  const colHighPoint   = Number(arrColumns[4]) - 1;
  const colAvgScore    = Number(arrColumns[5]) - 1;
  const colLowPoint    = Number(arrColumns[6]) - 1;
  const colResults     = Number(arrColumns[7]) - 1;
  const colToSort      = Number(arrColumns[8]) - 1;

  const arrControl        = controlSheet.getRange('rangeControl').getValues();
  const ctrlOriginalTable = Number(arrControl[0]) === 1;

  const resultsValues  = resultsSheet.getRange('rangeResults').getValues();
  const resultsHeader  = resultsValues[0];
  const arrResults     = resultsValues.slice(1); // strip header row
  const arrStarters  = resultsSheet.getRange('rangeStarters').getValues();
  const arrFinishers = resultsSheet.getRange('rangeFinishers').getValues();

  let csSR, hpSR, avgSR, lowSR;
  if (colCoxSprague > 0) csSR  = loadScoreRange(resultsSheet, 'rangeCoxSprague');
  if (colHighPoint  > 0) hpSR  = loadScoreRange(resultsSheet, 'rangeHighPoint');
  if (colAvgScore   > 0) avgSR = loadScoreRange(resultsSheet, 'rangeAvgScore');
  if (colLowPoint   > 0) lowSR = loadScoreRange(resultsSheet, 'rangeLowPoint');

  const arrCSTable = colCoxSprague > 0 ? buildCSTable(ctrlOriginalTable) : null;

  const arrCleaned = arrResults.map(row =>
    cleanScores(row.slice(colResults), arrStarters[0], arrFinishers[0])
  );

  const csDetails = colCoxSprague > 0 ? [] : null;

  for (let j = 0; j < arrResults.length; j++) {
    const hasName  = String(arrResults[j][colSkipperName]).trim().length > 0;
    const discards = arrResults[j][colDiscards];
    const cleaned  = arrCleaned[j];

    if (colCoxSprague > 0) {
      if (hasName) {
        const detail = csScoreDetail(cleaned, discards, arrCSTable);
        csSR.rows[j][0] = detail.score;
        csDetails.push({ name: String(arrResults[j][colSkipperName]).trim(), detail });
      } else {
        csSR.rows[j][0] = "";
        csDetails.push(null);
      }
    }
    if (colHighPoint  > 0) hpSR.rows[j][0]  = hasName ? hpScore(cleaned, discards)             : "";
    if (colAvgScore   > 0) avgSR.rows[j][0] = hasName ? avgScore(cleaned, discards)             : "";
    if (colLowPoint   > 0) lowSR.rows[j][0] = hasName ? lowScore(cleaned, discards)             : "";
  }

  if (colCoxSprague > 0) writeScoreRange(resultsSheet, 'rangeCoxSprague', csSR.rows,  csSR.header);
  if (colHighPoint  > 0) writeScoreRange(resultsSheet, 'rangeHighPoint',  hpSR.rows,  hpSR.header);
  if (colAvgScore   > 0) writeScoreRange(resultsSheet, 'rangeAvgScore',   avgSR.rows, avgSR.header);
  if (colLowPoint   > 0) writeScoreRange(resultsSheet, 'rangeLowPoint',   lowSR.rows, lowSR.header);

  if (colCoxSprague > 0) {
    const raceHeaders   = resultsHeader.slice(colResults);
    const activeSailors = csDetails.filter(d => d !== null);
    writeCSDetail(spreadSheet, activeSailors, arrStarters[0].length, arrStarters[0], raceHeaders);
  }

  // High-point systems rank high-to-low; low-point systems rank low-to-high.
  const resultsRange  = spreadSheet.getRange('rangeResults');
  const sortAscending = (colToSort !== colCoxSprague && colToSort !== colHighPoint);
  resultsRange.offset(1, 0, resultsRange.getNumRows() - 1).sort([
    { column: colQualified + 1, ascending: false },
    { column: colToSort + 1,    ascending: sortAscending }
  ]);
  spreadSheet.getRange('A3').activate();
}

// ─── SCORING FUNCTIONS ────────────────────────────────────────────────────────

// Cox-Sprague log formula for fleets > 20.
// DNF boats (place > starters) receive the floor value of 100 — same as last place.
function csLogScore(place, starters) {
  if (place > starters) return 100;
  return 100 + 200 * Math.log10(starters - place + 1) / Math.log10(starters - 1);
}

function csScore(referenceArray, discards, csGrid) {
  return csScoreDetail(referenceArray, discards, csGrid).score;
}

function csScoreDetail(referenceArray, discards, csGrid) {
  /**
   * Returns the Cox-Sprague score plus a per-race breakdown used by writeCSDetail.
   * Races with <= 20 starters use the lookup table; races with > 20 use the log formula.
   * Place is clamped to [1, starters+1] before the table lookup to guard against bad data.
   * cleanScores() tags each row with its original race index at position [7]; that index
   * survives the sort so discard decisions can be mapped back to the original race order.
   * @param {Array}  referenceArray - cleaned race data rows from cleanScores()
   * @param {number} discards       - number of worst races to drop
   * @param {Array}  csGrid         - lookup table from buildCSTable()
   * @return {{ score, numTotal, denTotal, racesMap }}
   *   racesMap is keyed by original race index → { num, den, ratio, discarded }
   *   Absent keys in racesMap indicate unsailed (DNC/DNS) races.
   */
  if (referenceArray.length === 0) return { score: 0, numTotal: 0, denTotal: 0, racesMap: {} };

  const cleanArray = referenceArray.map(row => [...row]);

  for (let i = 0; i < cleanArray.length; i++) {
    const place    = cleanArray[i][0];
    const starters = cleanArray[i][1];
    if (starters > 20) {
      cleanArray[i][4] = csLogScore(1, starters);
      cleanArray[i][5] = csLogScore(place, starters);
    } else {
      const lookupPlace = Math.min(Math.max(place, 1), starters + 1);
      cleanArray[i][4] = csGrid[1][starters] || 0;
      cleanArray[i][5] = csGrid[lookupPlace][starters] || 0;
    }
    cleanArray[i][6] = cleanArray[i][4] > 0 ? cleanArray[i][5] / cleanArray[i][4] : 0;
  }

  cleanArray.sort((a, b) => b[6] - a[6]); // best to worst

  const discardedIndices = new Set();
  if (discards > 0) {
    const dropCount = Math.min(discards, cleanArray.length);
    for (let i = cleanArray.length - dropCount; i < cleanArray.length; i++) {
      discardedIndices.add(cleanArray[i][7]);
    }
  }

  let numTotal = 0, denTotal = 0;
  for (let i = 0; i < cleanArray.length; i++) {
    if (!discardedIndices.has(cleanArray[i][7])) {
      numTotal += cleanArray[i][5];
      denTotal += cleanArray[i][4];
    }
  }

  const racesMap = {};
  for (let i = 0; i < cleanArray.length; i++) {
    const origIdx = cleanArray[i][7];
    racesMap[origIdx] = {
      num:       cleanArray[i][5],
      den:       cleanArray[i][4],
      ratio:     cleanArray[i][6],
      discarded: discardedIndices.has(origIdx)
    };
  }

  return {
    score:    denTotal === 0 ? 0 : numTotal / denTotal,
    numTotal,
    denTotal,
    racesMap
  };
}

function hpScore(referenceArray, discards) {
  /**
   * Returns the High Point of Perfection score.
   * Per-race value: (starters - place + 1) / starters.  Final: sum(numerators) / sum(denominators).
   * @param {Array}  referenceArray - cleaned race data rows from cleanScores()
   * @param {number} discards       - number of worst races to drop
   * @return {number} score <= 1
   */
  if (referenceArray.length === 0) return 0;

  const cleanArray = referenceArray.map(row => [...row]);

  for (let i = 0; i < cleanArray.length; i++) {
    cleanArray[i][3] = (cleanArray[i][1] - cleanArray[i][0] + 1) / cleanArray[i][1];
  }
  cleanArray.sort((a, b) => b[3] - a[3]); // best to worst

  if (discards > 0) {
    if (discards >= cleanArray.length) return 0;
    cleanArray.splice(-discards);
  }

  let numerator = 0, denominator = 0;
  for (let i = 0; i < cleanArray.length; i++) {
    numerator   += cleanArray[i][1] - cleanArray[i][0] + 1;
    denominator += cleanArray[i][1];
  }
  return denominator === 0 ? 0 : numerator / denominator;
}

function avgScore(referenceArray, discards) {
  /**
   * Returns the average low-point score (sum of places / races sailed, after discards).
   * @param {Array}  referenceArray - cleaned race data rows from cleanScores()
   * @param {number} discards       - number of worst races to drop
   * @return {number} average place
   */
  if (referenceArray.length === 0) return 0;
  const cleanArray = referenceArray.map(row => [...row]);
  applyDiscards(cleanArray, discards);
  if (cleanArray.length === 0) return 0;
  let total = 0;
  for (let i = 0; i < cleanArray.length; i++) total += cleanArray[i][0];
  return total / cleanArray.length;
}

function lowScore(referenceArray, discards) {
  /**
   * Returns the total low-point score (sum of places, after discards).
   * @param {Array}  referenceArray - cleaned race data rows from cleanScores()
   * @param {number} discards       - number of worst races to drop
   * @return {number} sum of places
   */
  if (referenceArray.length === 0) return 0;
  const cleanArray = referenceArray.map(row => [...row]);
  applyDiscards(cleanArray, discards);
  if (cleanArray.length === 0) return 0;
  let total = 0;
  for (let i = 0; i < cleanArray.length; i++) total += cleanArray[i][0];
  return total;
}

// ─── CLEANSCORES ─────────────────────────────────────────────────────────────

function cleanScores(resultsArray, startersArray, finishersArray) {
  /**
   * Converts letter penalty codes to numeric scores and removes un-sailed races.
   * Returns a 2D array where each row is [place, starters, finishers, 0, 0, 0, 0, origIndex].
   * The four zeros are working space used by the scoring functions; origIndex is the
   * 0-based position within resultsArray before filtering, used by csScoreDetail to
   * map discards back to their original race column on the CS Detail sheet.
   */
  const STARTERS_PLUS_ONE = new Set(["DNF", "DSQ", "RAF", "OCS", "BFD", "RET", "UFD", "NSC"]);

  for (let i = 0; i < resultsArray.length; i++) {
    const v = typeof resultsArray[i] === 'string' ? resultsArray[i].trim().toUpperCase() : resultsArray[i];
    if (STARTERS_PLUS_ONE.has(v)) {
      resultsArray[i] = startersArray[i] + 1; // penalty: starters + 1
    } else if (v === "TLE") {
      resultsArray[i] = finishersArray[i] + 2; // time limit expired: finishers + 2
    } else if (v === "DNC" || v === "DNS") {
      resultsArray[i] = 0; // not sailed — excluded from series
    } else if (typeof v !== 'number' || !isFinite(v)) {
      resultsArray[i] = 0; // unrecognised value (e.g. accidental space) — treat as not sailed
    }
  }

  return resultsArray
    .map((score, i) => [score, startersArray[i], finishersArray[i], 0, 0, 0, 0, i])
    .filter(row => row[0] !== 0 && row[0] !== null && row[0] !== undefined && row[0] !== "");
}

// ─── BUILDCSTABLE ─────────────────────────────────────────────────────────────

function buildCSTable(original) {
  /**
   * Returns the hard-coded Cox-Sprague lookup table as a 2D array [21 x 21].
   * Row index = place (1-based); column index = number of starters (1-based).
   * Row 0 and column 0 hold header values [0..20] so that place/starters integers
   * can be used directly as indices without offset arithmetic.
   *
   * When original=true, uses the original Cox-Sprague table values.
   * Default (original=false) uses the Y.R.A. of Long Island Sound modified table,
   * which adjusts 2-boat race values: 2nd place scores 7 (not 4) and DNF scores 5.
   *
   * Scores beyond 20 racers are handled by the calling function via a log approximation (TBD).
   */
  const csTableBuild = [];

  csTableBuild.push([0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20]); // column headings

  if (original) {
    csTableBuild.push([1, , 10, 31, 43, 52, 60, 66, 72, 76, 80, 84, 87, 90, 92, 94, 96, 97, 98, 99, 100]);
    csTableBuild.push([2, ,  4, 25, 37, 46, 54, 60, 66, 70, 74, 78, 81, 84, 86, 88, 90, 91, 92, 93,  94]);
    csTableBuild.push([3, ,  , 21, 33, 42, 50, 56, 62, 66, 70, 74, 77, 80, 82, 84, 86, 87, 88, 89,  90]);
  } else {
    csTableBuild.push([1, , 10, 31, 43, 52, 60, 66, 72, 76, 80, 84, 87, 90, 92, 94, 96, 97, 98, 99, 100]);
    csTableBuild.push([2, ,  7, 25, 37, 46, 54, 60, 66, 70, 74, 78, 81, 84, 86, 88, 90, 91, 92, 93,  94]);
    csTableBuild.push([3, ,  5, 21, 33, 42, 50, 56, 62, 66, 70, 74, 77, 80, 82, 84, 86, 87, 88, 89,  90]);
  }

  csTableBuild.push([4, , , 17, 29, 38, 46, 52, 58, 62, 66, 70, 73, 76, 78, 80, 82, 83, 84, 85, 86]);
  csTableBuild.push([5, , , , 26, 35, 43, 49, 55, 59, 63, 67, 70, 73, 75, 77, 79, 80, 81, 82, 83]);
  csTableBuild.push([6, , , , , 32, 40, 46, 52, 56, 60, 64, 67, 70, 72, 74, 76, 77, 78, 79, 80]);
  csTableBuild.push([7, , , , , , 38, 44, 50, 54, 58, 62, 65, 68, 70, 72, 74, 75, 76, 77, 78]);
  csTableBuild.push([8, , , , , , , 42, 48, 52, 56, 60, 63, 66, 68, 70, 72, 73, 74, 75, 76]);
  csTableBuild.push([9, , , , , , , , 46, 50, 54, 58, 61, 64, 66, 68, 70, 71, 72, 73, 74]);
  csTableBuild.push([10, , , , , , , , , 48, 52, 56, 59, 62, 64, 66, 68, 69, 70, 71, 72]);
  csTableBuild.push([11, , , , , , , , , , 50, 54, 57, 60, 62, 64, 66, 67, 68, 69, 70]);
  csTableBuild.push([12, , , , , , , , , , , 52, 55, 58, 60, 62, 64, 65, 66, 67, 68]);
  csTableBuild.push([13, , , , , , , , , , , , 53, 56, 58, 60, 62, 63, 64, 65, 66]);
  csTableBuild.push([14, , , , , , , , , , , , , 55, 57, 59, 61, 62, 63, 64, 65]);
  csTableBuild.push([15, , , , , , , , , , , , , , 56, 58, 60, 61, 62, 63, 64]);
  csTableBuild.push([16, , , , , , , , , , , , , , , 57, 59, 60, 61, 62, 63]);
  csTableBuild.push([17, , , , , , , , , , , , , , , , 58, 59, 60, 61, 62]);
  csTableBuild.push([18, , , , , , , , , , , , , , , , , 58, 59, 60, 61]);
  csTableBuild.push([19, , , , , , , , , , , , , , , , , , 58, 59, 60]);
  csTableBuild.push([20, , , , , , , , , , , , , , , , , , , 58, 59]);
  csTableBuild.push([21, , , , , , , , , , , , , , , , , , , , 58]); // DNF score for a 20-boat race

  return csTableBuild;
}

// ─── CS DETAIL SHEET ──────────────────────────────────────────────────────────

function writeCSDetail(spreadSheet, sailors, raceCount, startersRow, raceHeaders) {
  /**
   * Builds (or rebuilds) the 'CS Detail' sheet with a per-race Cox-Sprague breakdown.
   * Layout: races as rows, sailors as column groups of three (Num | Den | Ratio).
   * Discarded races are shown with strikethrough + grey fill on a per-sailor basis
   * (two sailors in the same race may discard different races).
   * Unsailed races (absent from racesMap) appear as blank rows so numbering stays
   * aligned with the RaceResults sheet.
   *
   * @param {Spreadsheet} spreadSheet  - active spreadsheet
   * @param {Array}       sailors      - [{ name, detail }] — nulls already filtered by caller
   * @param {number}      raceCount    - total number of race columns
   * @param {Array}       startersRow  - arrStarters[0]: starters count per race column
   * @param {Array}       raceHeaders  - race column header labels from rangeResults
   */
  if (sailors.length === 0 || raceCount === 0) return;

  let sheet = spreadSheet.getSheetByName('CS Detail');
  if (!sheet) {
    sheet = spreadSheet.insertSheet('CS Detail');
  } else {
    sheet.clear();
  }

  const totalCols    = 2 + sailors.length * 3;
  const dataStartRow = 3;                              // rows 1–2 are headers
  const totalsRow    = dataStartRow + raceCount + 1;  // +1 skips a blank separator row
  const totalRows    = totalsRow;

  // ── Build data grid ───────────────────────────────────────────────────────

  const grid = Array.from({ length: totalRows }, () => Array(totalCols).fill(''));

  // Header row 1: sailor names in the first column of each 3-column group
  for (let s = 0; s < sailors.length; s++) {
    grid[0][2 + s * 3] = sailors[s].name;
  }

  // Header row 2: column labels
  grid[1][0] = 'Race';
  grid[1][1] = 'Starters';
  for (let s = 0; s < sailors.length; s++) {
    grid[1][2 + s * 3]     = 'CS Score';
    grid[1][2 + s * 3 + 1] = 'CS Max';
    grid[1][2 + s * 3 + 2] = 'Ratio';
  }

  // Race rows — unsailed races (absent from racesMap) remain blank
  for (let r = 0; r < raceCount; r++) {
    grid[r + 2][0] = raceHeaders[r] || `Race ${r + 1}`;
    grid[r + 2][1] = startersRow[r];
    for (let s = 0; s < sailors.length; s++) {
      const rd = sailors[s].detail.racesMap[r];
      if (rd) {
        grid[r + 2][2 + s * 3]     = rd.num;
        grid[r + 2][2 + s * 3 + 1] = rd.den;
        grid[r + 2][2 + s * 3 + 2] = rd.ratio;
      }
    }
  }

  // Blank separator row (grid[raceCount + 2]) is already empty

  // Totals row — score appears in the Ratio column of the same row
  grid[totalsRow - 1][0] = 'Totals';
  for (let s = 0; s < sailors.length; s++) {
    grid[totalsRow - 1][2 + s * 3]     = sailors[s].detail.numTotal;
    grid[totalsRow - 1][2 + s * 3 + 1] = sailors[s].detail.denTotal;
    grid[totalsRow - 1][2 + s * 3 + 2] = sailors[s].detail.score;
  }

  // ── Write data in a single call ───────────────────────────────────────────

  sheet.getRange(1, 1, totalRows, totalCols).setValues(grid);

  // ── Merges ────────────────────────────────────────────────────────────────

  for (let s = 0; s < sailors.length; s++) {
    sheet.getRange(1, 3 + s * 3, 1, 3).merge(); // sailor name
  }

  // ── Header formatting ─────────────────────────────────────────────────────

  sheet.getRange(1, 1, 2, totalCols)
    .setFontWeight('bold')
    .setBackground('#c9daf8')
    .setHorizontalAlignment('center');
  sheet.getRange(2, 1, 1, 2).setHorizontalAlignment('left'); // Race / Starters labels

  // ── Number formats ────────────────────────────────────────────────────────
  // Ratio columns (one per sailor) + the score cells — all to 3 decimal places.

  const fmtA1s = [];
  for (let s = 0; s < sailors.length; s++) {
    fmtA1s.push(sheet.getRange(dataStartRow, 5 + s * 3, raceCount, 1).getA1Notation());
    fmtA1s.push(sheet.getRange(totalsRow,    5 + s * 3, 1,         1).getA1Notation());
  }
  sheet.getRangeList(fmtA1s).setNumberFormat('0.000');

  // ── Totals and score row formatting ───────────────────────────────────────

  sheet.getRange(totalsRow, 1, 1, totalCols).setFontWeight('bold');

  // ── Discard formatting ────────────────────────────────────────────────────
  // Collect all discarded cell ranges across all sailors into one RangeList call.

  const discardA1s = [];
  for (let s = 0; s < sailors.length; s++) {
    const numCol = 3 + s * 3;
    for (let r = 0; r < raceCount; r++) {
      const rd = sailors[s].detail.racesMap[r];
      if (rd && rd.discarded) {
        discardA1s.push(sheet.getRange(r + dataStartRow, numCol, 1, 3).getA1Notation());
      }
    }
  }
  if (discardA1s.length > 0) {
    sheet.getRangeList(discardA1s)
      .setFontLine('line-through')
      .setFontColor('#999999')
      .setBackground('#f5f5f5');
  }

  // ── Freeze and resize ─────────────────────────────────────────────────────

  sheet.setFrozenRows(2);
  sheet.setFrozenColumns(1);
  sheet.setColumnWidth(1, 125);
  sheet.setColumnWidths(2, totalCols - 1, 75);
}
