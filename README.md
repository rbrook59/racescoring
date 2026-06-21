# Stamford Yacht Club — Race Scoring Library

Version 3.1 — June 2026

Developed by Rick Brook (rick.brook@gmail.com) for Stamford Yacht Club.
The Cox-Sprague algorithm was originally developed in VBA by Witold Gesing (wgesing@gmail.com).

This software may be copied and re-distributed freely. If you make changes, please share them with the above authors.

---

## Overview

This Google Apps Script library calculates sailboat race series scores from results entered in a Google Sheet. It supports four scoring methods simultaneously:

- **Cox-Sprague** — high-point percentage of perfection (table-based for ≤ 20 boats, log formula for > 20)
- **High Point of Perfection** — percentage-based high-point score
- **Average Score** — low-point average place across races sailed
- **Low Point** — total of low-point places across races sailed

All four methods support optional discards (dropping the worst N races from the series total).

---

## Main Functions

### `scoreSailors()`

The main entry point, called from the sheet's Race Scoring menu. Reads all data from the active spreadsheet, calls the appropriate scoring functions for each sailor, writes the calculated scores back to the sheet, and sorts the results range.

- Reads column configuration from the `ControlPanel` sheet
- Reads race results, starters, and finishers from the `RaceResults` sheet
- Only calculates scoring methods whose output columns are configured (> 0)
- Sorts qualified sailors above non-qualified; sorts high-point scores descending, low-point scores ascending

### `csScore(referenceArray, discards, csGrid)`

Calculates the Cox-Sprague percentage-of-perfection score for a sailor's series.

- For races with **≤ 20 starters**: looks up values in the hard-coded CS table
- For races with **> 20 starters**: uses the log formula `100 + 200 × log₁₀(N − P + 1) / log₁₀(N − 1)`
- DNF boats in large-fleet races receive the floor score of 100 (same as last place)
- Discards are selected by dropping the races with the lowest actual/max ratio
- Returns a decimal value ≤ 1 (percentage of perfection)

### `hpScore(referenceArray, discards)`

Calculates the High Point of Perfection score for a sailor's series.

- Per-race value: `(starters − place + 1) / starters`
- Series score: `sum(numerators) / sum(denominators)` across kept races
- Discards are the races with the lowest per-race ratio
- Returns a decimal value ≤ 1

### `avgScore(referenceArray, discards)`

Calculates the average low-point score for a sailor's series.

- Sorts races best to worst (lowest place number = best)
- Removes the worst `discards` races
- Returns `sum(places) / number of races sailed` — lower is better

### `lowScore(referenceArray, discards)`

Calculates the total low-point score for a sailor's series.

- Sorts races best to worst, removes the worst `discards` races
- Returns the sum of place numbers — lower is better

### `cleanScores(resultsArray, startersArray, finishersArray)`

Converts raw result entries into a normalized numeric array for use by the scoring functions.

Penalty code conversions:

| Code                              | Meaning                 | Score assigned           |
| --------------------------------- | ----------------------- | ------------------------ |
| DNF, DSQ, RAF, OCS, BFD, RET, UFD | Penalty finish          | Starters + 1             |
| TLE                               | Time limit expired      | Finishers + 2            |
| DNC, DNS                          | Did not compete / start | 0 (excluded from series) |

Returns a 2D array where each row is `[place, starters, finishers, 0, 0, 0, 0]`. The four trailing zeros are working space used by the scoring functions.

### `buildCSTable(original)`

Builds and returns the hard-coded Cox-Sprague lookup table as a 22 × 21 array (rows 0–21, columns 0–20).

- Row index = finishing place (1-based)
- Column index = number of starters (1-based)
- Row 0 and column 0 hold index headings so that place and starters integers can be used directly as array indices
- When `original = true`, uses the original Cox-Sprague table values
- Default (`original = false`) uses the Y.R.A. of Long Island Sound modified table, which adjusts 2-boat race values: 2nd place scores 7 (not 4) and DNF scores 5

---

## Helper Functions

### `loadScoreRange(sheet, rangeName)`

Reads a named range from the sheet and splits it into a header row and data rows.

Returns `{ header, rows }` where `header` is the first row and `rows` is the remaining data. Used to preserve column headings when reading score output columns before writing results back.

### `writeScoreRange(sheet, rangeName, rows, header)`

Writes a header row and data rows back to a named range on the sheet.

Combines `header` and `rows` into a single 2D array and calls `setValues()` in one sheet operation.

### `applyDiscards(cleanArray, discards)`

Sorts a score array ascending by place (column 0) and removes the worst (highest) `discards` entries from the end.

Mutates the array in place. Always call this on a deep copy of the original data, not the source array.

### `csLogScore(place, starters)`

Calculates a single Cox-Sprague score using the log formula for fleets larger than 20 boats.

Formula: `100 + 200 × log₁₀(starters − place + 1) / log₁₀(starters − 1)`

Returns 100 (the floor) for DNF boats where `place > starters`.

---

## Google Sheet Named Ranges

The following named ranges must be defined in each spreadsheet that uses this library. All named ranges are set up via **Data → Named ranges** in Google Sheets.

### ControlPanel sheet

| Named Range    | Contents                     | Description                                                                                                                  |
| -------------- | ---------------------------- | ---------------------------------------------------------------------------------------------------------------------------- |
| `rangeColumns` | Single-column range, 9 rows  | Physical column numbers (1-based) for each data column in the results range. See order below.                                |
| `rangeControl` | Single-column range, 1+ rows | Control flags. Row 1: set to `1` to use the original Cox-Sprague table; `0` for the Y.R.A. Long Island Sound modified table. |

**`rangeColumns` row order** (each row contains a 1-based physical column number):

| Row | Column identified                              |
| --- | ---------------------------------------------- |
| 1   | Skipper / boat name                            |
| 2   | Qualified flag (used for sort priority)        |
| 3   | Number of discards for this sailor             |
| 4   | Cox-Sprague score output column (0 = not used) |
| 5   | High Point score output column (0 = not used)  |
| 6   | Average Score output column (0 = not used)     |
| 7   | Low Point score output column (0 = not used)   |
| 8   | First column of race result data               |
| 9   | Column to sort final standings on              |

### RaceResults sheet

| Named Range       | Contents                                | Description                                                                                   |
| ----------------- | --------------------------------------- | --------------------------------------------------------------------------------------------- |
| `rangeResults`    | Full results table including header row | All sailor rows plus one header row at the top. The header row is stripped before processing. |
| `rangeStarters`   | Single row                              | Number of starters for each race (one value per race column).                                 |
| `rangeFinishers`  | Single row                              | Number of finishers for each race (one value per race column).                                |
| `rangeCoxSprague` | Single-column range including header    | Output column for Cox-Sprague scores. Only required if Cox-Sprague column is configured.      |
| `rangeHighPoint`  | Single-column range including header    | Output column for High Point scores. Only required if High Point column is configured.        |
| `rangeAvgScore`   | Single-column range including header    | Output column for Average scores. Only required if Average Score column is configured.        |
| `rangeLowPoint`   | Single-column range including header    | Output column for Low Point scores. Only required if Low Point column is configured.          |

---

## Setting Up a New Series Spreadsheet

Each series spreadsheet needs a thin wrapper script that calls the shared library. The library contains all scoring logic; the wrapper provides the menu and connects to the library.

Scoring template found here: https://docs.google.com/spreadsheets/d/1gcHldfxAblRt53Q-iujgoTqNqKf6VgV2sYDlaO7oZ3w/edit?gid=302990953#gid=302990953

### Step 1 — Add the library

1. Open the series spreadsheet → **Extensions → Apps Script**
2. In the left panel, click **+** next to **Libraries**
3. Paste the Script ID: `1cWsrDWwDo_PlVuSJzRHibeFrC-0epofe_P-Zoc5JwXN9h6BvTyoP4xh_`
4. Click **Look up**
5. Set **Identifier** to `RaceScoringCodeLibrary`
6. Set **Version** to **Head** (always uses the latest code)
7. Click **Add**

### Step 2 — Add the wrapper script

Replace the entire contents of the script editor with the following:

```javascript
// Race Scoring wrapper — all scoring logic lives in the shared library.
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Race Scoring")
    .addItem("Calculate Scores", "scoreSailors")
    .addSeparator()
    .addToUi();
}

function scoreSailors() {
  RaceScoringCodeLibrary.scoreSailors();
}
```

### Step 3 — Set up named ranges

Define all named ranges listed above in the **ControlPanel** and **RaceResults** sheets using **Data → Named ranges**.

### Step 4 — Test

Reload the spreadsheet. The **Race Scoring** menu should appear. Select **Calculate Scores** to run.

---

## Updating the Library

When the library code is updated and pushed:

1. From the project directory, run: `clasp push`
2. Because dependent sheets use **Head** version, changes are live immediately — no version bumping required.

To release a controlled versioned update instead (e.g., when sharing with other organizations):

1. In the standalone script editor: **Deploy → Manage Deployments → Edit → New Version**
2. In each dependent sheet's script editor: **Libraries → pencil icon → update version number**
