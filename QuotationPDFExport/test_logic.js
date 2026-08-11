/**
 * Node-based unit tests for the PURE logic functions extracted from
 * PDFExport.gs (computeItemVisibility_ and computeSpacerExtraPx_).
 * These functions contain no Apps Script API calls, so they are copied
 * here verbatim to verify behavior before pasting into Apps Script.
 *
 * Run: node test_logic.js
 */

function computeItemVisibility_(cellValues, minVisible) {
  var totalRows = cellValues.length;

  var lastFilledOffset = -1;
  for (var i = 0; i < totalRows; i++) {
    var v = cellValues[i];
    if (v !== '' && v !== null && v !== undefined && String(v).trim() !== '') {
      lastFilledOffset = i;
    }
  }

  var filledCount = lastFilledOffset + 1;
  var visibleCount = Math.max(filledCount, minVisible);
  visibleCount = Math.min(visibleCount, totalRows);

  var hiddenCount = totalRows - visibleCount;

  return {
    filledCount: filledCount,
    visibleCount: visibleCount,
    hiddenCount: hiddenCount
  };
}

function computeSpacerExtraPx_(spaceLeftOnCurrentPage, footerHeight) {
  if (footerHeight <= spaceLeftOnCurrentPage) {
    return 0;
  }
  var extra = Math.max(0, spaceLeftOnCurrentPage - 1);
  return Math.round(extra);
}

// ---------- test helpers ----------
var passed = 0;
var failed = 0;

function assertEqual(actual, expected, label) {
  var ok = JSON.stringify(actual) === JSON.stringify(expected);
  if (ok) {
    passed++;
    console.log('PASS: ' + label);
  } else {
    failed++;
    console.log('FAIL: ' + label + ' -> expected ' + JSON.stringify(expected) + ', got ' + JSON.stringify(actual));
  }
}

function makeRows(filledCount, totalRows) {
  var arr = [];
  for (var i = 0; i < totalRows; i++) {
    arr.push(i < filledCount ? 'item ' + (i + 1) : '');
  }
  return arr;
}

var TOTAL_ITEM_ROWS = 42; // rows 26..67 inclusive

// ---------- Test 1: exactly the user's example -> 13 filled rows, min 5 ----------
(function () {
  var rows = makeRows(13, TOTAL_ITEM_ROWS);
  var result = computeItemVisibility_(rows, 5);
  assertEqual(result, { filledCount: 13, visibleCount: 13, hiddenCount: TOTAL_ITEM_ROWS - 13 },
    '13 filled rows -> 13 visible, rest hidden (matches user example)');
})();

// ---------- Test 2: fewer than min (e.g. 3 filled) -> still show minimum 5 ----------
(function () {
  var rows = makeRows(3, TOTAL_ITEM_ROWS);
  var result = computeItemVisibility_(rows, 5);
  assertEqual(result, { filledCount: 3, visibleCount: 5, hiddenCount: TOTAL_ITEM_ROWS - 5 },
    '3 filled rows -> minimum 5 stay visible (2 blank rows shown), rest hidden');
})();

// ---------- Test 3: zero filled rows -> minimum 5 still visible ----------
(function () {
  var rows = makeRows(0, TOTAL_ITEM_ROWS);
  var result = computeItemVisibility_(rows, 5);
  assertEqual(result, { filledCount: 0, visibleCount: 5, hiddenCount: TOTAL_ITEM_ROWS - 5 },
    '0 filled rows -> minimum 5 blank rows still visible');
})();

// ---------- Test 4: exactly 5 filled rows -> exactly 5 visible, rest hidden ----------
(function () {
  var rows = makeRows(5, TOTAL_ITEM_ROWS);
  var result = computeItemVisibility_(rows, 5);
  assertEqual(result, { filledCount: 5, visibleCount: 5, hiddenCount: TOTAL_ITEM_ROWS - 5 },
    '5 filled rows -> exactly 5 visible (boundary case)');
})();

// ---------- Test 5: all rows filled -> nothing hidden ----------
(function () {
  var rows = makeRows(TOTAL_ITEM_ROWS, TOTAL_ITEM_ROWS);
  var result = computeItemVisibility_(rows, 5);
  assertEqual(result, { filledCount: TOTAL_ITEM_ROWS, visibleCount: TOTAL_ITEM_ROWS, hiddenCount: 0 },
    'all 42 rows filled -> none hidden');
})();

// ---------- Test 6: gap in the middle (row filled after a blank) still counted via last-filled-offset ----------
(function () {
  var rows = makeRows(10, TOTAL_ITEM_ROWS);
  rows[15] = 'late entry'; // a filled row far beyond the first 10, with blanks in between
  var result = computeItemVisibility_(rows, 5);
  assertEqual(result, { filledCount: 16, visibleCount: 16, hiddenCount: TOTAL_ITEM_ROWS - 16 },
    'gap-filled row at offset 15 -> visible extends through it (16 visible)');
})();

// ---------- Test 7: whitespace-only cell treated as blank ----------
(function () {
  var rows = makeRows(6, TOTAL_ITEM_ROWS);
  rows[5] = '   '; // whitespace only, should be treated as blank, so last real filled = offset 4
  var result = computeItemVisibility_(rows, 5);
  assertEqual(result, { filledCount: 5, visibleCount: 5, hiddenCount: TOTAL_ITEM_ROWS - 5 },
    'whitespace-only cell not counted as filled');
})();

// ---------- Footer spacer logic ----------

// Test 8: footer fits fully in remaining page space -> no push needed
(function () {
  var result = computeSpacerExtraPx_(300, 200);
  assertEqual(result, 0, 'footer (200px) fits in remaining space (300px) -> no spacer extra');
})();

// Test 9: footer does NOT fit -> push footer to next page (fill remaining space)
(function () {
  var result = computeSpacerExtraPx_(150, 400);
  assertEqual(result, 149, 'footer (400px) does not fit in remaining 150px -> spacer pushes to next page');
})();

// Test 10: footer exactly equals remaining space -> fits exactly, no push
(function () {
  var result = computeSpacerExtraPx_(250, 250);
  assertEqual(result, 0, 'footer exactly equal to remaining space -> fits, no push needed');
})();

console.log('\n' + passed + ' passed, ' + failed + ' failed');
if (failed > 0) {
  process.exit(1);
}
