/**
 * Crop Sheet add-on for Google Sheets. Allows users to remove excess rows and
 * columns from their spreadsheet based on the current selection or the cells
 * that have data.
 * @OnlyCurrentDoc
 */

// ESLint config.
/* exported onOpen, cropToSelection, cropToData */

/**
 * Adds a custom menu when the spreadsheet is opened.
 */
function onOpen() {
  SpreadsheetApp.getUi().createAddonMenu()
    .addItem('Crop to Data', 'cropToData')
    .addItem('Crop to Selection', 'cropToSelection')
    .addToUi();
}

/**
 * Crops the current sheet to the user's selected range.
 */
function cropToSelection() {
  var selectedRange = SpreadsheetApp.getActiveSheet().getActiveRange();
  cropSheetToRange(selectedRange);
}

/**
 * Crops the current sheet to the data range.
 */
function cropToData() {
  var dataRange = SpreadsheetApp.getActiveSheet().getDataRange();
  cropSheetToRange(dataRange);
}

/**
 * Crops the sheet to the specified range.
 * @param {SpreadsheetApp.Range} range The range to crop to.
 */
function cropSheetToRange(range) {
  var sheet = range.getSheet();
  var maxRows = sheet.getMaxRows();
  var maxColumns = sheet.getMaxColumns();

  // Define the new range boundaries
  var firstRow = range.getRow();
  var lastRow = firstRow + range.getNumRows() - 1;
  var firstColumn = range.getColumn();
  var lastColumn = firstColumn + range.getNumColumns() - 1;

  // Delete excess rows below the range
  if (lastRow < maxRows) {
    sheet.deleteRows(lastRow + 1, maxRows - lastRow);
  }

  // Delete excess rows above the range
  if (firstRow > 1) {
    sheet.deleteRows(1, firstRow - 1);
  }

  // Delete excess columns to the right of the range
  if (lastColumn < maxColumns) {
    sheet.deleteColumns(lastColumn + 1, maxColumns - lastColumn);
  }

  // Delete excess columns to the left of the range
  if (firstColumn > 1) {
    sheet.deleteColumns(1, firstColumn - 1);
  }

  // Activate the cropped range
  sheet.setActiveRange(range);
}
