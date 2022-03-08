/**
 * inject extracted data into our output sheet
 *
 * @param {array} data.
 * @return {void}
 * @customfunction
 */
const fillOutputSheet = (data) => {
  const s = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  s.getRange(1, 1, data.length, 3).setValues(data);
}