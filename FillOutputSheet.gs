/**
 * inject extracted data into our output sheet
 *
 * @param {array} data.
 * @return {void}
 * @customfunction
 */
const fillOutputSheet = (data) => {
  try {
    const s = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    s.getRange(1, 1, data.length, 3).setValues(data);
  } catch (err) {
    customNotice(`Script failed to write data into the output sheet for the following reason : \n\n ${err.toString()}`);
  }
}