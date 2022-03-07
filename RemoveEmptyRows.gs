/**
 * Remove empty rows in the target converted excel sheet
 *
 * @param {array} data
 * @return {array}
 * @customfunction
 */
const removeEmptyRows = (data) => {
  try{
    let cleanedRows = [];
    data.forEach(row => {
      let counter = 0;
      row.forEach(cell => {
        if (cell === '') counter++;
      });
      if (row.length !== counter) cleanedRows.push(row);
    });
    return cleanedRows;
  } catch (err) {
    customNotice(`Script failed during empty rows removal for the following reason : \n\n ${err.toString()}`);
  }
}