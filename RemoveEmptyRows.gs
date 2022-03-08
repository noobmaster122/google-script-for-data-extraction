/**
 * Remove empty rows in the target converted excel sheet
 *
 * @param {array} data
 * @return {array}
 * @customfunction
 */
const removeEmptyRows = (data) => data.filter(row => !row.every(cell => cell === ('')));
