/**
 * Remove empty rows in the target converted excel sheet
 *
 * @param {array} data
 * @return {array}
 * @customfunction
 */
const removeEmptyRows = (data) => { 
  try { 
    return data.filter(row => !row.every(cell => cell === (''))) 
    } catch (err) {
       customNotice(`Script failed during empty rows removal for the following reason : \n\n ${err.toString()}`); 
  }
}