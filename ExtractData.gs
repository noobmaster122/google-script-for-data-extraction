/**
 * extract data out of converted excel files
 * format the data into an array
 * delete converted excel files
 * @uses removeEmptyRows() | formatData() | removeConvertedXls()
 *
 * @param {array} xlsIds.
 * @return {array}
 * @customfunction
 */
const extractData = (xlsIds) => {
  try {
    const ss = SpreadsheetApp;
    const cleaneData = [];
    xlsIds.forEach(id => {
      const convertedSheet = ss.openById(id);
      let data = convertedSheet.getDataRange().getValues();// get raw data
      data = formatData(removeEmptyRows(data), convertedSheet.getName());// remove empty columns and format data 
      cleaneData.push(...data);// save extracted data
      removeConvertedXls(id);// delete sheet after extracting data
    });
    cleaneData.unshift(['Project Name', 'Date', 'Value']);//append the table headers
    return cleaneData;
  } catch (err) {
    customNotice(__getStackTrace__(err));
  }
}