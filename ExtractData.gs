/**
 * extract data out of converted excel files
 * format the data into an array
 * delete converted excel files
 * @uses removeEmptyRows() | formatData() | removeConvertedXls()
 *
 * @param {array} array of files Ids.
 * @return {array} array of arrays
 * @customfunction
 */
const extractData = (xlsIds) => {
  try {
    const ss = SpreadsheetApp;
    let cleaneData = [];
    xlsIds.forEach(id => {
      const convertedSheet = ss.openById(id);
      let data = convertedSheet.getDataRange().getValues();// get raw data
      data = formatData(removeEmptyRows(data), convertedSheet.getName());// remove empty columns and format data 
      cleaneData.push(...data);// save extracted data
      removeConvertedXls(id);// delete sheet after extracting data
    });

    cleaneData.unshift(['File', 'Date', 'Value']);//append the table headers
    return cleaneData;
  } catch (err) {
    customNotice(`Script failed during data extraction for the following reason : \n\n ${err.toString()}`);
  }
}