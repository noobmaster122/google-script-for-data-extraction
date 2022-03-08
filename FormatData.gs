/**
 * remove empty columns from data rows
 * format data into 3 column arrays
 * the script will ignore extra columns
 * @uses getExcelExt()
 *
 * @param {array} xlsData
 * @param {string} fileName
 * @return {array}
 * @customfunction
 */
const formatData = (xlsData, fileName) => {
  const fileNameWithoutExt = fileName.split(fileNameTimestampAndExt(fileName))[0];// get file title without ext
  let data = []// save file rows
  if (xlsData[0].length === xlsData[1].length) {// start data extraction if date and value rows are equal
    for (let i = 0; i < xlsData[0].length; i++) {
      const emptyCellsCond = xlsData[0][i] === '' && xlsData[1][i] === '';// dont save empty columns
      const headerCellsCond = xlsData[0][i] === 'Date' && xlsData[1][i] === 'Value';// dont save header cells
      if (!emptyCellsCond && !headerCellsCond) data.push([fileNameWithoutExt, xlsData[0][i], xlsData[1][i]]);
    }
  } else {
    throw `Size mismatch between the Date and value row \n this file : ${filename} might be incorrectly formatted!`;// this will get catched in main()
  }

  return data;
}