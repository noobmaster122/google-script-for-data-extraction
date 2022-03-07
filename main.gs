/**
 * Convert excel sheets and inject their data into our output sheet
 * @uses importXLS() | extractData() | fillOutputSheet()
 *
 * @param {void}
 * @return {void}
 * @customfunction
 */
const main = () => {

  const excelSheetsFolder = scenario();
  const convertedXlsIds = importXLS(excelSheetsFolder);// convert xls into sheets

  fillOutputSheet(extractData(convertedXlsIds));// write data into result sheet
}