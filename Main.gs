/**
 * Convert excel sheets and inject their data into our output sheet
 * @uses customNotice() | importXLS() | extractData() | fillOutputSheet()
 *
 * @param {void}
 * @return {void}
 * @customfunction
 */
const main = () => {
    const excelSheetsFolder = scenario();
    if (excelSheetsFolder) {
      const convertedXlsIds = importXLS(excelSheetsFolder);// convert xls into sheets
      fillOutputSheet(extractData(convertedXlsIds));// write data into result sheet
    }
    cleanUp();// clear tmp folder
}