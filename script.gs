function onOpen(){
  main()// init script
}
/**
 * usage workflow
 * @uses customNotice()
 *
 * @param {void}
 * @return {void}
 * @customfunction
 */
 const scenario = () => {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const defaultEntryFolder = "GAP";

  let presetsResponse = ui.alert(`do you wish to use default presets? \n ${defaultEntryFolder} for reading the excel sheets`, ui.ButtonSet.YES_NO);

  // Process the user's response.
  if (presetsResponse != ui.Button.YES) {
    return grabInputFolderName(ui, defaultEntryFolder);
  } else {
    return defaultEntryFolder;
  }

}
/**
 * default folder where the raw excel sheets are!
 *
 * @param {object}
 * @param {string}
 * 
 * @return {string} 
 * @customfunction
 */
const grabInputFolderName = (ui, defaultFolder) => {
  let inputFolderRes = ui.prompt('Where have you stored the excel files ?');
  let res = '';
  // Process the user's response.
  if (inputFolderRes.getSelectedButton() == ui.Button.OK) {
    let x = inputFolderRes.getResponseText();
    if (x.length !== 0) res = inputFolderRes.getResponseText();
    if (x.length === 0) {
      customNotice(`Default folder (${defaultFolder}) will be used`);
      res = defaultFolder;
    }
  } else {
    customNotice(`Default folder (${defaultFolder}) will be used`);
    res = defaultFolder;
  }

  return res;

}
/**
 * Convert excel sheets and inject their data into our output sheet
 * @uses importXLS() | extractData() | fillOutputSheet()
 *
 * @param {void}
 * @return {void}
 * @customfunction
 */
 const main = () => {

  let excelSheetsFolder = scenario();
  const convertedXlsIds = importXLS(excelSheetsFolder);// convert xls into sheets

  fillOutputSheet(extractData(convertedXlsIds));// write data into result sheet
}
/**
 * wrap the script errors in a toast and display to user
 *
 * @param {string}
 * @return {void}
 * @customfunction
 */
 const customNotice = (msg) => {
  const ss = SpreadsheetApp;
  ss.getUi().alert(msg);
}
/**
 * create a tmp folder to hold converted xls files
 * @uses : customNotice()| isExcelSheet() | getTargetFiles() | excelToSheet()
 *
 * @param {string}
 * @return {array} of sheets Ids
 * @customfunction
 */
 function importXLS(readFromFolder) {

  try {
    let files = getTargetFiles(readFromFolder);// get all xls files 
    let convertedXlsIds = [];
    while (files.hasNext()) {
      let xFile = files.next();
      let name = xFile.getName();
      if (isExcelSheet(name)) {// only parse excel files
        convertedXlsIds.push(excelToSheet(xFile))// save converted file id
      }
    }
    return convertedXlsIds;// return Ids
  } catch (f) {
    customNotice(f.toString())
  }

}
/**
 * create a tmp folder to hold converted xls files
 * @see https://support.microsoft.com/en-us/office/file-formats-that-are-supported-in-excel-0943ff2c-6014-4e8d-aaea-b83d51d46247
 *
 * @param {object}
 * @return {string} 
 * @customfunction
 */
const isExcelSheet = (title) => {
  let reg = new RegExp("\.xl(s[xmb]|tx|[ta]m|s|t|a|w|r)$")
  return reg.exec(title) ? true : false;
}
/**
 * create a tmp folder to hold converted xls files
 * @uses geTmpFolderId()
 *
 * @param {object}
 * @return {string} 
 * @customfunction
 */
const excelToSheet = (file) => {

  let ID = file.getId();// get excel sheet id
  let name = file.getName();
  let xBlob = file.getBlob();// extract its blob
  let newFile = {
    title: name + '_converted',
    key: ID,
    "parents": [{ 'id': geTmpFolderId() }]// push the converted file into the output folder
  }
  convertedFile = Drive.Files.insert(newFile, xBlob, {
    convert: true,
  });

  return convertedFile.getId();
}
/**
 * retrieve files
 *
 * @param {string}
 * @return {object} 
 * @customfunction
 */
const getTargetFiles = (name) => {
  let folders = DriveApp.getFoldersByName(name);
  folder = folders.hasNext() ? folders.next() : undefined;
  if (!!!folder) throw 'Error: target folder not found! Enter an existing folder!';
  return folder
    .getFiles();
  //  .searchFiles('title != "nothing"');
}
/**
 * get id of tmp folder or create one
 *
 * @param {void}
 * @return {string} 
 * @customfunction
 */
const geTmpFolderId = () => {
  let folders = DriveApp.getFoldersByName('tmp');
  let tmpFolder = folders.hasNext() ? folders.next() : undefined;
  if (!tmpFolder) {
    tmpFolder = DriveApp.createFolder('tmp');// create the tmp folder in the root of the drive
  }

  return tmpFolder.getId();

}
/**
 * inject extracted data into our output sheet
 *
 * @param {array} array of arrays.
 * @return {void}
 * @customfunction
 */
 const fillOutputSheet = (data) => {
  try {
    const s = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    s.getRange(1, 1, data.length, 3).setValues(data);
  } catch (f) {
    customNotice(f.toString())
  }
}
/**
 * remove empty columns from data rows
 * format data into 3 column arrays
 *
 * @param {array} data array.
 * @param {string} converted excel file name.
 * @return {array | false}.
 * @customfunction
 */
 const formatData = (xlsData, fileName) => {
  try {
    let fileNameWithoutExt = fileName.split('.x')[0];// get file title without ext
    let data = []
    if (xlsData[0].length === xlsData[1].length) {
      for (let i = 0; i < xlsData[0].length; i++) {
        let emptyCellsCond = xlsData[0][i] === '' && xlsData[1][i] === '';// dont save empty columns
        let headerCellsCond = xlsData[0][i] === 'Date' && xlsData[1][i] === 'Value';// dont save header cells
        if (!emptyCellsCond && !headerCellsCond) data.push([fileNameWithoutExt, xlsData[0][i], xlsData[1][i]]);
      }
    } else {
      return false;
    }

    return data;

  } catch (f) {
    customNotice(f.toString())
  }
}
/**
 * remove converted xls file
 *
 * @param {string} file id.
 * @return {void}
 * @customfunction
 */
 const removeConvertedXls = (id) => {
  try {
    file = Drive.Files.remove(id);
  } catch (f) {
    customNotice(f.toString())
  }
}
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
    const cleaneData = [];
    xlsIds.forEach(id => {
      let convertedSheet = ss.openById(id);
      let data = convertedSheet.getDataRange().getValues();// get raw data
      data = formatData(removeEmptyRows(data), convertedSheet.getName());// remove empty columns and format data 
      cleaneData.push(...data);// save extracted data
      removeConvertedXls(id);// delete sheet after extracting data
    });

    cleaneData.unshift(['File', 'Date', 'Value']);//append the table headers
    return cleaneData;
  } catch (f) {
    customNotice(f.toString())
  }
}
/**
 * Remove empty rows in the target converted excel sheet
 *
 * @param {array} array of arrays (data rows).
 * @return {array} array of arrays
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
  } catch (f) {
    customNotice(f.toString())
  }
}