function onOpen(){
  // welcome to the matrix ;)
  main()// init script
}
/**
 * Remove empty rows in the target converted excel sheet
 *
 * @param {array} data
 * @return {array}
 * @customfunction
 */
const removeEmptyRows = (data) => data.filter(row => !row.every(cell => cell === ('')));
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
/**
 * remove converted xls file
 *
 * @param {string} id
 * @return {void}
 * @customfunction
 */
const removeConvertedXls = (id) => Drive.Files.remove(id);
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
  try{
    const fileNameWithoutExt = fileName.split(fileNameTimestampAndExt(fileName))[0];// get file title without ext
    const data = []// save file rows
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
  } catch(err){
    customNotice(__getStackTrace__(err));
  }
}
/**
 * inject extracted data into our output sheet
 *
 * @param {array} data.
 * @return {void}
 * @customfunction
 */
const fillOutputSheet = (data) => {
  const s = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  s.getRange(1, 1, data.length, 3).setValues(data);
}
/**
 * create a tmp folder to hold converted xls files
 * @uses : isExcelSheet() | getTargetFiles() | excelToSheet()
 *
 * @param {string} readFromFolder
 * @return {array}
 * @customfunction
 */
function importXLS(readFromFolder) {
  try {
    const files = getTargetFiles(readFromFolder);// get all xls files 
    const convertedXlsIds = [];
    while (files.hasNext()) {
      const xFile = files.next();
      const name = xFile.getName();
      if (isExcelSheet(name)) {// only parse excel files
        convertedXlsIds.push(excelToSheet(xFile))// save converted file id
      }
    }
    return convertedXlsIds;// return Ids
  } catch (err) {
    customNotice(__getStackTrace__(err));
  }
}
/**
 * check if file is valid excel sheet
 * @see https://support.microsoft.com/en-us/office/file-formats-that-are-supported-in-excel-0943ff2c-6014-4e8d-aaea-b83d51d46247
 *
 * @param {string} title
 * @return {bool} 
 * @customfunction
 */
const isExcelSheet = (title) => {
  const reg = new RegExp(/\.xl(?:s[xmb]|tx|[ta]m|s|t|a|w|r)$/);
  return !!reg.exec(title);
}
/**
 * get file extension
 *
 * @param {string} title
 * @return {string|null} 
 * @customfunction
 */
const fileNameTimestampAndExt = (title) => {
  const reg = new RegExp(/\s-\s\d*\.xl(?:s[xmb]|tx|[ta]m|s|t|a|w|r)_converted$/);
  return null || reg.exec(title)[0];
}
/**
 * create a tmp folder to hold converted xls files
 * @uses geTmpFolderId()
 *
 * @param {object} file
 * @return {string} 
 * @customfunction
 */
const excelToSheet = (file) => {
  const ID = file.getId();// get excel sheet id
  const name = file.getName();
  const xBlob = file.getBlob();// extract its blob
  const newFile = {
    title: name + '_converted',
    key: ID,
    "parents": [{ 'id': geTmpFolderId() }]// push the converted file into the tmp folder
  }
  convertedFile = Drive.Files.insert(newFile, xBlob, {
    convert: true,
  });

  return convertedFile.getId();
}
/**
 * retrieve files
 *
 * @param {string} name
 * @return {FileIterator} 
 * @customfunction
 */
const getTargetFiles = (name) => {
  const folders = DriveApp.getFoldersByName(name);
  const folder = folders.hasNext() ? folders.next() : undefined;
  if (!folder) throw 'target folder not found! Enter an existing folder!';// this will get catched in main()

  return folder
    .getFiles();
}
/**
 * get id of tmp folder or create one
 *
 * @param {void}
 * @return {string} 
 * @customfunction
 */
const geTmpFolderId = () => {
  const folders = DriveApp.getFoldersByName('tmp');
  let tmpFolder = folders.hasNext() ? folders.next() : undefined;
  if (!tmpFolder) {
    tmpFolder = DriveApp.createFolder('tmp');// create the tmp folder in the root of the drive
  }

  return tmpFolder.getId();
}
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
/**
 * wrap the script errors in a toast and display to user
 *
 * @param {string}
 * @return {void}
 * @customfunction
 */
const customNotice = (msg) => SpreadsheetApp.getUi().alert(msg);
/**
 * get error stack
 * @todo find a better way to display the whole stack trace and not just a shallow one
 *
 * @param {string}
 * @return {string}
 * @customfunction
 */
const __getStackTrace__ = function(message) {
  let s = `Error: ${message}\n`;
  (new Error()).stack
               .split('\n')
               .forEach((token)=>
               {s += `\t${token.trim()}\n`}
  );      
  return s;
}
/**
 * usage workflow
 * when user clicks the close button in the modal, the script will stop running
 * @uses customNotice()
 *
 * @param {void}
 * @return {string|null}
 * @customfunction
 */
const scenario = () => {

  const ui = SpreadsheetApp.getUi();
  //create menu
  const menu = ui.createMenu("⚙️ Custom scripts");
  menu.addItem("Extract excel data", "main");
  menu.addToUi();

  const defaultEntryFolder = "GAP";

  const presetsResponse = ui.alert(`do you wish to use default presets? \n\n 
                                  *- ${defaultEntryFolder} for reading the excel sheets
                                  *- tmp folder will be used to store converted excel sheets`, ui.ButtonSet.YES_NO_CANCEL);

  // Process the user's response.
  if (presetsResponse == ui.Button.YES) {
    return defaultEntryFolder;// default will be returned if no folder is chosen!
  } else if (presetsResponse == ui.Button.NO) {
    return grabInputFolderName(ui, defaultEntryFolder);// retrieve the folder name from the user
  } else {
    return null;// if user clicks cancel or exit button, script will stop running
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
  const inputFolderRes = ui.prompt('Where have you stored the excel files ?');
  let res = '';
  // Process the user's response.
  if (inputFolderRes.getSelectedButton() == ui.Button.OK) {
    const x = inputFolderRes.getResponseText();
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
 * the script wont clear the tmp folder if it fails
 * in which case, this function will do that
 * @uses listFolderFiles() | fileDeletionNotice() | removeConvertedXls()
 *
 * @param {void}
 * @return {void}
 * @customfunction
 */
const cleanUp = () => {
  const filesList = listFolderFiles();
  if (filesList.length === 0) return;
  if (fileDeletionNotice()) filesList.forEach(file => removeConvertedXls(file.id));// delete files inside tmp folder  
}
/**
 * retrieve data of all the files inside tmp folder
 * @uses getTargetFiles()
 * 
 * @param {void}
 * @return {array}
 * @customfunction
 */
const listFolderFiles = () => {
  const files = [];
  const defaultFolder = 'tmp';
  const filesEntry = getTargetFiles(defaultFolder);
  while (filesEntry.hasNext()) {
    const file = filesEntry.next();
    files.push({ id: file.getId(), title: file.getName() });
  }
  return files;
}
/**
 * display a modal informing the user of the files to be deleted, and giving the choice to proceed or not!
 * @uses listFolderFiles()
 * 
 * @param {void}
 * @return {bool}
 * @customfunction
 */
const fileDeletionNotice = () => {
  const ui = SpreadsheetApp.getUi();
  const files = listFolderFiles().map(file => file.title).join('\n file name : ');
  const presetsResponse = ui.alert(`These files inside the tmp folder will get deleted!Do you wish to continue ? \n\n 
                                  ${files}`, ui.ButtonSet.YES_NO);
  return presetsResponse == ui.Button.YES;
}


