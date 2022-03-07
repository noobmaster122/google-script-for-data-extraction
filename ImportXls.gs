/**
 * create a tmp folder to hold converted xls files
 * @uses : customNotice()| isExcelSheet() | getTargetFiles() | excelToSheet()
 *
 * @param {string} readFromFolder
 * @return {array}
 * @customfunction
 */
function importXLS(readFromFolder) {

  try {
    const files = getTargetFiles(readFromFolder);// get all xls files 
    let convertedXlsIds = [];
    while (files.hasNext()) {
      let xFile = files.next();
      let name = xFile.getName();
      if (isExcelSheet(name)) {// only parse excel files
        convertedXlsIds.push(excelToSheet(xFile))// save converted file id
      }
    }
    return convertedXlsIds;// return Ids
  } catch (err) {
    customNotice(`Script failed to convert excel sheets for the following reason : \n\n ${err.toString()}`)
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
  const reg = new RegExp("\.xl(?:s[xmb]|tx|[ta]m|s|t|a|w|r)$")
  return !!reg.exec(title);
}
/**
 * get file extension
 *
 * @param {string} title
 * @return {string|null} 
 * @customfunction
 */
const getExcelExt = (title) => {
  const reg = new RegExp("\.xl(?:s[xmb]|tx|[ta]m|s|t|a|w|r)_converted$")
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
 * @return {object} 
 * @customfunction
 */
const getTargetFiles = (name) => {
  const folders = DriveApp.getFoldersByName(name);
  let folder = folders.hasNext() ? folders.next() : undefined;
  if (!folder) throw 'Error: target folder not found! Enter an existing folder!';
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


