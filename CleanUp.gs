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


