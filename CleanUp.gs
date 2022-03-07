/**
 * the script wont clear the tmp folder if it fails
 * in which case, this function will do that
 *
 * @param {void}
 * @return {void}
 * @customfunction
 */
const cleanUp = () => {
  const filesList = listFolderFiles();
  if(filesList.length === 0) return;
  if(fileDeletionNotice()) filesList.forEach(file => removeConvertedXls(file.id));// delete files inside tmp folder  
}
/**
 *
 * @param {void}
 * @return {array}
 * @customfunction
 */
const listFolderFiles = () => {
  let files = [];
  const defaultFolder = 'tmp';
  const filesEntry = getTargetFiles(defaultFolder);
  while (filesEntry.hasNext()) {
    const file = filesEntry.next();
    files.push({id: file.getId(), title: file.getName()});
  }
  return files;
}
/**
 *
 * @param {void}
 * @return {bool}
 * @customfunction
 */
const fileDeletionNotice = () => {
    const ui = SpreadsheetApp.getUi();
    const files = listFolderFiles().map(file => file.title).join('\n file name : ');


  ui.alert(`These files inside the tmp folder will get deleted!Do you wish to continue ? \n\n 
                                  ${files}`, ui.ButtonSet.YES_NO);

  return ui.Button.YES ? true : false;
}
