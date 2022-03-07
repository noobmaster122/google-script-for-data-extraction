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

  const presetsResponse = ui.alert(`do you wish to use default presets? \n\n 
                                  *- ${defaultEntryFolder} for reading the excel sheets
                                  *- tmp folder will be used to store converted excel sheets`, ui.ButtonSet.YES_NO);

  // Process the user's response.
  if (presetsResponse != ui.Button.YES) {
    return grabInputFolderName(ui, defaultEntryFolder);
  } else {
    return defaultEntryFolder;// defaul will be returned if no folder is chosen!
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