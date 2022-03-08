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

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  //create menu
   const menu = SpreadsheetApp.getUi().createMenu("⚙️ Custom scripts");
   menu.addItem("Extract excel data", "main");
   menu.addToUi();

  const defaultEntryFolder = "GAP";

  const presetsResponse = ui.alert(`do you wish to use default presets? \n\n 
                                  *- ${defaultEntryFolder} for reading the excel sheets
                                  *- tmp folder will be used to store converted excel sheets`, ui.ButtonSet.YES_NO_CANCEL);

  // Process the user's response.
  if (presetsResponse == ui.Button.YES) {
    return defaultEntryFolder;// default will be returned if no folder is chosen!
  } else if (presetsResponse == ui.Button.NO){
    return grabInputFolderName(ui, defaultEntryFolder);// retrieve the folder name from the user
  } else{
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