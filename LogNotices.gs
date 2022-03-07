/**
 * wrap the script errors in a toast and display to user
 *
 * @param {string}
 * @return {void}
 * @customfunction
 */
const customNotice = (msg) => SpreadsheetApp.getUi().alert(msg);