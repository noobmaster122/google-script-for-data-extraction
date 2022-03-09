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