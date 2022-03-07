/**
 * remove converted xls file
 *
 * @param {string} file id.
 * @return {void}
 * @customfunction
 */
const removeConvertedXls = (id) => Drive.Files.remove(id);
