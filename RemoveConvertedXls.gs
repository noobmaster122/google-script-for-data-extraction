/**
 * remove converted xls file
 *
 * @param {string} id
 * @return {void}
 * @customfunction
 */
const removeConvertedXls = (id) => Drive.Files.remove(id);
