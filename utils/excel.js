const xlsx = require("xlsx");

/**
 * Read an Excel file and return the data as an array of objects
 * @param {string} file
 * @returns
 */
const readExcelFile = (file) => {
  try {
    console.log("Reading Excel file: ", file);

    // Read the workbook
    const workbook = xlsx.readFile(file);

    // Get the first sheet
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    // Convert the sheet to an array of objects
    const data = xlsx.utils.sheet_to_json(sheet);

    // log the data
    console.log("Data read successfully");

    // return the data
    return data;
  } catch (error) {
    console.error(error);
    return [];
  }
};

/**
 * Write data to an Excel file
 * @param {string} file
 * @param {Object[]} data
 * @param {string} sheetName
 */
const writeExcelFile = (file, data, sheetName = "sheet1") => {
  try {
    console.log("Writing Excel file: ", file);

    // Create a new workbook
    const workbook = xlsx.utils.book_new();

    // Convert the data to a sheet
    const sheet = xlsx.utils.json_to_sheet(data);

    // Add the sheet to the workbook
    xlsx.utils.book_append_sheet(workbook, sheet, sheetName);

    // Write the workbook to a file
    xlsx.writeFile(workbook, file);

    // log the data
    console.log("Data written successfully");
  } catch (error) {
    console.error(error);
  }
};

module.exports = {
  readExcelFile,
  writeExcelFile,
};
