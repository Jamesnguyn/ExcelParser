const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

// Specify the directory path
const privateDataExcelFilePath = 'privateData';

// Specify the sheet names you want to parse
const sheetsToParse = ['SalesOrders', 'Sale Financials'];


// Read the files in the directory
fs.readdir(privateDataExcelFilePath, (err, files) => {
  if (err) {
    console.error('Error reading directory:', err);
    return;
  }

  // Filter Excel files
  const excelFiles = files.filter(file => path.extname(file).toLowerCase() === '.xlsx');

  // Iterate through each Excel file
  excelFiles.forEach(excelFile => {
    const excelFilePath = path.join(privateDataExcelFilePath, excelFile);

    // Read the Excel file
    const privateDataExcel = XLSX.readFile(excelFilePath);

    // Iterate through each sheet
    sheetsToParse.forEach(sheetName => {

      // Check if sheet exists in workbook
      if(privateDataExcel.Sheets.hasOwnProperty(sheetName)) {
        // Get the data from the sheet
        const sheetData = XLSX.utils.sheet_to_json(privateDataExcel.Sheets[sheetName]);

        // Now 'sheetData' contains an array of objects representing the data in the current sheet
        console.log(`Data from sheet "${sheetName}" in file "${excelFile}":`, sheetData);
      }
      else {
        console.log(`Sheet "${sheetName}" not found in file "${excelFile}".`);
      }
    });
  });
});