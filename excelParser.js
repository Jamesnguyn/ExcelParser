const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

// Specify the directory path
const privateDataExcelFilePath = 'privateData';

// Specify the sheet names you want to parse
const sheetsToParse = [
  { sheetName: 'ble_stats', columns: ['toc', 'ttc', 'ttd', 'tic'] },
  { sheetName: 'sensordata', columns: ['timestamp', 'Display Time'] }
];

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
    sheetsToParse.forEach(sheetInfo => {
      const { sheetName, columns } = sheetInfo;

      // Check if sheet exists in workbook
      if(privateDataExcel.Sheets.hasOwnProperty(sheetName)) {
        // Get the data from the sheet
        const sheetData = XLSX.utils.sheet_to_json(privateDataExcel.Sheets[sheetName], { header: 1 });

        // Find column indices in the header row
        const headerRow = sheetData[0];

        // Extract specific columns from each row
        const parsedData = sheetData.slice(1).map(row => {
          const rowData = {};
          columns.forEach(col => {
            const columnIndex = headerRow.indexOf(col);
            rowData[col] = row[columnIndex];
          });
          return rowData;
        });

        // Now 'sheetData' contains an array of objects representing the data in the current sheet
        console.log(`Data from sheet "${sheetName}" in file "${excelFile}":`, parsedData);
      }
      else {
        console.log(`Sheet "${sheetName}" not found in file "${excelFile}".`);
      }
    });
  });
});