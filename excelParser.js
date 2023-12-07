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

let sensorDataFirstTimestamp;
let sensorDataLastTimestamp;
let lowerBound;
let upperBound;

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

    // Extract file name without extension
    const fileNameWithoutExtension = path.parse(excelFile).name;

    // Split fileNameWithoutExtension by underscores
    const fileNameParts = fileNameWithoutExtension.split('_');

    // find lower bound date and time
    const firstPart = fileNameParts.slice(0, 2).join('.');
    const firstDateAndTime = firstPart.split(".");
    const firstDatePart = firstDateAndTime[0];
    const firstFormattedDate = `${firstDatePart.slice(0, 2)}/${firstDatePart.slice(2, 4)}/${firstDatePart.slice(4)}`;
    const firstTimePart = firstDateAndTime[1];
    const firstFormattedTime = `${firstTimePart.slice(0, 2)}:${firstTimePart.slice(2)}:00`;
    lowerBound = firstFormattedDate.concat(".", firstFormattedTime);

    // find upper bound date and time
    const secondPart = fileNameParts.slice(2).join('.');
    const secondDateAndTime = secondPart.split(".");
    const secondDatePart = secondDateAndTime[0];
    const secondFormattedDate = `${secondDatePart.slice(0, 2)}/${secondDatePart.slice(2, 4)}/${secondDatePart.slice(4)}`;
    const secondTimePart = secondDateAndTime[1];
    const secondFormattedTime = `${secondTimePart.slice(0, 2)}:${secondTimePart.slice(2)}:00`;
    upperBound = secondFormattedDate.concat(".", secondFormattedTime);

    // Read the Excel file
    const privateDataExcel = XLSX.readFile(excelFilePath);

    let bleStatsData = [];
    let timestampColumnName;

    // Iterate through each sheet
    sheetsToParse.forEach(sheetInfo => {
      const { sheetName, columns } = sheetInfo;

      // Check if sheet exists in workbook
      if (privateDataExcel.Sheets.hasOwnProperty(sheetName)) {
        // Get the data from the sheet
        const sheetData = XLSX.utils.sheet_to_json(privateDataExcel.Sheets[sheetName], { header: 1 });

        // Find column indices in the header row
        const headerRow = sheetData[0];

        // Extract specific columns from each row
        const parsedData = sheetData.slice(1).map(row => {
          const rowData = {};
          columns.forEach(col => {
            const columnIndex = headerRow.indexOf(col);

            // Check if the current column is "Display Time"
            // Now 'sheetData' contains an array of objects representing the data in the current sheet
            if (col === "Display Time") {
              // Split and remove the last element
              const displayTimeParts = row[columnIndex].split(".");
              displayTimeParts.pop(); // Remove the last element
              rowData[col] = displayTimeParts.join(".");
            } else {
              rowData[col] = row[columnIndex];
            }
          });
          return rowData;
        });

        // Filter data between lowerBound and upperBound
        const filteredData = parsedData.filter(row => {
          const rowDisplayTime = row["Display Time"];
          return rowDisplayTime >= lowerBound && rowDisplayTime <= upperBound;
        });

        // Extract and print only the timestamp values from the "sensordata" sheet
        if (sheetName === "sensordata") {
          if (filteredData.length > 0) {
            const timestampValues = filteredData.map(row => row["timestamp"]);
            sensorDataFirstTimestamp = timestampValues[0];
            sensorDataLastTimestamp = timestampValues[timestampValues.length - 1];
          }
        }
        console.log("First TimeStamp: ", sensorDataFirstTimestamp);
        console.log("Last TimeStamp: ", sensorDataLastTimestamp);
        console.log("Lower Bound: ", lowerBound);
        console.log("Upper Bound: ", upperBound);

        if (sheetName === "ble_stats") {
          // Check if the "timestamp" column is present
          timestampColumnName = "toc"; // Change this to the actual column name
          if (!headerRow.includes(timestampColumnName)) {
            console.error(`Error: "${timestampColumnName}" column not found in sheet "${sheetName}" in file "${excelFile}".`);
            return;
          }

          bleStatsData = sheetData.slice(1).map(row => {
            const rowData = {};
            columns.forEach(col => {
              const columnIndex = headerRow.indexOf(col);
              rowData[col] = row[columnIndex];
            });
            return rowData;
          });
        }
        else {
          console.log(`Sheet "${sheetName}" not found in file "${excelFile}".`);
        }
      }
    });

    const filteredBleStatsData = bleStatsData.filter(row => {
      const rowTimestamp = row[timestampColumnName];
      return rowTimestamp >= sensorDataFirstTimestamp && rowTimestamp <= sensorDataLastTimestamp;
    });

    console.log('FINAL FORM', filteredBleStatsData)
    console.log('FINAL FORM Length', filteredBleStatsData.length)

  });
});