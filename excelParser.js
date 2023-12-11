const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

// Specify the directory path
const privateDataExcelFilePath = 'privateData';

// Specify the sheet names you want to parse
const sheetsToParse = [
  { sheetName: 'ble_stats', columns: ['toc', 'ttc', 'ttd', 'tic', "cmp"] },
  { sheetName: 'sensordata', columns: ['timestamp', 'Display Time'] },
  { sheetName: 'log_info', columns: ['Display Time', 'code0'] }
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
    // console.log(lowerBound);

    // find upper bound date and time
    const secondPart = fileNameParts.slice(2).join('.');
    const secondDateAndTime = secondPart.split(".");
    const secondDatePart = secondDateAndTime[0];
    const secondFormattedDate = `${secondDatePart.slice(0, 2)}/${secondDatePart.slice(2, 4)}/${secondDatePart.slice(4)}`;
    const secondTimePart = secondDateAndTime[1];
    const secondFormattedTime = `${secondTimePart.slice(0, 2)}:${secondTimePart.slice(2)}:00`;
    upperBound = secondFormattedDate.concat(".", secondFormattedTime);
    // console.log(upperBound);

    // Read the Excel file
    const privateDataExcel = XLSX.readFile(excelFilePath);

    let bleStatsData = [];
    let rssiValues = [];
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
            console.log(sensorDataFirstTimestamp);
            console.log(sensorDataLastTimestamp);
          }
        }

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
        
        else if (sheetName === "log_info"){
          rssiValues = filteredData.map(row => parseFloat(row["code0"]));

          const minRssi = Math.min(...rssiValues);
          const maxRssi = Math.max(...rssiValues);
          const avgRssi = rssiValues.reduce((sum, value) => sum + value, 0) / rssiValues.length;
          const roundedAvgRssi = avgRssi.toFixed(2);

          const below88Count = rssiValues.filter(value => value < 88).length;
          const equalOrAbove88Count = rssiValues.filter(value => value >= 88).length;
          const equalOrAbove95Count = rssiValues.filter(value => value >= 95).length;

          console.log("========================================");
          console.log("Total RSSI Count: ", rssiValues.length);
          console.log("Num less than 88: ", below88Count);
          console.log("Num greater than or equal to 88: ", equalOrAbove88Count);
          console.log("Num greater than or equal to 95: ", equalOrAbove95Count);
          console.log('Minimum RSSI:', minRssi);
          console.log('Maximum RSSI:', maxRssi);
          console.log('Average RSSI:', roundedAvgRssi);
          console.log("========================================");
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

    // =================TTC=============================
    // Extract "ttc" values from the filtered BLE Stats data
    const ttcValues = filteredBleStatsData.map(row => parseFloat(row['ttc']));

    // Calculate minimum, maximum, and average values
    const minTTC = Math.min(...ttcValues);
    const maxTTC = Math.max(...ttcValues);
    const avgTTC = ttcValues.reduce((sum, value) => sum + value, 0) / ttcValues.length;

    // Round average to two decimal points
    const roundedAvgTTC = avgTTC.toFixed(2);

    // Calculate Standard Deviation
    // Calculate the sum of squared differences from the average
    const ttcSumSquaredDifferences = ttcValues.reduce((sum, value) => sum + Math.pow(value - avgTTC, 2), 0);
    // Calculate the variance
    const ttcVariance = ttcSumSquaredDifferences / ttcValues.length;
    // Calculate the standard deviation
    const ttcStdDeviation = Math.sqrt(ttcVariance);
    // Round standard deviation to two decimal points
    const ttcRoundedStdDeviation = ttcStdDeviation.toFixed(2);

    // =================TTD=============================
    // Extract "ttd" values from the filtered BLE Stats data
    const ttdValues = filteredBleStatsData.map(row => parseFloat(row['ttd']));

    // Calculate minimum, maximum, and average values
    const minTTD = Math.min(...ttdValues);
    const maxTTD = Math.max(...ttdValues);
    const avgTTD = ttdValues.reduce((sum, value) => sum + value, 0) / ttdValues.length;

    // Round average to two decimal points
    const roundedAvgTTD = avgTTD.toFixed(2);

    // Calculate Standard Deviation
    // Calculate the sum of squared differences from the average
    const ttdSumSquaredDifferences = ttdValues.reduce((sum, value) => sum + Math.pow(value - avgTTD, 2), 0);
    // Calculate the variance
    const ttdVariance = ttdSumSquaredDifferences / ttdValues.length;
    // Calculate the standard deviation
    const ttdStdDeviation = Math.sqrt(ttdVariance);
    // Round standard deviation to two decimal points
    const ttdRoundedStdDeviation = ttdStdDeviation.toFixed(2);

    // =================TIC=============================
    // Extract "ttc" values from the filtered BLE Stats data
    const ticValues = filteredBleStatsData.map(row => parseFloat(row['tic']));

    // Calculate minimum, maximum, and average values
    const minTIC = Math.min(...ticValues);
    const maxTIC = Math.max(...ticValues);
    const avgTIC = ticValues.reduce((sum, value) => sum + value, 0) / ticValues.length;

    // Round average to two decimal points
    const roundedAvgTIC = avgTIC.toFixed(2);

    // Calculate Standard Deviation
    // Calculate the sum of squared differences from the average
    const ticSumSquaredDifferences = ticValues.reduce((sum, value) => sum + Math.pow(value - avgTIC, 2), 0);
    // Calculate the variance
    const ticVariance = ticSumSquaredDifferences / ticValues.length;
    // Calculate the standard deviation
    const ticStdDeviation = Math.sqrt(ticVariance);
    // Round standard deviation to two decimal points
    const ticRoundedStdDeviation = ticStdDeviation.toFixed(2);

    // =================CMP=============================
    // Extract "ttc" values from the filtered BLE Stats data
    const cmpValues = filteredBleStatsData.map(row => parseFloat(row['cmp']));

    // Calculate sum value
    const sumCMP = cmpValues.reduce((sum, value) => sum + value, 0);

    // console.log('Filtered BLE Data', filteredBleStatsData);
    // console.log('Filtered BLE Data Length', filteredBleStatsData.length + 1);

    console.log('Minimum TTC:', minTTC);
    console.log('Maximum TTC:', maxTTC);
    console.log('Average TTC:', roundedAvgTTC);
    console.log('STD TTC:', ttcRoundedStdDeviation);

    console.log("========================================");

    console.log('Minimum TTD:', minTTD);
    console.log('Maximum TTD:', maxTTD);
    console.log('Average TTD:', roundedAvgTTD);
    console.log('STD TTD:', ttdRoundedStdDeviation);

    console.log("========================================");

    console.log('Minimum TIC:', minTIC);
    console.log('Maximum TIC:', maxTIC);
    console.log('Average TIC:', roundedAvgTIC);
    console.log('STD TIC:', ticRoundedStdDeviation);

    console.log("========================================");

    console.log('Sum CMP:', sumCMP);

    console.log("========================================");
  });
});