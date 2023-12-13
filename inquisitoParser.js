const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

// Specify the directory path
const inquisitoDataExcelFilePath = 'inquisitoData';

const fileType = '.xlsx';

const sheetPhoneErrorLog = 'PhoneErrorLog';
const sheetPhoneUserActivity = 'PhoneUserActivity';

const columnMessage = 'Message';
const columnRecordedDisplayTime = 'RecordedDisplayTime';
const columnData = 'Data';

const messageCGMReadings = 'CGM readings:';
const messageReadEGV = 'readEGV';
const messageOnReadRssi = 'onReadRssi:';
const dataRssi = `"Rssi":`;

let lowerBound;
let upperBound;

const sheetsToParse = [
  { sheetName: sheetPhoneErrorLog, columns: [columnMessage, columnRecordedDisplayTime] }
];

// Read files in Directory
fs.readdir(inquisitoDataExcelFilePath, (err, files) => {
  if (err) {
    console.error('Error reading directory:', err);
    return;
  }

  // Filter Excel Files
  const excelFiles = files.filter(file => path.extname(file).toLowerCase() === fileType);

  //Iterate through each excel file
  excelFiles.forEach(excelFile => {
    const excelFilePath = path.join(inquisitoDataExcelFilePath, excelFile);

    // Extract file name without extension
    const fileNameWithoutExtension = path.parse(excelFile).name;

    // Split filewithoutExension by underscores
    const fileNameParts = fileNameWithoutExtension.split('_');

    // find lower bound date and time
    const firstPart = fileNameParts.slice(0, 2).join('.');
    const firstDateAndTime = firstPart.split(".");
    const firstDatePart = firstDateAndTime[0];
    const firstFormattedDate = `${firstDatePart.slice(4)}-${firstDatePart.slice(0, 2)}-${firstDatePart.slice(2, 4)}`;
    const firstTimePart = firstDateAndTime[1];
    const firstFormattedTime = `${firstTimePart.slice(0, 2)}:${firstTimePart.slice(2)}:00`;
    lowerBound = `${firstFormattedDate} ${firstFormattedTime}`;
    // console.log(lowerBound);

    // find upper bound date and time
    const secondPart = fileNameParts.slice(2).join('.');
    const secondDateAndTime = secondPart.split(".");
    const secondDatePart = secondDateAndTime[0];
    const secondFormattedDate = `${secondDatePart.slice(4)}-${secondDatePart.slice(0, 2)}-${secondDatePart.slice(2, 4)}`;
    const secondTimePart = secondDateAndTime[1];
    const secondFormattedTime = `${secondTimePart.slice(0, 2)}:${secondTimePart.slice(2)}:00`;
    upperBound = `${secondFormattedDate} ${secondFormattedTime}`;
    // console.log(upperBound);

    // Read the Excel file
    const inquisitoDataExcel = XLSX.readFile(excelFilePath);

    let capturedEGVs = 0; // Counter to track the number of rows printed

    // Iterate through each sheet
    sheetsToParse.forEach(sheetInfo => {
      const { sheetName, columns } = sheetInfo;

      // Check if sheet exists in workbook
      if (inquisitoDataExcel.Sheets.hasOwnProperty(sheetName)) {
        // Get the data from the sheet
        const sheetData = XLSX.utils.sheet_to_json(inquisitoDataExcel.Sheets[sheetName], { header: 1 });

        // Find column indices in the header row
        const headerRow = sheetData[0];

        // Extract specific columns from each row
        const parsedData = sheetData.slice(1).map(row => {
          const rowData = {};
          columns.forEach(col => {
            const columnIndex = headerRow.indexOf(col);

            // Check if the current column is "Recorded Display Time"
            if (col === columnRecordedDisplayTime) {
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
          const rowRecordedDisplayTime = row[columnRecordedDisplayTime];
          return (
            rowRecordedDisplayTime >= lowerBound &&
            rowRecordedDisplayTime <= upperBound &&
            (row[columnMessage].includes(messageCGMReadings) ||
              row[columnMessage].includes(messageReadEGV))
          );
        });

        // Print the values between the parameters
        filteredData.forEach(row => {
          // console.log(row[columnRecordedDisplayTime], row[columnMessage]);
          capturedEGVs++;
        });

        console.log('=============================================');
        console.log('File: ', fileNameWithoutExtension);
        console.log(`Number of captured EGVs: ${capturedEGVs}`);
      }
    })
  })
})