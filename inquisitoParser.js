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
let dataRssiValues = [];
let rssiCount = 0;

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
    rssiCount = 0;
    dataRssiValues = [];

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
    lowerBound = new Date(`${firstFormattedDate} ${firstFormattedTime}`).toISOString();
    console.log('Lower bound UTC: ', lowerBound);

    // find upper bound date and time
    const secondPart = fileNameParts.slice(2).join('.');
    const secondDateAndTime = secondPart.split(".");
    const secondDatePart = secondDateAndTime[0];
    const secondFormattedDate = `${secondDatePart.slice(4)}-${secondDatePart.slice(0, 2)}-${secondDatePart.slice(2, 4)}`;
    const secondTimePart = secondDateAndTime[1];
    const secondFormattedTime = `${secondTimePart.slice(0, 2)}:${secondTimePart.slice(2)}:00`;
    upperBound = new Date(`${secondFormattedDate} ${secondFormattedTime}`).toISOString();
    console.log('Upper bound UTC: ', upperBound);
    console.log('=============================================');

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
              const formattedDate = displayTimeParts.join(".");
              const dateObject = new Date(formattedDate);
              rowData[col] = dateObject.toISOString();
            } else {
              rowData[col] = row[columnIndex];
            }
          });
          return rowData;
        });

        const filteredData = parsedData.filter(row => {
          const rowRecordedDisplayTime = row[columnRecordedDisplayTime];
          const messageIncludesCGMReadings = row[columnMessage].includes(messageCGMReadings);
          const messageIncludesReadEGV = row[columnMessage].includes(messageReadEGV);
          const messageIncludesOnReadRssi = row[columnMessage].includes(messageOnReadRssi);

          if (
            rowRecordedDisplayTime >= lowerBound &&
            rowRecordedDisplayTime <= upperBound &&
            (messageIncludesCGMReadings || messageIncludesReadEGV || messageIncludesOnReadRssi)
          ) {
            // Print the values between the parameters
            // console.log(row[columnRecordedDisplayTime], row[columnMessage]);

            // Count captured EGVs only for CGMReadings and ReadEGV
            if (messageIncludesCGMReadings || messageIncludesReadEGV) {
              capturedEGVs++;
            }

            // Check if the message includes "onReadRssi" and print only the numbers in absolute value
            if (messageIncludesOnReadRssi) {
              const rssiMatch = row[columnMessage].match(/onReadRssi: (-?\d+)/);
              if (rssiMatch) {
                const rssiValue = parseInt(rssiMatch[1], 10);
                // console.log(`${Math.abs(rssiValue)}`);
                // rssiCount++;
              }
            }

            else {
              // If "onReadRssi" is not present, navigate to "sheetPhoneUserActivity" file
              const userActivitySheet = inquisitoDataExcel.Sheets[sheetPhoneUserActivity];
              if (userActivitySheet) {
                const userActivityData = XLSX.utils.sheet_to_json(userActivitySheet, { header: 1 });

                // Find column indices in the header row
                const headerRowUserActivity = userActivityData[0];
                const columnIndexRecordedDisplayTime = headerRowUserActivity ? headerRowUserActivity.indexOf(columnRecordedDisplayTime) : -1;
                const columnIndexData = headerRowUserActivity ? headerRowUserActivity.indexOf(columnData) : -1;

                if (columnIndexRecordedDisplayTime !== -1 && columnIndexData !== -1) {
                  // Extract values from the "columnData" column associated with "columnRecordedDisplayTime" that include "dataRssi"
                  dataRssiValues = userActivityData.slice(1)
                    .filter(row => {
                      const rowRecordedDisplayTime = row[columnIndexRecordedDisplayTime];
                      const rowData = row[columnIndexData];
                      if (rowData === undefined) {
                        console.log(`Undefined value in columnData for Recorded Display Time ${rowRecordedDisplayTime}`);
                        return false;
                      }

                      // Convert recorded display time to UTC
                      const utcRecordedDisplayTime = new Date(rowRecordedDisplayTime).toISOString();

                      return (
                        utcRecordedDisplayTime >= lowerBound &&
                        utcRecordedDisplayTime <= upperBound &&
                        rowData && rowData.includes(dataRssi)
                      );
                    })
                    .map(row => {
                      try {
                        const jsonData = JSON.parse(row[columnIndexData]);
                        return jsonData?.['Rssi'];
                      } catch (error) {
                        console.error(`Error parsing JSON in row with Recorded Display Time ${row[columnIndexRecordedDisplayTime]}`);
                        return null; // or handle it according to your needs
                      }
                    })
                    .filter(value => value !== null && value !== undefined); // Remove null and undefined values
                  
                }
                else {
                  console.log(`Column ${columnRecordedDisplayTime} or ${columnData} not found in sheetPhoneUserActivity`);
                }
                rssiCount++;
              }
              else {
                console.log(`Sheet ${sheetPhoneUserActivity} not found in the file`);
              }
            }
          }
        });

        // Print the values between the parameters
        filteredData.forEach(row => {
          // console.log(row[columnRecordedDisplayTime], row[columnMessage]);
          capturedEGVs++;
        });
      }
    })

    // Print the values of "Rssi" that include "dataRssi"
    dataRssiValues
      .filter(value => value !== null && value !== undefined)
      .forEach(value => {
        if (typeof value === 'number') {
          console.log(Math.abs(value));
        } else {
          console.error(`Invalid value for Rssi: ${value}`);
        }
      });

    console.log(`Number of RSSI values: ${rssiCount}`);
    console.log('=============================================');
    console.log(`Number of captured EGVs: ${capturedEGVs}`);
    console.log('File: ', fileNameWithoutExtension);
    console.log('=============================================');
  })
})