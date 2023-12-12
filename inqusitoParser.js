const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

// Specify the directory path
const inquisitoDataExcelFilePath = 'inquisitoData';

const fileType = '.xlsx';

const sheetPhoneErrorLog = 'PhoneErrorLog';

const columnMessage = 'Message';
const columnRecordedDisplayTime = 'RecordedDisplayTime';

const sheetsToParse = [
    { sheetName: sheetPhoneErrorLog, columns: [columnMessage, columnRecordedDisplayTime]}
];

// Read files in Directory
fs.readdir(inquisitoDataExcelFilePath, (err, files) => {
    if(err){
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
        
    })
})