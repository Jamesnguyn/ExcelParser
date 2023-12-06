const XLSX = require('xlsx');
const fs = require('fs');

const privateDataExcelFilePath = 'privateData\\SampleData.xlsx';

const privateDataExcel = XLSX.readFile(privateDataExcelFilePath);

privateDataExcel.SheetNames.forEach(sheetName => {
  const sheetData = XLSX.utils.sheet_to_json(privateDataExcel.Sheets[sheetName]);

  console.log(`Dta from sheet "{sheetName}":`, sheetData);
})