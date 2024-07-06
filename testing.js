const xlsx = require('xlsx');
const fs = require('fs');
const path = require('path');

// Function to extract all string values from a specified row, excluding values containing "DUMMY"
function extractRowValues(worksheet, rowNumber) {
  const rowValues = [];
  const range = xlsx.utils.decode_range(worksheet['!ref']);

  for (let col = range.s.c; col <= range.e.c; col++) {
    const cellAddress = xlsx.utils.encode_cell({ c: col, r: rowNumber });
    const cell = worksheet[cellAddress];
    const value = cell ? cell.v : null;

    if (value !== null && typeof value === 'string' && !value.includes('DUMMY')) {
      rowValues.push({ value, col });
    }
  }

  return rowValues;
}

// Function to extract data from a specific column starting from a given row for a certain number of rows
function extractColumnData(worksheet, col, startRow, numRows) {
  const columnData = [];
  for (let row = startRow; row < startRow + numRows; row++) {
    const cellAddress = xlsx.utils.encode_cell({ c: col, r: row });
    const cell = worksheet[cellAddress];
    const value = cell ? cell.v : 'N/A';
    columnData.push(value);
  }
  return columnData;
}

// Function to generate time intervals for 1440 rows
function generateTimeIntervals(numRows) {
  const timeIntervals = [];
  const startDate = new Date(2000, 0, 1, 0, 0, 0);

  for (let i = 0; i < numRows; i++) {
    const endDate = new Date(startDate);
    endDate.setMinutes(startDate.getMinutes() + 1);
    const startTime = startDate.toTimeString().slice(0, 5);
    const endTime = endDate.toTimeString().slice(0, 5);
    timeIntervals.push(`${startTime} - ${endTime}`);
    startDate.setMinutes(startDate.getMinutes() + 1);
  }

  return timeIntervals;
}

// Initializing output data array to store data from all files
let outputData = [['Zone', 'Name of Station', 'Date', 'Time', 'SCADA Tag', 'Data']];

function processFile(inputFilePath, outputFolder, batchIndex, zone) {
  // Read the Excel file
  const workbook = xlsx.readFile(inputFilePath);
  const sheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];

  // Extract string values from the third row
  const thirdRowValues = extractRowValues(worksheet, 2);

  // Log the SCADA Tags found in the third row
  console.log(`String values in the third row (excluding values containing "DUMMY") from ${inputFilePath}:`, thirdRowValues);

  // Extraction of Name of the Station from cell A1
  const nameOfStation = worksheet['A1'] ? worksheet['A1'].v : 'N/A';

  // Extraction of Date from cell B6
  let date = worksheet['B6'] ? worksheet['B6'].v : 'N/A';
  if (!isNaN(date)) {
    date = xlsx.SSF.format('yyyy-mm-dd', date);
  }

  // Generating time intervals for 1440 rows
  const timeIntervals = generateTimeIntervals(1440);

  // Function to add data to the output
  function addDataToOutput(rowValues, startRow) {
    rowValues.forEach(({ value, col }) => {
      const columnData = extractColumnData(worksheet, col, startRow, 1440);
      columnData.forEach((dataValue, index) => {
        outputData.push([zone, nameOfStation, date, timeIntervals[index], value, dataValue]);
      });
    });
  }

  // Add the data from the third row and 1440 rows of SCADA tag data
  addDataToOutput(thirdRowValues, 6); 

  // Create a new workbook and worksheet for the output data
  const outputWorkbook = xlsx.utils.book_new();
  const outputWorksheet = xlsx.utils.aoa_to_sheet(outputData);
  xlsx.utils.book_append_sheet(outputWorkbook, outputWorksheet, `Batch_${batchIndex + 1}`);

  if (!fs.existsSync(outputFolder)) {
    fs.mkdirSync(outputFolder, { recursive: true });
  }

  // Generate output file path
  const outputFilePath = path.join(outputFolder, `Batch_${batchIndex + 1}_Extracted_SCADA_Tag_Data.xlsx`);
  xlsx.writeFile(outputWorkbook, outputFilePath);

  console.log(`Batch ${batchIndex + 1} SCADA Tag data has been successfully written to ${outputFilePath}`);
}

// Main function to process files in multiple folders
function processFolders(inputFolders, outputFolders) {
  inputFolders.forEach((inputFolder, folderIndex) => {
    const outputFolder = outputFolders[folderIndex];
    const zone = path.basename(inputFolder).split('_')[0]; // Extracting zone name from the folder name

    fs.readdir(inputFolder, (err, files) => {
      if (err) {
        console.error(`Error reading the folder ${inputFolder}:`, err);
        return;
      }

      // Splitting the files into batches of 50
      const batchSize = 50;
      let batchCount = Math.ceil(files.length / batchSize);

      for (let batchIndex = 0; batchIndex < batchCount; batchIndex++) {
        outputData = [['Zone', 'Name of Station', 'Date', 'Time', 'SCADA Tag', 'Data']]; 

        const start = batchIndex * batchSize;
        const end = Math.min(start + batchSize, files.length);
        const filesToProcess = files.slice(start, end);

        filesToProcess.forEach(file => {
          if (path.extname(file) === '.xls' || path.extname(file) === '.xlsx') {
            const inputFilePath = path.join(inputFolder, file);
            processFile(inputFilePath, outputFolder, batchIndex, zone);
          }
        });
      }
    });
  });
}

// Define input and output folders
const inputFolders = ['./BGM_testing', './BGK_testing'];
const outputFolders = ['./output_BGM', './output_BGK'];

// Processing the folders
processFolders(inputFolders, outputFolders);
