const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const baseDirectory = __dirname;
const fileName = 'Assignment_Timecard.xlsx';
const excelFilePath = path.join(baseDirectory, fileName);

// console.log(excelFilePath);

const workbook = XLSX.readFile(excelFilePath);

const sheetName = workbook.SheetNames[0];
const sheet = workbook.Sheets[sheetName];

const data = XLSX.utils.sheet_to_json(sheet);

// Function to convert Excel date to JavaScript date
const excelDateToJSDate = (excelDate) => {
  return new Date((excelDate - (25567 + 1)) * 86400 * 1000);
};

// a) Who has worked for 7 consecutive days
const findEmployeesWorkingConsecutiveDays = () => {
  const sortedData = data.sort((a, b) => a['Time Out'] - b['Time Out']);
  const consecutiveDaysThreshold = 7;

  for (let i = 0; i < sortedData.length - consecutiveDaysThreshold + 1; i++) {
    const consecutiveDays = sortedData.slice(i, i + consecutiveDaysThreshold);

    const workedDays = new Set(consecutiveDays.map(entry => {
      const date = excelDateToJSDate(entry['Time Out']);
      return `${date.getMonth() + 1}-${date.getDate()}-${date.getFullYear()}`;
    }));

    if (workedDays.size >= consecutiveDaysThreshold) {
      console.log(`Employee ${consecutiveDays[0]['Employee Name']} (${consecutiveDays[0]['Position ID']}) has worked for 7 consecutive days.`);
    }
  }
};

// b) Who have less than 10 hours of time between shifts but greater than 1 hour
const findEmployeesWithShortBreaks = () => {
  const shortBreaksThreshold = 10 * 60 * 60; // 10 hours in seconds
  const minimumBreak = 1 * 60 * 60; // 1 hour in seconds

  for (let i = 1; i < data.length; i++) {
    const timeBetweenShifts = data[i]['Time Out'] - data[i - 1]['Time Out'];

    if (timeBetweenShifts < shortBreaksThreshold && timeBetweenShifts > minimumBreak) {
      console.log(`Employee ${data[i]['Employee Name']} (${data[i]['Position ID']}) has a short break between shifts.`);
    }
  }
};

// c) Who has worked for more than 14 hours in a single shift
const findEmployeesLongShifts = () => {
  const longShiftThreshold = 14 * 60 * 60; // 14 hours in seconds

  for (const entry of data) {
    if (entry['Time'] > longShiftThreshold) {
      console.log(`Employee ${entry['Employee Name']} (${entry['Position ID']}) has worked for more than 14 hours in a single shift.`);
    }
  }
};

// Execute the functions
findEmployeesWorkingConsecutiveDays();
findEmployeesWithShortBreaks();
findEmployeesLongShifts();
