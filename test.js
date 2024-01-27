const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const baseDirectory = __dirname;
const fileName = 'Assignment_Timecard.xlsx';
const excelFilePath = path.join(baseDirectory, fileName);

const workbook = XLSX.readFile(excelFilePath);

const sheetName = workbook.SheetNames[0];
const sheet = workbook.Sheets[sheetName];

const data = XLSX.utils.sheet_to_json(sheet);

const excelDateToJSDate = (excelDate) => {
    return new Date((excelDate - (25567 + 1)) * 86400 * 1000);
};

const results = [];
const uniqueEmployees = new Set(); // To track unique employees

const findEmployeesWorkingConsecutiveDays = () => {
    for (let i = 0; i < data.length; i++) {
        const startTime = excelDateToJSDate(data[i]['Pay Cycle Start Date']);
        const endTime = excelDateToJSDate(data[i]['Pay Cycle End Date']);

        const daysDifference = (endTime - startTime) / (1000 * 60 * 60 * 24);

        if (daysDifference >= 7 && !uniqueEmployees.has(data[i]['Employee Name'])) {
            results.push({
                employeeName: data[i]['Employee Name'],
                positionID: data[i]['Position ID'],
                message: `Employee ${data[i]['Employee Name']} (${data[i]['Position ID']}) has worked for 7 consecutive days.`,
            });
            uniqueEmployees.add(data[i]['Employee Name']);
        }
    }
};

findEmployeesWorkingConsecutiveDays();

const txtContent = results.map(result => `${result.message}\n`).join('');

const txtFilePath = 'sevendays.txt';
fs.writeFileSync(txtFilePath, txtContent, 'utf-8');

console.log(`Results written to ${txtFilePath}`);

//////////////////////////////

const res = [];
const uniqueEmp = new Set();

const findEmployeesWithShortBreaks = () => {
    for (let i = 1; i < data.length; i++) {
        // const timeBetweenShifts = data[i]['Time Out'] - data[i]['Time'];
        // const hoursBetweenShifts = timeBetweenShifts / 3600;
        // console.log(timeBetweenShifts);

        const take = data[i]['Timecard Hours (as Time)'];
        const [hours, minutes] = take.split(':');
        
        const hoursNumeric = parseInt(hours, 10);
        const minutesNumeric = parseInt(minutes, 10);

        const totalHours = hoursNumeric + minutesNumeric / 60;
        

        if ((totalHours > 1 && totalHours < 10)  && !uniqueEmp.has(data[i]['Employee Name'])) {
            res.push({
                employeeName: data[i]['Employee Name'],
                positionID: data[i]['Position ID'],
                message: `Employee ${data[i]['Employee Name']} (${data[i]['Position ID']}) has a short break between shifts.`,
            });
            uniqueEmp.add(data[i]['Employee Name']);
        }
    }
};

findEmployeesWithShortBreaks();

// console.log(res);
const txtCont = res.map(result => `${result.message}\n`).join('');
// console.log(txtCont);

fs.writeFileSync('shortbreaks.txt', txtCont, 'utf-8');

const resNew = [];
const uniqueEmpNew = new Set();

const findEmployeesWith14hr = () => {
    for (let i = 1; i < data.length; i++) {
        const take = data[i]['Timecard Hours (as Time)'];
        const [hours, minutes] = take.split(':');
        
        const hoursNumeric = parseInt(hours, 10);
        const minutesNumeric = parseInt(minutes, 10);

        const totalHours = hoursNumeric + minutesNumeric / 60;
        // console.log(totalHours);
        

        if (totalHours > 14  && !uniqueEmpNew.has(data[i]['Employee Name'])) {
            resNew.push({
                employeeName: data[i]['Employee Name'],
                positionID: data[i]['Position ID'],
                message: `Employee ${data[i]['Employee Name']} (${data[i]['Position ID']}) has worked more than 14 hrs.`,
            });
            uniqueEmpNew.add(data[i]['Employee Name']);
        }
    }
};

findEmployeesWith14hr();

const txtContNew = resNew.map(result => `${result.message}\n`).join('');
// console.log(txtCont);

fs.writeFileSync('morethan14hrs.txt', txtContNew, 'utf-8');