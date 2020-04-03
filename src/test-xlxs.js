const XLSX = require('xlsx');
const fs = require('fs');

var buf = fs.readFileSync("./sample/1.xlsx");
var xls = XLSX.read(buf, { type: 'buffer' });

var sheet1 = xls.Workbook.Sheets[0];
var sheetName = sheet1.name;
var rows = xls.Sheets[sheetName];

console.log(sheet1);
console.log(rows);

rows.K2.f = 'I2-J2+30'

console.log(rows);


XLSX.writeFile(xls, "./sample/2.xlsx")