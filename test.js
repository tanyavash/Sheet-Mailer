const xlsx = require('xlsx');
const fs = require('fs');
const fileContents = fs.readFileSync('./input.xlsx');
const wb = xlsx.read(fileContents);
//console.log(wb.SheetNames);
const ws = wb.Sheets["Sheet1"];
console.log(ws);

const data = xlsx.utils.sheet_to_json(ws);
//console.log(data);

fs.writeFileSync('./jsondata.json', JSON.stringify(data, null, 2));