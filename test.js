
// --- PACKAGE ---
const XLSX = require('xlsx');
const { Builder, By, Key, until } = require('selenium-webdriver');
const dotenv = require("dotenv").config();
// ---------------

// envoyer varibale (ex. node app.js 1)
const idx = parseInt(process.argv[2], 10);
// console.log(idx);

// read file excel
const workbook = XLSX.readFile("list.xlsx", {
    type: 'binary',
    cellDates: true,
    cellNF: false,
    cellText: false
});

// get sheet names
const sheet_name_list = workbook.SheetNames;
// finally JSON
const dataJSON = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]]);

// personnel with witch we will work
const unitData = dataJSON[idx - 1];


/* TEST 
1. Difference de date (resultat en millesecound)
2. pour voir emploi Militaire du rang fin de CDD ou Militaire sous-officier fin de CDD
3. 00 + NID[10]
*/
// 1
console.log(Math.round((unitData.date_FDC - unitData.date_eng) / (24 * 3600 * 1000) / 365))

// 2 
const emploi = (
    unitData.grade.includes("Major") ||
    unitData.grade.includes("Adjudant") ||
    unitData.grade.includes("Sergent")) ?
    "Militaire sous-officier fin de CDD" :
    "Militaire du rang fin de CDD";
console

// 3
console.log("00" + unitData.nid.slice(0, 10))