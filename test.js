
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
    cellText: false,
});

// get sheet names
const sheet_name_list = workbook.SheetNames;
// finally JSON with defval
const dataJSON = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]], { defval: "" });

// personnel with witch we will work
const unitData = dataJSON[idx - 1];

//function pour obtenir la date en string
function getDataString(numData) {
    let date = new Date(numData);

    // rajouté un jour
    date.setDate(date.getDate() + 1);

    let options = {
        year: "numeric",
        month: "numeric",
        day: "numeric",
    };

    // console.log(Intl.DateTimeFormat('fr-FR', options).format(date));
    return Intl.DateTimeFormat("fr-FR", options).format(date);
}

/* TEST
1. Difference de date (resultat en millesecound)
2. pour voir emploi Militaire du rang fin de CDD ou Militaire sous-officier fin de CDD
3. 00 + NID[10]
*/
// 1
// console.log(Math.round((unitData.date_FDC - unitData.date_eng) / (24 * 3600 * 1000) / 365))

// 2
const emploi = (
    unitData.grade.includes("Major") ||
    unitData.grade.includes("Adjudant") ||
    unitData.grade.includes("Sergent")) ?
    "Militaire sous-officier fin de CDD" :
    "Militaire du rang fin de CDD";

// 3
// console.log("00" + unitData.nid.slice(0, 10))

// unitData.find(el => el == "01/02/2018")

function getTwoDigit(number) {
    return (number < 10 ? '0' : '') + number
}

// exampl!
// var d = new Date();
// console.log(getTwoDigit(d.getDate()) + "/" + getTwoDigit(d.getMonth() + 1) + "/" + d.getFullYear())


// creation massiv de deux nivau
const arrUnitData = Object.entries(unitData)

// dernier année date de pay
// console.log("01/" + getTwoDigit(unitData.date_DJTP.getMonth() + 2) + "/" + (unitData.date_DJTP.getFullYear() - 3))
// console.log(arrUnitData[237])

// console.log(unitData.date_DJTP.getMonth())
// first mounth
/* console.log("01/" + getTwoDigit(unitData.date_DJTP.getMonth()) + "/" + (unitData.date_DJTP.getFullYear() - 3))
console.log(sbm + "1")
console.log(arrUnitData[237]) // dernier sbm */
// console.log(arrUnitData[122]) // sbm 12

/* console.log("01/" + getTwoDigit(unitData.date_DJTP.getMonth()) + "/" + (unitData.date_DJTP.getFullYear() - 3))
console.log(sbm + "1")
console.log(arrUnitData[237]) // dernier sbm */


/* 
CREATION DE ARRAY SOLD
Composition: 
    [
        0. Name de tag ou il faut inserer Solde de base mensuel brute,
        1. Name de tag ou il faut inserer Date de paie salaire
        2. Date de paie
        3. Solde de base brut mensuels
        4. Name de tag nombreTempsTravaillesSalaire_<n>
        5. Name de tag nombreTempsNonPayesSalaire_<n>
    ]
*/
const sbm = "salaireBrutMensuel_"
const ddp = "datePaieSalaire_"
const ntts = "nombreTempsTravaillesSalaire_"
const ntnps = "nombreTempsNonPayesSalaire_"
let arrSold = []
let buffer = []

// creation dernier date si année suivant ou non 
const year = (unitData.date_DJTP.getMonth() + 1) === 12 ? (unitData.date_DJTP.getFullYear() - 2) : (unitData.date_DJTP.getFullYear() - 3)
const mounth = (unitData.date_DJTP.getMonth() + 1) === 12 ? "01" : getTwoDigit(unitData.date_DJTP.getMonth() + 2)
buffer.push([
    (sbm + "1"),
    (ddp + "1"),
    ("01/" + mounth + "/" + year),
    arrUnitData[237][1],
    (ntts + "1"),
    (ntnps + "1")
])

let k = 0; // pas dans le massiv
console.log(unitData.date_DJTP.getMonth())
for (let i = 2; i <= 36 + unitData.date_DJTP.getMonth(); i++) {

    if (i < (12 - unitData.date_DJTP.getMonth() + 1)) {
        // console.log(sbm + i)
        buffer.push([
            (sbm + i),
            (ddp + i),
            getDataString(arrUnitData[233 - k][1]),
            arrUnitData[237 - k][1],
            (ntts + i),
            (ntnps + i)
        ])
        k += 5
    }
}

// console.log(buffer)
buffer.forEach(el => arrSold.push(el)) // ajouté 1 annéé
// console.log(arrSold)
// console.log(k)

buffer = [] // netoyage buffer

// console.log(buffer)
for (let i = 13; i <= 36; i++) {
    // console.log(sbm + i)
    // console.log(arrUnitData[237 - k], 237 - k)
    buffer.push([
        (sbm + i),
        (ddp + i),
        getDataString(arrUnitData[233 - k][1]),
        arrUnitData[237 - k][1],
        (ntts + i),
        (ntnps + i)
    ])
    if ((237 - k) <= 122) {
        k += 7
    } else {
        k += 5
    }
}

buffer.forEach(el => arrSold.push(el)) // ajouté 2 et 4 année

buffer = [] // netoyage buffer

// console.log(arrSold)

let iw = 36
while (iw < (36 + unitData.date_DJTP.getMonth())) {
    iw++
    // console.log(sbm + iw)
    // console.log(arrUnitData[237 - k], 237 - k)
    buffer.push([
        (sbm + iw),
        (ddp + iw),
        getDataString(arrUnitData[233 - k][1]),
        arrUnitData[237 - k][1],
        (ntts + iw),
        (ntnps + iw)
    ])
    if ((237 - k) <= 122) {
        k += 7
    } else {
        k += 5
    }
}

if (buffer.length !== 0) { // ajouté dernier anné si il y a
    buffer.forEach(el => arrSold.push(el))
}

/* arrSold pret! */
// console.log(arrSold)

/* TEST pour ETAP 4 */
// console.log(getDataString(arrUnitData[240][1]))
// console.log(arrUnitData[246][1].toFixed(2))
// console.log(arrUnitData[244][1].toFixed(2))

/*
CREATION DE ARRAY SOLD - END
Arr solde est pret!
Ex:
  [
    'salaireBrutMensuel_36',
    'datePaieSalaire_36',
    '01/01/2021',
    2286.78,
    'nombreTempsTravaillesSalaire_36',
    'nombreTempsNonPayesSalaire_36'
  ],
  [
    'salaireBrutMensuel_37',
    'datePaieSalaire_37',
    '01/02/2021',
    2286.78,
    'nombreTempsTravaillesSalaire_37',
    'nombreTempsNonPayesSalaire_37'
  ]
*/

// for (let i = 2; i <= 12 - unitData.date_DJTP.getMonth(); i++) {
//     console.log(sbm + i.toString());
//     let idx = (i - 2) * 5;
//     console.log(arrUnitData[237 - idx]);
//     if (i === (12 - unitData.date_DJTP.getMonth())) {
//         console.log("true 1 ")
//         let i2 = i
//         i += (unitData.date_DJTP.getMonth() + 1)

//         while (i <= 26) {
//             idx = (i2 - 1) * 5;
//             console.log(sbm + i.toString());
//             console.log(arrUnitData[237 - idx]);
//             i++
//             i2++
//         }
//         if (i === 27) {
//             console.log("true 2")
//             let i3 = i2
//             // i += (unitData.date_DJTP.getMonth() + 1)
//             // idx = ((i3 - 1) * 5) + 2;
//             idx = ((i3 - 1) * 5);
//             while (i <= (35 + unitData.date_DJTP.getMonth() + 1)) {

//                 console.log(237 - idx)
//                 console.log(sbm + i.toString());
//                 console.log(arrUnitData[237 - idx]);
//                 if (i >= 28) {
//                     idx = idx + 7;
//                 } else {
//                     idx = idx + 5;
//                 }
//                 i++
//                 i3++
//             }

//         }
//     }
// }


// console.log(unitData)