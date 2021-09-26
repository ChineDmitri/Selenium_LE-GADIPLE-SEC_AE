
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
// finally JSON with defval
const dataJSON = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]], { defval: "" });

// personnel with witch we will work (JSON)
const unitData = dataJSON[idx - 1];

// personnel with witch we will work (TWO-DIMENSIONAL ARRAY)
const arrUnitData = Object.entries(unitData)

// pour voir emploi Militaire du rang fin de CDD ou Militaire sous-officier fin de CDD
const emploi = (
    unitData.grade.includes("Major") ||
    unitData.grade.includes("Adjudant") ||
    unitData.grade.includes("Sergent")) ?
    "Militaire sous-officier fin de CDD" :
    "Militaire du rang fin de CDD"

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

// function pour obtenir ZERO (0) devant la date (Mounth or Day)
function getTwoDigit(number) {
    return (number < 10 ? '0' : '') + number
}

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

/*
START selenium-webdriver
TODO:
1. Creation session et auth
2. Acceder pour tout les attestation
3. Entrer en "Creation Attestation"
4. ETAP 1
5. ETAP 2
6. ETAP 3
7. ETAP 4
8. ETAP 5
*/
const driver = new Builder().forBrowser('firefox').build();

/* 1. Creation session et auth. */
// driver.get('file:///C:/Users/mrdim/Desktop/2.html');
// driver.findElement(By.id('footer_tc_privacy_button_2')).click();
driver.get('https://gestion.pole-emploi.fr/espaceemployeur/authentification/authentification/');
driver.findElement(By.name('identifiant')).sendKeys(process.env.identifiant)
driver.findElement(By.name('codeAcces')).sendKeys(process.env.pass)
driver.findElement(By.name('departement')).sendKeys(process.env.dep);
driver.findElement(By.id('footer_tc_privacy_button_2')).click();
driver.findElement(By.id('boutonValider')).click();

/* 2. Acceder pour tout les attestation */
driver.wait(until.titleIs('Espace Employeur - Votre compte'))
    .then(() => {
        driver.findElement(By.css('[href^="/employeur/gdc/compte/compte.liensaisieattestationrgbis"]')).click();
    })

/* 3. Entrer en "Creation Attestation" */
driver.wait(until.titleIs('Attestations en ligne - Pôle emploi'))
    .then(() => {
        driver.findElement(By.css('[href^="/entreprise-attestation/accueil.blocinformationsetablissement:creerattestation"]')).click();
    })

/* 4. ETAP 1
pour passe en etape suivant il faut clique manuellement */
driver.wait(until.titleIs('1 Salarié - Saisie attestation en ligne - Pôle emploi'))
    .then(() => {
        // checkbox
        driver.findElement(By.css('[for="civiliteMonsieur"]')).click();
        driver.findElement(By.css('[for="statutCategorielRetraiteComplementaireNon"]')).click();

        // inputs
        driver.findElement(By.id('prenom')).sendKeys(unitData.prenom);
        driver.findElement(By.id('etape1Nom')).sendKeys(unitData.nom);
        driver.findElement(By.id('dateNaissance')).sendKeys(getDataString(unitData.dateNaissance));
        // driver.findElement(By.id('paysNaissance')).sendKeys(data.paysNaissance)
        driver.findElement(By.id('lieuNaissance')).sendKeys(unitData.lieuNaissance);
        // driver.findElement(By.id('ressortissant')).sendKeys("Reste du Monde")
        driver.findElement(By.id('nir')).sendKeys(unitData.ss);
        driver.findElement(By.id('adresseSalarie')).sendKeys(unitData.adress);

        // Profession intermédiaire (technicien, contremaître, agent de maîtrise, clergé)
        driver.wait(
            driver.findElement(By.id('statutSalarie')).click())
            .then(() => {
                driver.findElement(By.css('[value="05"]')).click();
            })

        // IRCANTEC
        driver.wait(
            driver.findElement(By.id('categorieRetraiteComplementaire')).click())
            .then(() => {
                driver.findElement(By.css('[value="IRCANTEC"]')).click();
            })

        // Organisme de sécurité sociale
        driver.wait(driver.findElement(By.id('organismeRetraiteComplementaire')).click())
            .then(() => {
                driver.findElement(By.css('[value="200"]')).click();
            })

    })

/* 5. ETAP 2 */
driver.wait(until.titleIs('2 Emploi - Saisie attestation en ligne - Pôle emploi'))
    .then(() => {

        driver.findElement(By.name('numeroContrat')).sendKeys("000000");

        driver.findElement(By.name('dateDebutPeriodeEmploi')).sendKeys(getDataString(unitData.date_eng));

        driver.wait(until.elementsLocated(By.xpath('/html/body/main/div[1]/form/div[2]/div/div[3]/div/div[6]/section[2]/section[1]/div[3]/div[2]/div/input')))
            .then((result) => {
                // console.log(result);
                driver.wait(
                    driver.findElement(By.name('motifRecours')).click())
                    .then(() => {
                        driver.findElement(By.xpath('/html/body/main/div[1]/form/div[2]/div/div[3]/div/div[6]/section[2]/section[1]/div[1]/div[2]/select/option[7]')).click();
                    })

                driver.findElement(By.name('dateFinPrevisionnelle')).sendKeys(getDataString(unitData.date_FDC));

                driver.findElement(By.name('anciennete')).sendKeys(
                    Math.round((unitData.date_FDC - unitData.date_eng) / (24 * 3600 * 1000) / 365)
                )

                driver.wait(
                    driver.findElement(By.name('uniteHorairesAnciennete')).click())
                    .then(() => {
                        driver.findElement(By.xpath('/html/body/main/div[1]/form/div[2]/div/div[3]/div/div[6]/section[2]/section[1]/div[4]/div/fieldset/div/div/select/option[3]')).click();
                    })

                driver.findElement(By.name('dernierEmploiTenu')).sendKeys(emploi);

                driver.wait(
                    driver.findElement(By.name('emploisMultiples')).click())
                    .then(() => {
                        driver.findElement(By.xpath('/html/body/main/div[1]/form/div[2]/div/div[3]/div/div[6]/section[2]/section[1]/div[5]/div[2]/div/select/option[2]')).click();
                    })

                driver.findElement(By.name('horairesEntreprise')).sendKeys(151);

                driver.findElement(By.name('horairesSalarie')).sendKeys(151);

                driver.findElement(By.name('adresseLieuTravail')).sendKeys("580 Route de la Légion");

                driver.wait(
                    driver.findElement(By.name('natureLieuTravail')).click())
                    .then(() => {
                        driver.findElement(By.xpath('/html/body/main/div[1]/form/div[2]/div/div[3]/div/div[6]/section[2]/section[1]/div[12]/select/option[2]')).click();
                    })

                driver.wait(
                    driver.findElement(By.name('motifRupture')).click())
                    .then(() => {
                        driver.findElement(By.xpath('/html/body/main/div[1]/form/div[2]/div/div[3]/div/div[6]/section[2]/section[4]/div[2]/div[1]/div/select/option[6]')).click();
                    })

                driver.findElement(By.className('form-control input-date datefincontrat')).sendKeys(getDataString(unitData.date_FDC));

                driver.findElement(By.name('dernierJourTravaille')).sendKeys(getDataString(unitData.date_DJTP));

                // !Depreciated "* Préavis"
                // driver.wait(driver.findElement(By.name('typePreavis_1')).click())
                //     .then(() => {
                //         driver.findElement(By.css('[value="90"]')).click();
                //     })

                driver.wait(
                    driver.findElement(By.xpath('/html/body/main/div[1]/form/div[2]/div/div[3]/div/div[6]/section[2]/section[5]/div[2]/div/label')).click())
                    .then(() => {
                        driver.findElement(By.name('numeroConventionGestion')).sendKeys("1110DEFMIL")

                        driver.findElement(By.name('codeAffectationAC')).sendKeys(178011);

                        driver.findElement(By.name('numeroInterneEmployeur')).sendKeys("00" + unitData.nid.slice(0, 10));
                    })

            })

    })

/* 6. ETAP 3 */
driver.wait(until.titleIs('3 Salaires et primes - Saisie attestation en ligne - Pôle emploi'))
    .then(() => {
        console.log("Vous avez 10 seconds pour férmer la Pop-Under ;)")

        setTimeout(() => {

            driver.wait(
                driver.findElement(By.xpath('/html/body/main/div[1]/form/div[2]/div/div[3]/div/div[6]/section[3]/section[1]/div/div[2]/h3/a')).click())
                .then(() => {
                    // console.log("yes 1")
                })
                .catch(() => {
                    console.log("no located")
                })

            driver.wait(
                driver.findElement(By.xpath('/html/body/main/div[1]/form/div[2]/div/div[3]/div/div[6]/section[3]/section[1]/div/div[3]/h3/a')).click())
                .then(() => {
                    // console.log("yes 2")
                })
                .catch(() => {
                    console.log("no located")
                })

            driver.wait(
                driver.findElement(By.xpath('/html/body/main/div[1]/form/div[2]/div/div[3]/div/div[6]/section[3]/section[1]/div/div[4]/h3/a')).click())
                .then(() => {
                    // console.log("yes 3")
                })
                .catch(() => {
                    console.log("no located")
                })

            driver.wait(
                driver.findElement(By.xpath('/html/body/main/div[1]/form/div[2]/div/div[3]/div/div[6]/section[3]/section[1]/div/div[5]/h3/a')).click())
                .then(() => {
                    // console.log("yes 4");


                    for (let i = 0; i < arrSold.length; i++) {
                        driver.findElement(By.name(`${arrSold[i][0]}`)).sendKeys(arrSold[i][3])
                        driver.findElement(By.name(`${arrSold[i][1]}`)).sendKeys(arrSold[i][2])

                        // !Depreciated! ERRCONNECTED trop de request
                        // driver.findElement(By.name(`${arrSold[i][4]}`)).sendKeys(151)
                        // driver.findElement(By.name(`${arrSold[i][5]}`)).sendKeys(0)
                    }

                })
                .catch(() => {
                    console.log("no located")
                })

        }, 11500)

    })


/* 7. ETAP 4
pour passe en etape suivant il faut clique manuellement */
driver.wait(until.titleIs('4 Solde de tout compte - Saisie attestation en ligne - Pôle emploi'))
    .then(() => {
        driver.findElement(By.name('datePaiePeriodeSolde')).sendKeys(getDataString(arrUnitData[240][1]));

        driver.findElement(By.name('tempsTravailleSolde')).sendKeys(arrUnitData[246][1].toFixed(2));

        driver.findElement(By.name('tempsNonPayeSolde')).sendKeys(0);

        driver.findElement(By.name('salaireBrutMensuelSolde')).sendKeys(arrUnitData[244][1].toFixed(2));
    })


/* 8. ETAP 5
pour passe en etape suivant il faut clique manuellement */
driver.wait(until.titleIs('5 Validation - Saisie attestation en ligne - Pôle emploi'))
    .then(() => {
        driver.findElement(By.name('etape5Prenom')).sendKeys("José");

        driver.findElement(By.name('etape5Nom')).sendKeys("de QUINA");

        /* DEPRECIATED 
        driver.wait(
            driver.findElement(By.name('qualite')).click())
            .then(() => {
                driver.findElement(By.css('[value="06"]')).click();
            })

        driver.wait(
            driver.findElement(By.name('etape5motifRupture')).click())
            .then(() => {
                driver.findElement(By.css('[value="031"]')).click();
            })
        */

        driver.findElement(By.name('personneAJoindreNom')).sendKeys("Nico");

        driver.findElement(By.name('telephone')).sendKeys(112);

        driver.findElement(By.name('lieuSignature')).sendKeys("Aubagne");
    })