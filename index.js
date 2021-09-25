
// --- PACKAGE ---
const XLSX = require('xlsx');
const { Builder, By, Key, until } = require('selenium-webdriver');
const { elementIsSelected } = require('selenium-webdriver/lib/until');
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

/* 
START selenium-webdriver
TODO:
1. Creation session et auth
2. Acceder pour tout les attestation
3. Entrer en "Creation Attestation"
4. ETAP 1
5. ETAP 2
6. ETAP 3
*/
const driver = new Builder().forBrowser('firefox').build();

/* 1. Creation session et auth. */
driver.get('file:///C:/Users/mrdim/Desktop/2.html');
driver.findElement(By.id('footer_tc_privacy_button_2')).click();
// driver.get('https://gestion.pole-emploi.fr/espaceemployeur/authentification/authentification/');
// driver.findElement(By.name('identifiant')).sendKeys(process.env.identifiant)
// driver.findElement(By.name('codeAcces')).sendKeys(process.env.pass)
// driver.findElement(By.name('departement')).sendKeys(process.env.dep);
// driver.findElement(By.id('footer_tc_privacy_button_2')).click();
// driver.findElement(By.id('boutonValider')).click();

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

        // // IRCANTEC
        // driver.wait(driver.findElement(By.id('organismeRetraiteComplementaire')).click())
        //     .then(() => {
        //         driver.findElement(By.css('[value="200"]')).click();
        //     })

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
        

    })

