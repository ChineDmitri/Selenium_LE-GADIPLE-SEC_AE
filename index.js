
// --- PACKAGE ---
const XLSX = require('xlsx');
const { Builder, By, Key, until } = require('selenium-webdriver');
const dotenv = require("dotenv").config();
// ---------------

// envoyer varibale (ex. node app.js 1)
const idx = parseInt(process.argv[2], 10);
// console.log(idx);