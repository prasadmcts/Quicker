const csv = require("csvtojson");
const fs = require('fs');
let logger = require('../setup/logger');

async function readCreditLimitsData(filePath) {
    var data = [];
    await readFileData(filePath)
        .then(result => { data = result; })
        .catch((err) => { console.error('error'); logger.error(`Error Occured: '${err}'`); });
    return data;
}

async function readFileData(filePath) {
    return new Promise((resolve, reject) => {
        const creditLimitsData = csv().fromFile(filePath);
        resolve(creditLimitsData);
    });
}

function writeToFile(filePath, csvData) {
    let csv = `"${csvData.join('"\n"')}"`;
    fs.writeFile(filePath, csv, 'utf8', function (err) {
        if (err) {
            logger.error(`Some error occured - file either not saved or corrupted file saved.: '${err}'`);
        } else {
            logger.info(`Output File Saved to: '${filePath}'`);
        }
    });
}

module.exports.readCreditLimitsData = readCreditLimitsData;
module.exports.readFileData = readFileData;
module.exports.writeToFile = writeToFile;
