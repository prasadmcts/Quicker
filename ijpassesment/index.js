const csv = require("csvtojson");
let logger = require('./setup/logger');
let configHelper = require('./helper/configHelper');
let creditLimitHelper = require('./helper/creditLimitHelper');
let fileHelper = require('./helper/fileHelper');
let constants = require('./helper/constants');
const format = require('string-format');
require("dotenv").config();
format.extend(String.prototype, {});

var env = process.env.NODE_ENV || 'DEV';
var config = configHelper.getConfig()[env];

log.info(`App running in '${config.EnvName}' mode`);
log.info(`Credit Limits Input File Path : '${config.CreditLimitsInputCsvPath}'`);
log.info(`Credit Limits Output File Path : '${config.CreditLimitsOutputCsvPath}'`);

(async () => {
    try {
        let creditLimitsData = await fileHelper.readCreditLimitsData(config.CreditLimitsInputCsvPath);
        generateReport(creditLimitsData);
    } catch (err) {
        logger.error(`Error Occured: '${err}'`);
    }
})();

function generateReport(creditLimitsData) {
    log.info(`Total '${creditLimitsData.length}' Credit Limits found.`);

    var structuredCLs = creditLimitHelper.createCLsToDataTree(creditLimitsData);

    var outputData = [];
    structuredCLs.forEach(cl => {

        let entities = creditLimitHelper.getEntities(cl);
        outputData.push(`${constants.Constants.Entities}${entities}`);
        log.info(`${constants.Constants.Entities}${entities}`);

        // Check is limit breaches
        let totalUtilisation = creditLimitHelper.getSubEntitiesCombinedUtilisation(cl);
        let msg = Number(cl.limit) < totalUtilisation ?
            constants.Constants.LimitBreaches.format(cl.entity, cl.limit, cl.utilisation, totalUtilisation) :  // breached
            constants.Constants.NolimitBreaches; //  Not breached

        outputData.push(`   ${msg}`);
        log.info(`   ${msg}`);

        outputData.push("");
        outputData.push("");
    });

    fileHelper.writeToFile(config.CreditLimitsOutputCsvPath, outputData);
}

module.exports.generateReport = generateReport;