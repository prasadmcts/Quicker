let index = require('../index');
let mockData = require('../__mocks__/testData.mock');
let configHelper = require('../helper/configHelper');
let creditLimitHelper = require('../helper/creditLimitHelper');
let fileHelper = require('../helper/fileHelper');

require("dotenv").config();

describe('Test environmental variables', () => {
    const OLD_ENV = process.env;
    beforeEach(() => {
        jest.resetModules();
        process.env = { ...OLD_ENV };
        delete process.env.NODE_ENV;
    });

    afterEach(() => {
        process.env = OLD_ENV;
    });

    it('receive process.env variables', () => {
        process.env.NODE_ENV = "DEV";
        var config = configHelper.getConfig()[process.env.NODE_ENV];
        expect(config.EnvName).toEqual("development");
        expect(config.CreditLimitsInputCsvPath.length).toBeGreaterThan(0);
        expect(config.CreditLimitsOutputCsvPath.length).toBeGreaterThan(0);
    });
});

jest.mock('../helper/fileHelper');

describe('Generate Report', () => {
    it('Read Credit Limits data', () => {
        process.env.NODE_ENV = "DEV";
        var config = configHelper.getConfig()[process.env.NODE_ENV];
        const spy = jest.spyOn(fileHelper, 'readCreditLimitsData');
        spy.mockReturnValue(mockData.creditlimitsDataSet);
        expect(fileHelper.readCreditLimitsData(config.CreditLimitsInputCsvPath)).toBe(mockData.creditlimitsDataSet);
    });

    it('Generate Credit Limits output data', () => {
        // Arrange
        process.env.NODE_ENV = "DEV";
        var config = configHelper.getConfig()[process.env.NODE_ENV];
        const spy = jest.spyOn(creditLimitHelper, 'createCLsToDataTree');
        spy.mockReturnValue(mockData.expectedDataTreeCLDataSet);

        jest.spyOn(fileHelper, 'writeToFile').mockImplementation((path, data) => { });

        // Act
        index.generateReport(mockData.expectedDataTreeCLDataSet);

        // Assert
        expect(creditLimitHelper.createCLsToDataTree).toHaveBeenCalledTimes(1);
        expect(fileHelper.writeToFile).toHaveBeenCalledTimes(1);
        const outputData = ["Entities: A/B/C", "   No limit breaches", "", ""];
        expect(fileHelper.writeToFile).toBeCalledWith(config.CreditLimitsOutputCsvPath, outputData);
    });
});