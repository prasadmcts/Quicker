const fs = require('fs');
let fileHelper = require('../../helper/fileHelper');
let mockData = require('../../__mocks__/testData.mock');

describe('test readCreditLimitsData func', () => {
    it('test readFileData func', () => {
        const spy = jest.spyOn(fileHelper, 'readFileData');
        spy.mockReturnValue('[]');
        expect(fileHelper.readFileData()).toBe('[]');
    });

    it('test readCreditLimitsData func', () => {
        const spy = jest.spyOn(fileHelper, 'readCreditLimitsData');
        spy.mockReturnValue(mockData.creditlimitsDataSet);
        expect(fileHelper.readCreditLimitsData("")).toBe(mockData.creditlimitsDataSet);
        spy.mockRestore();
    });
});

describe('writeToFile tests', () => {
    afterEach(() => {
        jest.restoreAllMocks();
    });

    it('should read json and write file correctly', async () => {

        // Arrange
        let callback;
        const outputData = ['Entities: A/B/C'];
        jest.spyOn(fs, 'writeFile').mockImplementation((path, data, utf, cb) => {
            callback = cb
        });

        // Act
        await fileHelper.writeToFile("somePath", outputData);

        // Assert
        expect(fs.writeFile).toHaveBeenCalledTimes(1);
        expect(fs.writeFile).toBeCalledWith("somePath", JSON.stringify("Entities: A/B/C"), "utf8", callback);
    });
});
