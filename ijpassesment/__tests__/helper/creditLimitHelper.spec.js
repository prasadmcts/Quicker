let commonHelper = require('../../helper/creditLimitHelper');
let mockData = require('../../__mocks__/testData.mock');

describe('Credit limits graph tree', () => {
    it('creates a Credit limits graph tree based on input data', () => {
        expect(commonHelper.createCLsToDataTree(mockData.creditlimitsDataSet)).toEqual(mockData.expectedDataTreeCLDataSet)
    })
})

describe('Entity combination', () => {
    it('Get Entity combination for given Credit limits input data', () => {
        let expectedEntitySet = "A/B/C";
        let clDataTree = commonHelper.createCLsToDataTree(mockData.creditlimitsDataSet);
        expect(clDataTree.length).toBeGreaterThan(0);
        expect(commonHelper.getEntities(clDataTree[0])).toEqual(expectedEntitySet)
    })
})

describe('Sub Entities Combined Utilisation', () => {
    it('Get Sub Entities Combined Utilisation for given Credit limits input data', () => {
        let expectedEntitySet = 50;
        let clDataTree = commonHelper.createCLsToDataTree(mockData.creditlimitsDataSet);
        expect(clDataTree.length).toBe(1);
        expect(commonHelper.getSubEntitiesCombinedUtilisation(clDataTree[0])).toEqual(expectedEntitySet)
    })
})