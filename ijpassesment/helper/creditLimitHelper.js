let constants = require('./constants');

const createCLsToDataTree = dataset => {
    let hashTable = Object.create(null);
    dataset.forEach(aData => hashTable[aData.entity] = { ...aData, childNodes: [] });
    let dataGraphTree = [];
    dataset.forEach(aData => {
        if (aData.parent) hashTable[aData.parent].childNodes.push(hashTable[aData.entity]);
        else dataGraphTree.push(hashTable[aData.entity]);
    })
    return dataGraphTree;
}

function getEntities(object) {
    let entities = "";
    for (child of object.childNodes) {
        entities += constants.Constants.EntitySeparator + getEntities(child);
    }
    return object.entity + entities;
}

function getSubEntitiesCombinedUtilisation(object) {
    let totalUtilisation = 0;
    for (child of object.childNodes) {
        totalUtilisation += getSubEntitiesCombinedUtilisation(child);
    }
    return Number(object.utilisation) + Number(totalUtilisation);
}

module.exports.createCLsToDataTree = createCLsToDataTree;
module.exports.getEntities = getEntities;
module.exports.getSubEntitiesCombinedUtilisation = getSubEntitiesCombinedUtilisation;