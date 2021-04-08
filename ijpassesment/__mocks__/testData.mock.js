
const creditlimitsDataSet = [
    {
        "entity": "A",
        "parent": "",
        "limit": 100,
        "utilisation": 0
    },
    {
        "entity": "B",
        "parent": "A",
        "limit": 90,
        "utilisation": 10
    },
    {
        "entity": "C",
        "parent": "B",
        "limit": 40,
        "utilisation": 40
    }
]

const expectedDataTreeCLDataSet = [
    {
        "entity": "A",
        "parent": "",
        "limit": 100,
        "utilisation": 0,
        childNodes: [
            {
                "entity": "B",
                "parent": "A",
                "limit": 90,
                "utilisation": 10,
                childNodes: [
                    {
                        "entity": "C",
                        "parent": "B",
                        "limit": 40,
                        "utilisation": 40,
                        childNodes: []
                    }]
            }
        ]
    }
]

module.exports = { creditlimitsDataSet, expectedDataTreeCLDataSet };