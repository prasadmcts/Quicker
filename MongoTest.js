const { MongoClient } = require("mongodb");

const uri = `mongodb://localhost:27017/1Mos`;

//const client = new MongoClient(uri);

async function run() {
    try {
        MongoClient.connect(uri, function (err, db) {

            const dbo = db.db("myDb");

            // Here you get your lines from your .txt file
            //let line_to_execute = 'dbo.collection("users").findOne({}, (err, res)=>{console.log(res)});';
            let line_to_execute = 'dbo.collection("users").createIndex({ name: 1 });'

            // Launch the execution of line_to_execute
            var res = eval(line_to_execute).then(x => console.log(x));
            console.log(res);

        })

        //await client.connect();

        //const dbo = client.db("1Mos");

        // let line_to_execute = {
        //     dbStats: 1
        // };
        // eval(line_to_execute);
        // //const result = await db.command(line_to_execute);
        // //console.log(result);

        // let userNameIndex = {
        //     createIndexes: "users",
        //     indexes: [
        //         {
        //             key: { "name": 1 },
        //             name: "nameIndex"
        //         }
        //     ],
        //     comment: "Prasad test"
        // };
        // eval(userNameIndex);
        // const result2 = await db.command(userNameIndex);
        // console.log(result2);

    } finally {
        //await client.close();
    }
}
run().catch(console.dir);
