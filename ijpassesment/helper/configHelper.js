let fs = require('fs');
let yaml = require('js-yaml');

function getConfig() {
    try {
        const config = yaml.safeLoad(fs.readFileSync('./config/app_env.yml', 'utf8'));
        const indentedJson = JSON.stringify(config, null, 4);
        return config;
    } catch (e) {
        errHandler(e);
    }
}

var errHandler = function (err) {
    console.log(err);
}

global.getConfig = getConfig;
global.errHandler = errHandler;

module.exports.getConfig = getConfig;
module.exports.errHandler = errHandler;