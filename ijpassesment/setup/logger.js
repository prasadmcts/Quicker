'use strict';
const { createLogger, format, transports } = require('winston');
// const { combine, timestamp, label, printf } = format;
const fs = require('fs');
const path = require('path');

const env = process.env.NODE_ENV || 'DEV';
const logDir = 'log';

// Create the log directory if it does not exist
if (!fs.existsSync(logDir)) {
  fs.mkdirSync(logDir);
}

const filename = path.join(logDir, 'CredeitLimits.log');

const logger = createLogger({
  // change level if in dev environment versus production
  level: env === 'DEV' ? 'debug' : 'info',
  format: format.combine(
    format.timestamp({
      format: 'YYYY-MM-DD HH:mm:ss'
    }),
    format.printf(info => `${info.timestamp} ${info.level}: ${info.message}`)
  ),
  transports: [
    new transports.Console({
      level: 'info',
      format: format.combine(
        format.colorize(),
        format.printf(
          info => `${info.timestamp} ${info.level}: ${info.message}`
        )
      )
    }),
    new transports.File({ filename })
  ]
});

global.log = logger;
module.exports = logger;