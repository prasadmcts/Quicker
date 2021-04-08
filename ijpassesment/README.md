# Assesment

### Prerequisites

```
Install [Node](https://nodejs.org/en/), I've used Node: 12.18.0 LTS
```

### Installing

Follow below simple steps to run the project

```
git clone https://github.service.anz/kanaparp/ijpassesment.git
cd ijpassesment (Basically checkout folder)
npm install
npm start


npm test                // Run all unit test cases
npm run test:coverage   // Run all unit test cases and generates Coverage index.html report in (./coverage/unit) folder
```

## Input
  . Input Report(csv) file is in **data\input\CreditLimits.csv**

## Output Report Genearation
  Output data can be seen in 3 places
  
  1.  Output Report(csv) file will be saved to **data\output\CreditLimits.csv**
       ![Image of CSV](https://github.service.anz/kanaparp/ijpassesment/blob/master/images/CSV.JPG)
  2.  In Log file (**log\results.log**)
  3.  Console
       ![Image of Console](https://github.service.anz/kanaparp/ijpassesment/blob/master/images/Console.JPG)

## Unit Testing Code Coverage Report - Using JEST
  Run **npm run test:coverage**, the unit test code coverage report (index.html) and related files will be generating in *coverage\unit* path. Click on index.html then it'll be show like below.
  
  1.  Overall project code coverage
       ![Image of Main_Coverage](https://github.service.anz/kanaparp/ijpassesment/blob/master/images/Main_Coverage.jpg)
  2.  index.js file code coverage
	   ![Image of Index](https://github.service.anz/kanaparp/ijpassesment/blob/master/images/Index.jpg)
  3.  Helper folder code coverage
       ![Image of Helper](https://github.service.anz/kanaparp/ijpassesment/blob/master/images/Helper.JPG)
	   
## Deployment

The Project setup with PROD env mode

## Built With

* [nodemon](https://www.npmjs.com/package/nodemon) - To automatically restarting the application
* [dotenv](https://www.npmjs.com/package/dotenv) - Loads environment variables from .env file
* [js-yaml](https://www.npmjs.com/package/js-yaml) - Used to for env configuration file
* [csvtojson](https://www.npmjs.com/package/csvtojson) - Used to read input csv file and convert to json array
* [winston](https://rometools.github.io/rome/) - For general logging
* [jest](https://www.npmjs.com/package/jest) - For unit tesing (AAA Pattern) & Code coverage report generation.

## Code Logic High-Level

1. In **.env** file, set the Env. Ex: NODE_ENV = "PROD"
1. Read the Env config(**app_env.yml**) based on above env setup, It's file input/output paths.
1. Read CSV input file(Path is from #2) and convert data to Json Array.
1. Do the required code logic (Please see index.js)
1. Write to Output file path(Path is from #2)
