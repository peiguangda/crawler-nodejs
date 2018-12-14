var Excel = require('exceljs');
var workbook = new Excel.Workbook();

// create a new sheet writer with pageSetup settings for fit-to-page
var worksheetWriter = workbook.addWorksheet('My Sheet');

// adjust pageSetup settings afterwards
worksheetWriter.pageSetup.margins = {
    left: 0.7, right: 0.7,
    top: 0.75, bottom: 0.75,
    header: 0.3, footer: 0.3
};

// Set Print Area for a sheet
worksheetWriter.pageSetup.printArea = 'A1:G20';

// Repeat specific rows on every printed page
worksheetWriter.pageSetup.printTitlesRow = '1:3';

worksheetWriter.columns = [
    {header: 'Question', key: 'question', width: 30},
    {header: 'Ans1', key: 'ans1', width: 10},
    {header: 'Ans2', key: 'ans2', width: 10},
    {header: 'Ans3', key: 'ans3', width: 10},
    {header: 'Ans4', key: 'ans4', width: 10},
];

worksheetWriter.addRow(
    {
        question: "tao hoi cai",
        ans1: 'John Doe',
        ans2: 'John Doe',
        ans3: 'John Doe',
        ans4: 'John Doe'
    });

workbook.csv.writeFile("test.csv")
    .then(function () {
        console.log('ok');
    });
