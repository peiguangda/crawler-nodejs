var Crawler = require("crawler");
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
    {header: 'Question', key: 'question', width: 50},
    {header: 'Ans1', key: 'ans1', width: 10},
    {header: 'Ans2', key: 'ans2', width: 10},
    {header: 'Ans3', key: 'ans3', width: 10},
    {header: 'Ans4', key: 'ans4', width: 10},
];
var page = -10;
var c = new Crawler({
    rateLimit: 1000,
    maxConnections: 1,
    // This will be called for each crawled page
    callback: function (error, res, done) {
        if (error) {
            console.log(error);
        } else {
            var $ = res.$;
            page += 10;
            $('td[colspan="5"]').each(function (index, element) {
                ++index;
                var question = $(this).text();
                var number = page + index;
                var ans1 = $("label[for=" + "\"s_ans[" + number + "]1\"" + "]").text();
                var ans2 = $("label[for=" + "\"s_ans[" + number + "]2\"" + "]").text();
                var ans3 = $("label[for=" + "\"s_ans[" + number + "]3\"" + "]").text();
                var ans4 = $("label[for=" + "\"s_ans[" + number + "]4\"" + "]").text();
                var answer = $('tr td input[name="r_ans"]').toArray()[index - 1].attribs.value;

                if (ans1.includes(answer)) {
                    ans1 = ans1.replace("1)", "*.");
                    ans2 = ans2.replace("2)", "");
                    ans3 = ans3.replace("3)", "");
                    ans4 = ans4.replace("4)", "");
                } else if (ans2.includes(answer)) {
                    ans2 = ans2.replace("2)", "*.");
                    ans1 = ans1.replace("1)", "");
                    ans3 = ans3.replace("3)", "");
                    ans4 = ans4.replace("4)", "");
                    [ans2, ans1] = [ans1, ans2];
                } else if (ans3.includes(answer)) {
                    ans3 = ans3.replace("3)", "*.");
                    ans2 = ans2.replace("2)", "");
                    ans1 = ans1.replace("1)", "");
                    ans4 = ans4.replace("4)", "");
                    [ans3, ans1] = [ans1, ans3];
                } else {
                    ans4 = ans4.replace("4)", "*.");
                    ans2 = ans2.replace("2)", "");
                    ans3 = ans3.replace("3)", "");
                    ans1 = ans1.replace("1)", "");
                    [ans4, ans1] = [ans1, ans4];
                }
                question = question.trim();
                question = "#." + question;
                // console.log(question);
                // console.log(ans1);
                // console.log(ans2);
                // console.log(ans3);
                // console.log(ans4);
                console.log(number);

                worksheetWriter.addRow(
                    {
                        question: question,
                        ans1: ans1,
                        ans2: ans2,
                        ans3: ans3,
                        ans4: ans4
                    });
                worksheetWriter.addRow(
                    {
                        question: "",
                        ans1: "",
                        ans2: "",
                        ans3: "",
                        ans4: ""
                    });

            });
            workbook.xlsx.writeFile("goi_n2.xlsx")
                .then(function () {
                    console.log('ok');
                });

        }
        done();
    }
});

// Queue just one URL, with default callback
c.queue('http://www.n-lab.org/library/mondaidata/test.php?mode=html&dbupdate=0&target=all&type=and&start_num=1&show_num=10&sort=test_num&kyu=2&syu=%E8%AA%9E%E5%BD%99&word=&submit=%E6%A4%9C%E7%B4%A2');
for (var i = 10; i < 915; i += 10) {
    c.queue('http://www.n-lab.org/library/mondaidata/test.php?mode=html&dbupdate=0&data_count=915&jump_num=' + i + '&show_num=10&kyu=2&syu=%E8%AA%9E%E5%BD%99&target=all&word=&type=and&sort=test_num&test_num=');
}