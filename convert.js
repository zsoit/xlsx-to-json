const fileName = "xlsx/test.xlsx";

const util = require('util')
var XLSX = require("xlsx");
var workbook = XLSX.readFile(fileName);
var sheet_name_list = workbook.SheetNames;

function display_sheets(y) {
    var worksheet = workbook.Sheets[y];
    var headers = {};
    var data = [];
    for (z in worksheet) {
        if (z[0] === "!") continue;
        var col = z.substring(0, 1);
        var row = parseInt(z.substring(1));
        var value = worksheet[z].v;

        if (row == 1) {
            headers[col] = value;
            continue;
        }
        if (!data[row]) data[row] = {};

        data[row][headers[col]] = value;
    }

    data.shift();
    data.shift();

    console.log("[ ");
    for (var i = 0; i <= data.length - 1; i++) {
        var result = Object.keys(data).map((key) => [Number(key), data[key]]);
        result = JSON.stringify(result[i][1]);
        (i == data.length - 1) ? console.log(result): console.log(result + ", ");

    }
    console.log("] ");
}

function consoletable() {

    console.log('"tables": \n[\n');
    for (var i = 0; i <= workbook.SheetNames.length - 1; i++) {
        var name_sheets = JSON.stringify(workbook.SheetNames[i]);
        console.log(
            `{"ID":${i+1},"name":${name_sheets}}`

        );
        if (i != workbook.SheetNames.length - 1) {
            console.log(", ");

        }
    }
    console.log("\n],\n");

}

function consoleJSON() {
    console.log('{ ');
    consoletable();
    sheet_name_list.forEach(function(y, idx, array) {
        console.log(`"${y}": \n`);
        display_sheets(y);
        (idx === array.length - 1) ? console.log("  "): console.log(`, `);;
    });
    console.log('} ');
}

consoleJSON();