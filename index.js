
// import xlsx from 'node-xlsx';



// //const workSheetFromBuffer = xlsx.parse(fs.readFileSysc(`${__dirname}/test.xlsx`));
// const workSheetsFromFile = xlsx.parse(`${__dirname}/test.xlsx`);


// ES5

var xlsx = require('node-xlsx');
var fs = require('fs');

var obj = { "worksheets": [
    {"data": [["name","sex","age"],["li","male","24"]] }
]};

// var file = xlsx.build(obj);
/**
 *  ## xlsx.build(obj) 格式
 *   ```
 *   obj = [sheet1, sheet2, ...];
 *   sheet1 = { name: "sheetName", data: sheetData};
 *   sheetData = [ line1DataArray, line2DataArray, ...];
 *   line1DataArray = [data1, data2, data3, ...];
 *   ```
 *                     |------------------------|
 *   line1DataArray -> | data1 | data2| data3 |  .....|
 *                     | --------------------------|
 * 
 *                     | sheetName1 |
 */

const workSheetsFromBuffer = xlsx.parse(fs.readFileSync(`${__dirname}/04月汇总表.xls`));

console.log(JSON.stringify(workSheetsFromBuffer));

// const data = [[1, 2, 3], [true, false, null, 'sheetjs'], ['foo', 'bar', new Date('2014-02-19T14:30Z'), '0.3'], ['baz', null, 'qux']];
// var buffer = xlsx.build([{name: "mySheetName", data: data}]); // Returns a buffer

// fs.writeFileSync('user.xlsx', buffer, 'binary');