var xlsx = require('node-xlsx');
var fs = require('fs');

const originFilename = `${__dirname}/04月汇总表.xls`;
const originSheetname = '刷卡记录';

// 读取文件
const originWorkSheets = xlsx.parse(fs.readFileSync(originFilename));
// var originSheetData = null;
// var targetSheetData = null;


// // 转换函数
// // 输入原始数据中单元格的数据，拟合出目标Excel中一行的数据。
// function convert(id,name) {

// }

// // 获取指定sheet的数据 ，Sheet名称为：刷卡记录
// for( var i = 0; i < originWorkSheets.length; i++) {
//     if(originWorkSheets[i] && originWorkSheets[i].name && originWorkSheets[i].name === originSheetname ) {
//         console.log(originWorkSheets[i].name);
//         originSheetData = originWorkSheets[i].data;  
//     }
// }
// // 开始转换
// for( var i = 0; i < originSheetData.length; i++){
//      console.log(JSON.stringify(originSheetData[i]));
// }


// // 输出转换后的格式
// if (targetSheetData) {
//     console.log(JSON.stringify(targetSheetData));
// }
//const sheet = originWorkSheets["刷卡记录"];

//console.log(JSON.stringify(sheet));

//const data = [[1, 2, 3], [true, false, null, 'sheetjs'], ['foo', 'bar', new Date('2014-02-19T14:30Z'), '0.3'], ['baz', null, 'qux']];
var buffer = xlsx.build(JSON.stringify(originWorkSheets)); // Returns a buffer
// 保存文件
fs.writeFileSync('user1.xlsx', originWorkSheets, 'binary');