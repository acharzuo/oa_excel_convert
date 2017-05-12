var xlsx = require('node-xlsx');
var fs = require('fs');
var config = require("config");

// 获取参数
var arguments = process.argv.splice(2);


// 命令行参数判断
if(!arguments || arguments.length < 1){
    console.log("usage: node index <Excle Filename>")
    return;
}
const filename = arguments[0] ;

console.log(filename);

const originFilename = `${__dirname}/` + filename;
const originSheetname = '刷卡记录';
var originWorkSheets  = null


return 0;

// 读取文件
try {
  originWorkSheets = xlsx.parse(fs.readFileSync(originFilename)); 
} catch (error) {
    console.error("读取文件错误！\n" + originFilename + error);
    return;
}

var originSheetData = null;
var targetSheetData = null;


// 读取配置文件
const WORK_IN_TIME = config.get('workInTime');
const WORK_OUT_TIME = config.get('workOutTime');

// 人员对应规则
const ACCOUNT_PAIR = config.get('accountPair')

// 转换函数
// 输入原始数据Sheet的数据，拟合出目标Excel中新Sheet的数据。
function convert(oriObj) {

    if(typeof oriObj != "object") return;

    var ret = [];

   // 一行的数据
   var line = [
       "",  // id	编号
       "",  // account	用户
       "",  // date	日期
       "",  // signIn	签到
       "",  // signOut	签退
       "",  // status	状态
       "",  // ip	IP
       "",  // device	设备
       "",  // manualIn	签到时间
       "",  // manualOut	签退时间
       "",  // reason	原因
       "",  // desc	描述
       "",  // reviewStatus	补录状态
       "",  // reviewedBy	审核人
       ""  // reviewedDate	审核时间
   ];

   // 行号
   var index = 1 ;
   var month = "";
   const DATE_LINE_INDEX = 3;

   // 遍历所有行
   for( var i = 0; i < oriObj.length; i++){
        // 获取考勤日期的月份
        if( oriObj[i] && oriObj[i][0] && oriObj[i][0] === "考勤日期 : " && oriObj[i][2] ) {
            // 数据格式 模拟 ["2017/04/01 ~ 04/30 (ACHAR科技)"]
            var tmp = oriObj[i][2].split(" ");
            month = tmp[0].substr(0,7).replace("/","-");

        } else if (oriObj[i] && oriObj[i][0] && oriObj[i][0] === "工 号：" && oriObj[i][2]) {
            // 获取每个人每天的记录的记录
            for(var day = 0; day < 31; day++){
                var record = "";    // 打开记录
                if(oriObj[i+1][day]) {
                    record = oriObj[i+1][day];
                    record = record.trim();
                } else {
                    continue;
                }

                var account = oriObj[i][10];
                account = ACCOUNT_PAIR[account];
                if(!account) {
                    console.log(oriObj[i][10] + "没有对应的OA账号，已经智能跳过！");
                    break;
                }
                // record = "15:45\n18:59\n"
                var recordSplit = record.split("\n");
                var signIn = recordSplit[0];        // 第一次打卡时间为签到时间
                var signOut = recordSplit.length > 1 ? recordSplit[recordSplit.length-1] : "";    // 最后一次打卡时间为签退时间

                // 整理字符格式，使用5位的时间表示方法
                signIn = signIn.length === 4 ? "0" + signIn : signIn;
                signOut = signOut.length === 4 ? "0" + signOut : signOut;


                var status = "normal";
                if(signIn > WORK_IN_TIME && signOut < WORK_OUT_TIME && signIn != signOut ){
                    // 迟到早退 签到 > 8:50 && 签退 < 18:00 && 签到 != 签退
                    status = "both";
                } else if(signIn > WORK_IN_TIME && signIn < WORK_OUT_TIME && signOut >= WORK_OUT_TIME){
                    // 迟到判断 签到> 8:50 && 签到< 18:00 && 签退 >=18:00	
                    status = "late";
                } else if (signIn < WORK_IN_TIME && signOut > WORK_IN_TIME && signOut < WORK_OUT_TIME ){
                    // 早退判断 签到<= 8:50 && 签退>8:50 && 签退<=18:00
                    status = "early";
                }

                var line = [];
                line[0] = index++;
                line[1] = account;
                line[2] = month + "-" + oriObj[DATE_LINE_INDEX][day];
                line[3] = signIn;
                line[4] = signOut;
                line[5] = status;
                line[6] = "127.0.0.1";
                line[7] = "得力打卡机01";
                line[8] = "00:00:00";
                line[9] = "00:00:00";
                line[10] = "";
                line[11] = "";
                line[12] = "";
                line[13] = "";
                line[14] = "0/0/0000 00:00:00";
                ret.push(line);
            }
        }
   }
   
   return ret;

}

// 获取指定sheet的数据 ，Sheet名称为：刷卡记录
for( var i = 0; i < originWorkSheets.length; i++) {
    if(originWorkSheets[i] && originWorkSheets[i].name && originWorkSheets[i].name === originSheetname ) {
        originSheetData = originWorkSheets[i].data;  
    }
}

// 执行转换
targetSheetData = convert(originSheetData);

// 输出转换后的格式
if (targetSheetData) {
    //console.log(JSON.stringify(targetSheetData));

    var buffer = xlsx.build([{"name": "打卡记录", "data":targetSheetData}]); 
    // 保存文件
    fs.writeFileSync( filename.replace(".","_转换."), buffer, 'binary');
    console.log("转换完成！");
}

