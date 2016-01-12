var fs = require("fs")
var pathjs = require("path")
var xlsx = require('node-xlsx');
//var nt = require('time');
var moment = require('moment-timezone');

var outputDataPath = './historicoArtesp';
var xlsDataPath = './';
var historyJsonPath = './historyJson.json';
var dataSpreadsheet;
var historyJson;

Date.prototype.UTCyyyymmdd = function() {
    var yyyy = this.getUTCFullYear().toString();
    var mm = (this.getUTCMonth()+1).toString(); // getMonth() is zero-based
    var dd  = this.getUTCDate().toString();
    return yyyy + (mm.length===2?mm:"0"+mm[0]) + (dd.length===2?dd:"0"+dd[0]); // padding
};

function getAllXls(lastPath, callback){
    if(!lastPath) {
        var result = getAllXls(gmapsDataPath, callback);
        return result;
    } else {
        // console.log("lastPath:", lastPath);
        historyJson.path = lastPath;
        fs.writeFileSync(historyJsonPath, JSON.stringify(historyJson));
        var result = false;
        var stat = fs.statSync(lastPath);
        if(stat.isDirectory()){
            historyJson.dirRead = historyJson.dirRead || [];
            if(historyJson.dirRead.indexOf(lastPath) < 0) {
                historyJson.dirPath = lastPath;
                var allFiles = fs.readdirSync(lastPath);
                fs.writeFileSync(historyJsonPath, JSON.stringify(historyJson));
                var absolute_result = true;
                allFiles.every(function(file){
                    result = getAllXls(pathjs.join(lastPath, file), callback);
                    absolute_result &= result;
                    // console.log("result:", result);
                    return result;
                });
                // if(absolute_result) {
                //     historyJson.dirRead.push(lastPath);
                //     fs.writeFileSync(historyJsonPath, JSON.stringify(historyJson));
                // }
                result = absolute_result;
            } else {
                result = true;
                historyJson.dirPath = pathjs.dirname(lastPath);
                fs.writeFileSync(historyJsonPath, JSON.stringify(historyJson));
                if(typeof(callback) === 'function') process.nextTick(callback, "Directory fully read: " + lastPath, null);
                //process.nextTick(getAllXls, historyJson.dirPath, callback);
            }
        } else if (stat.isFile()) {
            historyJson.filesRead = historyJson.filesRead || [];
            if(historyJson.filesRead.indexOf(lastPath) < 0) {
                historyJson.filesRead.push(lastPath);
                fs.writeFileSync(historyJsonPath, JSON.stringify(historyJson));
                var fileData;
                if(pathjs.extname(lastPath) == ".zip"){
                    var zip = new AdmZip(lastPath);
                    var zipEntries = zip.getEntries(); // an array of ZipEntry records 
                    fileData = zip.readAsText(zipEntries[0]);
                }
                return false;
            } else {
                return true;
            }
        } else {
            return false;
        }
        active = false;
        return result;
    }
}

(function main () {
    if(!fs.existsSync(historyJsonPath)) {
        fs.writeFileSync(historyJsonPath, "");
    }
    try {
        historyJson = require(historyJsonPath);
    } catch(ex) {
        historyJson = {};
    }

    // ---------------- XLS Object Structure -------------------
    // xls
    // |── tab number (Ecovias tab no. 35 and 36)
    // |   |── name (Need to extract the road direction)
    // |   |── data
    // |   |   |── Row
    // |   |   |   |── Columns [Array]
    // ----------------------------------------------------------

    // makeDir(pathjs.join(outputDataPath, "SP-055","Leste"));

    console.log("---->" + excelDateToDate(41090, 3) );
    // var countAnalise = 1;
    // var filename = "248-263";
    // var obj = xlsx.parse(pathjs.join(xlsDataPath, '248-263.xls'));
    // Object.keys(obj).every(function (value){
    //     if (obj[value].name.indexOf("Análise") < 0) {
    //         return true;
    //     }
    //     else {
    //         var globalData = {};
    //         globalData.concessionaria = "Ecovias"
    //         var table = obj[value].data;
    //         for (var row = 0; row < table.length; row++) {
    //             if ((table[row][1]+"").toUpperCase() == "HORA" && 
    //                 (table[row][2]+"").toUpperCase() == "DATA" && 
    //                 (table[row][3]+"").toUpperCase() == "VOLUME LEVANTADO")
    //             {
    //                 var arrRoadDirection = table[row-12][2].split(' ');
    //                 globalData.road = arrRoadDirection[0];
    //                 globalData.direction = arrRoadDirection[1];
    //                 row += 1;
    //                 var dataRow = table[row+1];
    //                 var dateReport = excelDateToDate(dataRow[2]);
    //                 dateReport.setUTCHours(dateReport.getUTCHours()+(parseInt(dataRow[1])-1)%24);
    //                 console.log("%d, 2, 9, LOG, %s, %d, %d, %f, %f, %f, %f, %s, %s, passeio:\%d comercial:\%d tx_fluxo:\%f vp:\%f velocidade:\%f densidade:\%f ns:\%s concessionaria:\%s",
    //                             dateReport.getTime(), filename, dataRow[3], dataRow[4], dataRow[5], dataRow[6], dataRow[7], dataRow[8], dataRow[9], dataRow[10], globalData.concessionaria);
    //                 throw Error("just stop");
    //                 continue;
    //             }
    //         }
    //         return countAnalise++ < 2;
    //     }
    // });
    // console.log(keysFiltered);
    // console.log(typeof(obj[35].data));
    // console.log("is Array? %s", Array.isArray(obj[35].data));
    // console.log(JSON.stringify(obj[35].data));
    // console.log(obj[35].data['0']);
    // console.log(obj[35].data['1']);
    // console.log(obj[35].data['2']);
    // console.log(obj[35].data['3']);
    // console.log(obj[35].data['15']);
    // console.log(obj[35].data['16']);
    // console.log(obj[35].data['17']);

})();


function excelDateToDate(excelDate, excelHour){
    var result = new moment("1899-12-31").tz('America/Sao_Paulo'); //to offset to Unix epoch and multiply by milliseconds
    result.add({ days: excelDate});
    result.hour(excelHour);
    result.minute(0);
    result.second(0);

    result.utc(); //Convert to GMT-0

    console.log(result.toString());

    return result.unix();
}

function makeDir(str_path) {
    try {
        fs.mkdirSync(str_path);
    } catch(ex) {
        // console.log(JSON.stringify(ex));
        if(ex.errno != -4075) {
            if(ex.errno == -4058) {
                var dirPathSplit = str_path.split(pathjs.sep);
                var dirPath = "";
                dirPathSplit.forEach(function (value) {
                    dirPath = pathjs.join(dirPath, value);
                    makeDir(dirPath);
                });
            } else {
                throw ex;
            }
        }
    }
}
