var fs = require("fs");
var pathjs = require("path");
var xlsx = require('node-xlsx');
// var nt = require('time');
var moment = require('moment');

var outputDataPath = './historicoArtesp';
var xlsDataPath = './xls';
var historyJsonPath = './historyJson.json';
var dataSpreadsheet;
var historyJson;

function getAllXls(filePath, callback){
    if(!filePath) {
        var result = getAllXls(xlsDataPath, callback);
        return result;
    } else {
        console.log("filePath:", filePath);
        historyJson.path = filePath;
        fs.writeFileSync(historyJsonPath, JSON.stringify(historyJson, null, " "));
        var stat = fs.statSync(filePath);
        if(stat.isDirectory()){
            historyJson.dirRead = historyJson.dirRead || [];
            if(historyJson.dirRead.indexOf(filePath) < 0) {
                historyJson.dirPath = filePath;
                var allFiles = fs.readdirSync(filePath);
                fs.writeFileSync(historyJsonPath, JSON.stringify(historyJson, null, " "));
                allFiles.forEach(function(file){
                    result = getAllXls(pathjs.join(filePath, file), callback);
                });
            } else {
                historyJson.dirPath = pathjs.dirname(filePath);
                fs.writeFileSync(historyJsonPath, JSON.stringify(historyJson, null, " "));
            }
        } else if (stat.isFile()) {
            historyJson.filesRead = historyJson.filesRead || [];
            if(historyJson.filesRead.indexOf(filePath) < 0) {
                historyJson.filesRead.push(filePath);
                fs.writeFileSync(historyJsonPath, JSON.stringify(historyJson, null, " "));
                if (typeof(callback) === 'function') callback(filePath);
            }
        }
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

    // ---------------- Input Folder Structure ------------------
    // CONCESSIONARIA/
    // |── ANO/
    // |   |── MES/
    // |   |   ├── TRECHO.(xls, xlsm, xlsx)
    // ----------------------------------------------------------

    // --------------- Output Folder Structure ------------------
    // RODOVIA_CODIGO/
    // |── TRECHO/
    // |   |── SENTIDO/
    // |   |   ├── YYYYMMDD.log
    // ----------------------------------------------------------

    getAllXls(xlsDataPath, function(filePath) {
        var countAnalise = 1;
        var filename = pathjs.basename(filePath);
        filename = filename.substr(0, filename.length - pathjs.extname(filePath));
        var obj = xlsx.parse(filePath);
        Object.keys(obj).every(function (value){
            var tabName = obj[value].name;
            if (tabName.toUpperCase().indexOf("ANÁLISE") < 0 || tabName.indexOf("TH") < 0 ) {
                return true;
            }
            else {
                var globalData = {};
                var table = obj[value].data;
                for (var row = 0; row < table.length; row++) {
                    if ((table[row][1]+"").toUpperCase() == "HORA" && 
                        (table[row][2]+"").toUpperCase() == "DATA" && 
                        (table[row][3]+"").toUpperCase() == "VOLUME LEVANTADO")
                    {
                        var relativePath = pathjs.relative(xlsDataPath, filePath);
                        var relativeSplit = relativePath.split(pathjs.sep);
                        globalData.road = filename;
                        globalData.direction = table[row-12][2];
                        globalData.dealership = relativeSplit[0];
                        row += 1;
                        continue;
                    }
                    else if(Object.keys(globalData).length > 0 && table[row][1]){
                        var dataRow = table[row];
                        var dateReport = excelDateToDate(dataRow[2]);
                        // dateReport.setUTCHours(dateReport.getUTCHours()+(parseInt(dataRow[1])-1)%24);
                        dateReport.add({hours:((parseInt(dataRow[1])-1)%24)});
                        // console.log(globalData);
                        // console.log("%d, 2, 9, LOG, %s, %d, %d, %d, %d, %d, %d, %s, %s, passeio:\%d comercial:\%d tx_fluxo:\%d vp:\%d velocidade:\%d densidade:\%f ns:\%s concessionaria:\%s",
                        //             dateReport.unix(), filename, dataRow[3], dataRow[4], dataRow[5], dataRow[6], dataRow[7], dataRow[8], dataRow[9], dataRow[10], globalData.concessionaria);
                        console.log(`${dateReport.unix()}, 2, 9, LOG, ${filename}, ${dataRow[3]}, ${dataRow[4]}, ${dataRow[5]}, ${dataRow[6]}, ${dataRow[7]}, ${dataRow[8]}, ${dataRow[9]}, ${globalData.dealership}, passeio:\%d comercial:\%d tx_fluxo:\%d vp:\%d velocidade:\%d densidade:\%f ns:\%s concessionaria:\%s\n`);
                    } else if(Object.keys(globalData).length > 0 && !table[row][1]) {
                        // throw Error("just stop");
                        break;
                    }
                }
                return countAnalise++ < 2;
            }
        });
    });

    // console.log(excelDateToDate(40909).format("YYYYMMDD"));
    // console.log(excelDateToDate(40909).format("dddd, MMMM Do YYYY, h:mm:ss a"));
    // var countAnalise = 1;
    // var filename = "248-263";
    // var obj = xlsx.parse(pathjs.join(xlsDataPath, '248-263.xls'));
    // Object.keys(obj).every(function (value){
    //     if (obj[value].name.indexOf("Análise") < 0) {
    //         return true;
    //     }
    //     else {
    //         var globalData = {};
    //         var table = obj[value].data;
    //         for (var row = 0; row < table.length; row++) {
    //             if ((table[row][1]+"").toUpperCase() == "HORA" && 
    //                 (table[row][2]+"").toUpperCase() == "DATA" && 
    //                 (table[row][3]+"").toUpperCase() == "VOLUME LEVANTADO")
    //             {
    //                 var arrRoadDirection = table[row-12][2].split(' ');
    //                 globalData.road = arrRoadDirection[0];
    //                 globalData.direction = arrRoadDirection[1];
    //                 globalData.dealership = "Ecovias";
    //                 row += 1;
    //                 continue;
    //             }
    //             else if(Object.keys(globalData).length > 0 && table[row][1]){
    //                 var dataRow = table[row];
    //                 var dateReport = excelDateToDate(dataRow[2]);
    //                 // dateReport.setUTCHours(dateReport.getUTCHours()+(parseInt(dataRow[1])-1)%24);
    //                 dateReport.add({hours:((parseInt(dataRow[1])-1)%24)});
    //                 // console.log(globalData);
    //                 // console.log("%d, 2, 9, LOG, %s, %d, %d, %d, %d, %d, %d, %s, %s, passeio:\%d comercial:\%d tx_fluxo:\%d vp:\%d velocidade:\%d densidade:\%f ns:\%s concessionaria:\%s",
    //                 //             dateReport.unix(), filename, dataRow[3], dataRow[4], dataRow[5], dataRow[6], dataRow[7], dataRow[8], dataRow[9], dataRow[10], globalData.concessionaria);
    //                 console.log(`${dateReport.unix()}, 2, 9, LOG, ${filename}, ${dataRow[3]}, ${dataRow[4]}, ${dataRow[5]}, ${dataRow[6]}, ${dataRow[7]}, ${dataRow[8]}, ${dataRow[9]}, ${globalData.dealership}, passeio:\%d comercial:\%d tx_fluxo:\%d vp:\%d velocidade:\%d densidade:\%f ns:\%s concessionaria:\%s`);
    //             } else if(Object.keys(globalData).length > 0 && !table[row][1]) {
    //                 // throw Error("just stop");
    //                 break;
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


function excelDateToDate(excelDate){
    //var result = new moment(Date.UTC(1899, 11, 30)); //to offset to Unix epoch and multiply by milliseconds
    var result = new moment("1899-12-31T00:00:00+03:00"); //to offset to Unix epoch and multiply by milliseconds
    result.add({ days: excelDate});
    // var localtime = nt.localtime( result.toDate().getTime() / 1000);
    // if ( localtime['isDaylightSavings'] == true) {
    //     result.subtract({ hours: 1 });
    // }
    return result;
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
