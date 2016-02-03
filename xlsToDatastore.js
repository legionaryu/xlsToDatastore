// TODO: Check valid data using the folder month to compare with the date on the cell

var fs = require("fs")
var pathjs = require("path")
var xlsx = require('node-xlsx');
//var nt = require('time');
var moment = require('moment-timezone');

var outputDataPath = './historicoArtesp';
// var xlsDataPath = './xls';//'/servidor/Google IS/Materiais/ARTESP/DONE';
// var xlsDataPath = '/servidor/Google IS/Materiais/ARTESP/DONE';
// var xlsDataPath = '/servidor/Google IS/Materiais/ARTESP/DONE/ECOPISTAS';
var xlsDataPath = '/servidor/Google IS/Materiais/ARTESP/DONE/ECOVIAS';
// var xlsDataPath = "\\\\10.0.1.150\\servidor\\Google IS\\Materiais\\ARTESP\\DONE\\ECOPISTAS";
var historyJsonPath = './historyJson.json';
var warningFilesPath = './warningFiles.txt';
var dataSpreadsheet;
var historyJson;

Date.prototype.UTCyyyymmdd = function() {
    var yyyy = this.getUTCFullYear().toString();
    var mm = (this.getUTCMonth()+1).toString(); // getMonth() is zero-based
    var dd  = this.getUTCDate().toString();
    return yyyy + (mm.length===2?mm:"0"+mm[0]) + (dd.length===2?dd:"0"+dd[0]); // padding
};

function getAllXls(filePath, callback){
    if(!filePath) {
        var result = getAllXls(xlsDataPath, callback);
        return result;
    } else {
        console.log("filePath:", filePath);
        historyJson.path = filePath;
        // fs.writeFileSync(historyJsonPath, JSON.stringify(historyJson, null, " "));
        var stat = fs.statSync(filePath);
        if(stat.isDirectory()){
            historyJson.dirRead = historyJson.dirRead || [];
            var pathSplit = filePath.split(pathjs.sep);
            // if(pathSplit.length > 9 && (pathjs.basename(filePath) != "SP160" || pathjs.basename(filePath) != "SP150")) return;
            if(pathSplit.length > 9 && (pathjs.basename(filePath) != "SP160")) return;
            if(historyJson.dirRead.indexOf(filePath) < 0 && pathjs.basename(filePath)[0] != ".") {
                historyJson.dirPath = filePath;
                var allFiles = fs.readdirSync(filePath);
                // fs.writeFileSync(historyJsonPath, JSON.stringify(historyJson, null, " "));
                allFiles.forEach(function(file){
                    result = getAllXls(pathjs.join(filePath, file), callback);
                });
            } else {
                historyJson.dirPath = pathjs.dirname(filePath);
                // fs.writeFileSync(historyJsonPath, JSON.stringify(historyJson, null, " "));
            }
        } else if (stat.isFile()) {
            historyJson.filesRead = historyJson.filesRead || [];
            if(pathjs.basename(filePath).indexOf("61-65") < 0 && pathjs.basename(filePath).indexOf("11-16") < 0) return;
            if(historyJson.filesRead.indexOf(filePath) < 0 && pathjs.basename(filePath)[0] != "." && pathjs.extname(filePath) != ".db") {
                if (typeof(callback) === 'function') callback(filePath);
                // historyJson.filesRead.push(filePath);
                // fs.writeFileSync(historyJsonPath, JSON.stringify(historyJson, null, " "));
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
    // |   |   |── RODOVIA/
    // |   |   |   ├── TRECHO.(xls, xlsm, xlsx)
    // ----------------------------------------------------------

    // --------------- Output Folder Structure ------------------
    // RODOVIA_CODIGO/
    // |── TRECHO/
    // |   |── SENTIDO/
    // |   |   ├── YYYYMMDD.log
    // ----------------------------------------------------------

    getAllXls(xlsDataPath, function(filePath) {
    // getAllXls("/servidor/Google IS/Materiais/ARTESP/DONE/ECOPISTAS/2013/09-Setembro/SP070/112-130-RAMPAS-3.xlsx", function(filePath) {
        var countAnalise = 0;
        var filename = pathjs.basename(filePath);
        filename = filename.substr(0, filename.length - pathjs.extname(filePath).length);
        var obj = xlsx.parse(filePath);
        Object.keys(obj).every(function (value){
            var tabName = obj[value].name;
            if (tabName.toUpperCase().indexOf("ANÁLISE") < 0 && 
                tabName.toUpperCase().indexOf("ANALISE") < 0 && 
                tabName.toUpperCase().indexOf("TH") < 0 && 
                tabName.toUpperCase().indexOf("RAMPA") < 0 )
            {
                return true;
            }
            else {
                var result = ++countAnalise <= 2;
                var repeatedZeroCount = 0;
                var logFilenameList = [];
                var globalData = {};
                var table = obj[value].data;
                for (var row = 0; row < table.length; row++) {
                    if ((table[row][1]+"").toUpperCase() == "HORA" && 
                        (table[row][2]+"").toUpperCase() == "DATA" && 
                        (table[row][3]+"").toUpperCase() == "VOLUME LEVANTADO")
                    {
                        var relativePath = pathjs.relative(xlsDataPath, filePath);
                        var relativeSplit = relativePath.split(pathjs.sep);
                        globalData.road = relativeSplit[relativeSplit.length - 2].replace(/\./g, '_');
                        globalData.stretch = filename.replace(/\./g, '_');

                        if(tabName.toUpperCase().indexOf("INTERNA") >= 0 || 
                            tabName.toUpperCase().indexOf("INT") >= 0 || 
                           (table[row-12][2]+"").toUpperCase().indexOf("INTERNA") >= 0 ||
                           (table[row-11][2]+"").toUpperCase().indexOf("INTERNA") >= 0)
                        {
                            globalData.direction = "Interna";
                        }
                        else if(tabName.toUpperCase().indexOf("EXTERNA") >= 0 || 
                            tabName.toUpperCase().indexOf("EXT") >= 0 || 
                           (table[row-12][2]+"").toUpperCase().indexOf("EXTERNA") >= 0 || 
                           (table[row-11][2]+"").toUpperCase().indexOf("EXTERNA") >= 0)
                        {
                            globalData.direction = "Externa";
                        }
                        else if(tabName.toUpperCase().indexOf("LESTE") >= 0 || 
                           (table[row-12][2]+"").toUpperCase().indexOf("LESTE") >= 0)
                        {
                            globalData.direction = "Leste";
                        }
                        else if(tabName.toUpperCase().indexOf("OESTE") >= 0 || 
                           (table[row-12][2]+"").toUpperCase().indexOf("OESTE") >= 0)
                        {
                            globalData.direction = "Oeste";
                        }
                        else if(tabName.toUpperCase().indexOf("NORTE") >= 0 || 
                           (table[row-12][2]+"").toUpperCase().indexOf("NORTE") >= 0)
                        {
                            globalData.direction = "Norte";
                        }
                        else if(tabName.toUpperCase().indexOf("SUL") >= 0 || 
                           (table[row-12][2]+"").toUpperCase().indexOf("SUL") >= 0)
                        {
                            globalData.direction = "Sul";
                        }
                        else
                        {
                            throw Error("Couldn't find direction on the file " + filePath);
                        }
                        globalData.dealership = relativeSplit[0];
                        row += 1;
                        continue;
                    }
                    else if(Object.keys(globalData).length > 0 && parseInt(table[row][1])){
                        var dataRow = table[row];
                        if(dataRow[2] == '-' || !parseInt(dataRow[2])) {
                            break;
                        }
                        var dateReport = excelDateToDate(dataRow[2], ((parseInt(dataRow[1])-1)%24));
                        if (dateReport.year() < 2000) {
                            var pathSplit = filePath.split(pathjs.sep);
                            var month = parseInt(pathSplit[8]);
                            dateReport = excelDateToDate("1/" + month + "/" + pathSplit[7], (parseInt(dataRow[1])-1));
                             if (dateReport.year() < 2000) {
                                console.log(dataRow);
                                throw Error("Invalid year on the file " + filePath + " | table:" + tabName + " row:" + row);
                            }
                            else if(dateReport.month()+1 != month) {
                                break;
                            }
                        }
                        // console.log("zeros: %s | repeatedZeroCount: %d", (!dataRow[3] && !dataRow[4] && !dataRow[5]), repeatedZeroCount);
                        if (!dataRow[3] && !dataRow[4] && !dataRow[5]) {
                            if(repeatedZeroCount > 300) {
                                logFilenameList.forEach(function (file) {
                                    fs.unlinkSync(file);
                                });
                                fs.appendFileSync(warningFilesPath, filePath + " | too many zeros " + "\r\n");
                                return result;
                            }
                            repeatedZeroCount += 1;
                        } else {
                            repeatedZeroCount = 0;
                        }
                        // dateReport.setUTCHours(dateReport.getUTCHours()+(parseInt(dataRow[1])-1)%24);
                        // dateReport.add({hours:((parseInt(dataRow[1])-1)%24)});
                        // console.log(globalData);
                        // console.log("%d, 2, 9, LOG, %s, %d, %d, %d, %d, %d, %d, %s, %s, passeio:\%d comercial:\%d tx_fluxo:\%d vp:\%d velocidade:\%d densidade:\%f ns:\%s concessionaria:\%s",
                        //             dateReport.unix(), filename, dataRow[3], dataRow[4], dataRow[5], dataRow[6], dataRow[7], dataRow[8], dataRow[9], globalData.concessionaria);
                        var outputDir = pathjs.join(outputDataPath, globalData.road, globalData.stretch, globalData.direction);
                        var outputFilename = pathjs.join(outputDir, dateReport.format("YYYYMMDD[.log]"));
                        if(logFilenameList.indexOf(outputFilename) < 0) {
                            logFilenameList.push(outputFilename);
                        }
                        console.log(outputFilename);
                        // console.log(globalData);
                        makeDir(outputDir);
                        var roadClass = Math.max(Math.min((dataRow[9]+"").toUpperCase().charCodeAt(0) - 64, 7), 0) || 0;
                        var dataLog = `${dateReport.valueOf()}, 2, 7, LOG, ${globalData.road} ${globalData.stretch} ${globalData.direction},`;
                        for(var k=3; k < 9; k++){
                             dataLog += `${dataRow[k] || 0},`;
                        }
                        dataLog +=  `${roadClass}, passeio:\%d comercial:\%d tx_fluxo:\%d vp:\%d velocidade:\%d densidade:\%f ns:\%d [A-E] concessionaria:${globalData.dealership}\r\n`;
                        fs.appendFileSync(outputFilename, dataLog);
                        // console.log(`${dateReport.valueOf()}, 2, 9, LOG, ${filename}, ${dataRow[3]}, ${dataRow[4]}, ${dataRow[5]}, ${dataRow[6]}, ${dataRow[7]}, ${dataRow[8]}, ${dataRow[9]}, ${globalData.dealership}, passeio:\%d comercial:\%d tx_fluxo:\%d vp:\%d velocidade:\%d densidade:\%f ns:\%s concessionaria:\%s\n`);
                    } else if(Object.keys(globalData).length > 0 && !table[row][1]) {
                        // throw Error("just stop");
                        break;
                    }
                }
                return result;
            }
        });
        if (countAnalise > 0 && countAnalise < 2) {
            fs.appendFileSync(warningFilesPath, filePath + " | missing one tab " + "\r\n");
        }
        else if (countAnalise <= 0) {
            throw Error("Couldn't find the data tab on the file " + filePath);
        }
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
    var reg = new RegExp('^[0-9]$'); //number only string
    var result = new moment("1899-12-31").tz('America/Sao_Paulo'); //to offset to Unix epoch and multiply by milliseconds
    if(typeof(excelDate) == "number" || reg.test(excelDate)) {
        result.add({ days: excelDate});
    }
    else {
        result = new moment(excelDate, "DD/MM/YYYY").tz('America/Sao_Paulo');
    }
    result.hour(excelHour);
    result.minute(0);
    result.second(0);

    result.utc(); //Convert to GMT-0

    // console.log(result.toString());

    return result;
}

function makeDir(str_path) {
    try {
        fs.mkdirSync(str_path);
    } catch(ex) {
        // console.log(JSON.stringify(ex));
        if(ex.code != "EEXIST") {
            if(ex.code == "ENOENT") {
                var dirPathSplit = str_path.split(pathjs.sep);
                var dirPath = "";
                dirPathSplit.forEach(function (value) {
                    dirPath = pathjs.join(dirPath, value);
                    makeDir(dirPath);
                });
            } else {
                console.log(JSON.stringify(ex));
                throw ex;
            }
        }
    }
}
