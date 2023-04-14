"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.CreateResultFile = void 0;
var Data_1 = require("./Data");
var path = require("path");
var fs = require("fs");
var exceljs = require("exceljs");
var CreateResultFile = /** @class */ (function () {
    function CreateResultFile() {
    }
    CreateResultFile.Create = function () {
        var directoryPath = Data_1.Data.folder; // 파일 경로
        var extension = '.xls'; // 파일 확장자 명
        var xlsList = []; // 파일 경로에 있는 엑셀파일명을 담을 리스트
        fs.readdir(directoryPath, function (err, files) {
            // 파일 경로에 있는 엑셀파일의 절대경로를 모두 저장한다.
            if (err) {
                return console.log('Unable to scan directory: ' + err);
            }
            files.forEach(function (file) {
                if (path.extname(file) === extension) {
                    xlsList.push(path.join(directoryPath, file));
                }
            });
        });
        xlsList.forEach(function (xls) {
            // 기존의 엑셀 파일을 모두 삭제한다.
            fs.access(xls, fs.constants.F_OK, function (err) {
                if (err)
                    return console.log('error(파일 삭제 실패)');
                fs.unlink(xls, function (err) { return (err ? console.log(err) : console.log('${xls} 삭제 완료')); });
            });
        });
        var worksheet = new exceljs.Workbook();
        Data_1.Data.Dic.forEach(function (dic) {
            var workbook = worksheet.xlsx.readFile('입찰내역.xls');
            var sheet = workbook.GetSheetAt(0);
            var resultPath;
            var path;
            for (var i = 0; i < dic.values.length; i++) {
                if (dic.values[i].Item === '일반') {
                    //엑셀핸들링
                }
            }
        });
    };
    return CreateResultFile;
}());
exports.CreateResultFile = CreateResultFile;
