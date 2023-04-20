"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.ExcelHandling = void 0;
var fs = require("fs");
// import xls from 'exceljs'
// import { Xlsx } from 'exceljs'
var exceljs = require("exceljs");
var node_xj = require("xls-to-json");
var XLSX = require("xlsx");
var ExcelHandling = /** @class */ (function () {
    function ExcelHandling() {
    }
    ExcelHandling.GetRow = function (sheet, rownum) {
        var row = sheet.getRow(rownum); //rownum에 행이 있으면 그 행을 반환하고, 없으면 그 위치에 새로운 빈 행을 만듦
        return row;
    };
    ExcelHandling.GetCell2 = function (row, cellnum) {
        var cell = row.getCell(cellnum);
        // if (cell == null) {
        //     cell = row.getCell(cellnum);
        // }
        return cell;
    };
    /**해당 워크시트의 행, 열의 값을 return */
    ExcelHandling.GetCell = function (sheet, rownum, cellnum) {
        var row = this.GetRow(sheet, rownum);
        return this.GetCell2(row, cellnum);
    };
    /** exceljs의 read 메소드가 비동기 메소드이기에 return형이 Promise<exceljs.Workbook>임.
     *
     * 이를 사용하는 함수 혹은 메소드는 비동기(async / await)로 만들어야 제대로 된 exceljs.Workbook으로 쓸 수 있음.*/
    ExcelHandling.GetWorkbook = function (filename, version) {
        // 파일을 열고 파일 내용을 읽기/쓰기용 스트림으로 가져옴
        var workbook = new exceljs.Workbook();
        var stream = fs.createReadStream(filename + version, { flags: 'r+' });
        if (version === '.xls') {
            node_xj({ input: filename + version, output: filename + '.json' }, function (err, result) {
                if (err)
                    throw err;
                else {
                    //xls -> json으로 변환
                    var readJson = fs.readFileSync(filename + '.json', 'utf-8');
                    var xlsx = JSON.parse(readJson);
                    //json -> xlsx로 변환
                    var sheet = XLSX.utils.json_to_sheet(xlsx);
                    var book = XLSX.utils.book_new();
                    XLSX.utils.book_append_sheet(book, sheet, 'Sheet1');
                    XLSX.writeFile(book, filename + '.xlsx');
                    fs.rmSync(filename + '.json');
                    stream = fs.createReadStream(filename + '.xlsx', { flags: 'r+' });
                }
            });
            return workbook.xlsx.read(stream);
        }
        else if (version === '.xlsx')
            return workbook.xlsx.read(stream);
        throw new Error('올바른 Excel파일 형식(.xls/.xlsx)이 아닙니다. 파일을 다시 한 번 확인해주세요.');
    };
    ExcelHandling.WriteExcel = function (workbook, filepath) {
        var file = fs.createWriteStream(filepath);
        workbook.xlsx.write(file);
    };
    return ExcelHandling;
}());
exports.ExcelHandling = ExcelHandling;
//ExcelHandling.GetWorkbook('입찰내역_설비공사', '.xls');
