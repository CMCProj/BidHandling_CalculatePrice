"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.ExcelHandling = void 0;
var fs = require("fs");
var ExcelHandling = /** @class */ (function () {
    function ExcelHandling() {
    }
    ExcelHandling.GetRow = function (sheet, rownum) {
        var row = sheet.getRow(rownum);
        if (row == null) {
            row = sheet.addRow(rownum);
        }
        return row;
    };
    ExcelHandling.GetCell2 = function (row, cellnum) {
        var cell = row.GetCell(cellnum);
        if (cell == null) {
            cell = row.addCell(cellnum);
        }
        return cell;
    };
    ExcelHandling.GetCell = function (sheet, rownum, cellnum) {
        var row = this.GetRow(sheet, rownum);
        return this.GetCell2(row, cellnum);
    };
    ExcelHandling.GetWorkbook = function (filename, version) {
        // 파일을 열고 파일 내용을 읽기/쓰기용 스트림으로 가져옴
        var stream = fs.createReadStream(filename, { flags: 'r+' });
        if (version === '.xls') {
            // return xls.read(stream, { type: 'buffer' })
        }
        else if (version === '.xlsx') {
            // return Xlsx.read(stream, { type: 'buffer' })
        }
    };
    ExcelHandling.WriteExcel = function (workbook, filepath) {
        var file = fs.createWriteStream(filepath);
        workbook.Write(file);
    };
    return ExcelHandling;
}());
exports.ExcelHandling = ExcelHandling;
