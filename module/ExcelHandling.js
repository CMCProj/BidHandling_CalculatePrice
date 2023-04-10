"use strict";
var __createBinding = (this && this.__createBinding) || (Object.create ? (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    var desc = Object.getOwnPropertyDescriptor(m, k);
    if (!desc || ("get" in desc ? !m.__esModule : desc.writable || desc.configurable)) {
      desc = { enumerable: true, get: function() { return m[k]; } };
    }
    Object.defineProperty(o, k2, desc);
}) : (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    o[k2] = m[k];
}));
var __setModuleDefault = (this && this.__setModuleDefault) || (Object.create ? (function(o, v) {
    Object.defineProperty(o, "default", { enumerable: true, value: v });
}) : function(o, v) {
    o["default"] = v;
});
var __importStar = (this && this.__importStar) || function (mod) {
    if (mod && mod.__esModule) return mod;
    var result = {};
    if (mod != null) for (var k in mod) if (k !== "default" && Object.prototype.hasOwnProperty.call(mod, k)) __createBinding(result, mod, k);
    __setModuleDefault(result, mod);
    return result;
};
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.ExcelHandling = void 0;
var fs = __importStar(require("fs"));
var exceljs_1 = __importDefault(require("exceljs"));
var exceljs_2 = __importDefault(require("exceljs"));
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
            return exceljs_1.default.read(stream, { type: 'buffer' });
        }
        else if (version === '.xlsx') {
            return exceljs_2.default.read(stream, { type: 'buffer' });
        }
    };
    ExcelHandling.WriteExcel = function (workbook, filepath) {
        var file = fs.createWriteStream(filepath);
        workbook.Write(file);
    };
    return ExcelHandling;
}());
exports.ExcelHandling = ExcelHandling;
