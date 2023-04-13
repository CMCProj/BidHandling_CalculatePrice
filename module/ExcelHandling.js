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
Object.defineProperty(exports, "__esModule", { value: true });
exports.ExcelHandling = void 0;
const fs = __importStar(require("fs"));
// import xls from 'exceljs'
// import { Xlsx } from 'exceljs'
const exceljs = __importStar(require("exceljs"));
const node_xj = __importStar(require("xls-to-json"));
const XLSX = __importStar(require("xlsx"));
class ExcelHandling {
    static GetRow(sheet, rownum) {
        let row = sheet.getRow(rownum); //rownum에 행이 있으면 그 행을 반환하고, 없으면 그 위치에 새로운 빈 행을 만듦
        return row;
    }
    static GetCell2(row, cellnum) {
        let cell = row.getCell(cellnum);
        // if (cell == null) {
        //     cell = row.getCell(cellnum);
        // }
        return cell;
    }
    static GetCell(sheet, rownum, cellnum) {
        let row = this.GetRow(sheet, rownum);
        return this.GetCell2(row, cellnum);
    }
    static GetWorkbook(filename, version) {
        // 파일을 열고 파일 내용을 읽기/쓰기용 스트림으로 가져옴
        let workbook = new exceljs.Workbook();
        let stream = fs.createReadStream(filename + version, { flags: 'r+' });
        if (version === '.xls') {
            node_xj({ input: filename + version, output: filename + '.json' }, function (err, result) {
                if (err)
                    throw err;
                else {
                    //xls -> json으로 변환
                    const readJson = fs.readFileSync(filename + '.json', 'utf-8');
                    let xlsx = JSON.parse(readJson);
                    //json -> xlsx로 변환
                    const sheet = XLSX.utils.json_to_sheet(xlsx);
                    const book = XLSX.utils.book_new();
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
    }
    static WriteExcel(workbook, filepath) {
        const file = fs.createWriteStream(filepath);
        workbook.xlsx.write(file);
    }
}
exports.ExcelHandling = ExcelHandling;
//ExcelHandling.GetWorkbook('입찰내역_설비공사', '.xls');
