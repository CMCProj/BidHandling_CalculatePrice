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
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.CreateResultFile = void 0;
const Data_1 = require("./Data");
const path = __importStar(require("path"));
const fs = __importStar(require("fs"));
const ExcelHandling_1 = require("./ExcelHandling");
class CreateResultFile {
    static Create() {
        const directoryPath = Data_1.Data.folder; // 파일 경로
        const extension = '.xls'; // 파일 확장자 명
        const xlsList = []; // 파일 경로에 있는 엑셀파일명을 담을 리스트
        fs.readdir(directoryPath, function (err, files) {
            // 파일 경로에 있는 엑셀파일의 절대경로를 모두 저장한다.
            if (err) {
                return console.log('Unable to scan directory: ' + err);
            }
            files.forEach((file) => {
                if (path.extname(file) === extension) {
                    xlsList.push(path.join(directoryPath, file));
                }
            });
        });
        xlsList.forEach((xls) => {
            // 기존의 엑셀 파일을 모두 삭제한다.
            fs.access(xls, fs.constants.F_OK, (err) => {
                if (err)
                    return console.log('error(파일 삭제 실패)');
                fs.unlink(xls, (err) => (err ? console.log(err) : console.log('${xls} 삭제 완료')));
            });
        });
        Data_1.Data.Dic.forEach((value, key) => __awaiter(this, void 0, void 0, function* () {
            let workbook = yield ExcelHandling_1.ExcelHandling.GetWorkbook("입찰내역.xls", "xls");
            let sheet = workbook.getWorksheet(0);
            let resultPath;
            let Path;
            for (let i = 0; i < value.length; i++) {
                if (value[i].Item === '일반') {
                    sheet.getCell(i, 22).value = Number(value[i].PriceScore);
                    sheet.getCell(i, 23).value = Number(value[i].Score);
                }
                let materialUnit = Number(value[i].MaterialUnit);
                let laborUnit = Number(value[i].LaborUnit);
                let expenseUnit = Number(value[i].ExpenseUnit);
                let material = Number(value[i].Material);
                let labor = Number(value[i].Labor);
                let expense = Number(value[i].Expense);
                let unitpricesum = Number(value[i].UnitPriceSum);
                let pricesum = Number(value[i].PriceSum);
                sheet.getCell(i + 1, 1).value = (i + 1) * 100;
                sheet.getCell(i + 1, 3).value = value[i].Name;
                sheet.getCell(i + 1, 4).value = value[i].Standard;
                sheet.getCell(i + 1, 5).value = value[i].Unit;
                sheet.getCell(i + 1, 6).value = Number(value[i].Quantity);
                sheet.getCell(i + 1, 7).value = materialUnit;
                sheet.getCell(i + 1, 8).value = laborUnit;
                sheet.getCell(i + 1, 9).value = expenseUnit;
                sheet.getCell(i + 1, 10).value = unitpricesum;
                sheet.getCell(i + 1, 11).value = material;
                sheet.getCell(i + 1, 12).value = labor;
                sheet.getCell(i + 1, 13).value = expense;
                sheet.getCell(i + 1, 14).value = pricesum;
                sheet.getCell(i + 1, 18).value = value[i].Code;
                if (value[i].Item === '표준시장단가') {
                    sheet.getCell(i + 1, 16).value = Number(value[i].Item);
                    sheet.getCell(i + 1, 17).value = Number(value[i].Code);
                }
                else {
                    sheet.getCell(i + 1, 15).value = value[i].Item;
                }
            }
            resultPath = '입찰내역_' + Data_1.Data.ConstructionNums[key] + '.xls';
            Path = path.join(Data_1.Data.folder, resultPath);
            ExcelHandling_1.ExcelHandling.WriteExcel(workbook, Path);
        }));
    }
}
exports.CreateResultFile = CreateResultFile;
