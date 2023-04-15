"use strict";
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (g && (g = 0, op[0] && (_ = 0)), _) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.CreateResultFile = void 0;
var Data_1 = require("./Data");
var path = require("path");
var fs = require("fs");
var ExcelHandling_1 = require("./ExcelHandling");
var CreateResultFile = /** @class */ (function () {
    function CreateResultFile() {
    }
    CreateResultFile.Create = function () {
        var _this = this;
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
        Data_1.Data.Dic.forEach(function (value, key) { return __awaiter(_this, void 0, void 0, function () {
            var workbook, sheet, resultPath, Path, i, materialUnit, laborUnit, expenseUnit, material, labor, expense, unitpricesum, pricesum;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, ExcelHandling_1.ExcelHandling.GetWorkbook("입찰내역.xls", "xls")];
                    case 1:
                        workbook = _a.sent();
                        sheet = workbook.getWorksheet(0);
                        for (i = 0; i < value.length; i++) {
                            if (value[i].Item === '일반') {
                                sheet.getCell(i, 22).value = Number(value[i].PriceScore);
                                sheet.getCell(i, 23).value = Number(value[i].Score);
                            }
                            materialUnit = Number(value[i].MaterialUnit);
                            laborUnit = Number(value[i].LaborUnit);
                            expenseUnit = Number(value[i].ExpenseUnit);
                            material = Number(value[i].Material);
                            labor = Number(value[i].Labor);
                            expense = Number(value[i].Expense);
                            unitpricesum = Number(value[i].UnitPriceSum);
                            pricesum = Number(value[i].PriceSum);
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
                        return [2 /*return*/];
                }
            });
        }); });
    };
    return CreateResultFile;
}());
exports.CreateResultFile = CreateResultFile;
