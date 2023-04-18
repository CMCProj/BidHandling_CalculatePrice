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
exports.Setting = void 0;
var Data_1 = require("./Data");
var ExcelHandling_1 = require("./ExcelHandling");
//import Enumerable from 'linq'
var fs = require("fs");
//4.14 Setting.SetUnitPriceNoExcel에서 값 입력되지 않음.
var Setting = /** @class */ (function () {
    function Setting() {
    }
    Setting.GetData = function () {
        var bidString = fs.readFileSync(Data_1.Data.folder + '\\OutputDataFromBID.json', 'utf-8');
        Setting.docBID = JSON.parse(bidString);
        Setting.eleBID = this.docBID['data'];
        //세부공사별 번호 Data.ConstructionNums 딕셔너리에 저장
        Setting.GetConstructionNum();
        //세부공사별 리스트 생성(Dic -> key : 세부공사별 번호 / value : 세부공사별 리스트)
        Setting.AddConstructionList();
        //공내역 xml 파일 읽어들여 데이터 저장
        Setting.GetDataFromBID();
        if (Data_1.Data.XlsFiles !== undefined) {
            //실내역으로부터 Data 객체에 단가세팅
            Setting.SetUnitPrice();
        }
        else {
            Setting.SetUnitPriceNoExcel(); //실내역 없이 공내역만으로 Data 객체에 단가 세팅 (23.02.06)
        }
        //고정금액 및 적용비율 저장
        Setting.GetRate();
        //직공비 제외항목 및 고정금액 계산
        Setting.GetPrices();
        //표준시장단가 합계(조사금액) 저장
        Data_1.Data.InvestigateStandardMarket =
            Data_1.Data.StandardMaterial + Data_1.Data.StandardLabor + Data_1.Data.StandardExpense;
        //console.log(Data.StandardMaterial);
    };
    Setting.GetConstructionNum = function () {
        var constNums = Setting.eleBID['T2'];
        var index;
        var construction;
        for (var num in constNums) {
            index = constNums[num]['C1']['_text'];
            construction = constNums[num]['C3']['_text'];
            if (Data_1.Data.ConstructionNums.has(construction)) {
                construction += '2';
            }
            // console.log(
            //     '키:',
            //     index,
            //     '값:',
            //     construction,
            //     '으로 Data.ConstructionNums에 <string, string> 추가'
            // )
            Data_1.Data.ConstructionNums.set(index, construction);
        }
    };
    Setting.AddConstructionList = function () {
        Data_1.Data.ConstructionNums.forEach(function (value, key) {
            Data_1.Data.Dic.set(key, new Array());
            //console.log('키:', key, '값:', Data.Dic.get(key), '으로 Data.Dic에 <string, []> 추가')
        }); // Data.Dic자료구조 체크하기 <string. Data[]>가 맞는지, <string, Array[]>로 해야하는지
    };
    Setting.GetItem = function (bid) {
        var item = null;
        //해당 공종이 일반, 표준시장단가 및 공종(입력불가) 항목인 경우
        if (bid['C7']['_text'] === '0') {
            if (bid['C5']['_text'] === 'S') {
                if (bid['C10']['_text'] !== undefined)
                    item = '표준시장단가';
                else
                    item = '일반';
            }
            else
                item = '예외'; // null값이 발생해 예외 추가함
        }
        //해당 공종이 무대(입력불가)인 경우
        else if (bid['C7']['_text'] === '1')
            item = '공종(입력불가)';
        //해당 공종이 관급자재인 경우
        else if (bid['C7']['_text'] === '2')
            item = '관급자재';
        else if (bid['C7']['_text'] === '3')
            item = '관급공종';
        //해당 공종이 PS인 경우
        else if (bid['C7']['_text'] === '4')
            item = 'PS';
        //해당 공종이 제요율적용제외공종인 경우
        else if (bid['C7']['_text'] === '5')
            item = '재요율적용제외';
        //해당 공종이 PS내역인 경우
        else if (bid['C7']['_text'] === '7')
            item = 'PS내역';
        //해당 공종이 음의 가격 공종인 경우
        else if (bid['C7']['_text'] === '19')
            item = '음(-)의 입찰금액';
        //해당 공종이 PS(안전관리비)인 경우
        else if (bid['C7']['_text'] === '20')
            item = 'PS(안전관리비)';
        //해당 공종이 작업설인 경우
        else if (bid['C7']['_text'] === '22')
            item = '작업설';
        else
            item = '예외';
        return item;
    };
    Setting.GetDataFromBID = function () {
        // 수정 필요
        //4.10 주석처리한 부분 수정함
        var works = Setting.eleBID['T3'].map(function (work) {
            return new Data_1.Data(Setting.GetItem(work), work.C1._text, //null값 가능성O
            work.C2._text, //null값 가능성O
            work.C3._text, //null값 가능성O
            work.C9._text, //null값 가능성O
            work.C12._text, //null값 가능성O
            work.C13._text, //null값 가능성O
            work.C14._text, //null값 가능성O
            Number(work.C15._text), Number(work.C28._text), Number(work.C29._text), Number(work.C30._text));
        });
        works.forEach(function (work) {
            Data_1.Data.Dic.get(work.ConstructionNum).push(work);
            // console.log('Data.Dic의', 'arrray에', Data.Dic.get(work.ConstructionNum)[Data.Dic.get(work.ConstructionNum).length - 1], '추가')
        });
    };
    Setting.MatchConstructionNum = function (filePath) {
        return __awaiter(this, void 0, void 0, function () {
            var workbook, sheet, check //실내역 파일과 세부공사의 데이터가 일치하는 횟수
            ;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, ExcelHandling_1.ExcelHandling.GetWorkbook(filePath, '.xlsx')];
                    case 1:
                        workbook = _a.sent();
                        sheet = workbook.getWorksheet('내역서');
                        Data_1.Data.Dic.forEach(function (value, key) {
                            check = 0;
                            for (var i = 0; i < 5; i++) {
                                var row = sheet.getRow(i + 4);
                                var sameName = value[i].Name === row.getCell(4).value.toString();
                                if (sameName)
                                    check++;
                                if (check == 3) {
                                    Data_1.Data.MatchedConstNum.set(filePath, key);
                                    return;
                                }
                            }
                        });
                        Data_1.Data.IsFileMatch = false;
                        return [2 /*return*/];
                }
            });
        });
    };
    Setting.CopyFile = function (filePath) {
        return __awaiter(this, void 0, void 0, function () {
            var workbook, sheet, constNum, lastRowNum, rowIndex;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, ExcelHandling_1.ExcelHandling.GetWorkbook(filePath, '.xlsx')];
                    case 1:
                        workbook = _a.sent();
                        sheet = workbook.getWorksheet('내역서');
                        constNum = Data_1.Data.MatchedConstNum[filePath];
                        lastRowNum = sheet.rowCount;
                        rowIndex = 4;
                        Data_1.Data.Dic[constNum].forEach(function (curObj) {
                            var dcode = curObj.Code;
                            if (dcode == null || dcode.trim() === '') {
                                return;
                            }
                            var dname = curObj.Name;
                            var dunit = curObj.Unit;
                            var dquantity = curObj.Quantity;
                            while (true) {
                                var row = sheet.getRow(rowIndex);
                                var code = row.getCell(1).value.toString();
                                var name_1 = row.getCell(4).value.toString();
                                var unit = row.getCell(6).value.toString();
                                var quantity = 0.0;
                                try {
                                    var quantity_1 = Number(row.getCell(7).value);
                                }
                                catch (_a) {
                                    rowIndex++;
                                    if (rowIndex == lastRowNum)
                                        break;
                                    continue;
                                }
                                var sameCode = code === dcode;
                                var sameName = name_1 === dname;
                                var sameUnit = unit === dunit;
                                var sameQuantity = quantity === dquantity;
                                if ((sameName || sameCode) && (sameUnit || sameQuantity)) {
                                    curObj.MaterialUnit = Number(row.getCell(8).value);
                                    curObj.LaborUnit = Number(row.getCell(10).value);
                                    curObj.ExpenseUnit = Number(row.getCell(12).value);
                                    rowIndex++;
                                    break;
                                }
                                else {
                                    rowIndex++;
                                    if (rowIndex == lastRowNum)
                                        break;
                                    continue;
                                }
                            }
                        });
                        return [2 /*return*/];
                }
            });
        });
    };
    Setting.SetUnitPrice = function () {
        // 미완성(가장 마지막에 작성 예정)
        // let copiedFolder: string = Data.folder + "\\Actual Xlsx";
        // const dir = new DirectoryInfo(copiedFolder);
        // const files = dir.GetFiles();
    };
    Setting.SetUnitPriceNoExcel = function () {
        var bidT3 = this.eleBID['T3'];
        var code;
        var type;
        var _loop_1 = function (key) {
            code = JSON.stringify(bidT3[key]['C9']['_text']);
            type = JSON.stringify(bidT3[key]['C5']['_text'])[1];
            if (code !== undefined && type === 'S') {
                var constNum = bidT3[key]['C1']['_text'];
                var numVal_1 = bidT3[key]['C2']['_text'];
                var detailVal_1 = bidT3[key]['C3']['_text'];
                var curObject = Data_1.Data.Dic.get(constNum).find(function (x) { return x.WorkNum === numVal_1 && x.DetailWorkNum === detailVal_1; });
                if (curObject.Item === '일반' || curObject.Item === '재요율적용제외' || curObject.Item === '표준시장단가' && curObject !== undefined) {
                    bidT3[key]['C16']['_text'] = curObject.MaterialUnit.toString();
                    bidT3[key]['C17']['_text'] = curObject.LaborUnit.toString();
                    bidT3[key]['C18']['_text'] = curObject.ExpenseUnit.toString();
                    bidT3[key]['C19']['_text'] = curObject.UnitPriceSum.toString();
                    bidT3[key]['C20']['_text'] = curObject.Material.toString();
                    bidT3[key]['C21']['_text'] = curObject.Labor.toString();
                    bidT3[key]['C22']['_text'] = curObject.Expense.toString();
                    bidT3[key]['C23']['_text'] = curObject.PriceSum.toString();
                }
            }
        };
        for (var key in bidT3) {
            _loop_1(key);
        }
        if (fs.existsSync(Data_1.Data.work_path + '\\Setting_Json.json')) {
            fs.unlink(Data_1.Data.work_path + '\\Setting_Json.json', function (err) {
                if (err && err.code == 'ENOENT') {
                    console.log('파일 삭제 Error 발생');
                }
            });
        }
        fs.writeFileSync(Data_1.Data.work_path + '\\Setting_Json.json', JSON.stringify(this.docBID));
    };
    Setting.GetRate = function () {
        Data_1.Data.ConstructionTerm = Number(Setting.eleBID['T1']['C29']['_text']);
        var bidT5 = Setting.eleBID['T5'];
        //console.log()
        for (var key in bidT5) {
            var name_2 = JSON.stringify(bidT5[key]['C4']['_text']);
            var val1 = bidT5[key]['C6']['_text'];
            var val2 = bidT5[key]['C7']['_text'];
            if (bidT5[key]['C5']['_text'] === '7') {
                var fixedPrice = Number(bidT5[key]['C8']['_text']);
                Data_1.Data.Fixed.set(name_2, fixedPrice);
            }
            else {
                var applicationRate1 = Number(val1);
                var applicationRate2 = Number(val2);
                Data_1.Data.Rate1.set(name_2, applicationRate1);
                Data_1.Data.Rate2.set(name_2, applicationRate2);
            }
        }
    };
    Setting.GetPrices = function () {
        for (var _i = 0, _a = Array.from(Data_1.Data.Dic.values()); _i < _a.length; _i++) {
            var value = _a[_i];
            for (var _b = 0, value_1 = value; _b < value_1.length; _b++) {
                var item = value_1[_b];
                //console.log(item)
                if (item.Item === '관급') {
                    Data_1.Data.GovernmentMaterial += item.Material;
                    Data_1.Data.GovernmentLabor += item.Labor;
                    Data_1.Data.GovernmentExpense += item.Expense;
                }
                else if (item.Item === 'PS') {
                    Data_1.Data.PsMaterial += item.Material;
                    Data_1.Data.PsLabor += item.Labor;
                    Data_1.Data.PsExpense += item.Expense;
                }
                else if (item.Item === '제요율적용제외') {
                    Data_1.Data.ExcludingMaterial += item.Material;
                    Data_1.Data.ExcludingLabor += item.Labor;
                    Data_1.Data.ExcludingExpense += item.Expense;
                }
                //해당 공종이 PS(안전관리비)인 경우
                else if (item.Item === 'PS(안전관리비)') {
                    Data_1.Data.SafetyPrice += item.Expense;
                    //PS(안전관리비) 고정 단가에 추가 (23.02.06)
                    Data_1.Data.FixedPriceDirectMaterial += item.Material; //재료비 합 계산
                    Data_1.Data.FixedPriceDirectLabor += item.Labor; //노무비 합 계산
                    Data_1.Data.FixedPriceOutputExpense += item.Expense; //경비 합 계산
                    //직공비에 해당하는 각 객체의 재료비, 노무비, 경비를 직접재료비, 직접노무비, 산출 경비에 더해나감
                    Data_1.Data.RealDirectMaterial += item.Material;
                    Data_1.Data.RealDirectLabor += item.Labor;
                    Data_1.Data.RealOutputExpense += item.Expense;
                }
                //표준시장단가 품목인지 확인
                else if (item.Item === '표준시장단가') {
                    Data_1.Data.FixedPriceDirectMaterial += item.Material; //재료비 합 계산
                    Data_1.Data.FixedPriceDirectLabor += item.Labor; //노무비 합 계산
                    Data_1.Data.FixedPriceOutputExpense += item.Expense; //경비 합 계산
                    //직공비에 해당하는 각 객체의 재료비, 노무비, 경비를 직접재료비, 직접노무비, 산출 경비에 더해나감
                    Data_1.Data.RealDirectMaterial += item.Material;
                    Data_1.Data.RealDirectLabor += item.Labor;
                    Data_1.Data.RealOutputExpense += item.Expense;
                    //표준시장 단가 직공비 합계에 더해나감
                    Data_1.Data.StandardMaterial += item.Material;
                    Data_1.Data.StandardLabor += item.Labor;
                    Data_1.Data.StandardExpense += item.Expense;
                }
                //음(-)의 단가 품목인지 확인
                else if (item.Item === '음(-)의 입찰금액') {
                    Data_1.Data.FixedPriceDirectMaterial += item.Material;
                    Data_1.Data.FixedPriceDirectLabor += item.Labor;
                    Data_1.Data.FixedPriceOutputExpense += item.Expense;
                    //직공비에 해당하는 각 객체의 재료비, 노무비, 경비를 직접재료비, 직접노무비, 산출 경비에 더해나감
                    Data_1.Data.RealDirectMaterial += item.Material;
                    Data_1.Data.RealDirectLabor += item.Labor;
                    Data_1.Data.RealOutputExpense += item.Expense;
                }
                //직공비 중, 고정금액이 아닌 일반 항목들의 직공비 계산
                else if (item.Item === '일반') {
                    Data_1.Data.RealPriceDirectMaterial += item.Material;
                    Data_1.Data.RealPriceDirectLabor += item.Labor;
                    Data_1.Data.RealPriceOutputExpense += item.Expense;
                    //직공비에 해당하는 각 객체의 재료비, 노무비, 경비를 직접재료비, 직접노무비, 산출 경비에 더해나감
                    Data_1.Data.RealDirectMaterial += item.Material;
                    Data_1.Data.RealDirectLabor += item.Labor;
                    Data_1.Data.RealOutputExpense += item.Expense;
                }
                //작업설인지 확인
                else if (item.Item === '작업설') {
                    //작업설 가격 더해나감 (23.02.06)
                    Data_1.Data.ByProduct += item.Material;
                    Data_1.Data.ByProduct += item.Labor;
                    Data_1.Data.ByProduct += item.Expense;
                }
            }
        }
    };
    return Setting;
}());
exports.Setting = Setting;
Setting.GetData();
