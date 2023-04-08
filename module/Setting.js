"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
var Data_1 = require("./Data");
var fs_1 = __importDefault(require("fs"));
var SetUnitPriceByExcel;
(function (SetUnitPriceByExcel) {
    var Setting = /** @class */ (function () {
        function Setting() {
        }
        Setting.GetData = function () {
            var bidString = fs_1.default.readFileSync("C:\\Users\\USER\\Documents\\" + "\\OutputDataFromBID.json", 'utf-8');
            this.docBID = JSON.parse(bidString);
            this.eleBID = this.docBID['data'];
            //세부공사별 번호 Data.ConstructionNums 딕셔너리에 저장
            this.GetConstructionNum();
            //세부공사별 리스트 생성(Dic -> key : 세부공사별 번호 / value : 세부공사별 리스트)
            this.AddConstructionList();
            //공내역 xml 파일 읽어들여 데이터 저장
            this.GetDataFromBID();
            if (Data_1.Data.XlsFiles != null) //실내역으로부터 Data 객체에 단가세팅
             {
                this.SetUnitPrice();
            }
            else {
                this.SetUnitPriceNoExcel(); //실내역 없이 공내역만으로 Data 객체에 단가 세팅 (23.02.06)
            }
            //고정금액 및 적용비율 저장
            this.GetRate();
            //직공비 제외항목 및 고정금액 계산
            this.GetPrices();
            //표준시장단가 합계(조사금액) 저장
            Data_1.Data.InvestigateStandardMarket = Data_1.Data.StandardMaterial + Data_1.Data.StandardLabor + Data_1.Data.StandardExpense;
        };
        Setting.GetConstructionNum = function () {
            var constNums = this.eleBID['T2'];
            var index;
            var construction;
            for (var num in constNums) {
                index = JSON.stringify(constNums[num]['C1']['_text']);
                construction = JSON.stringify(constNums[num]['C3']['_text']);
                if (Data_1.Data.ConstructionNums.has(construction)) {
                    construction += "2";
                }
                Data_1.Data.ConstructionNums.set(index, construction);
            }
        };
        Setting.AddConstructionList = function () {
            Data_1.Data.ConstructionNums.forEach(function (key, value) { return Data_1.Data.Dic.set(key, new Array()); }); // Data.Dic자료구조 체크하기 <string. Data[]>가 맞는지, <string, Array[]>로 해야하는지
        };
        Setting.GetItem = function (bid) {
            var item = null;
            return item;
        };
        Setting.GetDataFromBID = function () {
            var _this = this;
            var works = this.eleBID.filter(function (work) { return work.Name === "T3"; }).map(function (work) {
                return {
                    Item: _this.GetItem(work),
                    ConstructionNum: work.Element("C1").Value.toString(),
                    WorkNum: work.Element("C2").Value.toString(),
                    DetailWorkNum: work.Element("C3").Value.toString(),
                    Code: work.Element("C9").Value.toString(),
                    Name: work.Element("C12").Value.toString(),
                    Standard: work.Element("C13").Value.toString(),
                    Unit: work.Element("C14").Value.toString(),
                    Quantity: parseFloat(work.Element("C15").Value),
                    MaterialUnit: parseFloat(work.Element("C28").Value),
                    LaborUnit: parseFloat(work.Element("C29").Value),
                    ExpenseUnit: parseFloat(work.Element("C30").Value)
                };
            });
        };
        Setting.MatchConstructionNum = function (filePath) {
            // public static Dic = new Map<string, Data[]>() //key : 세부공사별 번호 / value : 세부공사별 리스트
            var workbook = ExcelHandling.GetWorkbook(filePath, ".xlsx");
            var copySheetIndex = workbook.GetSheetIndex("내역서");
            var sheet = workbook.GetSheetAt(copySheetIndex);
            var check; //실내역 파일과 세부공사의 데이터가 일치하는 횟수
            Data_1.Data.Dic.forEach(function (value, key) {
                check = 0;
                for (var i = 0; i < 5; i++) {
                    var row = sheet.GetRow(i + 4);
                    var sameName = (value[i].Name === row.GetCell(4).StringCallValue);
                    if (sameName)
                        check++;
                    if (check == 3) {
                        Data_1.Data.MatchedConstNum.set(filePath, key);
                        return;
                    }
                }
            });
            Data_1.Data.IsFileMatch = false;
        };
        Setting.CopyFile = function (filePath) {
            var workbook = ExcelHandling.GetWorkbook(filePath, ".xlsx");
            var copySheetIndex = workbook.GetSheetIndex("내역서");
            var sheet = workbook.GetSheetAt(copySheetIndex);
            var constNum = Data_1.Data.MatchedConstNum[filePath];
            var lastRowNum = sheet.LastRowNum;
            var rowIndex = 4;
            Data_1.Data.Dic[constNum].forEach(function (curObj) {
                var dcode = curObj.Code;
                if (dcode == null || dcode.trim() === '') {
                    return;
                }
                var dname = curObj.Name;
                var dunit = curObj.Unit;
                var dquantity = curObj.Quantity;
                while (true) {
                    var row = sheet.GetRow(rowIndex);
                    var code = row.GetCell(1).StringCellValue;
                    var name_1 = row.GetCell(4).StringCellValue;
                    var unit = row.GetCell(6).StringCellValue;
                    var quantity = 0.0;
                    try {
                        var quantity_1 = Number(row.GetCell(7).NumericCellValue);
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
                        curObj.MaterialUnit = Number(row.GetCell(8).NumericCellValue);
                        curObj.LaborUnit = Number(row.GetCell(10).NumericCellValue);
                        curObj.ExpenseUnit = Number(row.GetCell(12).NumericCellValue);
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
        };
        Setting.SetUnitPrice = function () {
            var copiedFolder = Data_1.Data.folder + "\\Actual Xlsx";
            var dir = new DirectoryInfo(copiedFolder);
            var files = dir.GetFiles();
        };
        Setting.SetUnitPriceNoExcel = function () {
        };
        Setting.GetRate = function () {
            var bidT1 = this.eleBID['T1'];
            for (var key in bidT1) {
                Data_1.Data.ConstructionTerm = Number(JSON.stringify(bidT1[key]['C29']['_text']));
            }
            var bidT5 = this.eleBID['T5'];
            for (var key in bidT5) {
                var name_2 = JSON.stringify(bidT5[key]['C4']['_text']);
                var val1 = JSON.stringify(bidT5[key]['C6']['_text']);
                var val2 = JSON.stringify(bidT5[key]['C7']['_text']);
                if (JSON.stringify(bidT5[key]['C5']['_text']) === '7') {
                    var fixedPrice = Number(JSON.stringify(bidT5[key]['C8']['_text']));
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
            Data_1.Data.Dic.forEach(function (value, key) {
                value.forEach(function (item) {
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
                });
            });
        };
        return Setting;
    }());
})(SetUnitPriceByExcel || (SetUnitPriceByExcel = {}));
