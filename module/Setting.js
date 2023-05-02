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
exports.Setting = void 0;
const Data_1 = require("./Data");
const ExcelHandling_1 = require("./ExcelHandling");
//import Enumerable from 'linq'
const fs = __importStar(require("fs"));
//4.14 Setting.SetUnitPriceNoExcel에서 값 입력되지 않음.
class Setting {
    static GetData() {
        const bidString = fs.readFileSync(Data_1.Data.folder + '\\OutputDataFromBID.json', 'utf-8');
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
        console.log('세팉끝');
    }
    static GetConstructionNum() {
        const constNums = Setting.eleBID['T2'];
        let index;
        let construction;
        for (let num in constNums) {
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
    }
    static AddConstructionList() {
        Data_1.Data.ConstructionNums.forEach(function (value, key) {
            Data_1.Data.Dic.set(key, new Array());
            //console.log('키:', key, '값:', Data.Dic.get(key), '으로 Data.Dic에 <string, []> 추가')
        }); // Data.Dic자료구조 체크하기 <string. Data[]>가 맞는지, <string, Array[]>로 해야하는지
    }
    static GetItem(bid) {
        let item = null;
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
    }
    static GetDataFromBID() {
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
        works.forEach((work) => {
            Data_1.Data.Dic.get(work.ConstructionNum).push(work);
            // console.log('Data.Dic의', 'arrray에', Data.Dic.get(work.ConstructionNum)[Data.Dic.get(work.ConstructionNum).length - 1], '추가')
        });
    }
    static MatchConstructionNum(filePath) {
        return __awaiter(this, void 0, void 0, function* () {
            // public static Dic = new Map<string, Data[]>() //key : 세부공사별 번호 / value : 세부공사별 리스트
            const workbook = yield ExcelHandling_1.ExcelHandling.GetWorkbook(filePath, '.xlsx');
            const sheet = workbook.getWorksheet('내역서');
            let check; //실내역 파일과 세부공사의 데이터가 일치하는 횟수
            Data_1.Data.Dic.forEach((value, key) => {
                check = 0;
                for (let i = 0; i < 5; i++) {
                    let row = sheet.getRow(i + 4);
                    let sameName = value[i].Name === row.getCell(4).value.toString();
                    if (sameName)
                        check++;
                    if (check == 3) {
                        Data_1.Data.MatchedConstNum.set(filePath, key);
                        return;
                    }
                }
            });
            Data_1.Data.IsFileMatch = false;
        });
    }
    static CopyFile(filePath) {
        return __awaiter(this, void 0, void 0, function* () {
            const workbook = yield ExcelHandling_1.ExcelHandling.GetWorkbook(filePath, '.xlsx');
            const sheet = workbook.getWorksheet('내역서');
            const constNum = Data_1.Data.MatchedConstNum[filePath];
            const lastRowNum = sheet.rowCount;
            let rowIndex = 4;
            Data_1.Data.Dic[constNum].forEach((curObj) => {
                const dcode = curObj.Code;
                if (dcode == null || dcode.trim() === '') {
                    return;
                }
                const dname = curObj.Name;
                const dunit = curObj.Unit;
                const dquantity = curObj.Quantity;
                while (true) {
                    const row = sheet.getRow(rowIndex);
                    const code = row.getCell(1).value.toString();
                    const name = row.getCell(4).value.toString();
                    const unit = row.getCell(6).value.toString();
                    let quantity = 0.0;
                    try {
                        let quantity = Number(row.getCell(7).value);
                    }
                    catch (_a) {
                        rowIndex++;
                        if (rowIndex == lastRowNum)
                            break;
                        continue;
                    }
                    let sameCode = code === dcode;
                    let sameName = name === dname;
                    let sameUnit = unit === dunit;
                    let sameQuantity = quantity === dquantity;
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
        });
    }
    static SetUnitPrice() {
        // 미완성(가장 마지막에 작성 예정)
        // let copiedFolder: string = Data.folder + "\\Actual Xlsx";
        // const dir = new DirectoryInfo(copiedFolder);
        // const files = dir.GetFiles();
    }
    static SetUnitPriceNoExcel() {
        const bidT3 = this.eleBID['T3'];
        let code;
        let type;
        for (let key in bidT3) {
            code = JSON.stringify(bidT3[key]['C9']['_text']);
            type = JSON.stringify(bidT3[key]['C5']['_text'])[1];
            if (code !== undefined && type === 'S') {
                let constNum = bidT3[key]['C1']['_text'];
                let numVal = bidT3[key]['C2']['_text'];
                let detailVal = bidT3[key]['C3']['_text'];
                let curObject = Data_1.Data.Dic.get(constNum).find((x) => x.WorkNum === numVal && x.DetailWorkNum === detailVal);
                if (curObject.Item === '일반' ||
                    curObject.Item === '재요율적용제외' ||
                    (curObject.Item === '표준시장단가' && curObject !== undefined)) {
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
        }
        if (fs.existsSync(Data_1.Data.work_path + '\\Setting_Json.json')) {
            fs.unlink(Data_1.Data.work_path + '\\Setting_Json.json', (err) => {
                if (err && err.code == 'ENOENT') {
                    console.log('파일 삭제 Error 발생');
                }
            });
        }
        fs.writeFileSync(Data_1.Data.work_path + '\\Setting_Json.json', JSON.stringify(this.docBID));
    }
    static GetRate() {
        Data_1.Data.ConstructionTerm = Number(Setting.eleBID['T1']['C29']['_text']);
        const bidT5 = Setting.eleBID['T5'];
        //console.log()
        for (let key in bidT5) {
            let name = JSON.stringify(bidT5[key]['C4']['_text']);
            let val1 = bidT5[key]['C6']['_text'];
            let val2 = bidT5[key]['C7']['_text'];
            if (bidT5[key]['C5']['_text'] === '7') {
                let fixedPrice = Number(bidT5[key]['C8']['_text']);
                Data_1.Data.Fixed.set(name, fixedPrice);
            }
            else {
                let applicationRate1 = Number(val1);
                let applicationRate2 = Number(val2);
                Data_1.Data.Rate1.set(name, applicationRate1);
                Data_1.Data.Rate2.set(name, applicationRate2);
            }
        }
    }
    static GetPrices() {
        for (let value of Array.from(Data_1.Data.Dic.values())) {
            for (let item of value) {
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
    }
}
exports.Setting = Setting;
