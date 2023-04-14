'use strict'

//4.14 Setting.SetUnitPriceNoExcel에서 값 입력되지 않음.

var __assign =
    (this && this.__assign) ||
    function () {
        __assign =
            Object.assign ||
            function (t) {
                for (var s, i = 1, n = arguments.length; i < n; i++) {
                    s = arguments[i]
                    for (var p in s) if (Object.prototype.hasOwnProperty.call(s, p)) t[p] = s[p]
                }
                return t
            }
        return __assign.apply(this, arguments)
    }
Object.defineProperty(exports, '__esModule', { value: true })
exports.Setting = void 0
var Data_1 = require('./Data')
var ExcelHandling_1 = require('./ExcelHandling')
//import Enumerable from 'linq'
var fs = require('fs')
var Setting = /** @class */ (function () {
    function Setting() {}
    Setting.GetData = function () {
        var bidString = fs.readFileSync(Data_1.Data.folder + '\\OutputDataFromBID.json', 'utf-8')
        Setting.docBID = JSON.parse(bidString)
        Setting.eleBID = Setting.docBID['data']
        //세부공사별 번호 Data.ConstructionNums 딕셔너리에 저장
        Setting.GetConstructionNum()
        //세부공사별 리스트 생성(Dic -> key : 세부공사별 번호 / value : 세부공사별 리스트)
        Setting.AddConstructionList()
        //공내역 xml 파일 읽어들여 데이터 저장
        Setting.GetDataFromBID()
        if (Data_1.Data.XlsFiles != null) {
            //실내역으로부터 Data 객체에 단가세팅
            Setting.SetUnitPrice()
        } else {
            Setting.SetUnitPriceNoExcel() //실내역 없이 공내역만으로 Data 객체에 단가 세팅 (23.02.06)
        }
        //고정금액 및 적용비율 저장
        Setting.GetRate()
        //직공비 제외항목 및 고정금액 계산
        Setting.GetPrices()
        //표준시장단가 합계(조사금액) 저장
        Data_1.Data.InvestigateStandardMarket =
            Data_1.Data.StandardMaterial + Data_1.Data.StandardLabor + Data_1.Data.StandardExpense
    }
    Setting.GetConstructionNum = function () {
        var constNums = Setting.eleBID['T2']
        var index
        var construction
        for (var num in constNums) {
            index = constNums[num]['C1']['_text']
            construction = constNums[num]['C3']['_text']
            if (Data_1.Data.ConstructionNums.has(construction)) {
                construction += '2'
            }
            console.log(
                '키:',
                index,
                '값:',
                construction,
                '으로 Data.ConstructionNums에 <string, string> 추가'
            )
            Data_1.Data.ConstructionNums.set(index, construction)
        }
    }
    Setting.AddConstructionList = function () {
        Data_1.Data.ConstructionNums.forEach(function (value, key) {
            Data_1.Data.Dic.set(key, new Array())
            console.log(
                '키:',
                key,
                '값:',
                Data_1.Data.Dic.get(key),
                '으로 Data.Dic에 <string, []> 추가'
            )
        }) // Data.Dic자료구조 체크하기 <string. Data[]>가 맞는지, <string, Array[]>로 해야하는지
    }
    Setting.GetItem = function (bid) {
        var item = null
        //해당 공종이 일반, 표준시장단가 및 공종(입력불가) 항목인 경우
        if (bid['C7']['_text'] === '0') {
            if (bid['C5']['_text'] === 'S') {
                if (bid['C10']['_text'] === '') item = '표준시장단가'
                else item = '일반'
            } else item = '예외' // null값이 발생해 예외 추가함
        }
        //해당 공종이 무대(입력불가)인 경우
        else if (bid['C7']['_text'] === '1') item = '공종(입력불가)'
        //해당 공종이 관급자재인 경우
        else if (bid['C7']['_text'] === '2') item = '관급자재'
        else if (bid['C7']['_text'] === '3') item = '관급공종'
        //해당 공종이 PS인 경우
        else if (bid['C7']['_text'] === '4') item = 'PS'
        //해당 공종이 제요율적용제외공종인 경우
        else if (bid['C7']['_text'] === '5') item = '재요율적용제외'
        //해당 공종이 PS내역인 경우
        else if (bid['C7']['_text'] === '7') item = 'PS내역'
        //해당 공종이 음의 가격 공종인 경우
        else if (bid['C7']['_text'] === '19') item = '음(-)의 입찰금액'
        //해당 공종이 PS(안전관리비)인 경우
        else if (bid['C7']['_text'] === '20') item = 'PS(안전관리비)'
        //해당 공종이 작업설인 경우
        else if (bid['C7']['_text'] === '22') item = '작업설'
        else item = '예외'
        return item
    }
    Setting.GetDataFromBID = function () {
        // 수정 필요
        //4.10 주석처리한 부분 수정함
        var works = Setting.eleBID['T3'].map(function (work) {
            return new Data_1.Data(
                Setting.GetItem(work),
                work.C1._text, //null값 가능성O
                work.C2._text, //null값 가능성O
                work.C3._text, //null값 가능성O
                work.C9._text, //null값 가능성O
                work.C12._text, //null값 가능성O
                work.C13._text, //null값 가능성O
                work.C14._text, //null값 가능성O
                parseFloat(work.C15._text),
                parseFloat(work.C28._text),
                parseFloat(work.C29._text),
                parseFloat(work.C30._text)
            )
        })
        works.forEach(function (work) {
            Data_1.Data.Dic.get(work.ConstructionNum).push(__assign({}, work))
            console.log(
                'Data.Dic의',
                'arrray에',
                Data_1.Data.Dic.get(work.ConstructionNum)[
                    Data_1.Data.Dic.get(work.ConstructionNum).length - 1
                ],
                '추가'
            )
        })
    }
    Setting.MatchConstructionNum = function (filePath) {
        // public static Dic = new Map<string, Data[]>() //key : 세부공사별 번호 / value : 세부공사별 리스트
        var workbook = ExcelHandling_1.ExcelHandling.GetWorkbook(filePath, '.xlsx')
        var copySheetIndex = workbook.GetSheetIndex('내역서')
        var sheet = workbook.GetSheetAt(copySheetIndex)
        var check //실내역 파일과 세부공사의 데이터가 일치하는 횟수
        Data_1.Data.Dic.forEach(function (value, key) {
            check = 0
            for (var i = 0; i < 5; i++) {
                var row = sheet.GetRow(i + 4)
                var sameName = value[i].Name === row.GetCell(4).StringCallValue
                if (sameName) check++
                if (check == 3) {
                    Data_1.Data.MatchedConstNum.set(filePath, key)
                    return
                }
            }
        })
        Data_1.Data.IsFileMatch = false
    }
    Setting.CopyFile = function (filePath) {
        var workbook = ExcelHandling_1.ExcelHandling.GetWorkbook(filePath, '.xlsx')
        var copySheetIndex = workbook.GetSheetIndex('내역서')
        var sheet = workbook.GetSheetAt(copySheetIndex)
        var constNum = Data_1.Data.MatchedConstNum[filePath]
        var lastRowNum = sheet.LastRowNum
        var rowIndex = 4
        Data_1.Data.Dic[constNum].forEach(function (curObj) {
            var dcode = curObj.Code
            if (dcode == null || dcode.trim() === '') {
                return
            }
            var dname = curObj.Name
            var dunit = curObj.Unit
            var dquantity = curObj.Quantity
            while (true) {
                var row = sheet.GetRow(rowIndex)
                var code = row.GetCell(1).StringCellValue
                var name_1 = row.GetCell(4).StringCellValue
                var unit = row.GetCell(6).StringCellValue
                var quantity = 0.0
                try {
                    var quantity_1 = Number(row.GetCell(7).NumericCellValue)
                } catch (_a) {
                    rowIndex++
                    if (rowIndex == lastRowNum) break
                    continue
                }
                var sameCode = code === dcode
                var sameName = name_1 === dname
                var sameUnit = unit === dunit
                var sameQuantity = quantity === dquantity
                if ((sameName || sameCode) && (sameUnit || sameQuantity)) {
                    curObj.MaterialUnit = Number(row.GetCell(8).NumericCellValue)
                    curObj.LaborUnit = Number(row.GetCell(10).NumericCellValue)
                    curObj.ExpenseUnit = Number(row.GetCell(12).NumericCellValue)
                    rowIndex++
                    break
                } else {
                    rowIndex++
                    if (rowIndex == lastRowNum) break
                    continue
                }
            }
        })
    }
    Setting.SetUnitPrice = function () {
        // 미완성(가장 마지막에 작성 예정)
        // let copiedFolder: string = Data.folder + "\\Actual Xlsx";
        // const dir = new DirectoryInfo(copiedFolder);
        // const files = dir.GetFiles();
    }
    Setting.SetUnitPriceNoExcel = function () {
        var bidT3 = Setting.eleBID['T3']
        var _loop_1 = function (key) {
            if (bidT3[key]['C9']['_text'] != null && bidT3[key]['C5']['_text'] === 'S') {
                console.log('Setting UnitPrice!', key)

                var constNum = bidT3[key]['C1']['_text']
                var numVal_1 = bidT3[key]['C2']['_text']
                var detailVal_1 = bidT3[key]['C3']['_text']
                var curObject = Data_1.Data.Dic.get(constNum).find(function (x) {
                    return x.WorkNum === numVal_1 && x.DetailWorkNum === detailVal_1
                })
                if (
                    curObject.Item === '일반' ||
                    curObject.Item === '재요율적용제외' ||
                    curObject.Item === '표준시장단가'
                ) {
                    bidT3[key]['C16']['_text'] = curObject.MaterialUnit.toString()
                    bidT3[key]['C17']['_text'] = curObject.LaborUnit.toString()
                    bidT3[key]['C18']['_text'] = curObject.ExpenseUnit.toString()
                    bidT3[key]['C19']['_text'] = curObject.UnitPriceSum.toString()
                    bidT3[key]['C20']['_text'] = curObject.Material.toString()
                    bidT3[key]['C21']['_text'] = curObject.Labor.toString()
                    bidT3[key]['C22']['_text'] = curObject.Expense.toString()
                    bidT3[key]['C23']['_text'] = curObject.PriceSum.toString()
                    console.log(bidT3[key]['C21']['_text'])
                }
            }
        }
        for (var key in bidT3) {
            _loop_1(key)
        }
        if (fs.existsSync(Data_1.Data.work_path + '\\Setting_Json.json')) {
            fs.unlink(Data_1.Data.work_path + '\\Setting_Json.json', function (err) {
                if (err.code == 'ENOENT') {
                    console.log('파일 삭제 Error 발생')
                }
            })
        }
        fs.writeFileSync(
            Data_1.Data.work_path + '\\Setting_Json.json',
            JSON.stringify(Setting.docBID)
        )
    }
    Setting.GetRate = function () {
        console.log('start get rate')
        Data_1.Data.ConstructionTerm = Number(Setting.eleBID['T1']['C29']['_text'])
        var bidT5 = Setting.eleBID['T5']
        for (var key in bidT5) {
            console.log(
                bidT5[key]['C4']['_text'],
                bidT5[key]['C6']['_text'],
                bidT5[key]['C7']['_text'],
                bidT5[key]['C5']['_text'],
                bidT5[key]['C8']['_text']
            )
            var name_2 = JSON.stringify(bidT5[key]['C4']['_text'])
            var val1 = JSON.stringify(bidT5[key]['C6']['_text'])
            var val2 = JSON.stringify(bidT5[key]['C7']['_text'])
            if (JSON.stringify(bidT5[key]['C5']['_text']) === '7') {
                var fixedPrice = Number(JSON.stringify(bidT5[key]['C8']['_text']))
                Data_1.Data.Fixed.set(name_2, fixedPrice)
            } else {
                var applicationRate1 = Number(val1)
                var applicationRate2 = Number(val2)
                Data_1.Data.Rate1.set(name_2, applicationRate1)
                Data_1.Data.Rate2.set(name_2, applicationRate2)
            }
        }
    }
    Setting.GetPrices = function () {
        for (var _i = 0, _a = Array.from(Data_1.Data.Dic.values()); _i < _a.length; _i++) {
            var value = _a[_i]
            for (var _b = 0, value_1 = value; _b < value_1.length; _b++) {
                var item = value_1[_b]
                //console.log(item)
                if (item.Item === '관급') {
                    Data_1.Data.GovernmentMaterial += item.Material
                    Data_1.Data.GovernmentLabor += item.Labor
                    Data_1.Data.GovernmentExpense += item.Expense
                } else if (item.Item === 'PS') {
                    Data_1.Data.PsMaterial += item.Material
                    Data_1.Data.PsLabor += item.Labor
                    Data_1.Data.PsExpense += item.Expense
                } else if (item.Item === '제요율적용제외') {
                    Data_1.Data.ExcludingMaterial += item.Material
                    Data_1.Data.ExcludingLabor += item.Labor
                    Data_1.Data.ExcludingExpense += item.Expense
                }
                //해당 공종이 PS(안전관리비)인 경우
                else if (item.Item === 'PS(안전관리비)') {
                    Data_1.Data.SafetyPrice += item.Expense
                    //PS(안전관리비) 고정 단가에 추가 (23.02.06)
                    Data_1.Data.FixedPriceDirectMaterial += item.Material //재료비 합 계산
                    Data_1.Data.FixedPriceDirectLabor += item.Labor //노무비 합 계산
                    Data_1.Data.FixedPriceOutputExpense += item.Expense //경비 합 계산
                    //직공비에 해당하는 각 객체의 재료비, 노무비, 경비를 직접재료비, 직접노무비, 산출 경비에 더해나감
                    Data_1.Data.RealDirectMaterial += item.Material
                    Data_1.Data.RealDirectLabor += item.Labor
                    Data_1.Data.RealOutputExpense += item.Expense
                }
                //표준시장단가 품목인지 확인
                else if (item.Item === '표준시장단가') {
                    Data_1.Data.FixedPriceDirectMaterial += item.Material //재료비 합 계산
                    Data_1.Data.FixedPriceDirectLabor += item.Labor //노무비 합 계산
                    Data_1.Data.FixedPriceOutputExpense += item.Expense //경비 합 계산
                    //직공비에 해당하는 각 객체의 재료비, 노무비, 경비를 직접재료비, 직접노무비, 산출 경비에 더해나감
                    Data_1.Data.RealDirectMaterial += item.Material
                    Data_1.Data.RealDirectLabor += item.Labor
                    Data_1.Data.RealOutputExpense += item.Expense
                    //표준시장 단가 직공비 합계에 더해나감
                    Data_1.Data.StandardMaterial += item.Material
                    Data_1.Data.StandardLabor += item.Labor
                    Data_1.Data.StandardExpense += item.Expense
                }
                //음(-)의 단가 품목인지 확인
                else if (item.Item === '음(-)의 입찰금액') {
                    Data_1.Data.FixedPriceDirectMaterial += item.Material
                    Data_1.Data.FixedPriceDirectLabor += item.Labor
                    Data_1.Data.FixedPriceOutputExpense += item.Expense
                    //직공비에 해당하는 각 객체의 재료비, 노무비, 경비를 직접재료비, 직접노무비, 산출 경비에 더해나감
                    Data_1.Data.RealDirectMaterial += item.Material
                    Data_1.Data.RealDirectLabor += item.Labor
                    Data_1.Data.RealOutputExpense += item.Expense
                }
                //직공비 중, 고정금액이 아닌 일반 항목들의 직공비 계산
                else if (item.Item === '일반') {
                    Data_1.Data.RealPriceDirectMaterial += item.Material
                    Data_1.Data.RealPriceDirectLabor += item.Labor
                    Data_1.Data.RealPriceOutputExpense += item.Expense
                    //직공비에 해당하는 각 객체의 재료비, 노무비, 경비를 직접재료비, 직접노무비, 산출 경비에 더해나감
                    Data_1.Data.RealDirectMaterial += item.Material
                    Data_1.Data.RealDirectLabor += item.Labor
                    Data_1.Data.RealOutputExpense += item.Expense
                }
                //작업설인지 확인
                else if (item.Item === '작업설') {
                    //작업설 가격 더해나감 (23.02.06)
                    Data_1.Data.ByProduct += item.Material
                    Data_1.Data.ByProduct += item.Labor
                    Data_1.Data.ByProduct += item.Expense
                }
            }
        }
    }
    return Setting
})()
exports.Setting = Setting
