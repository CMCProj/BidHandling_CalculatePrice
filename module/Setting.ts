import { Data } from './Data'
import { ExcelHandling } from './ExcelHandling'
import Enumerable from 'linq'
import fs from 'fs'

export class Setting {
    private static docBID: JSON
    private static eleBID: JSON

    public static GetData(): void {
        const bidString: string = fs.readFileSync(Data.folder + '\\OutputDataFromBID.json', 'utf-8')
        this.docBID = JSON.parse(bidString)
        this.eleBID = this.docBID['data']

        //세부공사별 번호 Data.ConstructionNums 딕셔너리에 저장
        this.GetConstructionNum()

        //세부공사별 리스트 생성(Dic -> key : 세부공사별 번호 / value : 세부공사별 리스트)
        this.AddConstructionList()

        //공내역 xml 파일 읽어들여 데이터 저장
        this.GetDataFromBID()

        if (Data.XlsFiles != null) {
            //실내역으로부터 Data 객체에 단가세팅
            this.SetUnitPrice()
        } else {
            this.SetUnitPriceNoExcel() //실내역 없이 공내역만으로 Data 객체에 단가 세팅 (23.02.06)
        }

        //고정금액 및 적용비율 저장
        this.GetRate()

        //직공비 제외항목 및 고정금액 계산
        this.GetPrices()

        //표준시장단가 합계(조사금액) 저장
        Data.InvestigateStandardMarket =
            Data.StandardMaterial + Data.StandardLabor + Data.StandardExpense
    }

    public static GetConstructionNum(): void {
        const constNums: object = this.eleBID['T2']
        let index: string
        let construction: string
        for (let num in constNums) {
            index = JSON.stringify(constNums[num]['C1']['_text'])
            construction = JSON.stringify(constNums[num]['C3']['_text'])
            if (Data.ConstructionNums.has(construction)) {
                construction += '2'
            }
            Data.ConstructionNums.set(index, construction)
        }
    }

    public static AddConstructionList(): void {
        Data.ConstructionNums.forEach((key, value) => Data.Dic.set(key, new Array<Data>())) // Data.Dic자료구조 체크하기 <string. Data[]>가 맞는지, <string, Array[]>로 해야하는지
    }

    public static GetItem(bid): string {
        let item: string = null
        //해당 공종이 일반, 표준시장단가 및 공종(입력불가) 항목인 경우
        if (bid['C7']['_text'] === '0') {
            if (bid['C5']['_text'] === 'S') {
                if (bid['C10']['_text'] === '') item = '표준시장단가'
                else item = '일반'
            }
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

    public static GetDataFromBID(): void {
        // 수정 필요
        const works = (this.eleBID as unknown as any[])
            .filter((work) => work.Name === 'T3')
            .map((work) => {
                return {
                    Item: this.GetItem(work),
                    ConstructionNum: work.Element('C1').Value.toString(),
                    WorkNum: work.Element('C2').Value.toString(),
                    DetailWorkNum: work.Element('C3').Value.toString(),
                    Code: work.Element('C9').Value.toString(),
                    Name: work.Element('C12').Value.toString(),
                    Standard: work.Element('C13').Value.toString(),
                    Unit: work.Element('C14').Value.toString(),
                    Quantity: parseFloat(work.Element('C15').Value),
                    MaterialUnit: parseFloat(work.Element('C28').Value),
                    LaborUnit: parseFloat(work.Element('C29').Value),
                    ExpenseUnit: parseFloat(work.Element('C30').Value),
                }
            })
        works.forEach((work) => {
            Data.Dic[work.ConstructionNum].set(work)
        })
    }

    public static MatchConstructionNum(filePath: string): void {
        // public static Dic = new Map<string, Data[]>() //key : 세부공사별 번호 / value : 세부공사별 리스트
        const workbook = ExcelHandling.GetWorkbook(filePath, '.xlsx')
        const copySheetIndex = workbook.GetSheetIndex('내역서')
        const sheet = workbook.GetSheetAt(copySheetIndex)
        let check: number //실내역 파일과 세부공사의 데이터가 일치하는 횟수
        Data.Dic.forEach((value, key) => {
            check = 0
            for (let i = 0; i < 5; i++) {
                let row = sheet.GetRow(i + 4)
                let sameName: boolean = value[i].Name === row.GetCell(4).StringCallValue
                if (sameName) check++
                if (check == 3) {
                    Data.MatchedConstNum.set(filePath, key)
                    return
                }
            }
        })
        Data.IsFileMatch = false
    }

    public static CopyFile(filePath: string): void {
        const workbook = ExcelHandling.GetWorkbook(filePath, '.xlsx')
        const copySheetIndex = workbook.GetSheetIndex('내역서')
        const sheet = workbook.GetSheetAt(copySheetIndex)

        const constNum = Data.MatchedConstNum[filePath]
        const lastRowNum = sheet.LastRowNum
        let rowIndex = 4
        Data.Dic[constNum].forEach((curObj) => {
            const dcode: string = curObj.Code
            if (dcode == null || dcode.trim() === '') {
                return
            }
            const dname = curObj.Name
            const dunit = curObj.Unit
            const dquantity = curObj.Quantity
            while (true) {
                const row = sheet.GetRow(rowIndex)
                const code = row.GetCell(1).StringCellValue
                const name = row.GetCell(4).StringCellValue
                const unit = row.GetCell(6).StringCellValue
                let quantity: number = 0.0

                try {
                    let quantity: number = Number(row.GetCell(7).NumericCellValue)
                } catch {
                    rowIndex++
                    if (rowIndex == lastRowNum) break
                    continue
                }

                let sameCode: boolean = code === dcode
                let sameName: boolean = name === dname
                let sameUnit: boolean = unit === dunit
                let sameQuantity: boolean = quantity === dquantity

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

    public static SetUnitPrice(): void { // 미완성(가장 마지막에 작성 예정)
        // let copiedFolder: string = Data.folder + "\\Actual Xlsx";
        // const dir = new DirectoryInfo(copiedFolder);
        // const files = dir.GetFiles();
    }

    public static SetUnitPriceNoExcel(): void {
        const bidT3: object = this.eleBID['T3']
        for (let key in bidT3) {
            if (
                JSON.stringify(bidT3[key]['C9']['_text']) != null &&
                JSON.stringify(bidT3[key]['C5']['_text']) === 'S'
            ) {
                let constNum: string = bidT3[key]['C1']['_text']
                let numVal: string = bidT3[key]['C2']['_text']
                let detailVal: string = bidT3[key]['C3']['_text']
                let curObject = Data.Dic[constNum].Find(
                    (x) => x.WorkNum === numVal && x.DetailWorkNum === detailVal
                )

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
                }
            }
        }
        if (fs.existsSync(Data.work_path + '\\Setting_Xml.xml')) {
            fs.unlink(Data.work_path + '\\Setting_Xml.xml', (err) => {
                if (err.code == 'ENOENT') {
                    console.log('파일 삭제 Error 발생')
                }
            })
        }
        fs.writeFileSync(Data.work_path + '\\Setting_Xml.xml', JSON.stringify(this.docBID))
    }

    public static GetRate(): void {
        const bidT1: object = this.eleBID['T1']
        for (let key in bidT1) {
            Data.ConstructionTerm = Number(JSON.stringify(bidT1[key]['C29']['_text']))
        }
        const bidT5: object = this.eleBID['T5']
        for (let key in bidT5) {
            let name: string = JSON.stringify(bidT5[key]['C4']['_text'])
            let val1: string = JSON.stringify(bidT5[key]['C6']['_text'])
            let val2: string = JSON.stringify(bidT5[key]['C7']['_text'])
            if (JSON.stringify(bidT5[key]['C5']['_text']) === '7') {
                let fixedPrice: number = Number(JSON.stringify(bidT5[key]['C8']['_text']))
                Data.Fixed.set(name, fixedPrice)
            } else {
                let applicationRate1: number = Number(val1)
                let applicationRate2: number = Number(val2)
                Data.Rate1.set(name, applicationRate1)
                Data.Rate2.set(name, applicationRate2)
            }
        }
    }

    public static GetPrices(): void {
        Data.Dic.forEach((value, key) => {
            value.forEach((item) => {
                if (item.Item === '관급') {
                    Data.GovernmentMaterial += item.Material
                    Data.GovernmentLabor += item.Labor
                    Data.GovernmentExpense += item.Expense
                } else if (item.Item === 'PS') {
                    Data.PsMaterial += item.Material
                    Data.PsLabor += item.Labor
                    Data.PsExpense += item.Expense
                } else if (item.Item === '제요율적용제외') {
                    Data.ExcludingMaterial += item.Material
                    Data.ExcludingLabor += item.Labor
                    Data.ExcludingExpense += item.Expense
                }
                //해당 공종이 PS(안전관리비)인 경우
                else if (item.Item === 'PS(안전관리비)') {
                    Data.SafetyPrice += item.Expense
                    //PS(안전관리비) 고정 단가에 추가 (23.02.06)
                    Data.FixedPriceDirectMaterial += item.Material //재료비 합 계산
                    Data.FixedPriceDirectLabor += item.Labor //노무비 합 계산
                    Data.FixedPriceOutputExpense += item.Expense //경비 합 계산
                    //직공비에 해당하는 각 객체의 재료비, 노무비, 경비를 직접재료비, 직접노무비, 산출 경비에 더해나감
                    Data.RealDirectMaterial += item.Material
                    Data.RealDirectLabor += item.Labor
                    Data.RealOutputExpense += item.Expense
                }
                //표준시장단가 품목인지 확인
                else if (item.Item === '표준시장단가') {
                    Data.FixedPriceDirectMaterial += item.Material //재료비 합 계산
                    Data.FixedPriceDirectLabor += item.Labor //노무비 합 계산
                    Data.FixedPriceOutputExpense += item.Expense //경비 합 계산
                    //직공비에 해당하는 각 객체의 재료비, 노무비, 경비를 직접재료비, 직접노무비, 산출 경비에 더해나감
                    Data.RealDirectMaterial += item.Material
                    Data.RealDirectLabor += item.Labor
                    Data.RealOutputExpense += item.Expense
                    //표준시장 단가 직공비 합계에 더해나감
                    Data.StandardMaterial += item.Material
                    Data.StandardLabor += item.Labor
                    Data.StandardExpense += item.Expense
                }
                //음(-)의 단가 품목인지 확인
                else if (item.Item === '음(-)의 입찰금액') {
                    Data.FixedPriceDirectMaterial += item.Material
                    Data.FixedPriceDirectLabor += item.Labor
                    Data.FixedPriceOutputExpense += item.Expense
                    //직공비에 해당하는 각 객체의 재료비, 노무비, 경비를 직접재료비, 직접노무비, 산출 경비에 더해나감
                    Data.RealDirectMaterial += item.Material
                    Data.RealDirectLabor += item.Labor
                    Data.RealOutputExpense += item.Expense
                }
                //직공비 중, 고정금액이 아닌 일반 항목들의 직공비 계산
                else if (item.Item === '일반') {
                    Data.RealPriceDirectMaterial += item.Material
                    Data.RealPriceDirectLabor += item.Labor
                    Data.RealPriceOutputExpense += item.Expense
                    //직공비에 해당하는 각 객체의 재료비, 노무비, 경비를 직접재료비, 직접노무비, 산출 경비에 더해나감
                    Data.RealDirectMaterial += item.Material
                    Data.RealDirectLabor += item.Labor
                    Data.RealOutputExpense += item.Expense
                }
                //작업설인지 확인
                else if (item.Item === '작업설') {
                    //작업설 가격 더해나감 (23.02.06)
                    Data.ByProduct += item.Material
                    Data.ByProduct += item.Labor
                    Data.ByProduct += item.Expense
                }
            })
        })
    }
}
