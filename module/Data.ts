import * as path from 'path'
import Decimal from 'big.js'
export class Data {
    constructor(
        item: string,
        constructionNum: string,
        workNum: string,
        detailWorkNum: string,
        code: string,
        name: string,
        standard: string,
        unit: string,
        quantity: number,
        materialUnit: number,
        laborUnit: number,
        expenseUnit: number
    ) {
        this.Item = item
        this.ConstructionNum = constructionNum
        this.WorkNum = workNum
        this.DetailWorkNum = detailWorkNum
        this.Code = code
        this.Name = name
        this.Standard = standard
        this.Unit = unit
        this.Quantity = quantity
        this.MaterialUnit = materialUnit
        this.LaborUnit = laborUnit
        this.ExpenseUnit = expenseUnit
    }
    public static folder: string = path.resolve(__dirname, '../AutoBid') //내 문서 폴더의 AutoBID 폴더로 지정 (23.02.02)
    // WPF 앱 파일 관리 변수
    public static XlsText: string = ''
    public static XlsFiles: File[]
    public static BidText: string
    public static BidFile: File
    public static CanCovertFile: boolean = false // 새로운 파일 업로드 시 변환 가능
    public static IsConvert: boolean = false // 변환을 했는지 안했는지
    public static IsBidFileOk: boolean = true // 정상적인 공내역 파일인지
    public static IsFileMatch: boolean = true // 공내역 파일과 실내역 파일의 공사가 일치하는지
    public static CompanyRegistrationNum: string = '' //1.31 사업자등록번호 추가
    public static CompanyRegistrationName: string = '' // 2.02 회사명 추가
    public static PersonalRateNum?: number // 내 사정율 변수
    //C#에서는 BalanceRate 속성에서 get을 통해 ToDecimal(BalanceRateNum) 리턴받음.
    //js는 number가 부동소수점
    public static BalanceRateNum?: number // 업체 평균 사정율 변수
    // 프로그램 폴더로 위치 변경
    public static work_path: string = path.join(Data.folder, 'WORK DIRECTORY') //작업폴더(WORK DIRECTORY) 경로
    private materialUnit: number = 0 //재료비 단가
    private laborUnit: number = 0 //노무비 단가
    private expenseUnit: number = 0 //경비 단가

    public Item: string = '' //항목 구분(공종(입력불가), 무대(입력불가), 일반, 관급, PS, 제요율적용제외, 안전관리비, PS내역, 표준시장단가)
    public ConstructionNum: string = '' //공사 인덱스
    public WorkNum: string = '' //세부 공사 인덱스
    public DetailWorkNum: string = '' //세부 공종 인덱스
    public Code: string = '' //코드
    public Name: string = '' //품명
    public Standard: string = '' //규격
    public Unit: string = '' //단위
    public Quantity: number = 0 //수량

    /**재료비 단가*/
    public get MaterialUnit(): number {
        //사용자가 단가 정수처리를 원한다면("2") 정수 값으로 return / Reset 함수를 쓰지 않은 경우의 조건 추가 (23.02.06)
        if (Data.UnitPriceTrimming === '2' && Data.ExecuteReset === '0')
            return Number(new Decimal(this.materialUnit).toFixed(0, 3))
        else if (Data.UnitPriceTrimming === '1' || Data.ExecuteReset === '1')
            // 사용자가 단가 소수점 처리를 원하거나 Reset 함수를 썼다면 소수 첫째 자리 아래로 절사 (23.02.06)
            return Number(new Decimal(this.materialUnit).toFixed(1, 0))
        return this.materialUnit //Default는 있는 그대로의 값을 return
    }
    public set MaterialUnit(value: number) {
        this.materialUnit = value
    }

    //노무비 단가
    public get LaborUnit() {
        if (Data.UnitPriceTrimming === '2' && Data.ExecuteReset === '0')
            return Number(new Decimal(this.laborUnit).toFixed(0, 3))
        else if (Data.UnitPriceTrimming === '1' || Data.ExecuteReset === '1')
            return Number(new Decimal(this.laborUnit).toFixed(1, 0))
        return this.laborUnit
    }
    public set LaborUnit(value: number) {
        this.laborUnit = value
    }

    //경비 단가
    public get ExpenseUnit() {
        var decimal = new Decimal(0)

        if (Data.UnitPriceTrimming === '2' && Data.ExecuteReset === '0')
            return Number(new Decimal(this.expenseUnit).toFixed(0, 3))
        else if (Data.UnitPriceTrimming === '1' || Data.ExecuteReset === '1')
            return Number(new Decimal(this.expenseUnit).toFixed(1, 0))
        return this.expenseUnit
    }
    public set ExpenseUnit(value: number) {
        this.expenseUnit = value
    }

    public get Material() {
        var decimal = new Decimal(this.Quantity).times(this.MaterialUnit)
        return Number(decimal.toFixed(0, 0))
    } //재료비 (수량 x 단가)
    public get Labor() {
        var decimal = new Decimal(this.Quantity).times(this.LaborUnit)
        return Number(decimal.toFixed(0, 0))
    } //노무비
    public get Expense() {
        var decimal = new Decimal(this.Quantity).times(this.ExpenseUnit)
        return Number(decimal.toFixed(0, 0))
    } //경비
    public get UnitPriceSum() {
        var decimal = new Decimal(this.MaterialUnit).plus(this.LaborUnit).plus(this.ExpenseUnit)
        return decimal.toNumber()
    } //합계단가
    public get PriceSum() {
        var decimal = new Decimal(this.Material).plus(this.Labor).plus(this.Expense)
        return decimal.toNumber()
    } //합계(세부공종별 금액의 합계)

    public Weight: number = 0 //가중치
    public PriceScore: number = 0 //세부 점수
    public get Score() {
        return this.PriceScore * this.Weight
    } //단가 점수(세부 점수 * 가중치)

    //원가계산서에 필요한 데이터
    public static ConstructionTerm: number //공사 기간
    public static RealDirectMaterial: number = 0 //실내역 직접 재료비(일반, - , 표준시장단가)
    public static RealDirectLabor: number = 0 //실내역 직접 노무비(일반, - , 표준시장단가)
    public static RealOutputExpense: number = 0 //실내역 산출 경비(일반, - , 표준시장단가)
    public static FixedPriceDirectMaterial: number = 0 //고정금액 항목 직접 재료비
    public static FixedPriceDirectLabor: number = 0 //고정금액 항목 직접 노무비
    public static FixedPriceOutputExpense: number = 0 //고정금액 항목 산출 경비
    public static RealPriceDirectMaterial: number = 0 //일반항목 직접 재료비
    public static RealPriceDirectLabor: number = 0 //일반항목 직접 노무비
    public static RealPriceOutputExpense: number = 0 //일반항목 산출 경비
    public static InvestigateFixedPriceDirectMaterial: number = 0 //고정금액 항목 직접 재료비(조사금액)
    public static InvestigateFixedPriceDirectLabor: number = 0 //고정금액 항목 직접 노무비(조사금액)
    public static InvestigateFixedPriceOutputExpense: number = 0 //고정금액 항목 산출 경비(조사금액)
    public static InvestigateStandardMaterial: number = 0 //표준시장단가 재료비(조사금액)
    public static InvestigateStandardLabor: number = 0 //표준시장단가 노무비(조사금액)
    public static InvestigateStandardExpense: number = 0 //표준시장단가 산출경비(조사금액)
    public static PsMaterial: number = 0 //PS(재료비) 금액(직접 재료비에서 제외)
    public static PsLabor: number = 0 //PS(노무비) 금액(직접 노무비에서 제외)
    public static PsExpense: number = 0 //PS(경비) 금액(산출 경비에서 제외)
    public static ExcludingMaterial: number = 0 //제요율적용제외(재료비) 금액(직접 재료비에서 제외)
    public static ExcludingLabor: number = 0 //제요율적용제외(노무비) 금액(직접 노무비에서 제외)
    public static ExcludingExpense: number = 0 //제요율적용제외(경비) 금액(산출 경비에서 제외)
    public static AdjustedExMaterial: number = 0 //사정율 적용한 제요율적용제외 금액(재료비)
    public static AdjustedExLabor: number = 0 //사정율 적용한 제요율적용제외 금액(노무비)
    public static AdjustedExExpense: number = 0 //사정율 적용한 제요율적용제외 금액(경비)
    public static GovernmentMaterial: number = 0 //관급자재요소(재료비) 금액(직접 재료비에서 제외)
    public static GovernmentLabor: number = 0 //관급자재요소(노무비) 금액(직접 노무비에서 제외)
    public static GovernmentExpense: number = 0 //관급자재요소(경비) 금액(산출 경비에서 제외)
    public static SafetyPrice: number = 0 //안전관리비(산출 경비에서 제외)
    public static StandardMaterial: number = 0 //표준시장단가 재료비
    public static StandardLabor: number = 0 //표준시장단가 노무비
    public static StandardExpense: number = 0 //표준시장단가 산출경비
    public static InvestigateStandardMarket: number = 0 //표준시장단가 합계(조사내역)
    public static FixedPricePercent: number = 0 //고정금액 비중
    public static ByProduct: number = 0 //작업설

    public static Dic = new Map<string, Array<Data>>() //key : 세부공사별 번호 / value : 세부공사별 리스트
    public static ConstructionNums = new Map<string, string>() //세부 공사별 번호 저장
    public static MatchedConstNum = new Map<string, string>() //실내역과 세부공사별 번호의 매칭 결과

    public static Fixed = new Map<string, number>() //고정금액 항목별 금액 저장
    public static Rate1 = new Map<string, number>() //적용비율1 저장
    public static Rate2 = new Map<string, number>() //적용비율2 저장

    public static RealPrices = new Map<string, number>() //원가계산서 항목별 금액 저장

    public static Investigation = new Map<string, number>() //세부결과_원가계산서 항목별 조사금액 저장
    public static Bidding = new Map<string, number>() //세부결과_원가계산서 항목별 입찰금액 저장
    public static Correction = new Map<string, number>() //원가계산서 조사금액 보정 항목 저장

    //사용자의 옵션 및 사정률 데이터
    public static UnitPriceTrimming: string = '0' //단가 소수 처리 (defalut = "0")
    public static StandardMarketDeduction: string = '2' //표준시장단가 99.7% 적용
    public static ZeroWeightDeduction: string = '2' //가중치 0% 공종 50% 적용
    public static CostAccountDeduction: string = '2' //원가계산 제경비 99.7% 적용
    public static BidPriceRaise: string = '2' //투찰금액 천원 절상
    public static LaborCostLowBound: string = '2' //노무비 하한 80%

    //업체 평균 예측율
    public static get BalancedRate() {
        //decimal로 반환. 소수 계산 라이브러리 필요
        return Data.BalanceRateNum //입력받은 BalancedRateNum(double? 형)을 decimal로 바꿈
    }
    //내 예가 사정률
    public static get PersonalRate() {
        //decimal로 반환. 소수 계산 라이브러리 필요
        return Data.PersonalRateNum //입력받은 PersonalRateNum(double? 형)을 decimal로 바꿈
    }
    public static ExecuteReset: string = '0' //Reset 함수 사용시 단가 소수처리 옵션과 별개로 소수 첫째자리 아래로 절사

    public static reset_data(){
        this.XlsText = ''
        this.CanCovertFile = false // 새로운 파일 업로드 시 변환 가능
        this.IsConvert = false // 변환을 했는지 안했는지
        this.IsBidFileOk = true // 정상적인 공내역 파일인지
        this.IsFileMatch = true // 공내역 파일과 실내역 파일의 공사가 일치하는지
        this.CompanyRegistrationNum = '' //1.31 사업자등록번호 추가
        this.CompanyRegistrationName = '' // 2.02 회사명 추가
        this.PersonalRateNum = 0 // 내 사정율 변수
        //C#에서는 BalanceRate 속성에서 get을 통해 ToDecimal(BalanceRateNum) 리턴받음.
        //js는 number가 부동소수점
        this.BalanceRateNum = 0 // 업체 평균 사정율 변수

        this.ConstructionTerm = 0 //공사 기간
        this.RealDirectMaterial = 0 //실내역 직접 재료비(일반, - , 표준시장단가)
        this.RealDirectLabor = 0 //실내역 직접 노무비(일반, - , 표준시장단가)
        this.RealOutputExpense = 0 //실내역 산출 경비(일반, - , 표준시장단가)
        this.FixedPriceDirectMaterial = 0 //고정금액 항목 직접 재료비
        this.FixedPriceDirectLabor = 0 //고정금액 항목 직접 노무비
        this.FixedPriceOutputExpense = 0 //고정금액 항목 산출 경비
        this.RealPriceDirectMaterial = 0 //일반항목 직접 재료비
        this.RealPriceDirectLabor = 0 //일반항목 직접 노무비
        this.RealPriceOutputExpense = 0 //일반항목 산출 경비
        this.InvestigateFixedPriceDirectMaterial = 0 //고정금액 항목 직접 재료비(조사금액)
        this.InvestigateFixedPriceDirectLabor = 0 //고정금액 항목 직접 노무비(조사금액)
        this.InvestigateFixedPriceOutputExpense = 0 //고정금액 항목 산출 경비(조사금액)
        this.InvestigateStandardMaterial = 0 //표준시장단가 재료비(조사금액)
        this.InvestigateStandardLabor = 0 //표준시장단가 노무비(조사금액)
        this.InvestigateStandardExpense = 0 //표준시장단가 산출경비(조사금액)
        this.PsMaterial = 0 //PS(재료비) 금액(직접 재료비에서 제외)
        this.PsLabor = 0 //PS(노무비) 금액(직접 노무비에서 제외)
        this.PsExpense = 0 //PS(경비) 금액(산출 경비에서 제외)
        this.ExcludingMaterial = 0 //제요율적용제외(재료비) 금액(직접 재료비에서 제외)
        this.ExcludingLabor = 0 //제요율적용제외(노무비) 금액(직접 노무비에서 제외)
        this.ExcludingExpense = 0 //제요율적용제외(경비) 금액(산출 경비에서 제외)
        this.AdjustedExMaterial = 0 //사정율 적용한 제요율적용제외 금액(재료비)
        this.AdjustedExLabor = 0 //사정율 적용한 제요율적용제외 금액(노무비)
        this.AdjustedExExpense = 0 //사정율 적용한 제요율적용제외 금액(경비)
        this.GovernmentMaterial = 0 //관급자재요소(재료비) 금액(직접 재료비에서 제외)
        this.GovernmentLabor = 0 //관급자재요소(노무비) 금액(직접 노무비에서 제외)
        this.GovernmentExpense = 0 //관급자재요소(경비) 금액(산출 경비에서 제외)
        this.SafetyPrice = 0 //안전관리비(산출 경비에서 제외)
        this.StandardMaterial = 0 //표준시장단가 재료비
        this.StandardLabor = 0 //표준시장단가 노무비
        this.StandardExpense = 0 //표준시장단가 산출경비
        this.InvestigateStandardMarket = 0 //표준시장단가 합계(조사내역)
        this.FixedPricePercent = 0 //고정금액 비중
        this.ByProduct = 0 //작업설

        this.Dic = new Map<string, Array<Data>>() //key : 세부공사별 번호 / value : 세부공사별 리스트
        this.ConstructionNums = new Map<string, string>() //세부 공사별 번호 저장
        this.MatchedConstNum = new Map<string, string>() //실내역과 세부공사별 번호의 매칭 결과

        this.Fixed = new Map<string, number>() //고정금액 항목별 금액 저장
        this.Rate1 = new Map<string, number>() //적용비율1 저장
        this.Rate2 = new Map<string, number>() //적용비율2 저장

        this.RealPrices = new Map<string, number>() //원가계산서 항목별 금액 저장

        this.Investigation = new Map<string, number>() //세부결과_원가계산서 항목별 조사금액 저장
        this.Bidding = new Map<string, number>() //세부결과_원가계산서 항목별 입찰금액 저장
        this.Correction = new Map<string, number>() //원가계산서 조사금액 보정 항목 저장

        //사용자의 옵션 및 사정률 데이터
        this.UnitPriceTrimming = '0' //단가 소수 처리 (defalut = "0")
        this.StandardMarketDeduction = '2' //표준시장단가 99.7% 적용
        this.ZeroWeightDeduction= '2' //가중치 0% 공종 50% 적용
        this.CostAccountDeduction = '2' //원가계산 제경비 99.7% 적용
        this.BidPriceRaise = '2' //투찰금액 천원 절상
        this.LaborCostLowBound = '2' //노무비 하한 80%
    }
}
