'use strict'
var __createBinding =
    (this && this.__createBinding) ||
    (Object.create
        ? function (o, m, k, k2) {
              if (k2 === undefined) k2 = k
              var desc = Object.getOwnPropertyDescriptor(m, k)
              if (!desc || ('get' in desc ? !m.__esModule : desc.writable || desc.configurable)) {
                  desc = {
                      enumerable: true,
                      get: function () {
                          return m[k]
                      },
                  }
              }
              Object.defineProperty(o, k2, desc)
          }
        : function (o, m, k, k2) {
              if (k2 === undefined) k2 = k
              o[k2] = m[k]
          })
var __setModuleDefault =
    (this && this.__setModuleDefault) ||
    (Object.create
        ? function (o, v) {
              Object.defineProperty(o, 'default', { enumerable: true, value: v })
          }
        : function (o, v) {
              o['default'] = v
          })
var __importStar =
    (this && this.__importStar) ||
    function (mod) {
        if (mod && mod.__esModule) return mod
        var result = {}
        if (mod != null)
            for (var k in mod)
                if (k !== 'default' && Object.prototype.hasOwnProperty.call(mod, k))
                    __createBinding(result, mod, k)
        __setModuleDefault(result, mod)
        return result
    }
Object.defineProperty(exports, '__esModule', { value: true })
exports.Data = void 0
const path = __importStar(require('path'))
class Data {
    constructor(
        item,
        constructionNum,
        workNum,
        detailWorkNum,
        code,
        name,
        standard,
        unit,
        quantity,
        materialUnit,
        laborUnit,
        expenseUnit
    ) {
        this.materialUnit = 0 //재료비 단가
        this.laborUnit = 0 //노무비 단가
        this.expenseUnit = 0 //경비 단가
        this.Item = '' //항목 구분(공종(입력불가), 무대(입력불가), 일반, 관급, PS, 제요율적용제외, 안전관리비, PS내역, 표준시장단가)
        this.ConstructionNum = '' //공사 인덱스
        this.WorkNum = '' //세부 공사 인덱스
        this.DetailWorkNum = '' //세부 공종 인덱스
        this.Code = '' //코드
        this.Name = '' //품명
        this.Standard = '' //규격
        this.Unit = '' //단위
        this.Quantity = 0 //수량
        this.Weight = 0 //가중치
        this.PriceScore = 0 //세부 점수
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
    /**재료비 단가*/
    get MaterialUnit() {
        //사용자가 단가 정수처리를 원한다면("2") 정수 값으로 return / Reset 함수를 쓰지 않은 경우의 조건 추가 (23.02.06)
        if (Data.UnitPriceTrimming === '2' && Data.ExecuteReset === '0')
            return Math.ceil(this.materialUnit)
        else if (Data.UnitPriceTrimming === '1' || Data.ExecuteReset === '1')
            // 사용자가 단가 소수점 처리를 원하거나 Reset 함수를 썼다면 소수 첫째 자리 아래로 절사 (23.02.06)
            return Math.floor(this.materialUnit * 10) / 10
        return this.materialUnit //Default는 있는 그대로의 값을 return
    }
    set MaterialUnit(value) {
        this.materialUnit = value
    }
    //노무비 단가
    get LaborUnit() {
        if (Data.UnitPriceTrimming === '2' && Data.ExecuteReset === '0')
            return Math.ceil(this.laborUnit)
        else if (Data.UnitPriceTrimming === '1' || Data.ExecuteReset === '1')
            return Math.floor(this.laborUnit * 10) / 10
        return this.laborUnit
    }
    set LaborUnit(value) {
        this.laborUnit = value
    }
    //경비 단가
    get ExpenseUnit() {
        if (Data.UnitPriceTrimming === '2' && Data.ExecuteReset === '0')
            return Math.ceil(this.expenseUnit)
        else if (Data.UnitPriceTrimming === '1' || Data.ExecuteReset === '1')
            return Math.floor(this.expenseUnit * 10) / 10
        return this.expenseUnit
    }
    set ExpenseUnit(value) {
        this.expenseUnit = value
    }
    get Material() {
        return Math.floor(this.Quantity * this.MaterialUnit)
    } //재료비 (수량 x 단가)
    get Labor() {
        return Math.floor(this.Quantity * this.LaborUnit)
    } //노무비
    get Expense() {
        return Math.floor(this.Quantity * this.ExpenseUnit)
    } //경비
    get UnitPriceSum() {
        return (
            Math.round((this.MaterialUnit + this.LaborUnit + this.ExpenseUnit) * 10000000) /
            10000000
        ) //자바스크립트 소수점 계산 오류를 막기 위해 소수점 8자리에서 반올림 처리
    } //합계단가
    get PriceSum() {
        return Math.round((this.Material + this.Labor + this.Expense) * 10000000) / 10000000 //자바스크립트 소수점 계산 오류를 막기 위해 소수점 8자리에서 반올림 처리
    } //합계(세부공종별 금액의 합계)
    get Score() {
        return this.PriceScore * this.Weight
    } //단가 점수(세부 점수 * 가중치)
    //업체 평균 예측율
    static get BalancedRate() {
        //decimal로 반환. 소수 계산 라이브러리 필요
        return Data.BalanceRateNum //입력받은 BalancedRateNum(double? 형)을 decimal로 바꿈
    }
    //내 예가 사정률
    static get PersonalRate() {
        //decimal로 반환. 소수 계산 라이브러리 필요
        return Data.PersonalRateNum //입력받은 PersonalRateNum(double? 형)을 decimal로 바꿈
    }
}
Data.folder = path.resolve(__dirname, '../AutoBid') //내 문서 폴더의 AutoBID 폴더로 지정 (23.02.02)
// WPF 앱 파일 관리 변수
Data.XlsText = ''
Data.CanCovertFile = false // 새로운 파일 업로드 시 변환 가능
Data.IsConvert = false // 변환을 했는지 안했는지
Data.IsBidFileOk = true // 정상적인 공내역 파일인지
Data.IsFileMatch = true // 공내역 파일과 실내역 파일의 공사가 일치하는지
Data.CompanyRegistrationNum = '' //1.31 사업자등록번호 추가
Data.CompanyRegistrationName = '' // 2.02 회사명 추가
// 프로그램 폴더로 위치 변경
Data.work_path = path.join(Data.folder, 'WORK DIRECTORY') //작업폴더(WORK DIRECTORY) 경로
Data.RealDirectMaterial = 0 //실내역 직접 재료비(일반, - , 표준시장단가)
Data.RealDirectLabor = 0 //실내역 직접 노무비(일반, - , 표준시장단가)
Data.RealOutputExpense = 0 //실내역 산출 경비(일반, - , 표준시장단가)
Data.FixedPriceDirectMaterial = 0 //고정금액 항목 직접 재료비
Data.FixedPriceDirectLabor = 0 //고정금액 항목 직접 노무비
Data.FixedPriceOutputExpense = 0 //고정금액 항목 산출 경비
Data.RealPriceDirectMaterial = 0 //일반항목 직접 재료비
Data.RealPriceDirectLabor = 0 //일반항목 직접 노무비
Data.RealPriceOutputExpense = 0 //일반항목 산출 경비
Data.InvestigateFixedPriceDirectMaterial = 0 //고정금액 항목 직접 재료비(조사금액)
Data.InvestigateFixedPriceDirectLabor = 0 //고정금액 항목 직접 노무비(조사금액)
Data.InvestigateFixedPriceOutputExpense = 0 //고정금액 항목 산출 경비(조사금액)
Data.InvestigateStandardMaterial = 0 //표준시장단가 재료비(조사금액)
Data.InvestigateStandardLabor = 0 //표준시장단가 노무비(조사금액)
Data.InvestigateStandardExpense = 0 //표준시장단가 산출경비(조사금액)
Data.PsMaterial = 0 //PS(재료비) 금액(직접 재료비에서 제외)
Data.PsLabor = 0 //PS(노무비) 금액(직접 노무비에서 제외)
Data.PsExpense = 0 //PS(경비) 금액(산출 경비에서 제외)
Data.ExcludingMaterial = 0 //제요율적용제외(재료비) 금액(직접 재료비에서 제외)
Data.ExcludingLabor = 0 //제요율적용제외(노무비) 금액(직접 노무비에서 제외)
Data.ExcludingExpense = 0 //제요율적용제외(경비) 금액(산출 경비에서 제외)
Data.AdjustedExMaterial = 0 //사정율 적용한 제요율적용제외 금액(재료비)
Data.AdjustedExLabor = 0 //사정율 적용한 제요율적용제외 금액(노무비)
Data.AdjustedExExpense = 0 //사정율 적용한 제요율적용제외 금액(경비)
Data.GovernmentMaterial = 0 //관급자재요소(재료비) 금액(직접 재료비에서 제외)
Data.GovernmentLabor = 0 //관급자재요소(노무비) 금액(직접 노무비에서 제외)
Data.GovernmentExpense = 0 //관급자재요소(경비) 금액(산출 경비에서 제외)
Data.SafetyPrice = 0 //안전관리비(산출 경비에서 제외)
Data.StandardMaterial = 0 //표준시장단가 재료비
Data.StandardLabor = 0 //표준시장단가 노무비
Data.StandardExpense = 0 //표준시장단가 산출경비
Data.InvestigateStandardMarket = 0 //표준시장단가 합계(조사내역)
Data.FixedPricePercent = 0 //고정금액 비중
Data.ByProduct = 0 //작업설
Data.Dic = new Map() //key : 세부공사별 번호 / value : 세부공사별 리스트
Data.ConstructionNums = new Map() //세부 공사별 번호 저장
Data.MatchedConstNum = new Map() //실내역과 세부공사별 번호의 매칭 결과
Data.Fixed = new Map() //고정금액 항목별 금액 저장
Data.Rate1 = new Map() //적용비율1 저장
Data.Rate2 = new Map() //적용비율2 저장
Data.RealPrices = new Map() //원가계산서 항목별 금액 저장
Data.Investigation = new Map() //세부결과_원가계산서 항목별 조사금액 저장
Data.Bidding = new Map() //세부결과_원가계산서 항목별 입찰금액 저장
Data.Correction = new Map() //원가계산서 조사금액 보정 항목 저장
//사용자의 옵션 및 사정률 데이터
Data.UnitPriceTrimming = '0' //단가 소수 처리 (defalut = "0")
Data.StandardMarketDeduction = '2' //표준시장단가 99.7% 적용
Data.ZeroWeightDeduction = '2' //가중치 0% 공종 50% 적용
Data.CostAccountDeduction = '2' //원가계산 제경비 99.7% 적용
Data.BidPriceRaise = '2' //투찰금액 천원 절상
Data.LaborCostLowBound = '2' //노무비 하한 80%
Data.ExecuteReset = '0' //Reset 함수 사용시 단가 소수처리 옵션과 별개로 소수 첫째자리 아래로 절사
exports.Data = Data
