"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.Data = void 0;
var path = require("path");
var big_js_1 = require("big.js");
var Data = exports.Data = /** @class */ (function () {
    function Data(item, constructionNum, workNum, detailWorkNum, code, name, standard, unit, quantity, materialUnit, laborUnit, expenseUnit) {
        this.materialUnit = 0; //재료비 단가
        this.laborUnit = 0; //노무비 단가
        this.expenseUnit = 0; //경비 단가
        this.Item = ''; //항목 구분(공종(입력불가), 무대(입력불가), 일반, 관급, PS, 제요율적용제외, 안전관리비, PS내역, 표준시장단가)
        this.ConstructionNum = ''; //공사 인덱스
        this.WorkNum = ''; //세부 공사 인덱스
        this.DetailWorkNum = ''; //세부 공종 인덱스
        this.Code = ''; //코드
        this.Name = ''; //품명
        this.Standard = ''; //규격
        this.Unit = ''; //단위
        this.Quantity = 0; //수량
        this.Weight = 0; //가중치
        this.PriceScore = 0; //세부 점수
        this.Item = item;
        this.ConstructionNum = constructionNum;
        this.WorkNum = workNum;
        this.DetailWorkNum = detailWorkNum;
        this.Code = code;
        this.Name = name;
        this.Standard = standard;
        this.Unit = unit;
        this.Quantity = quantity;
        this.MaterialUnit = materialUnit;
        this.LaborUnit = laborUnit;
        this.ExpenseUnit = expenseUnit;
    }
    Object.defineProperty(Data.prototype, "MaterialUnit", {
        /**재료비 단가*/
        get: function () {
            //사용자가 단가 정수처리를 원한다면("2") 정수 값으로 return / Reset 함수를 쓰지 않은 경우의 조건 추가 (23.02.06)
            if (Data.UnitPriceTrimming === '2' && Data.ExecuteReset === '0')
                return Number(new big_js_1.default(this.materialUnit).toFixed(0, 3));
            else if (Data.UnitPriceTrimming === '1' || Data.ExecuteReset === '1')
                // 사용자가 단가 소수점 처리를 원하거나 Reset 함수를 썼다면 소수 첫째 자리 아래로 절사 (23.02.06)
                return Number(new big_js_1.default(this.materialUnit).toFixed(1, 0));
            return this.materialUnit; //Default는 있는 그대로의 값을 return
        },
        set: function (value) {
            this.materialUnit = value;
        },
        enumerable: false,
        configurable: true
    });
    Object.defineProperty(Data.prototype, "LaborUnit", {
        //노무비 단가
        get: function () {
            if (Data.UnitPriceTrimming === '2' && Data.ExecuteReset === '0')
                return Number(new big_js_1.default(this.laborUnit).toFixed(0, 3));
            else if (Data.UnitPriceTrimming === '1' || Data.ExecuteReset === '1')
                return Number(new big_js_1.default(this.laborUnit).toFixed(1, 0));
            return this.laborUnit;
        },
        set: function (value) {
            this.laborUnit = value;
        },
        enumerable: false,
        configurable: true
    });
    Object.defineProperty(Data.prototype, "ExpenseUnit", {
        //경비 단가
        get: function () {
            var decimal = new big_js_1.default(0);
            if (Data.UnitPriceTrimming === '2' && Data.ExecuteReset === '0')
                return Number(new big_js_1.default(this.expenseUnit).toFixed(0, 3));
            else if (Data.UnitPriceTrimming === '1' || Data.ExecuteReset === '1')
                return Number(new big_js_1.default(this.expenseUnit).toFixed(1, 0));
            return this.expenseUnit;
        },
        set: function (value) {
            this.expenseUnit = value;
        },
        enumerable: false,
        configurable: true
    });
    Object.defineProperty(Data.prototype, "Material", {
        get: function () {
            var decimal = new big_js_1.default(this.Quantity).times(this.MaterialUnit);
            return Number(decimal.toFixed(0, 0));
        } //재료비 (수량 x 단가)
        ,
        enumerable: false,
        configurable: true
    });
    Object.defineProperty(Data.prototype, "Labor", {
        get: function () {
            var decimal = new big_js_1.default(this.Quantity).times(this.LaborUnit);
            return Number(decimal.toFixed(0, 0));
        } //노무비
        ,
        enumerable: false,
        configurable: true
    });
    Object.defineProperty(Data.prototype, "Expense", {
        get: function () {
            var decimal = new big_js_1.default(this.Quantity).times(this.ExpenseUnit);
            return Number(decimal.toFixed(0, 0));
        } //경비
        ,
        enumerable: false,
        configurable: true
    });
    Object.defineProperty(Data.prototype, "UnitPriceSum", {
        get: function () {
            var decimal = new big_js_1.default(this.MaterialUnit).plus(this.LaborUnit).plus(this.ExpenseUnit);
            return decimal.toNumber();
        } //합계단가
        ,
        enumerable: false,
        configurable: true
    });
    Object.defineProperty(Data.prototype, "PriceSum", {
        get: function () {
            var decimal = new big_js_1.default(this.Material).plus(this.Labor).plus(this.Expense);
            return decimal.toNumber();
        } //합계(세부공종별 금액의 합계)
        ,
        enumerable: false,
        configurable: true
    });
    Object.defineProperty(Data.prototype, "Score", {
        get: function () {
            return this.PriceScore * this.Weight;
        } //단가 점수(세부 점수 * 가중치)
        ,
        enumerable: false,
        configurable: true
    });
    Object.defineProperty(Data, "BalancedRate", {
        //업체 평균 예측율
        get: function () {
            //decimal로 반환. 소수 계산 라이브러리 필요
            return Data.BalanceRateNum; //입력받은 BalancedRateNum(double? 형)을 decimal로 바꿈
        },
        enumerable: false,
        configurable: true
    });
    Object.defineProperty(Data, "PersonalRate", {
        //내 예가 사정률
        get: function () {
            //decimal로 반환. 소수 계산 라이브러리 필요
            return Data.PersonalRateNum; //입력받은 PersonalRateNum(double? 형)을 decimal로 바꿈
        },
        enumerable: false,
        configurable: true
    });
    Data.reset_data = function () {
        this.XlsText = '';
        this.CanCovertFile = false; // 새로운 파일 업로드 시 변환 가능
        this.IsConvert = false; // 변환을 했는지 안했는지
        this.IsBidFileOk = true; // 정상적인 공내역 파일인지
        this.IsFileMatch = true; // 공내역 파일과 실내역 파일의 공사가 일치하는지
        this.CompanyRegistrationNum = ''; //1.31 사업자등록번호 추가
        this.CompanyRegistrationName = ''; // 2.02 회사명 추가
        this.PersonalRateNum = 0; // 내 사정율 변수
        //C#에서는 BalanceRate 속성에서 get을 통해 ToDecimal(BalanceRateNum) 리턴받음.
        //js는 number가 부동소수점
        this.BalanceRateNum = 0; // 업체 평균 사정율 변수
        this.ConstructionTerm = 0; //공사 기간
        this.RealDirectMaterial = 0; //실내역 직접 재료비(일반, - , 표준시장단가)
        this.RealDirectLabor = 0; //실내역 직접 노무비(일반, - , 표준시장단가)
        this.RealOutputExpense = 0; //실내역 산출 경비(일반, - , 표준시장단가)
        this.FixedPriceDirectMaterial = 0; //고정금액 항목 직접 재료비
        this.FixedPriceDirectLabor = 0; //고정금액 항목 직접 노무비
        this.FixedPriceOutputExpense = 0; //고정금액 항목 산출 경비
        this.RealPriceDirectMaterial = 0; //일반항목 직접 재료비
        this.RealPriceDirectLabor = 0; //일반항목 직접 노무비
        this.RealPriceOutputExpense = 0; //일반항목 산출 경비
        this.InvestigateFixedPriceDirectMaterial = 0; //고정금액 항목 직접 재료비(조사금액)
        this.InvestigateFixedPriceDirectLabor = 0; //고정금액 항목 직접 노무비(조사금액)
        this.InvestigateFixedPriceOutputExpense = 0; //고정금액 항목 산출 경비(조사금액)
        this.InvestigateStandardMaterial = 0; //표준시장단가 재료비(조사금액)
        this.InvestigateStandardLabor = 0; //표준시장단가 노무비(조사금액)
        this.InvestigateStandardExpense = 0; //표준시장단가 산출경비(조사금액)
        this.PsMaterial = 0; //PS(재료비) 금액(직접 재료비에서 제외)
        this.PsLabor = 0; //PS(노무비) 금액(직접 노무비에서 제외)
        this.PsExpense = 0; //PS(경비) 금액(산출 경비에서 제외)
        this.ExcludingMaterial = 0; //제요율적용제외(재료비) 금액(직접 재료비에서 제외)
        this.ExcludingLabor = 0; //제요율적용제외(노무비) 금액(직접 노무비에서 제외)
        this.ExcludingExpense = 0; //제요율적용제외(경비) 금액(산출 경비에서 제외)
        this.AdjustedExMaterial = 0; //사정율 적용한 제요율적용제외 금액(재료비)
        this.AdjustedExLabor = 0; //사정율 적용한 제요율적용제외 금액(노무비)
        this.AdjustedExExpense = 0; //사정율 적용한 제요율적용제외 금액(경비)
        this.GovernmentMaterial = 0; //관급자재요소(재료비) 금액(직접 재료비에서 제외)
        this.GovernmentLabor = 0; //관급자재요소(노무비) 금액(직접 노무비에서 제외)
        this.GovernmentExpense = 0; //관급자재요소(경비) 금액(산출 경비에서 제외)
        this.SafetyPrice = 0; //안전관리비(산출 경비에서 제외)
        this.StandardMaterial = 0; //표준시장단가 재료비
        this.StandardLabor = 0; //표준시장단가 노무비
        this.StandardExpense = 0; //표준시장단가 산출경비
        this.InvestigateStandardMarket = 0; //표준시장단가 합계(조사내역)
        this.FixedPricePercent = 0; //고정금액 비중
        this.ByProduct = 0; //작업설
        this.Dic = new Map(); //key : 세부공사별 번호 / value : 세부공사별 리스트
        this.ConstructionNums = new Map(); //세부 공사별 번호 저장
        this.MatchedConstNum = new Map(); //실내역과 세부공사별 번호의 매칭 결과
        this.Fixed = new Map(); //고정금액 항목별 금액 저장
        this.Rate1 = new Map(); //적용비율1 저장
        this.Rate2 = new Map(); //적용비율2 저장
        this.RealPrices = new Map(); //원가계산서 항목별 금액 저장
        this.Investigation = new Map(); //세부결과_원가계산서 항목별 조사금액 저장
        this.Bidding = new Map(); //세부결과_원가계산서 항목별 입찰금액 저장
        this.Correction = new Map(); //원가계산서 조사금액 보정 항목 저장
        //사용자의 옵션 및 사정률 데이터
        this.UnitPriceTrimming = '0'; //단가 소수 처리 (defalut = "0")
        this.StandardMarketDeduction = '2'; //표준시장단가 99.7% 적용
        this.ZeroWeightDeduction = '2'; //가중치 0% 공종 50% 적용
        this.CostAccountDeduction = '2'; //원가계산 제경비 99.7% 적용
        this.BidPriceRaise = '2'; //투찰금액 천원 절상
        this.LaborCostLowBound = '2'; //노무비 하한 80%
    };
    Data.folder = path.resolve(__dirname, '../AutoBid'); //내 문서 폴더의 AutoBID 폴더로 지정 (23.02.02)
    // WPF 앱 파일 관리 변수
    Data.XlsText = '';
    Data.CanCovertFile = false; // 새로운 파일 업로드 시 변환 가능
    Data.IsConvert = false; // 변환을 했는지 안했는지
    Data.IsBidFileOk = true; // 정상적인 공내역 파일인지
    Data.IsFileMatch = true; // 공내역 파일과 실내역 파일의 공사가 일치하는지
    Data.CompanyRegistrationNum = ''; //1.31 사업자등록번호 추가
    Data.CompanyRegistrationName = ''; // 2.02 회사명 추가
    // 프로그램 폴더로 위치 변경
    Data.work_path = path.join(Data.folder, 'WORK DIRECTORY'); //작업폴더(WORK DIRECTORY) 경로
    Data.RealDirectMaterial = 0; //실내역 직접 재료비(일반, - , 표준시장단가)
    Data.RealDirectLabor = 0; //실내역 직접 노무비(일반, - , 표준시장단가)
    Data.RealOutputExpense = 0; //실내역 산출 경비(일반, - , 표준시장단가)
    Data.FixedPriceDirectMaterial = 0; //고정금액 항목 직접 재료비
    Data.FixedPriceDirectLabor = 0; //고정금액 항목 직접 노무비
    Data.FixedPriceOutputExpense = 0; //고정금액 항목 산출 경비
    Data.RealPriceDirectMaterial = 0; //일반항목 직접 재료비
    Data.RealPriceDirectLabor = 0; //일반항목 직접 노무비
    Data.RealPriceOutputExpense = 0; //일반항목 산출 경비
    Data.InvestigateFixedPriceDirectMaterial = 0; //고정금액 항목 직접 재료비(조사금액)
    Data.InvestigateFixedPriceDirectLabor = 0; //고정금액 항목 직접 노무비(조사금액)
    Data.InvestigateFixedPriceOutputExpense = 0; //고정금액 항목 산출 경비(조사금액)
    Data.InvestigateStandardMaterial = 0; //표준시장단가 재료비(조사금액)
    Data.InvestigateStandardLabor = 0; //표준시장단가 노무비(조사금액)
    Data.InvestigateStandardExpense = 0; //표준시장단가 산출경비(조사금액)
    Data.PsMaterial = 0; //PS(재료비) 금액(직접 재료비에서 제외)
    Data.PsLabor = 0; //PS(노무비) 금액(직접 노무비에서 제외)
    Data.PsExpense = 0; //PS(경비) 금액(산출 경비에서 제외)
    Data.ExcludingMaterial = 0; //제요율적용제외(재료비) 금액(직접 재료비에서 제외)
    Data.ExcludingLabor = 0; //제요율적용제외(노무비) 금액(직접 노무비에서 제외)
    Data.ExcludingExpense = 0; //제요율적용제외(경비) 금액(산출 경비에서 제외)
    Data.AdjustedExMaterial = 0; //사정율 적용한 제요율적용제외 금액(재료비)
    Data.AdjustedExLabor = 0; //사정율 적용한 제요율적용제외 금액(노무비)
    Data.AdjustedExExpense = 0; //사정율 적용한 제요율적용제외 금액(경비)
    Data.GovernmentMaterial = 0; //관급자재요소(재료비) 금액(직접 재료비에서 제외)
    Data.GovernmentLabor = 0; //관급자재요소(노무비) 금액(직접 노무비에서 제외)
    Data.GovernmentExpense = 0; //관급자재요소(경비) 금액(산출 경비에서 제외)
    Data.SafetyPrice = 0; //안전관리비(산출 경비에서 제외)
    Data.StandardMaterial = 0; //표준시장단가 재료비
    Data.StandardLabor = 0; //표준시장단가 노무비
    Data.StandardExpense = 0; //표준시장단가 산출경비
    Data.InvestigateStandardMarket = 0; //표준시장단가 합계(조사내역)
    Data.FixedPricePercent = 0; //고정금액 비중
    Data.ByProduct = 0; //작업설
    Data.Dic = new Map(); //key : 세부공사별 번호 / value : 세부공사별 리스트
    Data.ConstructionNums = new Map(); //세부 공사별 번호 저장
    Data.MatchedConstNum = new Map(); //실내역과 세부공사별 번호의 매칭 결과
    Data.Fixed = new Map(); //고정금액 항목별 금액 저장
    Data.Rate1 = new Map(); //적용비율1 저장
    Data.Rate2 = new Map(); //적용비율2 저장
    Data.RealPrices = new Map(); //원가계산서 항목별 금액 저장
    Data.Investigation = new Map(); //세부결과_원가계산서 항목별 조사금액 저장
    Data.Bidding = new Map(); //세부결과_원가계산서 항목별 입찰금액 저장
    Data.Correction = new Map(); //원가계산서 조사금액 보정 항목 저장
    //사용자의 옵션 및 사정률 데이터
    Data.UnitPriceTrimming = '0'; //단가 소수 처리 (defalut = "0")
    Data.StandardMarketDeduction = '2'; //표준시장단가 99.7% 적용
    Data.ZeroWeightDeduction = '2'; //가중치 0% 공종 50% 적용
    Data.CostAccountDeduction = '2'; //원가계산 제경비 99.7% 적용
    Data.BidPriceRaise = '2'; //투찰금액 천원 절상
    Data.LaborCostLowBound = '2'; //노무비 하한 80%
    Data.ExecuteReset = '0'; //Reset 함수 사용시 단가 소수처리 옵션과 별개로 소수 첫째자리 아래로 절사
    return Data;
}());
