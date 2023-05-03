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
Object.defineProperty(exports, "__esModule", { value: true });
exports.CalculatePrice = void 0;
//ApplyStandardPriceOption 소수점 0.1 더하기에서 오차 발생
const Data_1 = require("./Data");
const FillCostAccount_1 = require("./FillCostAccount");
const CreateResultFile_1 = require("./CreateResultFile");
const fs = __importStar(require("fs"));
const AdmZip = require('adm-zip');
//실행 위해 프로그램 내 폴더로 경로 변경
class CalculatePrice {
    static Calculation() {
        const bidString = fs.readFileSync(Data_1.Data.folder + '\\WORK DIRECTORY\\Setting_Json.json', 'utf-8');
        this.docBID = JSON.parse(bidString);
        this.eleBID = this.docBID['data'];
        FillCostAccount_1.FillCostAccount.CheckKeyNotFound();
        FillCostAccount_1.FillCostAccount.CalculateInvestigationCosts(Data_1.Data.Correction);
        FillCostAccount_1.FillCostAccount.CalculateInvestigationCosts(Data_1.Data.Correction);
        //FillConstAccount.FillInvestigationCosts() 내용
        let FM = Data_1.Data.FixedPriceDirectMaterial;
        let FL = Data_1.Data.FixedPriceDirectLabor;
        let FOE = Data_1.Data.FixedPriceOutputExpense;
        let SM = Data_1.Data.StandardMaterial;
        let SL = Data_1.Data.StandardLabor;
        let SOE = Data_1.Data.StandardExpense;
        Data_1.Data.InvestigateFixedPriceDirectMaterial = FM;
        Data_1.Data.InvestigateFixedPriceDirectLabor = FL;
        Data_1.Data.InvestigateFixedPriceOutputExpense = FOE;
        Data_1.Data.InvestigateStandardMaterial = SM;
        Data_1.Data.InvestigateStandardLabor = SL;
        Data_1.Data.InvestigateStandardExpense = SOE;
        //가격 재세팅 후 리셋 함수 실행 횟수 증가
        this.Reset();
        //최저네고단가율 계산 전, 표준시장단가 99.7% 적용옵션에 따른 분기처리
        if (Data_1.Data.StandardMarketDeduction === '1')
            this.ApplyStandardPriceOption();
        this.GetFixedPriceRate(); //직공비 대비 고정금액 비중 계산
        this.FindMyPercent(); //최저네고단가율 계산
        this.GetWeight(); //가중치 계산
        this.CalculateRate(Data_1.Data.PersonalRate, Data_1.Data.BalancedRate); //Target Rate 계산
        this.Recalculation(); //사정율에 따른 재계산
        if (this.exCount !== 0) {
            this.SetExcludingPrice(); //제요율적용제외공종 항목 Target Rate 적용
            this.GetAdjustedExcludePrice(); //사정율 적용한 제요율적용제외 금액 저장
        }
        this.SetPriceOfSuperConstruction(); //공종 합계 bid에 저장 (23.02.07)
        FillCostAccount_1.FillCostAccount.CalculateBiddingCosts(); //원가계산서 사정율적용(입찰) 금액 계산 및 저장
        this.SetBusinessInfo(); //사업자등록번호 <T1></C17></T1>에 추가
        this.SubstitutePrice(); //원가계산서 사정율 적용하여 계산한 금액들 BID 파일에도 반영
        this.CreateFile(); //입찰내역 파일 생성
    }
    static Reset() {
        Data_1.Data.ExecuteReset = '1'; //Reset 함수 사용 여부
        let DM = Data_1.Data.Investigation.get('직접재료비');
        let DL = Data_1.Data.Investigation.get('직접노무비');
        let OE = Data_1.Data.Investigation.get('산출경비');
        let FM = Data_1.Data.InvestigateFixedPriceDirectMaterial;
        let FL = Data_1.Data.InvestigateFixedPriceDirectLabor;
        let FOE = Data_1.Data.InvestigateFixedPriceOutputExpense;
        let SM = Data_1.Data.InvestigateStandardMaterial;
        let SL = Data_1.Data.InvestigateStandardLabor;
        let SOE = Data_1.Data.InvestigateStandardExpense;
        //조사 내역서 정보 백업
        Data_1.Data.RealDirectMaterial = DM;
        Data_1.Data.RealDirectLabor = DL;
        Data_1.Data.RealOutputExpense = OE;
        Data_1.Data.FixedPriceDirectMaterial = FM;
        Data_1.Data.FixedPriceDirectLabor = FL;
        Data_1.Data.FixedPriceOutputExpense = FOE;
        Data_1.Data.StandardMaterial = SM;
        Data_1.Data.StandardLabor = SL;
        Data_1.Data.StandardExpense = SOE;
        //사정율 재적용을 위한 초기화
        const bidT3 = this.eleBID['T3'];
        let code;
        let type;
        for (let key in bidT3) {
            //Dictionary 초기화
            code = JSON.stringify(bidT3[key]['C9']['_text']);
            type = JSON.stringify(bidT3[key]['C5']['_text'])[1];
            if (code !== undefined && type === 'S') {
                let constNum = JSON.stringify(bidT3[key]['C1']['_text']).slice(1, -1); //세부공사 번호
                let numVal = JSON.stringify(bidT3[key]['C2']['_text']).slice(1, -1); //세부공종 번호
                let detailVal = JSON.stringify(bidT3[key]['C3']['_text']).slice(1, -1); //세부 공종 번호
                //현재 탐색 공종
                let curObject = Data_1.Data.Dic.get(constNum).find((x) => x.WorkNum === numVal && x.DetailWorkNum === detailVal);
                curObject.MaterialUnit = Number(bidT3[key]['C16']['_text']);
                curObject.LaborUnit = Number(bidT3[key]['C17']['_text']);
                curObject.ExpenseUnit = Number(bidT3[key]['C18']['_text']);
            }
        }
        Data_1.Data.ExecuteReset = '0'; //Reset 함수 사용이 끝나면 다시 0으로 초기화
    }
    static ApplyStandardPriceOption() {
        const bidT3 = this.eleBID['T3'];
        let code;
        let type;
        for (let key in bidT3) {
            code = JSON.stringify(bidT3[key]['C9']['_text']);
            type = JSON.stringify(bidT3[key]['C5']['_text'])[1];
            //표준시장단가 항목인경우 99.7% 적용
            if (code !== undefined && type === 'S') {
                let constNum = JSON.stringify(bidT3[key]['C1']['_text']).slice(1, -1); //세부공사 번호
                let numVal = JSON.stringify(bidT3[key]['C2']['_text']).slice(1, -1); //세부공종 번호
                let detailVal = JSON.stringify(bidT3[key]['C3']['_text']).slice(1, -1); //세부 공종 번호
                let curObject = Data_1.Data.Dic.get(constNum).find((x) => x.WorkNum === numVal && x.DetailWorkNum === detailVal);
                if (curObject.Item === '표준시장단가') {
                    //직공비, 고정금액, 표준시장단가 금액 재계산
                    Data_1.Data.RealDirectMaterial -= Number(bidT3[key]['C20']['_text']);
                    Data_1.Data.RealDirectLabor -= Number(bidT3[key]['C21']['_text']);
                    Data_1.Data.RealOutputExpense -= Number(bidT3[key]['C22']['_text']);
                    Data_1.Data.FixedPriceDirectMaterial -= Number(bidT3[key]['C20']['_text']);
                    Data_1.Data.FixedPriceDirectLabor -= Number(bidT3[key]['C21']['_text']);
                    Data_1.Data.FixedPriceOutputExpense -= Number(bidT3[key]['C22']['_text']);
                    Data_1.Data.StandardMaterial -= Number(bidT3[key]['C20']['_text']);
                    Data_1.Data.StandardLabor -= Number(bidT3[key]['C21']['_text']);
                    Data_1.Data.StandardExpense -= Number(bidT3[key]['C22']['_text']);
                    //표준시장단가 99.7% 적용
                    if (curObject.MaterialUnit !== 0)
                        curObject.MaterialUnit =
                            Math.trunc(curObject.MaterialUnit * 0.997 * 10 + 1) / 10;
                    if (curObject.LaborUnit !== 0)
                        curObject.LaborUnit = Math.trunc(curObject.LaborUnit * 0.997 * 10 + 1) / 10;
                    if (curObject.ExpenseUnit !== 0)
                        curObject.ExpenseUnit =
                            Math.trunc(curObject.ExpenseUnit * 0.997 * 10 + 1) / 10;
                    //단가 변경사항 JSON에 적용
                    bidT3[key]['C16']['_text'] = curObject.MaterialUnit.toString(); //재료비 단가
                    bidT3[key]['C17']['_text'] = curObject.LaborUnit.toString(); //노무비 단가
                    bidT3[key]['C18']['_text'] = curObject.ExpenseUnit.toString(); //경비 단가
                    bidT3[key]['C19']['_text'] = curObject.UnitPriceSum.toString(); //합계 단가
                    bidT3[key]['C20']['_text'] = curObject.Material.toString(); //재료비
                    bidT3[key]['C21']['_text'] = curObject.Labor.toString(); //노무비
                    bidT3[key]['C22']['_text'] = curObject.Expense.toString(); //경비
                    bidT3[key]['C23']['_text'] = curObject.PriceSum.toString(); //합계
                    //붙여넣기한 각 객체의 재료비, 노무비, 경비를 직접재료비, 직접노무비, 산출 경비에 더해나감
                    Data_1.Data.RealDirectMaterial += Number(bidT3[key]['C20']['_text']);
                    Data_1.Data.RealDirectLabor += Number(bidT3[key]['C21']['_text']);
                    Data_1.Data.RealOutputExpense += Number(bidT3[key]['C22']['_text']);
                    Data_1.Data.FixedPriceDirectMaterial += Number(bidT3[key]['C20']['_text']);
                    Data_1.Data.FixedPriceDirectLabor += Number(bidT3[key]['C21']['_text']);
                    Data_1.Data.FixedPriceOutputExpense += Number(bidT3[key]['C22']['_text']);
                    Data_1.Data.StandardMaterial += Number(bidT3[key]['C20']['_text']);
                    Data_1.Data.StandardLabor += Number(bidT3[key]['C21']['_text']);
                    Data_1.Data.StandardExpense += Number(bidT3[key]['C22']['_text']);
                }
            }
        }
    }
    static GetFixedPriceRate() {
        //고정금액 비율 계산
        const directConstPrice = Data_1.Data.Investigation.get('직공비');
        const fixCostSum = Data_1.Data.InvestigateFixedPriceDirectMaterial +
            Data_1.Data.InvestigateFixedPriceDirectLabor +
            Data_1.Data.InvestigateFixedPriceOutputExpense;
        Data_1.Data.FixedPricePercent = Math.trunc((fixCostSum / directConstPrice) * 100 * 10000) / 10000; // 고정금액 비중 계산 / 고정금액 소수점 5자리 수에서 절사 (23.02.06)
    }
    static FindMyPercent() {
        //고정금액 비중에 따른 최저네고단가율 계산
        if (Data_1.Data.FixedPricePercent < 20.0)
            //고정금액 < 20%
            this.myPercent = 0.85;
        else if (Data_1.Data.FixedPricePercent < 25.0)
            //고정금액 < 25%
            this.myPercent = 0.84;
        else if (Data_1.Data.FixedPricePercent < 30.0)
            //고정금액 < 30%
            this.myPercent = 0.83;
        else
            this.myPercent = 0.82; //고정금액 > 30%
    }
    static GetWeight() {
        //가중치 계산
        let varCostSum = Data_1.Data.RealPriceDirectMaterial + Data_1.Data.RealPriceDirectLabor + Data_1.Data.RealPriceOutputExpense; //총 합계금액(-,PS,표준시장단가 제외)
        let weight;
        let maxWeight = 0;
        let weightSum = 0;
        //Data 인스턴스 생성시 생성자 함수로 프로퍼티 값을 설정하는 것으로 변경
        //->max를 빈 변수로 선언한 뒤 후에 Data인스턴스를 넣도록 수정
        let max;
        Data_1.Data.Dic.forEach((value, _) => {
            for (let idx in value) {
                if (value[idx].Item === '일반') {
                    let material = value[idx].Material;
                    let labor = value[idx].Labor;
                    let expense = value[idx].Expense;
                    weight =
                        Math.round(((material + labor + expense) / varCostSum) * 1000000) / 1000000; //소숫점 일곱 자리 반올림
                    weightSum += weight; //가중치를 더함
                    if (maxWeight < weight) {
                        //최대 가중치 갱신
                        maxWeight = weight;
                        max = value[idx];
                    }
                    value[idx].Weight = weight;
                }
            }
        });
        if (weightSum !== 1.0) {
            //가중치의 합이 1이 되지 않으면 가중치가 가장 큰 항목에 부족한 양을 더한다
            let lack = 1.0 - weightSum;
            max.Weight += lack;
        }
    }
    static CalculateRate(presonalRate, balancedRate) {
        //Target Rate 계산
        const unitPrice = 100;
        this.balancedUnitPriceRate =
            (0.9 * unitPrice * (1.0 + balancedRate / 100) * this.myPercent) /
                (1.0 - 0.1 * this.myPercent) /
                100; //균형단가율
        this.targetRate =
            ((unitPrice * (1.0 + presonalRate / 100) * 0.9 +
                unitPrice * this.balancedUnitPriceRate * 0.1) *
                this.myPercent) /
                100; //Target_Rate
        this.targetRate = Math.trunc(this.targetRate * 1000000) / 1000000;
    }
    static RoundOrTruncate(Rate, Object, refMyMaterialUnit, refMyLaborUnit, refMyExpenseUnit) {
        //절사,반올림 옵션
        if (Data_1.Data.UnitPriceTrimming === '1') {
            refMyMaterialUnit.value = Math.trunc(Object.MaterialUnit * Rate * 10) / 10;
            refMyLaborUnit.value = Math.trunc(Object.LaborUnit * Rate * 10) / 10;
            refMyExpenseUnit.value = Math.trunc(Object.ExpenseUnit * Rate * 10) / 10;
        }
        else if (Data_1.Data.UnitPriceTrimming === '2') {
            refMyMaterialUnit.value = Math.ceil(Object.MaterialUnit * Rate);
            refMyLaborUnit.value = Math.ceil(Object.LaborUnit * Rate);
            refMyExpenseUnit.value = Math.ceil(Object.ExpenseUnit * Rate);
        }
    }
    static CheckLaborLimit80(Object, refMyMaterialUnit, refMyLaborUnit, refMyExpenseUnit) {
        //2.8 노무비 80%미만일 경우 조정하는 메소드
        if (Object.LaborUnit * 0.8 > refMyLaborUnit.value) {
            const deficiency = Object.LaborUnit * 0.8 - refMyLaborUnit.value;
            if (refMyLaborUnit.value !== 0)
                refMyMaterialUnit.value -= deficiency;
            else if (refMyExpenseUnit.value !== 0)
                refMyExpenseUnit.value -= deficiency;
            refMyLaborUnit.value = Object.LaborUnit * 0.8;
        }
    }
    static Recalculation() {
        //사정율에 따라 재계산된 가격을 비드파일에 복사
        const bidT3 = this.eleBID['T3'];
        let code;
        let type;
        this.exCount = 0;
        this.exSum = 0;
        for (let key in bidT3) {
            code = JSON.stringify(bidT3[key]['C9']['_text']);
            type = JSON.stringify(bidT3[key]['C5']['_text'])[1];
            if (code !== undefined && type === 'S') {
                let constNum = JSON.stringify(bidT3[key]['C1']['_text']).slice(1, -1); //세부공사 번호
                let numVal = JSON.stringify(bidT3[key]['C2']['_text']).slice(1, -1); //세부공종 번호
                let detailVal = JSON.stringify(bidT3[key]['C3']['_text']).slice(1, -1); //세부 공종 번호
                let curObject = Data_1.Data.Dic.get(constNum).find((x) => x.WorkNum === numVal && x.DetailWorkNum === detailVal);
                //일반 항목인 경우
                if (curObject.Item === '일반') {
                    //직접공사비 재계산
                    Data_1.Data.RealDirectMaterial -= Number(bidT3[key]['C20']['_text']);
                    Data_1.Data.RealDirectLabor -= Number(bidT3[key]['C21']['_text']);
                    Data_1.Data.RealOutputExpense -= Number(bidT3[key]['C22']['_text']);
                    let targetPrice = (curObject.MaterialUnit + curObject.LaborUnit + curObject.ExpenseUnit) *
                        this.targetRate; //Target 단가 합계
                    //my 단가를 구하는 과정도 사용자의 옵션에 따라 소수 첫째 자리 아래로 절사(1) / 정수(2)로 나뉜다.
                    let myMaterialUnit = { value: 0 }; //reference로 값을 넘겨주기 위해 number를 object로 감싸서 만듬
                    let myLaborUnit = { value: 0 }; //reference로 값을 넘겨주기 위해 number를 object로 감싸서 만듬
                    let myExpenseUnit = { value: 0 }; //reference로 값을 넘겨주기 위해 number를 object로 감싸서 만듬
                    let myPrice;
                    if (Data_1.Data.ZeroWeightDeduction === '1') {
                        //최소단가율 50% 적용 O
                        if (curObject.Weight === 0 && curObject.LaborUnit === 0) {
                            //공종 가중치 0%이고 노무비 단가가 0원인 경우 사용자의 소수처리 옵션과 상관없이 50% 적용후 소수첫째자리에서 올림 (23.2.23)
                            curObject.MaterialUnit = Math.ceil(curObject.MaterialUnit * 0.5);
                            curObject.ExpenseUnit = Math.ceil(curObject.ExpenseUnit * 0.5);
                            //최종 단가 및 합계 계산
                            bidT3[key]['C16']['_text'] = curObject.MaterialUnit.toString(); //재료비 단가
                            bidT3[key]['C17']['_text'] = curObject.LaborUnit.toString(); //노무비 단가
                            bidT3[key]['C18']['_text'] = curObject.ExpenseUnit.toString(); //경비 단가
                            bidT3[key]['C19']['_text'] = curObject.UnitPriceSum.toString(); //합계 단가
                            bidT3[key]['C20']['_text'] = curObject.Material.toString(); //재료비
                            bidT3[key]['C21']['_text'] = curObject.Labor.toString(); //노무비
                            bidT3[key]['C22']['_text'] = curObject.Expense.toString(); //경비
                            bidT3[key]['C23']['_text'] = curObject.PriceSum.toString(); //합계
                            //붙여넣기한 각 객체의 재료비, 노무비, 경비를 직접재료비, 직접노무비, 산출 경비에 더해나감
                            Data_1.Data.RealDirectMaterial += Number(bidT3[key]['C20']['_text']);
                            Data_1.Data.RealDirectLabor += Number(bidT3[key]['C21']['_text']);
                            Data_1.Data.RealOutputExpense += Number(bidT3[key]['C22']['_text']);
                            continue;
                        }
                        else {
                            this.RoundOrTruncate(this.targetRate, curObject, myMaterialUnit, myLaborUnit, myExpenseUnit);
                            this.CheckLaborLimit80(curObject, myMaterialUnit, myLaborUnit, myExpenseUnit);
                        }
                    }
                    else if (Data_1.Data.ZeroWeightDeduction === '2') {
                        //최소단가율 50% 적용 X
                        this.RoundOrTruncate(this.targetRate, curObject, myMaterialUnit, myLaborUnit, myExpenseUnit);
                        this.CheckLaborLimit80(curObject, myMaterialUnit, myLaborUnit, myExpenseUnit);
                    }
                    myPrice = myMaterialUnit.value + myLaborUnit.value + myExpenseUnit.value;
                    if (Data_1.Data.LaborCostLowBound === '1') {
                        //노무비 하한 80% 적용 O
                        //여유분 조정 가능(조사노무비 대비 My노무비 비율에 따라 조정) <- ?
                        let Excess = myPrice - targetPrice;
                        let laborExcess = myLaborUnit.value - curObject.LaborUnit * 0.8;
                        laborExcess = Math.trunc(laborExcess * 10) / 10;
                        if (laborExcess > 0) {
                            if (myExpenseUnit.value !== 0) {
                                myLaborUnit.value -= laborExcess;
                                myExpenseUnit.value += laborExcess + Excess;
                            }
                            else {
                                if (myMaterialUnit.value !== 0) {
                                    myLaborUnit.value -= laborExcess;
                                    myMaterialUnit.value += laborExcess + Excess;
                                }
                                else {
                                    myLaborUnit.value -= laborExcess;
                                    myExpenseUnit.value += laborExcess + Excess;
                                }
                            }
                        }
                        else if (laborExcess < 0) {
                            myLaborUnit.value = curObject.LaborUnit * 0.8;
                            if (myMaterialUnit.value !== 0) {
                                myMaterialUnit.value += laborExcess + Excess;
                            }
                            else {
                                myExpenseUnit.value += laborExcess + Excess;
                            }
                        }
                    }
                    curObject.MaterialUnit = myMaterialUnit.value;
                    curObject.LaborUnit = myLaborUnit.value;
                    curObject.ExpenseUnit = myExpenseUnit.value;
                    //최종 단가 및 합계 계산
                    bidT3[key]['C16']['_text'] = curObject.MaterialUnit.toString(); //재료비 단가
                    bidT3[key]['C17']['_text'] = curObject.LaborUnit.toString(); //노무비 단가
                    bidT3[key]['C18']['_text'] = curObject.ExpenseUnit.toString(); //경비 단가
                    bidT3[key]['C19']['_text'] = curObject.UnitPriceSum.toString(); //합계 단가
                    bidT3[key]['C20']['_text'] = curObject.Material.toString(); //재료비
                    bidT3[key]['C21']['_text'] = curObject.Labor.toString(); //노무비
                    bidT3[key]['C22']['_text'] = curObject.Expense.toString(); //경비
                    bidT3[key]['C23']['_text'] = curObject.PriceSum.toString(); //합계
                    //붙여넣기한 각 객체의 재료비, 노무비, 경비를 직접재료비, 직접노무비, 산출 경비에 더해나감
                    Data_1.Data.RealDirectMaterial += Number(bidT3[key]['C20']['_text']);
                    Data_1.Data.RealDirectLabor += Number(bidT3[key]['C21']['_text']);
                    Data_1.Data.RealOutputExpense += Number(bidT3[key]['C22']['_text']);
                }
                //제요율적용제외공종 단가 재세팅
                else if (curObject.Item === '제요율적용제외') {
                    curObject.MaterialUnit =
                        Math.trunc(curObject.MaterialUnit * this.targetRate * 10) / 10;
                    curObject.LaborUnit =
                        Math.trunc(curObject.LaborUnit * this.targetRate * 10) / 10;
                    curObject.ExpenseUnit =
                        Math.trunc(curObject.ExpenseUnit * this.targetRate * 10) / 10;
                    console.log(curObject);
                    this.exSum += curObject.PriceSum; //사정율을 적용한 제요율적용제외공종 항목의 합계
                    this.exCount++; //제요율적용제외공종 항목 수
                }
            }
        }
    }
    static SetExcludingPrice() {
        //제요율적용제외공종 단가 처리 및 재세팅
        let TempInvestDirectSum = Data_1.Data.Investigation.get('직공비'); //조사직공비
        let TempRealDirectSum = FillCostAccount_1.FillCostAccount.ToLong(Data_1.Data.RealDirectMaterial + Data_1.Data.RealDirectLabor + Data_1.Data.RealOutputExpense); //사정율적용 직공비
        let InvestExSum = Data_1.Data.ExcludingMaterial + Data_1.Data.ExcludingLabor + Data_1.Data.ExcludingExpense; //조사 제요율적용제외공종
        let TempExRate = Math.round((InvestExSum / TempInvestDirectSum) * 100000) / 100000; //조사 직공비 대비 조사 제요율적용제외공종 비율
        let TempExPrice = Math.ceil(TempRealDirectSum * TempExRate); //사정율적용 제요율적용제외공종
        let keyFound = 0; //금액이 가장 높은 항목에 부족분을 더하는 방법과 모든 항목에 분배해서 더하는 방법 분기 점
        const bidT3 = this.eleBID['T3'];
        let code;
        let type;
        if (Data_1.Data.CostAccountDeduction === '1') {
            TempExPrice = Math.ceil(Math.ceil(TempExPrice * 0.997));
            //제경비 99.7% 옵션 적용시 TempExPrice 업데이트
        }
        for (let key in bidT3) {
            code = JSON.stringify(bidT3[key]['C9']['_text']);
            type = JSON.stringify(bidT3[key]['C5']['_text'])[1];
            if (code !== undefined && type === 'S') {
                let constNum = JSON.stringify(bidT3[key]['C1']['_text']).slice(1, -1); //세부공사 번호
                let numVal = JSON.stringify(bidT3[key]['C2']['_text']).slice(1, -1); //세부공종 번호
                let detailVal = JSON.stringify(bidT3[key]['C3']['_text']).slice(1, -1); //세부 공종 번호
                let curObject = Data_1.Data.Dic.get(constNum).find((x) => x.WorkNum === numVal && x.DetailWorkNum === detailVal);
                //제요율적용제외공종 단가 재세팅
                if (curObject.Item === '제요율적용제외') {
                    if (this.maxBID === undefined) {
                        //maxBid 초기화
                        this.maxBID = bidT3[key];
                    }
                    if (bidT3[key]['C15']['_text'].slice(1, -1) === '1' &&
                        bidT3[key]['C15']['_text'].slice(1, -1) !== '0') {
                        if (Number(bidT3[key]['C19']['_text']) > Number(this.maxBID['C19']['_text'])) {
                            if (Number(bidT3[key]['C19']['_text']) * 1.5 >
                                curObject.PriceSum + (TempExPrice - this.exSum)) {
                                keyFound = 1;
                                this.maxBID = bidT3[key];
                            }
                        }
                    } //수량이 1이고 합계단가가 0이 아닐 때, 조정된 금액이 조사금액의 150% 미만이면 maxBid 업데이트
                    bidT3[key]['C16']['_text'] = curObject.MaterialUnit.toString(); //재료비 단가
                    bidT3[key]['C17']['_text'] = curObject.LaborUnit.toString(); //노무비 단가
                    bidT3[key]['C18']['_text'] = curObject.ExpenseUnit.toString(); //경비 단가
                    bidT3[key]['C19']['_text'] = curObject.UnitPriceSum.toString(); //합계 단가
                    bidT3[key]['C20']['_text'] = curObject.Material.toString(); //재료비
                    bidT3[key]['C21']['_text'] = curObject.Labor.toString(); //노무비
                    bidT3[key]['C22']['_text'] = curObject.Expense.toString(); //경비
                    bidT3[key]['C23']['_text'] = curObject.PriceSum.toString(); //합계
                }
            }
        }
        if (keyFound === 0) {
            //조건에 부합하는 maxBid를 찾지 못하면 모든 제요율적용제외공종 항목에 값을 분배하여 적용
            let divisionPrice = Math.trunc((TempExPrice - this.exSum) / this.exCount); //항목의 수에 따라 분배한 금액
            let deficiency = Math.ceil(TempExPrice - this.exSum - divisionPrice * this.exCount); //절사, 반올림에 따른 부족분
            let count = 0;
            while (count !== this.exCount) {
                const bidT3 = this.eleBID['T3'];
                let code;
                let type;
                for (let key in bidT3) {
                    code = JSON.stringify(bidT3[key]['C9']['_text']);
                    type = JSON.stringify(bidT3[key]['C5']['_text'])[1];
                    if (code !== undefined && type === 'S') {
                        let constNum = JSON.stringify(bidT3[key]['C1']['_text']).slice(1, -1); //세부공사 번호
                        let numVal = JSON.stringify(bidT3[key]['C2']['_text']).slice(1, -1); //세부공종 번호
                        let detailVal = JSON.stringify(bidT3[key]['C3']['_text']).slice(1, -1); //세부 공종 번호
                        let curObject = Data_1.Data.Dic.get(constNum).find((x) => x.WorkNum === numVal && x.DetailWorkNum === detailVal);
                        if (curObject.Item === '제요율적용제외' && curObject.Quantity === 1) {
                            if (curObject.LaborUnit !== 0) {
                                if (Number(bidT3[key]['C19']['_text']) * 1.5 >
                                    curObject.LaborUnit + divisionPrice) {
                                    curObject.LaborUnit += divisionPrice;
                                    bidT3[key]['C17']['_text'] = curObject.LaborUnit.toString(); //노무비 단가
                                    bidT3[key]['C19']['_text'] = curObject.UnitPriceSum.toString(); //합계 단가
                                    bidT3[key]['C21']['_text'] = curObject.Labor.toString(); //노무비
                                    bidT3[key]['C23']['_text'] = curObject.PriceSum.toString(); //합계
                                    count++;
                                }
                                if (count === this.exCount) {
                                    //절사, 반올림에 따른 부족분 조정
                                    bidT3[key]['C17']['_text'] = (deficiency + curObject.LaborUnit).toString(); //노무비 단가
                                    bidT3[key]['C19']['_text'] = (deficiency + curObject.UnitPriceSum).toString(); //합계 단가
                                    bidT3[key]['C21']['_text'] = (deficiency + curObject.Labor).toString(); //노무비
                                    bidT3[key]['C23']['_text'] = (deficiency + curObject.PriceSum).toString(); //합계
                                    break;
                                }
                            }
                            else {
                                if (curObject.ExpenseUnit !== 0) {
                                    if (Number(bidT3[key]['C19']['_text']) * 1.5 >
                                        curObject.ExpenseUnit + divisionPrice) {
                                        curObject.ExpenseUnit += divisionPrice;
                                        bidT3[key]['C18']['_text'] =
                                            curObject.ExpenseUnit.toString(); //경비 단가
                                        bidT3[key]['C19']['_text'] =
                                            curObject.UnitPriceSum.toString(); //합계 단가
                                        bidT3[key]['C22']['_text'] = curObject.Expense.toString(); //경비
                                        bidT3[key]['C23']['_text'] = curObject.PriceSum.toString(); //합계
                                        count++;
                                    }
                                    if (count === this.exCount) {
                                        //절사, 반올림에 따른 부족분 조정
                                        bidT3[key]['C18']['_text'] = (deficiency + curObject.ExpenseUnit).toString(); //경비 단가
                                        bidT3[key]['C19']['_text'] = (deficiency + curObject.UnitPriceSum).toString(); //합계 단가
                                        bidT3[key]['C22']['_text'] = (deficiency + curObject.Expense).toString(); //경비
                                        bidT3[key]['C23']['_text'] = (deficiency + curObject.PriceSum).toString(); //합계
                                        break;
                                    }
                                }
                                else {
                                    if (Number(bidT3[key]['C19']['_text']) * 1.5 >
                                        curObject.MaterialUnit + divisionPrice) {
                                        curObject.MaterialUnit += divisionPrice;
                                        bidT3[key]['C16']['_text'] =
                                            curObject.MaterialUnit.toString(); //재료비 단가
                                        bidT3[key]['C19']['_text'] =
                                            curObject.UnitPriceSum.toString(); //합계 단가
                                        bidT3[key]['C20']['_text'] = curObject.Material.toString(); //재료비
                                        bidT3[key]['C23']['_text'] = curObject.PriceSum.toString(); //합계
                                        count++;
                                    }
                                    if (count === this.exCount) {
                                        //절사, 반올림에 따른 부족분 조정
                                        bidT3[key]['C18']['_text'] = (deficiency + curObject.MaterialUnit).toString(); //재료비 단가
                                        bidT3[key]['C19']['_text'] = (deficiency + curObject.UnitPriceSum).toString(); //합계 단가
                                        bidT3[key]['C22']['_text'] = (deficiency + curObject.Material).toString(); //재료비
                                        bidT3[key]['C23']['_text'] = (deficiency + curObject.PriceSum).toString(); //합계
                                        break;
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
        if (keyFound === 1 && this.exSum < TempExPrice) {
            this.maxBID['C17']['_text'] = (Number(this.maxBID['C17']['_text']) +
                TempExPrice -
                this.exSum).toString();
            this.maxBID['C19']['_text'] = (Number(this.maxBID['C19']['_text']) +
                TempExPrice -
                this.exSum).toString();
            this.maxBID['C21']['_text'] = (Number(this.maxBID['C21']['_text']) +
                TempExPrice -
                this.exSum).toString();
            this.maxBID['C23']['_text'] = (Number(this.maxBID['C23']['_text']) +
                TempExPrice -
                this.exSum).toString();
            //소수부분 차이에 의한 99.7% 이하 위반 문제에 대한 처리 (노무비에 보정)
        }
    }
    static GetAdjustedExcludePrice() {
        //사정율 적용한 제요율적용제외 금액 저장
        const bidT3 = this.eleBID['T3'];
        let code;
        let type;
        for (let key in bidT3) {
            code = JSON.stringify(bidT3[key]['C9']['_text']);
            type = JSON.stringify(bidT3[key]['C5']['_text'])[1];
            if (code !== undefined && type === 'S') {
                let constNum = JSON.stringify(bidT3[key]['C1']['_text']).slice(1, -1); //세부공사 번호
                let numVal = JSON.stringify(bidT3[key]['C2']['_text']).slice(1, -1); //세부공종 번호
                let detailVal = JSON.stringify(bidT3[key]['C3']['_text']).slice(1, -1); //세부 공종 번호
                let curObject = Data_1.Data.Dic.get(constNum).find((x) => x.WorkNum === numVal && x.DetailWorkNum === detailVal);
                if (curObject.Item === '제요율적용제외') {
                    Data_1.Data.AdjustedExMaterial += Number(bidT3[key]['C20']['_text']);
                    Data_1.Data.AdjustedExLabor += Number(bidT3[key]['C21']['_text']);
                    Data_1.Data.AdjustedExExpense += Number(bidT3[key]['C22']['_text']);
                }
            }
        }
    }
    static SetPriceOfSuperConstruction() {
        //상위 공종의 각 단가 합 및 합계 세팅 (23.02.07)
        let firstConstruction = undefined; //가장 상위 공종
        let secondConstruction = undefined; //중간 상위 공종
        let thirdConstruction = undefined; //마지막 상위 공종
        const bidT3 = this.eleBID['T3'];
        let code;
        let type;
        for (let key in bidT3) {
            code = JSON.stringify(bidT3[key]['C9']['_text']);
            type = JSON.stringify(bidT3[key]['C5']['_text'])[1];
            if (type === 'G') {
                //공종이면
                if (bidT3[key]['C23']['_text'] === '0') {
                    //이미 합계가 세팅되어 있는지 확인 (중복 계산을 막기 위함)
                    if (firstConstruction === undefined || bidT3[key]['C3']['_text'] === '0') {
                        //C3이 0이면 가장 상위 공종
                        firstConstruction = bidT3[key]; //현재 보고있는 object가 가장 상위 공종
                        secondConstruction = undefined; //중간 상위 공종 초기화
                        thirdConstruction = undefined; //마지막 상위 공종 초기화
                    }
                    else if (bidT3[key]['C3']['_text'] === firstConstruction['C2']['_text'] &&
                        firstConstruction !== undefined) {
                        //C3이 가장 상위 공종의 C2와 같다면 중간 상위 공종
                        secondConstruction = bidT3[key]; //현재 보고있는 object가 중간 상위 공종
                        thirdConstruction = undefined; //마지막 상위 공종 초기화
                    }
                    else if (bidT3[key]['C3']['_text'] === secondConstruction['C2']['_text'] &&
                        secondConstruction !== undefined)
                        // C3이 중간 상위 공종의 C2와 같다면 마지막 상위 공종
                        thirdConstruction = bidT3[key]; //현재 보고있는 object가 마지막 상위 공종
                }
                else {
                    //공종에 합계가 이미 세팅되어 있다면 전부 초기화
                    firstConstruction = undefined;
                    secondConstruction = undefined;
                    thirdConstruction = undefined;
                }
            }
            else if (code !== undefined && type === 'S') {
                //공종이 아니면
                if (firstConstruction !== undefined) {
                    //현재 보는 object가 가장 상위 공종에 포함되어 있다면 단가별 합과 합계를 더해나감
                    firstConstruction['C20']['_text'] = (Number(firstConstruction['C20']['_text']) +
                        Number(bidT3[key]['C20']['_text'])).toString(); //재료비
                    firstConstruction['C21']['_text'] = (Number(firstConstruction['C21']['_text']) +
                        Number(bidT3[key]['C21']['_text'])).toString(); //노무비
                    firstConstruction['C22']['_text'] = (Number(firstConstruction['C22']['_text']) +
                        Number(bidT3[key]['C22']['_text'])).toString(); //경비
                    firstConstruction['C23']['_text'] = (Number(firstConstruction['C23']['_text']) +
                        Number(bidT3[key]['C23']['_text'])).toString(); //합계
                }
                if (secondConstruction !== undefined) {
                    //현재 보는 object가 중간 상위 공종에 포함되어 있다면 단가별 합과 합계를 더해나감
                    secondConstruction['C20']['_text'] = (Number(secondConstruction['C20']['_text']) +
                        Number(bidT3[key]['C20']['_text'])).toString(); //재료비
                    secondConstruction['C21']['_text'] = (Number(secondConstruction['C21']['_text']) +
                        Number(bidT3[key]['C21']['_text'])).toString(); //노무비
                    secondConstruction['C22']['_text'] = (Number(secondConstruction['C22']['_text']) +
                        Number(bidT3[key]['C22']['_text'])).toString(); //경비
                    secondConstruction['C23']['_text'] = (Number(secondConstruction['C23']['_text']) +
                        Number(bidT3[key]['C23']['_text'])).toString(); //합계
                }
                if (thirdConstruction !== undefined) {
                    //현재 보는 object가 마지막 상위 공종에 포함되어 있다면 단가별 합과 합계를 더해나감
                    thirdConstruction['C20']['_text'] = (Number(thirdConstruction['C20']['_text']) +
                        Number(bidT3[key]['C20']['_text'])).toString(); //재료비
                    thirdConstruction['C21']['_text'] = (Number(thirdConstruction['C21']['_text']) +
                        Number(bidT3[key]['C21']['_text'])).toString(); //노무비
                    thirdConstruction['C22']['_text'] = (Number(thirdConstruction['C22']['_text']) +
                        Number(bidT3[key]['C22']['_text'])).toString(); //경비
                    thirdConstruction['C23']['_text'] = (Number(thirdConstruction['C23']['_text']) +
                        Number(bidT3[key]['C23']['_text'])).toString(); //합계
                }
            }
        }
    }
    static SetBusinessInfo() {
        const bidT1 = this.eleBID['T1'];
        bidT1['C17']['_text'] = Data_1.Data.CompanyRegistrationNum;
        bidT1['C18']['_text'] = Data_1.Data.CompanyRegistrationName;
    }
    static SubstitutePrice() {
        //BID 파일 내 원가계산서 관련 금액 세팅
        const bidT5 = this.eleBID['T5']; //bid.Name이 T5인지를 확인함으로 간단하게 원가 계산서부분의 element 인지를 판별. Tag는 T3가 아닌 T5 기준을 따른다. (23.01.31 수정)
        for (let key in bidT5) {
            if (bidT5[key]['C4']['_text'] !== '이윤' &&
                Data_1.Data.Bidding.get(bidT5[key]['C4']['_text'])) {
                bidT5[key]['C8']['_text'] = Data_1.Data.Bidding.get(bidT5[key]['C4']['_text']).toString();
            }
            else if (Data_1.Data.Rate1[bidT5[key]['C4']['_text']]) {
                bidT5[key]['C8']['_text'] = Data_1.Data.Bidding.get(bidT5[key]['C4']['_text']).toString();
            }
        }
        // 단순하게 JSON 포멧인 docBID를 string으로 변환하고 OutputDataFromBID.json 파일에 덮어쓰기
        fs.writeFileSync(Data_1.Data.folder + '\\OutputDataFromBID.json', JSON.stringify(this.docBID));
        //fs.writeFileSync(Data.folder + '\\OutputDataFromBID.json', JSON.stringify(this.docBID))
    }
    static CreateZipFile(xlsfiles) {
        if (fs.existsSync(Data_1.Data.folder + '\\WORK DIRCETORY\\입찰내역.zip')) {
            //기존 입찰내역.zip 파일은 삭제한다. (23.02.02)
            fs.rmSync(Data_1.Data.folder + '\\WORK DIRCETORY\\입찰내역.zip');
        }
        if (fs.existsSync(Data_1.Data.folder + '\\WORK DIRCETORY\\입찰내역.zip')) {
            //기존 입찰내역.zip 파일은 삭제한다. (23.02.02)
            fs.rmSync(Data_1.Data.folder + '\\WORK DIRCETORY\\입찰내역.zip');
        }
        const zip = new AdmZip();
        for (let idx in xlsfiles) {
            zip.addLocalFile(Data_1.Data.folder + '\\' + xlsfiles[idx]); //만들어진 xls파일들을 압축
            zip.addLocalFile(Data_1.Data.folder + '\\' + xlsfiles[idx]); //만들어진 xls파일들을 압축
        }
        zip.writeZip(Data_1.Data.folder + '\\WORK DIRECTORY\\입찰내역.zip');
        //zip.writeZip(Data.folder + '\\WORK DIRECTORY\\입찰내역.zip')
    }
    static CreateFile() {
        //최종 입찰내역 파일 세부공사별로 생성
        CreateResultFile_1.CreateResultFile.Create();
        //생성된 입찰내역 파일 압축
        const files = fs.readdirSync(Data_1.Data.folder);
        const xlsFiles = files.filter((file) => file.substring(file.length - 4, file.length).toLowerCase() === '.xls');
        this.CreateZipFile(xlsFiles);
    }
}
CalculatePrice.exSum = 0; //단가조정된 제요율적용제외공종 항목 합계
CalculatePrice.exCount = 0; //제요율적용제외공종 항목 수량 합계
exports.CalculatePrice = CalculatePrice;
