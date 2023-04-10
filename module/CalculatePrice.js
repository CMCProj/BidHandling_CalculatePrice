"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.CalculatePrice = void 0;
var Data_1 = require("./Data");
var fillCostAccount_1 = require("./fillCostAccount");
var CreateResultFile_1 = require("./CreateResultFile");
var fs_1 = __importDefault(require("fs"));
var adm_zip_1 = __importDefault(require("adm-zip"));
var big_js_1 = __importDefault(require("big.js"));
//실행 위해 프로그램 내 폴더로 경로 변경
var CalculatePrice = exports.CalculatePrice = /** @class */ (function () {
    function CalculatePrice() {
    }
    CalculatePrice.Calculation = function () {
        var bidString = fs_1.default.readFileSync(Data_1.Data.folder + '\\OutputDataFromBID.json', 'utf-8');
        this.docBID = JSON.parse(bidString);
        this.eleBID = this.docBID['data'];
        this.Reset();
        if (Data_1.Data.StandardMarketDeduction.localeCompare('1'))
            this.ApplyStandardPriceOption();
        this.GetFixedPriceRate();
        this.FindMyPercent();
        this.GetWeight();
        this.CalculateRate(Data_1.Data.PersonalRate, Data_1.Data.BalancedRate);
        this.Recalculation();
        if (this.exCount !== 0) {
            this.SetExcludingPrice();
            this.GetAdjustedExcludePrice();
        }
        this.SetPriceOfSuperConstruction();
        fillCostAccount_1.FillCostAccount.CalculateBiddingCosts();
        this.SetBusinessInfo();
        this.SubstitutePrice();
        this.CreateFile();
    };
    CalculatePrice.Reset = function () {
        Data_1.Data.ExecuteReset = '1';
        var DM = Data_1.Data.Investigation['직접재료비'];
        var DL = Data_1.Data.Investigation['직접노무비'];
        var OE = Data_1.Data.Investigation['산출경비'];
        var FM = Data_1.Data.InvestigateFixedPriceDirectMaterial;
        var FL = Data_1.Data.InvestigateFixedPriceDirectLabor;
        var FOE = Data_1.Data.InvestigateFixedPriceOutputExpense;
        var SM = Data_1.Data.InvestigateStandardMaterial;
        var SL = Data_1.Data.InvestigateStandardLabor;
        var SOE = Data_1.Data.InvestigateStandardExpense;
        Data_1.Data.RealDirectMaterial = DM;
        Data_1.Data.RealDirectLabor = DL;
        Data_1.Data.RealOutputExpense = OE;
        Data_1.Data.FixedPriceDirectMaterial = FM;
        Data_1.Data.FixedPriceDirectLabor = FL;
        Data_1.Data.FixedPriceOutputExpense = FOE;
        Data_1.Data.StandardMaterial = SM;
        Data_1.Data.StandardLabor = SL;
        Data_1.Data.StandardExpense = SOE;
        var bidT3 = this.eleBID['T3'];
        var code;
        var type;
        var _loop_1 = function (key) {
            code = JSON.stringify(bidT3[key]['C9']['_text']);
            type = JSON.stringify(bidT3[key]['C5']['_text'])[1];
            if (code !== undefined && type === 'S') {
                var constNum = JSON.stringify(bidT3[key]['C1']['_text']).slice(1, -1);
                var numVal_1 = JSON.stringify(bidT3[key]['C2']['_text']).slice(1, -1);
                var detailVal_1 = JSON.stringify(bidT3[key]['C3']['_text']).slice(1, -1);
                var curObject = Data_1.Data.Dic.get(constNum).find(function (x) { return x.WorkNum === numVal_1 && x.DetailWorkNum === detailVal_1; });
                curObject.MaterialUnit = Number(bidT3[key]['C16']['_text'].slice(1, -1));
                curObject.LaborUnit = Number(bidT3[key]['C17']['_text'].slice(1, -1));
                curObject.ExpenseUnit = Number(bidT3[key]['C18']['_text'].slice(1, -1));
            }
        };
        for (var key in bidT3) {
            _loop_1(key);
        }
        Data_1.Data.ExecuteReset = '0';
    };
    CalculatePrice.ApplyStandardPriceOption = function () {
        var bidT3 = this.eleBID['T3'];
        var code;
        var type;
        var _loop_2 = function (key) {
            code = JSON.stringify(bidT3[key]['C9']['_text']);
            type = JSON.stringify(bidT3[key]['C5']['_text'])[1];
            if (code !== undefined && type === 'S') {
                var constNum = JSON.stringify(bidT3[key]['C1']['_text']).slice(1, -1);
                var numVal_2 = JSON.stringify(bidT3[key]['C2']['_text']).slice(1, -1);
                var detailVal_2 = JSON.stringify(bidT3[key]['C3']['_text']).slice(1, -1);
                var curObject = Data_1.Data.Dic.get(constNum).find(function (x) { return x.WorkNum === numVal_2 && x.DetailWorkNum === detailVal_2; });
                if (curObject.Item.localeCompare('표준시장단가')) {
                    Data_1.Data.RealDirectMaterial -= +JSON.stringify(bidT3[key]['C20']['_text']).slice(1, -1);
                    Data_1.Data.RealDirectLabor -= +JSON.stringify(bidT3[key]['C21']['_text']).slice(1, -1);
                    Data_1.Data.RealOutputExpense -= +JSON.stringify(bidT3[key]['C22']['_text']).slice(1, -1);
                    Data_1.Data.FixedPriceDirectMaterial -= +JSON.stringify(bidT3[key]['C20']['_text']).slice(1, -1);
                    Data_1.Data.FixedPriceDirectLabor -= +JSON.stringify(bidT3[key]['C21']['_text']).slice(1, -1);
                    Data_1.Data.FixedPriceOutputExpense -= +JSON.stringify(bidT3[key]['C22']['_text']).slice(1, -1);
                    Data_1.Data.StandardMaterial -= +JSON.stringify(bidT3[key]['C20']['_text']).slice(1, -1);
                    Data_1.Data.StandardLabor -= +JSON.stringify(bidT3[key]['C21']['_text']).slice(1, -1);
                    Data_1.Data.StandardExpense -= +JSON.stringify(bidT3[key]['C22']['_text']).slice(1, -1);
                    if (curObject.MaterialUnit !== 0)
                        curObject.MaterialUnit =
                            Math.trunc(curObject.MaterialUnit * 0.997 * 10) / 10 + 1.0;
                    if (curObject.LaborUnit !== 0)
                        curObject.LaborUnit =
                            Math.trunc(curObject.LaborUnit * 0.997 * 10) / 10 + 0.1;
                    if (curObject.ExpenseUnit !== 0)
                        curObject.ExpenseUnit =
                            Math.trunc(curObject.ExpenseUnit * 0.997 * 10) / 10 + 0.1;
                    bidT3[key]['C16']['_text'] = curObject.MaterialUnit.toString();
                    bidT3[key]['C17']['_text'] = curObject.LaborUnit.toString();
                    bidT3[key]['C18']['_text'] = curObject.ExpenseUnit.toString();
                    bidT3[key]['C19']['_text'] = curObject.UnitPriceSum.toString();
                    bidT3[key]['C20']['_text'] = curObject.Material.toString();
                    bidT3[key]['C21']['_text'] = curObject.Labor.toString();
                    bidT3[key]['C22']['_text'] = curObject.Expense.toString();
                    bidT3[key]['C23']['_text'] = curObject.PriceSum.toString();
                    Data_1.Data.RealDirectMaterial += +JSON.stringify(bidT3[key]['C20']['_text']).slice(1, -1);
                    Data_1.Data.RealDirectLabor += +JSON.stringify(bidT3[key]['C21']['_text']).slice(1, -1);
                    Data_1.Data.RealOutputExpense += +JSON.stringify(bidT3[key]['C22']['_text']).slice(1, -1);
                    Data_1.Data.FixedPriceDirectMaterial += +JSON.stringify(bidT3[key]['C20']['_text']).slice(1, -1);
                    Data_1.Data.FixedPriceDirectLabor += +JSON.stringify(bidT3[key]['C21']['_text']).slice(1, -1);
                    Data_1.Data.FixedPriceOutputExpense += +JSON.stringify(bidT3[key]['C22']['_text']).slice(1, -1);
                    Data_1.Data.StandardMaterial += +JSON.stringify(bidT3[key]['C20']['_text']).slice(1, -1);
                    Data_1.Data.StandardLabor += +JSON.stringify(bidT3[key]['C21']['_text']).slice(1, -1);
                    Data_1.Data.StandardExpense += +JSON.stringify(bidT3[key]['C22']['_text']).slice(1, -1);
                }
            }
        };
        for (var key in bidT3) {
            _loop_2(key);
        }
    };
    CalculatePrice.GetFixedPriceRate = function () {
        var directConstPrice = Data_1.Data.Investigation['직공비'];
        var fixCostSum = Data_1.Data.InvestigateFixedPriceDirectMaterial +
            Data_1.Data.InvestigateFixedPriceDirectLabor +
            Data_1.Data.InvestigateFixedPriceOutputExpense;
        Data_1.Data.FixedPricePercent = Math.trunc((fixCostSum / directConstPrice) * 100 * 10000) / 10000;
    };
    CalculatePrice.FindMyPercent = function () {
        if (Data_1.Data.FixedPricePercent < 20.0)
            this.myPercent = 0.85;
        else if (Data_1.Data.FixedPricePercent < 25.0)
            this.myPercent = 0.84;
        else if (Data_1.Data.FixedPricePercent < 30.0)
            this.myPercent = 0.83;
        else
            this.myPercent = 0.82;
    };
    CalculatePrice.GetWeight = function () {
        var varCostSum = Data_1.Data.RealPriceDirectMaterial + Data_1.Data.RealPriceDirectLabor + Data_1.Data.RealPriceOutputExpense;
        var weight;
        var maxWeight = 0;
        var weightSum = 0;
        //Data 인스턴스 생성시 생성자 함수로 프로퍼티 값을 설정하는 것으로 변경
        //->max를 빈 변수로 선언한 뒤 후에 Data인스턴스를 넣도록 수정
        var max;
        Data_1.Data.Dic.forEach(function (value, _) {
            for (var idx in value) {
                if (value[idx].Item.localeCompare('일반')) {
                    var material = value[idx].Material;
                    var labor = value[idx].Labor;
                    var expense = value[idx].Expense;
                    weight =
                        Math.round(((material + labor + expense) / varCostSum) * 1000000) / 1000000;
                    weightSum += weight;
                    if (maxWeight < weight) {
                        maxWeight = weight;
                        max = value[idx];
                    }
                    value[idx].Weight = weight;
                }
            }
        });
        if (weightSum !== 1.0) {
            var lack = 1.0 - weightSum;
            max.Weight += lack;
        }
    };
    CalculatePrice.CalculateRate = function (presonalRate, balancedRate) {
        var unitPrice = (0, big_js_1.default)(100);
        presonalRate = (0, big_js_1.default)(presonalRate);
        balancedRate = (0, big_js_1.default)(balancedRate);
        this.balancedUnitPriceRate =
            (0.9 * unitPrice * (1.0 + balancedRate / 100) * this.myPercent) /
                (1.0 - 0.1 * this.myPercent) /
                100;
        this.targetRate =
            ((unitPrice * (1.0 + presonalRate / 100) * 0.9 +
                unitPrice * this.balancedUnitPriceRate * 0.1) *
                this.myPercent) /
                100;
        this.targetRate = Math.trunc(this.targetRate * 1000000) / 1000000;
        console.log(this.targetRate);
    };
    CalculatePrice.RoundOrTruncate = function (Rate, Object, refMyMaterialUnit, refMyLaborUnit, refMyExpenseUnit) {
        if (Data_1.Data.UnitPriceTrimming.localeCompare('1')) {
            refMyMaterialUnit.value = Math.trunc(Object.MaterialUnit * Rate * 10) / 10;
            refMyLaborUnit.value = Math.trunc(Object.LaborUnit * Rate * 10) / 10;
            refMyExpenseUnit.value = Math.trunc(Object.ExpenseUnit * Rate * 10) / 10;
        }
        else if (Data_1.Data.UnitPriceTrimming.localeCompare('2')) {
            refMyMaterialUnit.value = Math.ceil(Object.MaterialUnit * Rate);
            refMyLaborUnit.value = Math.ceil(Object.LaborUnit * Rate);
            refMyExpenseUnit.value = Math.ceil(Object.ExpenseUnit * Rate);
        }
    };
    CalculatePrice.CheckLaborLimit80 = function (Object, refMyMaterialUnit, refMyLaborUnit, refMyExpenseUnit) {
        if (Object.LaborUnit * 0.8 > refMyLaborUnit.value) {
            var deficiency = Object.LaborUnit * 0.8 - refMyLaborUnit.value;
            if (refMyLaborUnit.value !== 0)
                refMyMaterialUnit.value -= deficiency;
            else if (refMyExpenseUnit.value !== 0)
                refMyExpenseUnit.value -= deficiency;
            refMyLaborUnit.value = Object.LaborUnit * 0.8;
        }
    };
    CalculatePrice.Recalculation = function () {
        var bidT3 = this.eleBID['T3'];
        var code;
        var type;
        this.exCount = 0;
        this.exSum = 0;
        var _loop_3 = function (key) {
            code = JSON.stringify(bidT3[key]['C9']['_text']);
            type = JSON.stringify(bidT3[key]['C5']['_text'])[1];
            if (code !== undefined && type === 'S') {
                var constNum = JSON.stringify(bidT3[key]['C1']['_text']).slice(1, -1);
                var numVal_3 = JSON.stringify(bidT3[key]['C2']['_text']).slice(1, -1);
                var detailVal_3 = JSON.stringify(bidT3[key]['C3']['_text']).slice(1, -1);
                var curObject = Data_1.Data.Dic.get(constNum).find(function (x) { return x.WorkNum === numVal_3 && x.DetailWorkNum === detailVal_3; });
                if (curObject.Item.localeCompare('일반')) {
                    Data_1.Data.RealDirectMaterial -= +JSON.stringify(bidT3[key]['C20']['_text']).slice(1, -1);
                    Data_1.Data.RealDirectLabor -= +JSON.stringify(bidT3[key]['C21']['_text']).slice(1, -1);
                    Data_1.Data.RealOutputExpense -= +JSON.stringify(bidT3[key]['C22']['_text']).slice(1, -1);
                    var targetPrice = (curObject.MaterialUnit + curObject.LaborUnit + curObject.ExpenseUnit) *
                        this_1.targetRate;
                    var myMaterialUnit = { value: 0 };
                    var myLaborUnit = { value: 0 };
                    var myExpenseUnit = { value: 0 };
                    var myPrice = void 0;
                    if (Data_1.Data.ZeroWeightDeduction.localeCompare('1')) {
                        if (curObject.Weight === 0 && curObject.LaborUnit === 0) {
                            curObject.MaterialUnit = Math.ceil(curObject.MaterialUnit * 0.5);
                            curObject.ExpenseUnit = Math.ceil(curObject.ExpenseUnit * 0.5);
                            bidT3[key]['C16']['_text'] = curObject.MaterialUnit.toString();
                            bidT3[key]['C17']['_text'] = curObject.LaborUnit.toString();
                            bidT3[key]['C18']['_text'] = curObject.ExpenseUnit.toString();
                            bidT3[key]['C19']['_text'] = curObject.UnitPriceSum.toString();
                            bidT3[key]['C20']['_text'] = curObject.Material.toString();
                            bidT3[key]['C21']['_text'] = curObject.Labor.toString();
                            bidT3[key]['C22']['_text'] = curObject.Expense.toString();
                            bidT3[key]['C23']['_text'] = curObject.PriceSum.toString();
                            Data_1.Data.RealDirectMaterial += +JSON.stringify(bidT3[key]['C20']['_text']).slice(1, -1);
                            Data_1.Data.RealDirectLabor += +JSON.stringify(bidT3[key]['C21']['_text']).slice(1, -1);
                            Data_1.Data.RealOutputExpense += +JSON.stringify(bidT3[key]['C22']['_text']).slice(1, -1);
                            return "continue";
                        }
                        else {
                            this_1.RoundOrTruncate(this_1.targetRate, curObject, myMaterialUnit, myLaborUnit, myExpenseUnit);
                            this_1.CheckLaborLimit80(curObject, myMaterialUnit, myLaborUnit, myExpenseUnit);
                        }
                    }
                    else if (Data_1.Data.ZeroWeightDeduction.localeCompare('2')) {
                        this_1.RoundOrTruncate(this_1.targetRate, curObject, myMaterialUnit, myLaborUnit, myExpenseUnit);
                        this_1.CheckLaborLimit80(curObject, myMaterialUnit, myLaborUnit, myExpenseUnit);
                    }
                    myPrice = myMaterialUnit.value + myLaborUnit.value + myExpenseUnit.value;
                    if (Data_1.Data.LaborCostLowBound.localeCompare('1')) {
                        var Excess = myPrice - targetPrice;
                        var laborExcess = myLaborUnit.value - curObject.LaborUnit * 0.8;
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
                    bidT3[key]['C16']['_text'] = curObject.MaterialUnit.toString();
                    bidT3[key]['C17']['_text'] = curObject.LaborUnit.toString();
                    bidT3[key]['C18']['_text'] = curObject.ExpenseUnit.toString();
                    bidT3[key]['C19']['_text'] = curObject.UnitPriceSum.toString();
                    bidT3[key]['C20']['_text'] = curObject.Material.toString();
                    bidT3[key]['C21']['_text'] = curObject.Labor.toString();
                    bidT3[key]['C22']['_text'] = curObject.Expense.toString();
                    bidT3[key]['C23']['_text'] = curObject.PriceSum.toString();
                    Data_1.Data.RealDirectMaterial += +JSON.stringify(bidT3[key]['C20']['_text']).slice(1, -1);
                    Data_1.Data.RealDirectLabor += +JSON.stringify(bidT3[key]['C21']['_text']).slice(1, -1);
                    Data_1.Data.RealOutputExpense += +JSON.stringify(bidT3[key]['C22']['_text']).slice(1, -1);
                }
                else if (curObject.Item === '제요율적용제외') {
                    curObject.MaterialUnit =
                        Math.trunc(curObject.MaterialUnit * this_1.targetRate * 10) / 10;
                    curObject.LaborUnit =
                        Math.trunc(curObject.LaborUnit * this_1.targetRate * 10) / 10;
                    curObject.ExpenseUnit =
                        Math.trunc(curObject.ExpenseUnit * this_1.targetRate * 10) / 10;
                    this_1.exSum += curObject.PriceSum;
                    this_1.exCount++;
                }
            }
        };
        var this_1 = this;
        for (var key in bidT3) {
            _loop_3(key);
        }
    };
    CalculatePrice.SetExcludingPrice = function () {
        var TempInvestDirectSum = Data_1.Data.Investigation['직공비'];
        var TempRealDirectSum = fillCostAccount_1.FillCostAccount.ToLong(Data_1.Data.RealDirectMaterial + Data_1.Data.RealDirectLabor + Data_1.Data.RealOutputExpense);
        var InvestExSum = Data_1.Data.ExcludingMaterial + Data_1.Data.ExcludingLabor + Data_1.Data.ExcludingExpense;
        var TempExRate = Math.round((InvestExSum / TempInvestDirectSum) * 100000) / 100000;
        var TempExPrice = Math.ceil(TempRealDirectSum * TempExRate);
        var keyFound = 0;
        var bidT3 = this.eleBID['T3'];
        var code;
        var type;
        if (Data_1.Data.CostAccountDeduction.localeCompare('1')) {
            TempExPrice = Math.ceil(Math.ceil(TempExPrice * 0.997));
        }
        var _loop_4 = function (key) {
            code = JSON.stringify(bidT3[key]['C9']['_text']);
            type = JSON.stringify(bidT3[key]['C5']['_text'])[1];
            if (code !== undefined && type === 'S') {
                var constNum = JSON.stringify(bidT3[key]['C1']['_text']).slice(1, -1);
                var numVal_4 = JSON.stringify(bidT3[key]['C2']['_text']).slice(1, -1);
                var detailVal_4 = JSON.stringify(bidT3[key]['C3']['_text']).slice(1, -1);
                var curObject = Data_1.Data.Dic.get(constNum).find(function (x) { return x.WorkNum === numVal_4 && x.DetailWorkNum === detailVal_4; });
                if (curObject.Item.localeCompare('제요율적용제외')) {
                    if (this_2.maxBID === undefined) {
                        this_2.maxBID = bidT3[key];
                    }
                    if (bidT3[key]['C15']['_text'].slice(1, -1) === '1' &&
                        bidT3[key]['C15']['_text'].slice(1, -1) !== '0') {
                        if (Number(bidT3[key]['C19']['_text'].slice(1, -1)) >
                            Number(this_2.maxBID['C19']['_text'].slice(1, -1))) {
                            if (Number(bidT3[key]['C19']['_text'].slice(1, -1)) * 1.5 >
                                curObject.PriceSum + (TempExPrice - this_2.exSum)) {
                                keyFound = 1;
                                this_2.maxBID = bidT3[key];
                            }
                        }
                    }
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
        var this_2 = this;
        for (var key in bidT3) {
            _loop_4(key);
        }
        if (keyFound === 0) {
            var divisionPrice = Math.trunc((TempExPrice - this.exSum) / this.exCount);
            var deficiency = Math.ceil(TempExPrice - this.exSum - divisionPrice * this.exCount);
            var count = 0;
            while (count !== this.exCount) {
                var bidT3_1 = this.eleBID['T3'];
                var code_1 = void 0;
                var type_1 = void 0;
                var _loop_5 = function (key) {
                    code_1 = JSON.stringify(bidT3_1[key]['C9']['_text']);
                    type_1 = JSON.stringify(bidT3_1[key]['C5']['_text'])[1];
                    if (code_1 !== undefined && type_1 === 'S') {
                        var constNum = JSON.stringify(bidT3_1[key]['C1']['_text']).slice(1, -1);
                        var numVal_5 = JSON.stringify(bidT3_1[key]['C2']['_text']).slice(1, -1);
                        var detailVal_5 = JSON.stringify(bidT3_1[key]['C3']['_text']).slice(1, -1);
                        var curObject = Data_1.Data.Dic.get(constNum).find(function (x) { return x.WorkNum === numVal_5 && x.DetailWorkNum === detailVal_5; });
                        if (curObject.Item.localeCompare('제요율적용제외') &&
                            curObject.Quantity === 1) {
                            if (curObject.LaborUnit !== 0) {
                                if (Number(bidT3_1[key]['C19']['_text'].slice(1, -1)) * 1.5 >
                                    curObject.LaborUnit + divisionPrice) {
                                    curObject.LaborUnit += divisionPrice;
                                    bidT3_1[key]['C17']['_text'] = curObject.LaborUnit.toString();
                                    bidT3_1[key]['C19']['_text'] = curObject.UnitPriceSum.toString();
                                    bidT3_1[key]['C21']['_text'] = curObject.Labor.toString();
                                    bidT3_1[key]['C23']['_text'] = curObject.PriceSum.toString();
                                    count++;
                                }
                                if (count === this_3.exCount) {
                                    bidT3_1[key]['C17']['_text'] = (deficiency + curObject.LaborUnit).toString();
                                    bidT3_1[key]['C19']['_text'] = (deficiency + curObject.UnitPriceSum).toString();
                                    bidT3_1[key]['C21']['_text'] = (deficiency + curObject.Labor).toString();
                                    bidT3_1[key]['C23']['_text'] = (deficiency + curObject.PriceSum).toString();
                                    return "break";
                                }
                            }
                            else {
                                if (curObject.ExpenseUnit !== 0) {
                                    if (Number(bidT3_1[key]['C19']['_text'].slice(1, -1)) * 1.5 >
                                        curObject.ExpenseUnit + divisionPrice) {
                                        curObject.ExpenseUnit += divisionPrice;
                                        bidT3_1[key]['C18']['_text'] =
                                            curObject.ExpenseUnit.toString();
                                        bidT3_1[key]['C19']['_text'] =
                                            curObject.UnitPriceSum.toString();
                                        bidT3_1[key]['C22']['_text'] = curObject.Expense.toString();
                                        bidT3_1[key]['C23']['_text'] = curObject.PriceSum.toString();
                                        count++;
                                    }
                                    if (count === this_3.exCount) {
                                        bidT3_1[key]['C18']['_text'] = (deficiency + curObject.ExpenseUnit).toString();
                                        bidT3_1[key]['C19']['_text'] = (deficiency + curObject.UnitPriceSum).toString();
                                        bidT3_1[key]['C22']['_text'] = (deficiency + curObject.Expense).toString();
                                        bidT3_1[key]['C23']['_text'] = (deficiency + curObject.PriceSum).toString();
                                        return "break";
                                    }
                                }
                                else {
                                    if (Number(bidT3_1[key]['C19']['_text'].slice(1, -1)) * 1.5 >
                                        curObject.MaterialUnit + divisionPrice) {
                                        curObject.MaterialUnit += divisionPrice;
                                        bidT3_1[key]['C16']['_text'] =
                                            curObject.MaterialUnit.toString();
                                        bidT3_1[key]['C19']['_text'] =
                                            curObject.UnitPriceSum.toString();
                                        bidT3_1[key]['C20']['_text'] = curObject.Material.toString();
                                        bidT3_1[key]['C23']['_text'] = curObject.PriceSum.toString();
                                        count++;
                                    }
                                    if (count === this_3.exCount) {
                                        bidT3_1[key]['C18']['_text'] = (deficiency + curObject.MaterialUnit).toString();
                                        bidT3_1[key]['C19']['_text'] = (deficiency + curObject.UnitPriceSum).toString();
                                        bidT3_1[key]['C22']['_text'] = (deficiency + curObject.Material).toString();
                                        bidT3_1[key]['C23']['_text'] = (deficiency + curObject.PriceSum).toString();
                                        return "break";
                                    }
                                }
                            }
                        }
                    }
                };
                var this_3 = this;
                for (var key in bidT3_1) {
                    var state_1 = _loop_5(key);
                    if (state_1 === "break")
                        break;
                }
            }
        }
        if (keyFound === 1 && this.exSum < TempExPrice) {
            this.maxBID['C17']['_text'] = (Number(this.maxBID['C17']['_text'].slice(1, -1)) +
                TempExPrice -
                this.exSum).toString();
            this.maxBID['C19']['_text'] = (Number(this.maxBID['C19']['_text'].slice(1, -1)) +
                TempExPrice -
                this.exSum).toString();
            this.maxBID['C21']['_text'] = (Number(this.maxBID['C21']['_text'].slice(1, -1)) +
                TempExPrice -
                this.exSum).toString();
            this.maxBID['C23']['_text'] = (Number(this.maxBID['C23']['_text'].slice(1, -1)) +
                TempExPrice -
                this.exSum).toString();
        }
    };
    CalculatePrice.GetAdjustedExcludePrice = function () {
        var bidT3 = this.eleBID['T3'];
        var code;
        var type;
        var _loop_6 = function (key) {
            code = JSON.stringify(bidT3[key]['C9']['_text']);
            type = JSON.stringify(bidT3[key]['C5']['_text'])[1];
            if (code !== undefined && type === 'S') {
                var constNum = JSON.stringify(bidT3[key]['C1']['_text']).slice(1, -1);
                var numVal_6 = JSON.stringify(bidT3[key]['C2']['_text']).slice(1, -1);
                var detailVal_6 = JSON.stringify(bidT3[key]['C3']['_text']).slice(1, -1);
                var curObject = Data_1.Data.Dic.get(constNum).find(function (x) { return x.WorkNum === numVal_6 && x.DetailWorkNum === detailVal_6; });
                if (curObject.Item.localeCompare('제요율적용제외')) {
                    Data_1.Data.AdjustedExMaterial += Number(bidT3[key]['C20']['_text'].slice(1, -1));
                    Data_1.Data.AdjustedExLabor += Number(bidT3[key]['C21']['_text'].slice(1, -1));
                    Data_1.Data.AdjustedExExpense += Number(bidT3[key]['C22']['_text'].slice(1, -1));
                }
            }
        };
        for (var key in bidT3) {
            _loop_6(key);
        }
    };
    CalculatePrice.SetPriceOfSuperConstruction = function () {
        var firstConstruction = undefined;
        var secondConstruction = undefined;
        var thirdConstruction = undefined;
        var bidT3 = this.eleBID['T3'];
        var code;
        var type;
        for (var key in bidT3) {
            code = JSON.stringify(bidT3[key]['C9']['_text']);
            type = JSON.stringify(bidT3[key]['C5']['_text'])[1];
            if (code !== undefined && type === 'G') {
                if (bidT3[key]['C23']['_text'].slice(1, -1) === '0') {
                    if (firstConstruction === undefined ||
                        bidT3[key]['C3']['_text'].slice(1, -1) === '0') {
                        firstConstruction = bidT3[key];
                        secondConstruction = undefined;
                        thirdConstruction = undefined;
                    }
                    else if (bidT3[key]['C3']['_text'].slice(1, -1) ===
                        firstConstruction['C2']['_text'].slice(1, -1) &&
                        firstConstruction !== undefined) {
                        secondConstruction = bidT3[key];
                        thirdConstruction = undefined;
                    }
                    else if (bidT3[key]['C3']['_text'].slice(1, -1) ===
                        secondConstruction['C2']['_text'].slice(1, -1) &&
                        secondConstruction !== undefined)
                        thirdConstruction = bidT3[key];
                }
                else {
                    firstConstruction = undefined;
                    secondConstruction = undefined;
                    thirdConstruction = undefined;
                }
            }
            else if (code !== undefined && type === 'S') {
                if (firstConstruction !== undefined) {
                    firstConstruction['C20']['_text'] = (Number(firstConstruction['C20']['_text'].slice(1, -1)) +
                        Number(bidT3[key]['C20']['_text'].slice(1, -1))).toString();
                    firstConstruction['C21']['_text'] = (Number(firstConstruction['C21']['_text'].slice(1, -1)) +
                        Number(bidT3[key]['C21']['_text'].slice(1, -1))).toString();
                    firstConstruction['C22']['_text'] = (Number(firstConstruction['C22']['_text'].slice(1, -1)) +
                        Number(bidT3[key]['C22']['_text'].slice(1, -1))).toString();
                    firstConstruction['C23']['_text'] = (Number(firstConstruction['C23']['_text'].slice(1, -1)) +
                        Number(bidT3[key]['C23']['_text'].slice(1, -1))).toString();
                }
                if (secondConstruction !== undefined) {
                    secondConstruction['C20']['_text'] = (Number(secondConstruction['C20']['_text'].slice(1, -1)) +
                        Number(bidT3[key]['C20']['_text'].slice(1, -1))).toString();
                    secondConstruction['C21']['_text'] = (Number(secondConstruction['C21']['_text'].slice(1, -1)) +
                        Number(bidT3[key]['C21']['_text'].slice(1, -1))).toString();
                    secondConstruction['C22']['_text'] = (Number(secondConstruction['C22']['_text'].slice(1, -1)) +
                        Number(bidT3[key]['C22']['_text'].slice(1, -1))).toString();
                    secondConstruction['C23']['_text'] = (Number(secondConstruction['C23']['_text'].slice(1, -1)) +
                        Number(bidT3[key]['C23']['_text'].slice(1, -1))).toString();
                }
                if (thirdConstruction !== undefined) {
                    thirdConstruction['C20']['_text'] = (Number(thirdConstruction['C20']['_text'].slice(1, -1)) +
                        Number(bidT3[key]['C20']['_text'].slice(1, -1))).toString();
                    thirdConstruction['C21']['_text'] = (Number(thirdConstruction['C21']['_text'].slice(1, -1)) +
                        Number(bidT3[key]['C21']['_text'].slice(1, -1))).toString();
                    thirdConstruction['C22']['_text'] = (Number(thirdConstruction['C22']['_text'].slice(1, -1)) +
                        Number(bidT3[key]['C22']['_text'].slice(1, -1))).toString();
                    thirdConstruction['C23']['_text'] = (Number(thirdConstruction['C23']['_text'].slice(1, -1)) +
                        Number(bidT3[key]['C23']['_text'].slice(1, -1))).toString();
                }
            }
        }
    };
    CalculatePrice.SetBusinessInfo = function () {
        var bidT1 = this.eleBID['T1'];
        bidT1['C17']['_text'] = Data_1.Data.CompanyRegistrationNum;
        bidT1['C18']['_text'] = Data_1.Data.CompanyRegistrationName;
    };
    CalculatePrice.SubstitutePrice = function () {
        var bidT5 = this.eleBID['T5'];
        for (var key in bidT5) {
            if (bidT5[key]['C4']['_text'].slice(1, -1) !== '이윤' &&
                Data_1.Data.Bidding[bidT5[key]['C4']['_text'].slice(1, -1)]) {
                bidT5[key]['C8']['_text'] =
                    Data_1.Data.Bidding[bidT5[key]['C4']['_text'].slice(1, -1)].toString();
            }
            else if (Data_1.Data.Rate1[bidT5[key]['C4']['_text'].slice(1, -1)]) {
                bidT5[key]['C8']['_text'] =
                    Data_1.Data.Bidding[bidT5[key]['C4']['_text'].slice(1, -1)].toString();
            }
        }
        fs_1.default.writeFileSync(Data_1.Data.folder + '\\OutputDataFromBID.json', JSON.stringify(this.docBID));
    };
    CalculatePrice.CreateZipFile = function (xlsfiles) {
        if (fs_1.default.existsSync(Data_1.Data.folder + '\\WORK DIRCETORY\\입찰내역.zip')) {
            fs_1.default.rmSync(Data_1.Data.folder + '\\WORK DIRCETORY\\입찰내역.zip');
        }
        var zip = new adm_zip_1.default();
        for (var idx in xlsfiles) {
            zip.addLocalFile(Data_1.Data.folder + '\\' + xlsfiles[idx]);
        }
        zip.writeZip(Data_1.Data.work_path + '\\입찰내역.zip');
    };
    CalculatePrice.CreateFile = function () {
        CreateResultFile_1.CreateResultFile.Create();
        var files = fs_1.default.readdirSync(Data_1.Data.folder);
        var xlsFiles = files.filter(function (file) { return file.substring(file.length - 4, file.length).toLowerCase() === '.xls'; });
        this.CreateZipFile(xlsFiles);
    };
    CalculatePrice.exSum = 0;
    CalculatePrice.exCount = 0;
    return CalculatePrice;
}());
// CalculatePrice.Calculation();
