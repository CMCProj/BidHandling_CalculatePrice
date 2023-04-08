var DATA = require('./Data')
var FCA = require('./fillCostAccount')
//var CRF = require('./CreateResultFile');
var fs = require('fs')
var AdmZip = require('adm-zip')
var Big = require('big.js')
var CalculatePrice = /** @class */ (function () {
    function CalculatePrice() {}
    CalculatePrice.Calculation = function () {
        var bidString = fs.readFileSync(DATA.folder + '\\OutputDATAFromBID.json', 'utf-8')
        this.docBID = JSON.parse(bidString)
        this.eleBID = this.docBID['data']
        this.Reset()
        if (DATA.StandardMarketDeduction.localeCompare('1')) this.ApplyStandardPriceOption()
        this.GetFixedPriceRate()
        this.FindMyPercent()
        this.GetWeight()
        this.CalculateRate(DATA.PersonalRate, DATA.BalancedRate)
        this.Recalculation()
        if (this.exCount !== 0) {
            this.SetExcludingPrice()
            this.GetAdjustedExcludePrice()
        }
        this.SetPriceOfSuperConstruction()
        FCA.CalculateBiddingCosts()
        this.SetBusinessInfo()
        this.SubstitutePrice()
        this.CreateFile()
    }
    CalculatePrice.Reset = function () {
        DATA.ExecuteReset = '1'
        var DM = DATA.Investigation['직접재료비']
        var DL = DATA.Investigation['직접노무비']
        var OE = DATA.Investigation['산출경비']
        var FM = DATA.InvestigateFixedPriceDirectMaterial
        var FL = DATA.InvestigateFixedPriceDirectLabor
        var FOE = DATA.InvestigateFixedPriceOutputExpense
        var SM = DATA.InvestigateStandardMaterial
        var SL = DATA.InvestigateStandardLabor
        var SOE = DATA.InvestigateStandardExpense
        DATA.RealDirectMaterial = DM
        DATA.RealDirectLabor = DL
        DATA.RealOutputExpense = OE
        DATA.FixedPriceDirectMaterial = FM
        DATA.FixedPriceDirectLabor = FL
        DATA.FixedPriceOutputExpense = FOE
        DATA.StandardMaterial = SM
        DATA.StandardLabor = SL
        DATA.StandardExpense = SOE
        var bidT3 = this.eleBID['T3']
        var code
        var type
        var _loop_1 = function (key) {
            code = JSON.stringify(bidT3[key]['C9']['_text'])
            type = JSON.stringify(bidT3[key]['C5']['_text'])[1]
            if (code !== undefined && type === 'S') {
                var constNum = JSON.stringify(bidT3[key]['C1']['_text']).slice(1, -1)
                var numVal_1 = JSON.stringify(bidT3[key]['C2']['_text']).slice(1, -1)
                var detailVal_1 = JSON.stringify(bidT3[key]['C3']['_text']).slice(1, -1)
                var curObject = DATA.Dic.get(constNum).find(function (x) {
                    return x.WorkNum === numVal_1 && x.DetailWorkNum === detailVal_1
                })
                curObject.MaterialUnit = Number(bidT3[key]['C16']['_text'].slice(1, -1))
                curObject.LaborUnit = Number(bidT3[key]['C17']['_text'].slice(1, -1))
                curObject.ExpenseUnit = Number(bidT3[key]['C18']['_text'].slice(1, -1))
            }
        }
        for (var key in bidT3) {
            _loop_1(key)
        }
        DATA.ExecuteReset = '0'
    }
    CalculatePrice.ApplyStandardPriceOption = function () {
        var bidT3 = this.eleBID['T3']
        var code
        var type
        var _loop_2 = function (key) {
            code = JSON.stringify(bidT3[key]['C9']['_text'])
            type = JSON.stringify(bidT3[key]['C5']['_text'])[1]
            if (code !== undefined && type === 'S') {
                var constNum = JSON.stringify(bidT3[key]['C1']['_text']).slice(1, -1)
                var numVal_2 = JSON.stringify(bidT3[key]['C2']['_text']).slice(1, -1)
                var detailVal_2 = JSON.stringify(bidT3[key]['C3']['_text']).slice(1, -1)
                var curObject = DATA.Dic.get(constNum).find(function (x) {
                    return x.WorkNum === numVal_2 && x.DetailWorkNum === detailVal_2
                })
                if (curObject.Item.localeCompare('표준시장단가')) {
                    DATA.RealDirectMaterial -= +JSON.stringify(bidT3[key]['C20']['_text']).slice(
                        1,
                        -1
                    )
                    DATA.RealDirectLabor -= +JSON.stringify(bidT3[key]['C21']['_text']).slice(1, -1)
                    DATA.RealOutputExpense -= +JSON.stringify(bidT3[key]['C22']['_text']).slice(
                        1,
                        -1
                    )
                    DATA.FixedPriceDirectMaterial -= +JSON.stringify(
                        bidT3[key]['C20']['_text']
                    ).slice(1, -1)
                    DATA.FixedPriceDirectLabor -= +JSON.stringify(bidT3[key]['C21']['_text']).slice(
                        1,
                        -1
                    )
                    DATA.FixedPriceOutputExpense -= +JSON.stringify(
                        bidT3[key]['C22']['_text']
                    ).slice(1, -1)
                    DATA.StandardMaterial -= +JSON.stringify(bidT3[key]['C20']['_text']).slice(
                        1,
                        -1
                    )
                    DATA.StandardLabor -= +JSON.stringify(bidT3[key]['C21']['_text']).slice(1, -1)
                    DATA.StandardExpense -= +JSON.stringify(bidT3[key]['C22']['_text']).slice(1, -1)
                    if (curObject.MaterialUnit !== 0)
                        curObject.MaterialUnit =
                            Math.trunc(curObject.MaterialUnit * 0.997 * 10) / 10 + 1.0
                    if (curObject.LaborUnit !== 0)
                        curObject.LaborUnit =
                            Math.trunc(curObject.LaborUnit * 0.997 * 10) / 10 + 0.1
                    if (curObject.ExpenseUnit !== 0)
                        curObject.ExpenseUnit =
                            Math.trunc(curObject.ExpenseUnit * 0.997 * 10) / 10 + 0.1
                    bidT3[key]['C16']['_text'] = curObject.MaterialUnit.toString()
                    bidT3[key]['C17']['_text'] = curObject.LaborUnit.toString()
                    bidT3[key]['C18']['_text'] = curObject.ExpenseUnit.toString()
                    bidT3[key]['C19']['_text'] = curObject.UnitPriceSum.toString()
                    bidT3[key]['C20']['_text'] = curObject.Material.toString()
                    bidT3[key]['C21']['_text'] = curObject.Labor.toString()
                    bidT3[key]['C22']['_text'] = curObject.Expense.toString()
                    bidT3[key]['C23']['_text'] = curObject.PriceSum.toString()
                    DATA.RealDirectMaterial += +JSON.stringify(bidT3[key]['C20']['_text']).slice(
                        1,
                        -1
                    )
                    DATA.RealDirectLabor += +JSON.stringify(bidT3[key]['C21']['_text']).slice(1, -1)
                    DATA.RealOutputExpense += +JSON.stringify(bidT3[key]['C22']['_text']).slice(
                        1,
                        -1
                    )
                    DATA.FixedPriceDirectMaterial += +JSON.stringify(
                        bidT3[key]['C20']['_text']
                    ).slice(1, -1)
                    DATA.FixedPriceDirectLabor += +JSON.stringify(bidT3[key]['C21']['_text']).slice(
                        1,
                        -1
                    )
                    DATA.FixedPriceOutputExpense += +JSON.stringify(
                        bidT3[key]['C22']['_text']
                    ).slice(1, -1)
                    DATA.StandardMaterial += +JSON.stringify(bidT3[key]['C20']['_text']).slice(
                        1,
                        -1
                    )
                    DATA.StandardLabor += +JSON.stringify(bidT3[key]['C21']['_text']).slice(1, -1)
                    DATA.StandardExpense += +JSON.stringify(bidT3[key]['C22']['_text']).slice(1, -1)
                }
            }
        }
        for (var key in bidT3) {
            _loop_2(key)
        }
    }
    CalculatePrice.GetFixedPriceRate = function () {
        var directConstPrice = DATA.Investigation['직공비']
        var fixCostSum =
            DATA.InvestigateFixedPriceDirectMaterial +
            DATA.InvestigateFixedPriceDirectLabor +
            DATA.InvestigateFixedPriceOutputExpense
        DATA.FixedPricePercent = Math.trunc((fixCostSum / directConstPrice) * 100 * 10000) / 10000
    }
    CalculatePrice.FindMyPercent = function () {
        if (DATA.FixedPricePercent < 20.0) this.myPercent = 0.85
        else if (DATA.FixedPricePercent < 25.0) this.myPercent = 0.84
        else if (DATA.FixedPricePercent < 30.0) this.myPercent = 0.83
        else this.myPercent = 0.82
    }
    CalculatePrice.GetWeight = function () {
        var varCostSum =
            DATA.RealPriceDirectMaterial + DATA.RealPriceDirectLabor + DATA.RealPriceOutputExpense
        var weight
        var maxWeight = 0
        var weightSum = 0
        var max = new DATA()
        DATA.Dic.forEach(function (value, _) {
            for (var idx in value) {
                if (value[idx].Item.localeCompare('일반')) {
                    var material = value[idx].Material
                    var labor = value[idx].Labor
                    var expense = value[idx].Expense
                    weight =
                        Math.round(((material + labor + expense) / varCostSum) * 1000000) / 1000000
                    weightSum += weight
                    if (maxWeight < weight) {
                        maxWeight = weight
                        max = value[idx]
                    }
                    value[idx].Weight = weight
                }
            }
        })
        if (weightSum !== 1.0) {
            var lack = 1.0 - weightSum
            max.Weight += lack
        }
    }
    CalculatePrice.CalculateRate = function (presonalRate, balancedRate) {
        var unitPrice = Big(100)
        presonalRate = Big(presonalRate)
        balancedRate = Big(balancedRate)
        this.balancedUnitPriceRate =
            (0.9 * unitPrice * (1.0 + balancedRate / 100) * this.myPercent) /
            (1.0 - 0.1 * this.myPercent) /
            100
        this.targetRate =
            ((unitPrice * (1.0 + presonalRate / 100) * 0.9 +
                unitPrice * this.balancedUnitPriceRate * 0.1) *
                this.myPercent) /
            100
        this.targetRate = Math.trunc(this.targetRate * 1000000) / 1000000
        console.log(this.targetRate)
    }
    CalculatePrice.RoundOrTruncate = function (
        Rate,
        Object,
        refMyMaterialUnit,
        refMyLaborUnit,
        refMyExpenseUnit
    ) {
        if (DATA.UnitPriceTrimming.localeCompare('1')) {
            refMyMaterialUnit.value = Math.trunc(Object.MaterialUnit * Rate * 10) / 10
            refMyLaborUnit.value = Math.trunc(Object.LaborUnit * Rate * 10) / 10
            refMyExpenseUnit.value = Math.trunc(Object.ExpenseUnit * Rate * 10) / 10
        } else if (DATA.UnitPriceTrimming.localeCompare('2')) {
            refMyMaterialUnit.value = Math.ceil(Object.MaterialUnit * Rate)
            refMyLaborUnit.value = Math.ceil(Object.LaborUnit * Rate)
            refMyExpenseUnit.value = Math.ceil(Object.ExpenseUnit * Rate)
        }
    }
    CalculatePrice.CheckLaborLimit80 = function (
        Object,
        refMyMaterialUnit,
        refMyLaborUnit,
        refMyExpenseUnit
    ) {
        if (Object.LaborUnit * 0.8 > refMyLaborUnit.value) {
            var deficiency = Object.LaborUnit * 0.8 - refMyLaborUnit.value
            if (refMyLaborUnit.value !== 0) refMyMaterialUnit.value -= deficiency
            else if (refMyExpenseUnit.value !== 0) refMyExpenseUnit.value -= deficiency
            refMyLaborUnit.value = Object.LaborUnit * 0.8
        }
    }
    CalculatePrice.Recalculation = function () {
        var bidT3 = this.eleBID['T3']
        var code
        var type
        this.exCount = 0
        this.exSum = 0
        var _loop_3 = function (key) {
            code = JSON.stringify(bidT3[key]['C9']['_text'])
            type = JSON.stringify(bidT3[key]['C5']['_text'])[1]
            if (code !== undefined && type === 'S') {
                var constNum = JSON.stringify(bidT3[key]['C1']['_text']).slice(1, -1)
                var numVal_3 = JSON.stringify(bidT3[key]['C2']['_text']).slice(1, -1)
                var detailVal_3 = JSON.stringify(bidT3[key]['C3']['_text']).slice(1, -1)
                var curObject = DATA.Dic.get(constNum).find(function (x) {
                    return x.WorkNum === numVal_3 && x.DetailWorkNum === detailVal_3
                })
                if (curObject.Item.localeCompare('일반')) {
                    DATA.RealDirectMaterial -= +JSON.stringify(bidT3[key]['C20']['_text']).slice(
                        1,
                        -1
                    )
                    DATA.RealDirectLabor -= +JSON.stringify(bidT3[key]['C21']['_text']).slice(1, -1)
                    DATA.RealOutputExpense -= +JSON.stringify(bidT3[key]['C22']['_text']).slice(
                        1,
                        -1
                    )
                    var targetPrice =
                        (curObject.MaterialUnit + curObject.LaborUnit + curObject.ExpenseUnit) *
                        this_1.targetRate
                    var myMaterialUnit = { value: 0 }
                    var myLaborUnit = { value: 0 }
                    var myExpenseUnit = { value: 0 }
                    var myPrice = void 0
                    if (DATA.ZeroWeightDeduction.localeCompare('1')) {
                        if (curObject.Weight === 0 && curObject.LaborUnit === 0) {
                            curObject.MaterialUnit = Math.ceil(curObject.MaterialUnit * 0.5)
                            curObject.ExpenseUnit = Math.ceil(curObject.ExpenseUnit * 0.5)
                            bidT3[key]['C16']['_text'] = curObject.MaterialUnit.toString()
                            bidT3[key]['C17']['_text'] = curObject.LaborUnit.toString()
                            bidT3[key]['C18']['_text'] = curObject.ExpenseUnit.toString()
                            bidT3[key]['C19']['_text'] = curObject.UnitPriceSum.toString()
                            bidT3[key]['C20']['_text'] = curObject.Material.toString()
                            bidT3[key]['C21']['_text'] = curObject.Labor.toString()
                            bidT3[key]['C22']['_text'] = curObject.Expense.toString()
                            bidT3[key]['C23']['_text'] = curObject.PriceSum.toString()
                            DATA.RealDirectMaterial += +JSON.stringify(
                                bidT3[key]['C20']['_text']
                            ).slice(1, -1)
                            DATA.RealDirectLabor += +JSON.stringify(
                                bidT3[key]['C21']['_text']
                            ).slice(1, -1)
                            DATA.RealOutputExpense += +JSON.stringify(
                                bidT3[key]['C22']['_text']
                            ).slice(1, -1)
                            return 'continue'
                        } else {
                            this_1.RoundOrTruncate(
                                this_1.targetRate,
                                curObject,
                                myMaterialUnit,
                                myLaborUnit,
                                myExpenseUnit
                            )
                            this_1.CheckLaborLimit80(
                                curObject,
                                myMaterialUnit,
                                myLaborUnit,
                                myExpenseUnit
                            )
                        }
                    } else if (DATA.ZeroWeightDeduction.localeCompare('2')) {
                        this_1.RoundOrTruncate(
                            this_1.targetRate,
                            curObject,
                            myMaterialUnit,
                            myLaborUnit,
                            myExpenseUnit
                        )
                        this_1.CheckLaborLimit80(
                            curObject,
                            myMaterialUnit,
                            myLaborUnit,
                            myExpenseUnit
                        )
                    }
                    myPrice = myMaterialUnit.value + myLaborUnit.value + myExpenseUnit.value
                    if (DATA.LaborCostLowBound.localeCompare('1')) {
                        var Excess = myPrice - targetPrice
                        var laborExcess = myLaborUnit.value - curObject.LaborUnit * 0.8
                        laborExcess = Math.trunc(laborExcess * 10) / 10
                        if (laborExcess > 0) {
                            if (myExpenseUnit.value !== 0) {
                                myLaborUnit.value -= laborExcess
                                myExpenseUnit.value += laborExcess + Excess
                            } else {
                                if (myMaterialUnit.value !== 0) {
                                    myLaborUnit.value -= laborExcess
                                    myMaterialUnit.value += laborExcess + Excess
                                } else {
                                    myLaborUnit.value -= laborExcess
                                    myExpenseUnit.value += laborExcess + Excess
                                }
                            }
                        } else if (laborExcess < 0) {
                            myLaborUnit.value = curObject.LaborUnit * 0.8
                            if (myMaterialUnit.value !== 0) {
                                myMaterialUnit.value += laborExcess + Excess
                            } else {
                                myExpenseUnit.value += laborExcess + Excess
                            }
                        }
                    }
                    curObject.MaterialUnit = myMaterialUnit.value
                    curObject.LaborUnit = myLaborUnit.value
                    curObject.ExpenseUnit = myExpenseUnit.value
                    bidT3[key]['C16']['_text'] = curObject.MaterialUnit.toString()
                    bidT3[key]['C17']['_text'] = curObject.LaborUnit.toString()
                    bidT3[key]['C18']['_text'] = curObject.ExpenseUnit.toString()
                    bidT3[key]['C19']['_text'] = curObject.UnitPriceSum.toString()
                    bidT3[key]['C20']['_text'] = curObject.Material.toString()
                    bidT3[key]['C21']['_text'] = curObject.Labor.toString()
                    bidT3[key]['C22']['_text'] = curObject.Expense.toString()
                    bidT3[key]['C23']['_text'] = curObject.PriceSum.toString()
                    DATA.RealDirectMaterial += +JSON.stringify(bidT3[key]['C20']['_text']).slice(
                        1,
                        -1
                    )
                    DATA.RealDirectLabor += +JSON.stringify(bidT3[key]['C21']['_text']).slice(1, -1)
                    DATA.RealOutputExpense += +JSON.stringify(bidT3[key]['C22']['_text']).slice(
                        1,
                        -1
                    )
                } else if (curObject.Item === '제요율적용제외') {
                    curObject.MaterialUnit =
                        Math.trunc(curObject.MaterialUnit * this_1.targetRate * 10) / 10
                    curObject.LaborUnit =
                        Math.trunc(curObject.LaborUnit * this_1.targetRate * 10) / 10
                    curObject.ExpenseUnit =
                        Math.trunc(curObject.ExpenseUnit * this_1.targetRate * 10) / 10
                    this_1.exSum += curObject.PriceSum
                    this_1.exCount++
                }
            }
        }
        var this_1 = this
        for (var key in bidT3) {
            _loop_3(key)
        }
    }
    CalculatePrice.SetExcludingPrice = function () {
        var TempInvestDirectSum = DATA.Investigation['직공비']
        var TempRealDirectSum = FCA.ToLong(
            DATA.RealDirectMaterial + DATA.RealDirectLabor + DATA.RealOutputExpense
        )
        var InvestExSum = DATA.ExcludingMaterial + DATA.ExcludingLabor + DATA.ExcludingExpense
        var TempExRate = Math.round((InvestExSum / TempInvestDirectSum) * 100000) / 100000
        var TempExPrice = Math.ceil(TempRealDirectSum * TempExRate)
        var keyFound = 0
        var bidT3 = this.eleBID['T3']
        var code
        var type
        if (DATA.CostAccountDeduction.localeCompare('1')) {
            TempExPrice = Math.ceil(Math.ceil(TempExPrice * 0.997))
        }
        var _loop_4 = function (key) {
            code = JSON.stringify(bidT3[key]['C9']['_text'])
            type = JSON.stringify(bidT3[key]['C5']['_text'])[1]
            if (code !== undefined && type === 'S') {
                var constNum = JSON.stringify(bidT3[key]['C1']['_text']).slice(1, -1)
                var numVal_4 = JSON.stringify(bidT3[key]['C2']['_text']).slice(1, -1)
                var detailVal_4 = JSON.stringify(bidT3[key]['C3']['_text']).slice(1, -1)
                var curObject = DATA.Dic.get(constNum).find(function (x) {
                    return x.WorkNum === numVal_4 && x.DetailWorkNum === detailVal_4
                })
                if (curObject.Item.localeCompare('제요율적용제외')) {
                    if (this_2.maxBID === undefined) {
                        this_2.maxBID = bidT3[key]
                    }
                    if (
                        bidT3[key]['C15']['_text'].slice(1, -1) === '1' &&
                        bidT3[key]['C15']['_text'].slice(1, -1) !== '0'
                    ) {
                        if (
                            Number(bidT3[key]['C19']['_text'].slice(1, -1)) >
                            Number(this_2.maxBID['C19']['_text'].slice(1, -1))
                        ) {
                            if (
                                Number(bidT3[key]['C19']['_text'].slice(1, -1)) * 1.5 >
                                curObject.PriceSum + (TempExPrice - this_2.exSum)
                            ) {
                                keyFound = 1
                                this_2.maxBID = bidT3[key]
                            }
                        }
                    }
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
        var this_2 = this
        for (var key in bidT3) {
            _loop_4(key)
        }
        if (keyFound === 0) {
            var divisionPrice = Math.trunc((TempExPrice - this.exSum) / this.exCount)
            var deficiency = Math.ceil(TempExPrice - this.exSum - divisionPrice * this.exCount)
            var count = 0
            while (count !== this.exCount) {
                var bidT3_1 = this.eleBID['T3']
                var code_1 = void 0
                var type_1 = void 0
                var _loop_5 = function (key) {
                    code_1 = JSON.stringify(bidT3_1[key]['C9']['_text'])
                    type_1 = JSON.stringify(bidT3_1[key]['C5']['_text'])[1]
                    if (code_1 !== undefined && type_1 === 'S') {
                        var constNum = JSON.stringify(bidT3_1[key]['C1']['_text']).slice(1, -1)
                        var numVal_5 = JSON.stringify(bidT3_1[key]['C2']['_text']).slice(1, -1)
                        var detailVal_5 = JSON.stringify(bidT3_1[key]['C3']['_text']).slice(1, -1)
                        var curObject = DATA.Dic.get(constNum).find(function (x) {
                            return x.WorkNum === numVal_5 && x.DetailWorkNum === detailVal_5
                        })
                        if (
                            curObject.Item.localeCompare('제요율적용제외') &&
                            curObject.Quantity === 1
                        ) {
                            if (curObject.LaborUnit !== 0) {
                                if (
                                    Number(bidT3_1[key]['C19']['_text'].slice(1, -1)) * 1.5 >
                                    curObject.LaborUnit + divisionPrice
                                ) {
                                    curObject.LaborUnit += divisionPrice
                                    bidT3_1[key]['C17']['_text'] = curObject.LaborUnit.toString()
                                    bidT3_1[key]['C19']['_text'] = curObject.UnitPriceSum.toString()
                                    bidT3_1[key]['C21']['_text'] = curObject.Labor.toString()
                                    bidT3_1[key]['C23']['_text'] = curObject.PriceSum.toString()
                                    count++
                                }
                                if (count === this_3.exCount) {
                                    bidT3_1[key]['C17']['_text'] = (
                                        deficiency + curObject.LaborUnit
                                    ).toString()
                                    bidT3_1[key]['C19']['_text'] = (
                                        deficiency + curObject.UnitPriceSum
                                    ).toString()
                                    bidT3_1[key]['C21']['_text'] = (
                                        deficiency + curObject.Labor
                                    ).toString()
                                    bidT3_1[key]['C23']['_text'] = (
                                        deficiency + curObject.PriceSum
                                    ).toString()
                                    return 'break'
                                }
                            } else {
                                if (curObject.ExpenseUnit !== 0) {
                                    if (
                                        Number(bidT3_1[key]['C19']['_text'].slice(1, -1)) * 1.5 >
                                        curObject.ExpenseUnit + divisionPrice
                                    ) {
                                        curObject.ExpenseUnit += divisionPrice
                                        bidT3_1[key]['C18']['_text'] =
                                            curObject.ExpenseUnit.toString()
                                        bidT3_1[key]['C19']['_text'] =
                                            curObject.UnitPriceSum.toString()
                                        bidT3_1[key]['C22']['_text'] = curObject.Expense.toString()
                                        bidT3_1[key]['C23']['_text'] = curObject.PriceSum.toString()
                                        count++
                                    }
                                    if (count === this_3.exCount) {
                                        bidT3_1[key]['C18']['_text'] = (
                                            deficiency + curObject.ExpenseUnit
                                        ).toString()
                                        bidT3_1[key]['C19']['_text'] = (
                                            deficiency + curObject.UnitPriceSum
                                        ).toString()
                                        bidT3_1[key]['C22']['_text'] = (
                                            deficiency + curObject.Expense
                                        ).toString()
                                        bidT3_1[key]['C23']['_text'] = (
                                            deficiency + curObject.PriceSum
                                        ).toString()
                                        return 'break'
                                    }
                                } else {
                                    if (
                                        Number(bidT3_1[key]['C19']['_text'].slice(1, -1)) * 1.5 >
                                        curObject.MaterialUnit + divisionPrice
                                    ) {
                                        curObject.MaterialUnit += divisionPrice
                                        bidT3_1[key]['C16']['_text'] =
                                            curObject.MaterialUnit.toString()
                                        bidT3_1[key]['C19']['_text'] =
                                            curObject.UnitPriceSum.toString()
                                        bidT3_1[key]['C20']['_text'] = curObject.Material.toString()
                                        bidT3_1[key]['C23']['_text'] = curObject.PriceSum.toString()
                                        count++
                                    }
                                    if (count === this_3.exCount) {
                                        bidT3_1[key]['C18']['_text'] = (
                                            deficiency + curObject.MaterialUnit
                                        ).toString()
                                        bidT3_1[key]['C19']['_text'] = (
                                            deficiency + curObject.UnitPriceSum
                                        ).toString()
                                        bidT3_1[key]['C22']['_text'] = (
                                            deficiency + curObject.Material
                                        ).toString()
                                        bidT3_1[key]['C23']['_text'] = (
                                            deficiency + curObject.PriceSum
                                        ).toString()
                                        return 'break'
                                    }
                                }
                            }
                        }
                    }
                }
                var this_3 = this
                for (var key in bidT3_1) {
                    var state_1 = _loop_5(key)
                    if (state_1 === 'break') break
                }
            }
        }
        if (keyFound === 1 && this.exSum < TempExPrice) {
            this.maxBID['C17']['_text'] = (
                Number(this.maxBID['C17']['_text'].slice(1, -1)) +
                TempExPrice -
                this.exSum
            ).toString()
            this.maxBID['C19']['_text'] = (
                Number(this.maxBID['C19']['_text'].slice(1, -1)) +
                TempExPrice -
                this.exSum
            ).toString()
            this.maxBID['C21']['_text'] = (
                Number(this.maxBID['C21']['_text'].slice(1, -1)) +
                TempExPrice -
                this.exSum
            ).toString()
            this.maxBID['C23']['_text'] = (
                Number(this.maxBID['C23']['_text'].slice(1, -1)) +
                TempExPrice -
                this.exSum
            ).toString()
        }
    }
    CalculatePrice.GetAdjustedExcludePrice = function () {
        var bidT3 = this.eleBID['T3']
        var code
        var type
        var _loop_6 = function (key) {
            code = JSON.stringify(bidT3[key]['C9']['_text'])
            type = JSON.stringify(bidT3[key]['C5']['_text'])[1]
            if (code !== undefined && type === 'S') {
                var constNum = JSON.stringify(bidT3[key]['C1']['_text']).slice(1, -1)
                var numVal_6 = JSON.stringify(bidT3[key]['C2']['_text']).slice(1, -1)
                var detailVal_6 = JSON.stringify(bidT3[key]['C3']['_text']).slice(1, -1)
                var curObject = DATA.Dic.get(constNum).find(function (x) {
                    return x.WorkNum === numVal_6 && x.DetailWorkNum === detailVal_6
                })
                if (curObject.Item.localeCompare('제요율적용제외')) {
                    DATA.AdjustedExMaterial += Number(bidT3[key]['C20']['_text'].slice(1, -1))
                    DATA.AdjustedExLabor += Number(bidT3[key]['C21']['_text'].slice(1, -1))
                    DATA.AdjustedExExpense += Number(bidT3[key]['C22']['_text'].slice(1, -1))
                }
            }
        }
        for (var key in bidT3) {
            _loop_6(key)
        }
    }
    CalculatePrice.SetPriceOfSuperConstruction = function () {
        var firstConstruction = undefined
        var secondConstruction = undefined
        var thirdConstruction = undefined
        var bidT3 = this.eleBID['T3']
        var code
        var type
        for (var key in bidT3) {
            code = JSON.stringify(bidT3[key]['C9']['_text'])
            type = JSON.stringify(bidT3[key]['C5']['_text'])[1]
            if (code !== undefined && type === 'G') {
                if (bidT3[key]['C23']['_text'].slice(1, -1) === '0') {
                    if (
                        firstConstruction === undefined ||
                        bidT3[key]['C3']['_text'].slice(1, -1) === '0'
                    ) {
                        firstConstruction = bidT3[key]
                        secondConstruction = undefined
                        thirdConstruction = undefined
                    } else if (
                        bidT3[key]['C3']['_text'].slice(1, -1) ===
                            firstConstruction['C2']['_text'].slice(1, -1) &&
                        firstConstruction !== undefined
                    ) {
                        secondConstruction = bidT3[key]
                        thirdConstruction = undefined
                    } else if (
                        bidT3[key]['C3']['_text'].slice(1, -1) ===
                            secondConstruction['C2']['_text'].slice(1, -1) &&
                        secondConstruction !== undefined
                    )
                        thirdConstruction = bidT3[key]
                } else {
                    firstConstruction = undefined
                    secondConstruction = undefined
                    thirdConstruction = undefined
                }
            } else if (code !== undefined && type === 'S') {
                if (firstConstruction !== undefined) {
                    firstConstruction['C20']['_text'] = (
                        Number(firstConstruction['C20']['_text'].slice(1, -1)) +
                        Number(bidT3[key]['C20']['_text'].slice(1, -1))
                    ).toString()
                    firstConstruction['C21']['_text'] = (
                        Number(firstConstruction['C21']['_text'].slice(1, -1)) +
                        Number(bidT3[key]['C21']['_text'].slice(1, -1))
                    ).toString()
                    firstConstruction['C22']['_text'] = (
                        Number(firstConstruction['C22']['_text'].slice(1, -1)) +
                        Number(bidT3[key]['C22']['_text'].slice(1, -1))
                    ).toString()
                    firstConstruction['C23']['_text'] = (
                        Number(firstConstruction['C23']['_text'].slice(1, -1)) +
                        Number(bidT3[key]['C23']['_text'].slice(1, -1))
                    ).toString()
                }
                if (secondConstruction !== undefined) {
                    secondConstruction['C20']['_text'] = (
                        Number(secondConstruction['C20']['_text'].slice(1, -1)) +
                        Number(bidT3[key]['C20']['_text'].slice(1, -1))
                    ).toString()
                    secondConstruction['C21']['_text'] = (
                        Number(secondConstruction['C21']['_text'].slice(1, -1)) +
                        Number(bidT3[key]['C21']['_text'].slice(1, -1))
                    ).toString()
                    secondConstruction['C22']['_text'] = (
                        Number(secondConstruction['C22']['_text'].slice(1, -1)) +
                        Number(bidT3[key]['C22']['_text'].slice(1, -1))
                    ).toString()
                    secondConstruction['C23']['_text'] = (
                        Number(secondConstruction['C23']['_text'].slice(1, -1)) +
                        Number(bidT3[key]['C23']['_text'].slice(1, -1))
                    ).toString()
                }
                if (thirdConstruction !== undefined) {
                    thirdConstruction['C20']['_text'] = (
                        Number(thirdConstruction['C20']['_text'].slice(1, -1)) +
                        Number(bidT3[key]['C20']['_text'].slice(1, -1))
                    ).toString()
                    thirdConstruction['C21']['_text'] = (
                        Number(thirdConstruction['C21']['_text'].slice(1, -1)) +
                        Number(bidT3[key]['C21']['_text'].slice(1, -1))
                    ).toString()
                    thirdConstruction['C22']['_text'] = (
                        Number(thirdConstruction['C22']['_text'].slice(1, -1)) +
                        Number(bidT3[key]['C22']['_text'].slice(1, -1))
                    ).toString()
                    thirdConstruction['C23']['_text'] = (
                        Number(thirdConstruction['C23']['_text'].slice(1, -1)) +
                        Number(bidT3[key]['C23']['_text'].slice(1, -1))
                    ).toString()
                }
            }
        }
    }
    CalculatePrice.SetBusinessInfo = function () {
        var bidT1 = this.eleBID['T1']
        bidT1['C17']['_text'] = DATA.CompanyRegistrationNum
        bidT1['C18']['_text'] = DATA.CompanyRegistrationName
    }
    CalculatePrice.SubstitutePrice = function () {
        var bidT5 = this.eleBID['T5']
        for (var key in bidT5) {
            if (
                bidT5[key]['C4']['_text'].slice(1, -1) !== '이윤' &&
                DATA.Bidding[bidT5[key]['C4']['_text'].slice(1, -1)]
            ) {
                bidT5[key]['C8']['_text'] =
                    DATA.Bidding[bidT5[key]['C4']['_text'].slice(1, -1)].toString()
            } else if (DATA.Rate1[bidT5[key]['C4']['_text'].slice(1, -1)]) {
                bidT5[key]['C8']['_text'] =
                    DATA.Bidding[bidT5[key]['C4']['_text'].slice(1, -1)].toString()
            }
        }
        fs.writeFileSync(DATA.folder + '\\OutputDATAFromBID.json', JSON.stringify(this.docBID))
    }
    CalculatePrice.CreateZipFile = function (xlsfiles) {
        if (fs.existsSync(DATA.folder + '\\WORK DIRCETORY\\입찰내역.zip')) {
            fs.rmSync(DATA.folder + '\\WORK DIRCETORY\\입찰내역.zip')
        }
        var zip = new AdmZip()
        for (var idx in xlsfiles) {
            zip.addLocalFile(DATA.folder + '\\' + xlsfiles[idx])
        }
        zip.writeZip(DATA.work_path + '\\입찰내역.zip')
    }
    CalculatePrice.CreateFile = function () {
        CRF.Create()
        var files = fs.readdirSync(DATA.folder)
        var xlsFiles = files.filter(function (file) {
            return file.substring(file.length - 4, file.length).toLowerCase() === '.xls'
        })
        this.CreateZipFile(xlsFiles)
    }
    CalculatePrice.exSum = 0
    CalculatePrice.exCount = 0
    return CalculatePrice
})()
module.exports = CalculatePrice
// CalculatePrice.Calculation();
