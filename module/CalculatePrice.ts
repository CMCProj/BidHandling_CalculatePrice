const DATA = require('./Data')
const FCA = require('./fillCostAccount')
const CRF = require('./CreateResultFile')

const fs = require('fs')
const AdmZip = require('adm-zip')
const Big = require('big.js')

class CalculatePrice {
    private static docBID: JSON
    private static eleBID: JSON
    private static maxBID: JSON
    public static myPercent: number
    public static balancedUnitPriceRate: number
    public static targetRate: number
    private static exSum: number = 0
    private static exCount: number = 0

    public static Calculation(): void {
        const bidString: string = fs.readFileSync(DATA.folder + '\\OutputDATAFromBID.json', 'utf-8')
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

    public static Reset(): void {
        DATA.ExecuteReset = '1'

        let DM: number = DATA.Investigation['직접재료비']
        let DL: number = DATA.Investigation['직접노무비']
        let OE: number = DATA.Investigation['산출경비']
        let FM: number = DATA.InvestigateFixedPriceDirectMaterial
        let FL: number = DATA.InvestigateFixedPriceDirectLabor
        let FOE: number = DATA.InvestigateFixedPriceOutputExpense
        let SM: number = DATA.InvestigateStandardMaterial
        let SL: number = DATA.InvestigateStandardLabor
        let SOE: number = DATA.InvestigateStandardExpense

        DATA.RealDirectMaterial = DM
        DATA.RealDirectLabor = DL
        DATA.RealOutputExpense = OE
        DATA.FixedPriceDirectMaterial = FM
        DATA.FixedPriceDirectLabor = FL
        DATA.FixedPriceOutputExpense = FOE
        DATA.StandardMaterial = SM
        DATA.StandardLabor = SL
        DATA.StandardExpense = SOE

        const bidT3: object = this.eleBID['T3']
        let code: string
        let type: string

        for (let key in bidT3) {
            code = JSON.stringify(bidT3[key]['C9']['_text'])
            type = JSON.stringify(bidT3[key]['C5']['_text'])[1]

            if (code !== undefined && type === 'S') {
                let constNum: string = JSON.stringify(bidT3[key]['C1']['_text']).slice(1, -1)
                let numVal: string = JSON.stringify(bidT3[key]['C2']['_text']).slice(1, -1)
                let detailVal: string = JSON.stringify(bidT3[key]['C3']['_text']).slice(1, -1)
                let curObject = DATA.Dic.get(constNum).find(
                    (x) => x.WorkNum === numVal && x.DetailWorkNum === detailVal
                )

                curObject.MaterialUnit = Number(bidT3[key]['C16']['_text'].slice(1, -1))
                curObject.LaborUnit = Number(bidT3[key]['C17']['_text'].slice(1, -1))
                curObject.ExpenseUnit = Number(bidT3[key]['C18']['_text'].slice(1, -1))
            }
        }

        DATA.ExecuteReset = '0'
    }

    public static ApplyStandardPriceOption(): void {
        const bidT3: object = this.eleBID['T3']
        let code: string
        let type: string

        for (let key in bidT3) {
            code = JSON.stringify(bidT3[key]['C9']['_text'])
            type = JSON.stringify(bidT3[key]['C5']['_text'])[1]

            if (code !== undefined && type === 'S') {
                let constNum: string = JSON.stringify(bidT3[key]['C1']['_text']).slice(1, -1)
                let numVal: string = JSON.stringify(bidT3[key]['C2']['_text']).slice(1, -1)
                let detailVal: string = JSON.stringify(bidT3[key]['C3']['_text']).slice(1, -1)
                let curObject = DATA.Dic.get(constNum).find(
                    (x) => x.WorkNum === numVal && x.DetailWorkNum === detailVal
                )

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
    }

    public static GetFixedPriceRate(): void {
        const directConstPrice: number = DATA.Investigation['직공비']
        const fixCostSum: number =
            DATA.InvestigateFixedPriceDirectMaterial +
            DATA.InvestigateFixedPriceDirectLabor +
            DATA.InvestigateFixedPriceOutputExpense

        DATA.FixedPricePercent = Math.trunc((fixCostSum / directConstPrice) * 100 * 10000) / 10000
    }

    public static FindMyPercent(): void {
        if (DATA.FixedPricePercent < 20.0) this.myPercent = 0.85
        else if (DATA.FixedPricePercent < 25.0) this.myPercent = 0.84
        else if (DATA.FixedPricePercent < 30.0) this.myPercent = 0.83
        else this.myPercent = 0.82
    }

    public static GetWeight(): void {
        let varCostSum =
            DATA.RealPriceDirectMaterial + DATA.RealPriceDirectLabor + DATA.RealPriceOutputExpense
        let weight: number
        let maxWeight: number = 0
        let weightSum: number = 0
        let max: typeof DATA = new DATA()

        DATA.Dic.forEach((value, _) => {
            for (let idx in value) {
                if (value[idx].Item.localeCompare('일반')) {
                    let material = value[idx].Material
                    let labor = value[idx].Labor
                    let expense = value[idx].Expense

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
            let lack: number = 1.0 - weightSum
            max.Weight += lack
        }
    }

    public static CalculateRate(presonalRate: any, balancedRate: any): void {
        const unitPrice = Big(100)
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

    public static RoundOrTruncate(
        Rate: number,
        Object: typeof DATA,
        refMyMaterialUnit: { value: number },
        refMyLaborUnit: { value: number },
        refMyExpenseUnit: { value: number }
    ): void {
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

    public static CheckLaborLimit80(
        Object: typeof DATA,
        refMyMaterialUnit: { value: number },
        refMyLaborUnit: { value: number },
        refMyExpenseUnit: { value: number }
    ): void {
        if (Object.LaborUnit * 0.8 > refMyLaborUnit.value) {
            const deficiency: number = Object.LaborUnit * 0.8 - refMyLaborUnit.value

            if (refMyLaborUnit.value !== 0) refMyMaterialUnit.value -= deficiency
            else if (refMyExpenseUnit.value !== 0) refMyExpenseUnit.value -= deficiency

            refMyLaborUnit.value = Object.LaborUnit * 0.8
        }
    }

    public static Recalculation(): void {
        const bidT3: object = this.eleBID['T3']
        let code: string
        let type: string
        this.exCount = 0
        this.exSum = 0

        for (let key in bidT3) {
            code = JSON.stringify(bidT3[key]['C9']['_text'])
            type = JSON.stringify(bidT3[key]['C5']['_text'])[1]

            if (code !== undefined && type === 'S') {
                let constNum: string = JSON.stringify(bidT3[key]['C1']['_text']).slice(1, -1)
                let numVal: string = JSON.stringify(bidT3[key]['C2']['_text']).slice(1, -1)
                let detailVal: string = JSON.stringify(bidT3[key]['C3']['_text']).slice(1, -1)
                let curObject = DATA.Dic.get(constNum).find(
                    (x) => x.WorkNum === numVal && x.DetailWorkNum === detailVal
                )

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

                    let targetPrice: number =
                        (curObject.MaterialUnit + curObject.LaborUnit + curObject.ExpenseUnit) *
                        this.targetRate

                    let myMaterialUnit = { value: 0 }
                    let myLaborUnit = { value: 0 }
                    let myExpenseUnit = { value: 0 }
                    let myPrice: number

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

                            continue
                        } else {
                            this.RoundOrTruncate(
                                this.targetRate,
                                curObject,
                                myMaterialUnit,
                                myLaborUnit,
                                myExpenseUnit
                            )
                            this.CheckLaborLimit80(
                                curObject,
                                myMaterialUnit,
                                myLaborUnit,
                                myExpenseUnit
                            )
                        }
                    } else if (DATA.ZeroWeightDeduction.localeCompare('2')) {
                        this.RoundOrTruncate(
                            this.targetRate,
                            curObject,
                            myMaterialUnit,
                            myLaborUnit,
                            myExpenseUnit
                        )
                        this.CheckLaborLimit80(
                            curObject,
                            myMaterialUnit,
                            myLaborUnit,
                            myExpenseUnit
                        )
                    }

                    myPrice = myMaterialUnit.value + myLaborUnit.value + myExpenseUnit.value

                    if (DATA.LaborCostLowBound.localeCompare('1')) {
                        let Excess: number = myPrice - targetPrice
                        let laborExcess: number = myLaborUnit.value - curObject.LaborUnit * 0.8
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
                        Math.trunc(curObject.MaterialUnit * this.targetRate * 10) / 10
                    curObject.LaborUnit =
                        Math.trunc(curObject.LaborUnit * this.targetRate * 10) / 10
                    curObject.ExpenseUnit =
                        Math.trunc(curObject.ExpenseUnit * this.targetRate * 10) / 10

                    this.exSum += curObject.PriceSum
                    this.exCount++
                }
            }
        }
    }

    public static SetExcludingPrice(): void {
        let TempInvestDirectSum: number = DATA.Investigation['직공비']
        let TempRealDirectSum: number = FCA.ToLong(
            DATA.RealDirectMaterial + DATA.RealDirectLabor + DATA.RealOutputExpense
        )
        let InvestExSum: number =
            DATA.ExcludingMaterial + DATA.ExcludingLabor + DATA.ExcludingExpense
        let TempExRate: number = Math.round((InvestExSum / TempInvestDirectSum) * 100000) / 100000
        let TempExPrice: number = Math.ceil(TempRealDirectSum * TempExRate)
        let keyFound: number = 0

        const bidT3: object = this.eleBID['T3']
        let code: string
        let type: string

        if (DATA.CostAccountDeduction.localeCompare('1')) {
            TempExPrice = Math.ceil(Math.ceil(TempExPrice * 0.997))
        }

        for (let key in bidT3) {
            code = JSON.stringify(bidT3[key]['C9']['_text'])
            type = JSON.stringify(bidT3[key]['C5']['_text'])[1]

            if (code !== undefined && type === 'S') {
                let constNum: string = JSON.stringify(bidT3[key]['C1']['_text']).slice(1, -1)
                let numVal: string = JSON.stringify(bidT3[key]['C2']['_text']).slice(1, -1)
                let detailVal: string = JSON.stringify(bidT3[key]['C3']['_text']).slice(1, -1)
                let curObject = DATA.Dic.get(constNum).find(
                    (x) => x.WorkNum === numVal && x.DetailWorkNum === detailVal
                )

                if (curObject.Item.localeCompare('제요율적용제외')) {
                    if (this.maxBID === undefined) {
                        this.maxBID = bidT3[key]
                    }

                    if (
                        bidT3[key]['C15']['_text'].slice(1, -1) === '1' &&
                        bidT3[key]['C15']['_text'].slice(1, -1) !== '0'
                    ) {
                        if (
                            Number(bidT3[key]['C19']['_text'].slice(1, -1)) >
                            Number(this.maxBID['C19']['_text'].slice(1, -1))
                        ) {
                            if (
                                Number(bidT3[key]['C19']['_text'].slice(1, -1)) * 1.5 >
                                curObject.PriceSum + (TempExPrice - this.exSum)
                            ) {
                                keyFound = 1
                                this.maxBID = bidT3[key]
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

        if (keyFound === 0) {
            let divisionPrice: number = Math.trunc((TempExPrice - this.exSum) / this.exCount)
            let deficiency: number = Math.ceil(
                TempExPrice - this.exSum - divisionPrice * this.exCount
            )
            let count: number = 0

            while (count !== this.exCount) {
                const bidT3: object = this.eleBID['T3']
                let code: string
                let type: string

                for (let key in bidT3) {
                    code = JSON.stringify(bidT3[key]['C9']['_text'])
                    type = JSON.stringify(bidT3[key]['C5']['_text'])[1]

                    if (code !== undefined && type === 'S') {
                        let constNum: string = JSON.stringify(bidT3[key]['C1']['_text']).slice(
                            1,
                            -1
                        )
                        let numVal: string = JSON.stringify(bidT3[key]['C2']['_text']).slice(1, -1)
                        let detailVal: string = JSON.stringify(bidT3[key]['C3']['_text']).slice(
                            1,
                            -1
                        )
                        let curObject = DATA.Dic.get(constNum).find(
                            (x) => x.WorkNum === numVal && x.DetailWorkNum === detailVal
                        )

                        if (
                            curObject.Item.localeCompare('제요율적용제외') &&
                            curObject.Quantity === 1
                        ) {
                            if (curObject.LaborUnit !== 0) {
                                if (
                                    Number(bidT3[key]['C19']['_text'].slice(1, -1)) * 1.5 >
                                    curObject.LaborUnit + divisionPrice
                                ) {
                                    curObject.LaborUnit += divisionPrice

                                    bidT3[key]['C17']['_text'] = curObject.LaborUnit.toString()
                                    bidT3[key]['C19']['_text'] = curObject.UnitPriceSum.toString()
                                    bidT3[key]['C21']['_text'] = curObject.Labor.toString()
                                    bidT3[key]['C23']['_text'] = curObject.PriceSum.toString()
                                    count++
                                }

                                if (count === this.exCount) {
                                    bidT3[key]['C17']['_text'] = (
                                        deficiency + curObject.LaborUnit
                                    ).toString()
                                    bidT3[key]['C19']['_text'] = (
                                        deficiency + curObject.UnitPriceSum
                                    ).toString()
                                    bidT3[key]['C21']['_text'] = (
                                        deficiency + curObject.Labor
                                    ).toString()
                                    bidT3[key]['C23']['_text'] = (
                                        deficiency + curObject.PriceSum
                                    ).toString()
                                    break
                                }
                            } else {
                                if (curObject.ExpenseUnit !== 0) {
                                    if (
                                        Number(bidT3[key]['C19']['_text'].slice(1, -1)) * 1.5 >
                                        curObject.ExpenseUnit + divisionPrice
                                    ) {
                                        curObject.ExpenseUnit += divisionPrice

                                        bidT3[key]['C18']['_text'] =
                                            curObject.ExpenseUnit.toString()
                                        bidT3[key]['C19']['_text'] =
                                            curObject.UnitPriceSum.toString()
                                        bidT3[key]['C22']['_text'] = curObject.Expense.toString()
                                        bidT3[key]['C23']['_text'] = curObject.PriceSum.toString()
                                        count++
                                    }

                                    if (count === this.exCount) {
                                        bidT3[key]['C18']['_text'] = (
                                            deficiency + curObject.ExpenseUnit
                                        ).toString()
                                        bidT3[key]['C19']['_text'] = (
                                            deficiency + curObject.UnitPriceSum
                                        ).toString()
                                        bidT3[key]['C22']['_text'] = (
                                            deficiency + curObject.Expense
                                        ).toString()
                                        bidT3[key]['C23']['_text'] = (
                                            deficiency + curObject.PriceSum
                                        ).toString()
                                        break
                                    }
                                } else {
                                    if (
                                        Number(bidT3[key]['C19']['_text'].slice(1, -1)) * 1.5 >
                                        curObject.MaterialUnit + divisionPrice
                                    ) {
                                        curObject.MaterialUnit += divisionPrice

                                        bidT3[key]['C16']['_text'] =
                                            curObject.MaterialUnit.toString()
                                        bidT3[key]['C19']['_text'] =
                                            curObject.UnitPriceSum.toString()
                                        bidT3[key]['C20']['_text'] = curObject.Material.toString()
                                        bidT3[key]['C23']['_text'] = curObject.PriceSum.toString()
                                        count++
                                    }

                                    if (count === this.exCount) {
                                        bidT3[key]['C18']['_text'] = (
                                            deficiency + curObject.MaterialUnit
                                        ).toString()
                                        bidT3[key]['C19']['_text'] = (
                                            deficiency + curObject.UnitPriceSum
                                        ).toString()
                                        bidT3[key]['C22']['_text'] = (
                                            deficiency + curObject.Material
                                        ).toString()
                                        bidT3[key]['C23']['_text'] = (
                                            deficiency + curObject.PriceSum
                                        ).toString()
                                        break
                                    }
                                }
                            }
                        }
                    }
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

    public static GetAdjustedExcludePrice(): void {
        const bidT3: object = this.eleBID['T3']
        let code: string
        let type: string

        for (let key in bidT3) {
            code = JSON.stringify(bidT3[key]['C9']['_text'])
            type = JSON.stringify(bidT3[key]['C5']['_text'])[1]

            if (code !== undefined && type === 'S') {
                let constNum: string = JSON.stringify(bidT3[key]['C1']['_text']).slice(1, -1)
                let numVal: string = JSON.stringify(bidT3[key]['C2']['_text']).slice(1, -1)
                let detailVal: string = JSON.stringify(bidT3[key]['C3']['_text']).slice(1, -1)
                let curObject = DATA.Dic.get(constNum).find(
                    (x) => x.WorkNum === numVal && x.DetailWorkNum === detailVal
                )

                if (curObject.Item.localeCompare('제요율적용제외')) {
                    DATA.AdjustedExMaterial += Number(bidT3[key]['C20']['_text'].slice(1, -1))
                    DATA.AdjustedExLabor += Number(bidT3[key]['C21']['_text'].slice(1, -1))
                    DATA.AdjustedExExpense += Number(bidT3[key]['C22']['_text'].slice(1, -1))
                }
            }
        }
    }

    public static SetPriceOfSuperConstruction(): void {
        let firstConstruction: JSON = undefined
        let secondConstruction: JSON = undefined
        let thirdConstruction: JSON = undefined

        const bidT3: object = this.eleBID['T3']
        let code: string
        let type: string

        for (let key in bidT3) {
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

    public static SetBusinessInfo(): void {
        const bidT1: object = this.eleBID['T1']

        bidT1['C17']['_text'] = DATA.CompanyRegistrationNum
        bidT1['C18']['_text'] = DATA.CompanyRegistrationName
    }

    public static SubstitutePrice(): void {
        const bidT5: object = this.eleBID['T5']

        for (let key in bidT5) {
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

    public static CreateZipFile(xlsfiles: Array<string>): void {
        if (fs.existsSync(DATA.folder + '\\WORK DIRCETORY\\입찰내역.zip')) {
            fs.rmSync(DATA.folder + '\\WORK DIRCETORY\\입찰내역.zip')
        }

        const zip = new AdmZip()

        for (let idx in xlsfiles) {
            zip.addLocalFile(DATA.folder + '\\' + xlsfiles[idx])
        }

        zip.writeZip(DATA.work_path + '\\입찰내역.zip')
    }

    public static CreateFile(): void {
        CRF.Create()
        const files: Array<string> = fs.readdirSync(DATA.folder)
        const xlsFiles: Array<string> = files.filter(
            (file) => file.substring(file.length - 4, file.length).toLowerCase() === '.xls'
        )
        this.CreateZipFile(xlsFiles)
    }
}

module.exports = CalculatePrice

// CalculatePrice.Calculation();
