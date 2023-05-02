var Data_1 = require('./module/Data')
var CalculatePrice_1 = require('./module/CalculatePrice')
var BidHandling_1 = require('./module/BidHandling')
var Setting_1 = require('./module/Setting')
var FillCostAccount_1 = require('./module/FillCostAccount')

// '\AutoBid\EmpthBid'에 있는 공내역 BID 파일 읽어 '\AutoBid'에 json 파일 저장
//Calculate에 필요한 Data 프로퍼티 세팅
//Calculation()실행

function RadioDecimal_Check(char) {
    //소수 1자리 or 정수 (1: 소수 2: 정수)
    Data_1.Data.UnitPriceTrimming = char
}

function CheckStandardPrice(char) {
    //표준시장단가(1: 체크 2: 체크X)
    Data_1.Data.StandardMarketDeduction = char
}

function CheckWeightValue(char) {
    //공종가중치(1: 체크 2: 체크X)
    Data_1.Data.ZeroWeightDeduction = char
}

function CheckCAD_Click(char) {
    //법정요율('1': 체크 '2': 체크X)
    Data_1.Data.CostAccountDeduction = char
}

function CheckCeiling_Click(char) {
    //원단위 체크('1': 천원절상체크 '2': 체크X)
    Data_1.Data.BidPriceRaise = char
}

function CheckLaborCost_Click(char) {
    //노무비 하한율 체크('1': 체크 '2': 체크X)
    Data_1.Data.LaborCostLowBound = char
}

function SetBusinessInfoBtnClick(CompanyName, CompanyNum) {
    Data_1.Data.CompanyRegistrationName = CompanyName
    Data_1.Data.CompanyRegistrationNum = CompanyNum
}
console.log('sdfsfd')

BidHandling_1.BidHandling.BidToJson()
//Setting_1.Setting.GetData() //BidToJson 생략하고 Setting 테스트

RadioDecimal_Check('1')
CheckStandardPrice('1')
CheckWeightValue('1')
CheckCAD_Click('1')
CheckCeiling_Click('2')
CheckLaborCost_Click('1')
SetBusinessInfoBtnClick('회사이름', '1234567890')
Data_1.Data.BalanceRateNum = 0.1
Data_1.Data.PersonalRateNum = 0.2

CalculatePrice_1.CalculatePrice.Calculation()
