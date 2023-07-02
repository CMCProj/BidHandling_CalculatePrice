var Data_1 = require('./module/Data')
var CalculatePrice_1 = require('./module/CalculatePrice')
var BidHandling_1 = require('./module/BidHandling')
var Setting_1 = require('./module/Setting')
var FillCostAccount_1 = require('./module/fillCostAccount')

// '\AutoBid\EmpthBid'에 있는 공내역 BID 파일 읽어 '\AutoBid'에 json 파일 저장
//Calculate에 필요한 Data 프로퍼티 세팅
//Calculation()실행
function execute(RadioDecimal, StandardPrice, WeightValue, CAD_Click, Ceiling_Click, LaborCost_Click, CompanyName, CompanyNum, BalanceRate, PersonalRate){

    Data_1.Data.reset_data();

    console.log(Data_1.Data);

    Data_1.Data.UnitPriceTrimming = RadioDecimal;
    Data_1.Data.StandardMarketDeduction = StandardPrice;
    Data_1.Data.ZeroWeightDeduction = WeightValue;
    Data_1.Data.CostAccountDeduction = CAD_Click;
    Data_1.Data.BidPriceRaise = Ceiling_Click;
    Data_1.Data.LaborCostLowBound = LaborCost_Click;
    Data_1.Data.CompanyRegistrationName = CompanyName
    Data_1.Data.CompanyRegistrationNum = CompanyNum
    Data_1.Data.BalanceRateNum = BalanceRate;
    Data_1.Data.PersonalRateNum = PersonalRate;

    BidHandling_1.BidHandling.BidToJson()
    CalculatePrice_1.CalculatePrice.Calculation()
    BidHandling_1.BidHandling.JsonToBid()
    
    console.log(Data_1.Data);
}

exports.execute = execute;
