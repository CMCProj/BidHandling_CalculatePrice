"use strict";
//"object is possibly 'undefined'"" -> tsconfig.json 으로 이동하여 "strictNullChecks":false를 추가
//decimal, round사용 부분 수정 필요(라이브러리)
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.FillCostAccount = void 0;
var ExcelHandling_1 = require("./ExcelHandling");
var Data_1 = require("./Data");
var fs_1 = __importDefault(require("fs"));
var path_1 = __importDefault(require("path"));
var big_js_1 = __importDefault(require("big.js"));
var FillCostAccount = /** @class */ (function () {
    function FillCostAccount() {
    }
    //원가계산서 항목별 조사금액 채움(관리자 보정 후)
    FillCostAccount.FillInvestigationCosts = function () {
        var costStatementPath = '';
        //원가 계산서 양식 불러오기
        var workbook = ExcelHandling_1.ExcelHandling.GetWorkbook('세부결과_원가계산서.xlsx', '.xlsx');
        var sheet = workbook.GetSheetAt(0);
        //적용비율1 작성
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 7, 6).SetCellValue(Data_1.Data.Rate1.get('간접노무비') + ' %');
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 10, 6).SetCellValue(Data_1.Data.Rate1.get('산재보험료') + ' %');
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 11, 6).SetCellValue(Data_1.Data.Rate1.get('고용보험료') + ' %');
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 19, 6).SetCellValue(Data_1.Data.Rate1.get('환경보전비') + ' %');
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 20, 6).SetCellValue(Data_1.Data.Rate1.get('공사이행보증서발급수수료') + ' %');
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 21, 6).SetCellValue(Data_1.Data.Rate1.get('건설하도급보증수수료') + ' %');
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 22, 6).SetCellValue(Data_1.Data.Rate1.get('건설기계대여대금 지급보증서발급금액') + ' %');
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 23, 6).SetCellValue(Data_1.Data.Rate1.get('기타경비') + ' %');
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 24, 6).SetCellValue(Data_1.Data.Rate1.get('일반관리비') + ' %');
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 29, 6).SetCellValue(Data_1.Data.Rate1.get('공사손해보험료') + ' %');
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 31, 6).SetCellValue(Data_1.Data.Rate1.get('부가가치세') + ' %');
        //적용비율 2 작성
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 7, 7).SetCellValue(Data_1.Data.Rate2.get('간접노무비') + ' %');
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 10, 7).SetCellValue(Data_1.Data.Rate2.get('산재보험료') + ' %');
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 11, 7).SetCellValue(Data_1.Data.Rate2.get('고용보험료') + ' %');
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 23, 7).SetCellValue(Data_1.Data.Rate2.get('기타경비') + ' %');
        //금액 세팅
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 2, 8).SetCellValue(Data_1.Data.Investigation.get('순공사원가')); //1. 순공사원가
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 3, 8).SetCellValue(Data_1.Data.Investigation.get('직접재료비')); //가. 재료비
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 4, 8).SetCellValue(Data_1.Data.Investigation.get('직접재료비')); //가-1. 직접재료비
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 5, 8).SetCellValue(Data_1.Data.Investigation.get('노무비')); //나. 노무비
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 6, 8).SetCellValue(Data_1.Data.Investigation.get('직접노무비')); //나-1. 직접노무비
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 7, 8).SetCellValue(Data_1.Data.Investigation.get('간접노무비')); //나-2. 간접노무비
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 8, 8).SetCellValue(Data_1.Data.Investigation.get('경비')); //다. 경비
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 9, 8).SetCellValue(Data_1.Data.Investigation.get('산출경비')); //다-1. 산출경비
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 10, 8).SetCellValue(Data_1.Data.Investigation.get('산재보험료')); //다-2. 산재보험료
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 11, 8).SetCellValue(Data_1.Data.Investigation.get('고용보험료')); //다-3. 고용보험료
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 12, 8).SetCellValue(Data_1.Data.Fixed.get('국민건강보험료')); //다-4. 국민건강보험료
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 13, 8).SetCellValue(Data_1.Data.Fixed.get('노인장기요양보험')); //다-5. 노인장기요양보험
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 14, 8).SetCellValue(Data_1.Data.Fixed.get('국민연금보험료')); //다-6. 국민연금보험료
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 15, 8).SetCellValue(Data_1.Data.Fixed.get('퇴직공제부금')); //다-7. 퇴직공제부금
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 16, 8).SetCellValue(Data_1.Data.Fixed.get('산업안전보건관리비')); //다-8. 산업안전보건관리비
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 17, 8).SetCellValue(Data_1.Data.Fixed.get('안전관리비')); //다-9. 안전관리비
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 18, 8).SetCellValue(Data_1.Data.Fixed.get('품질관리비')); //다-10. 품질관리비
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 19, 8).SetCellValue(Data_1.Data.Investigation.get('환경보전비')); //다-11. 환경보전비
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 20, 8).SetCellValue(Data_1.Data.Investigation.get('공사이행보증서발급수수료')); //다-12. 공사이행보증수수료
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 21, 8).SetCellValue(Data_1.Data.Investigation.get('건설하도급보증수수료')); //다-13. 하도급대금지급 보증수수료
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 22, 8).SetCellValue(Data_1.Data.Investigation.get('건설기계대여대금 지급보증서발급금액')); //다-14. 건설기계대여대금 지급보증서 발급금액
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 23, 8).SetCellValue(Data_1.Data.Investigation.get('기타경비')); //다-15. 기타경비
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 24, 8).SetCellValue(Data_1.Data.Investigation.get('일반관리비')); //2. 일반관리비
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 25, 8).SetCellValue(Data_1.Data.Investigation.get('이윤')); //3. 이윤
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 26, 8).SetCellValue(Data_1.Data.Investigation.get('PS')); //3.1 PS
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 27, 8).SetCellValue(Data_1.Data.Investigation.get('제요율적용제외공종')); //3.2 제요율적용제외공종
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 28, 8).SetCellValue(Data_1.Data.Investigation.get('총원가')); //4. 총원가
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 29, 8).SetCellValue(Data_1.Data.Investigation.get('공사손해보험료')); //5. 공사손해보험료
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 30, 8).SetCellValue(Data_1.Data.Investigation.get('소계')); //6. 소계
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 31, 8).SetCellValue(Data_1.Data.Investigation.get('부가가치세')); //7. 부가가치세
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 32, 8).SetCellValue(0); //8. 매입세
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 33, 8).SetCellValue(Data_1.Data.Investigation.get('도급비계')); //9. 도급비계
        //원가계산서 조사금액 세팅 시점에 CalculatePrice.cs에서 재계산 시, 초기화를 위한 조사금액 저장
        var FM = Data_1.Data.FixedPriceDirectMaterial;
        var FL = Data_1.Data.FixedPriceDirectLabor;
        var FOE = Data_1.Data.FixedPriceOutputExpense;
        var SM = Data_1.Data.StandardMaterial;
        var SL = Data_1.Data.StandardLabor;
        var SOE = Data_1.Data.StandardExpense;
        Data_1.Data.InvestigateFixedPriceDirectMaterial = FM;
        Data_1.Data.InvestigateFixedPriceDirectLabor = FL;
        Data_1.Data.InvestigateFixedPriceOutputExpense = FOE;
        Data_1.Data.InvestigateStandardMaterial = SM;
        Data_1.Data.InvestigateStandardLabor = SL;
        Data_1.Data.InvestigateStandardExpense = SOE;
        if (fs_1.default.existsSync(Data_1.Data.work_path + '원가계산서.xlsx')) {
            //먼저 기존 원가계산서 파일이 있다면 삭제한다. (23.02.02)
            fs_1.default.unlinkSync(Data_1.Data.work_path + '원가계산서.xlsx');
        }
        costStatementPath = path_1.default.join(Data_1.Data.work_path, '원가계산서.xlsx');
        ExcelHandling_1.ExcelHandling.WriteExcel(workbook, costStatementPath);
    };
    //원가계산서 항목별 입찰금액 채움
    FillCostAccount.FillBiddingCosts = function () {
        //조사금액을 채운 원가계산서_세부결과.xlsx의 경로
        var costStatementPath = path_1.default.join(Data_1.Data.work_path, '원가계산서.xlsx');
        //원가계산서_세부결과 파일 불러오기
        var workbook = ExcelHandling_1.ExcelHandling.GetWorkbook(costStatementPath, '.xlsx');
        var sheet = workbook.GetSheetAt(0);
        //적용비율 1, 2 적용금액 원가계산서 반영
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 7, 9).SetCellValue(Data_1.Data.Bidding.get('간접노무비1'));
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 10, 9).SetCellValue(Data_1.Data.Bidding.get('산재보험료1'));
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 11, 9).SetCellValue(Data_1.Data.Bidding.get('고용보험료1'));
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 23, 9).SetCellValue(Data_1.Data.Bidding.get('기타경비1'));
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 7, 10).SetCellValue(Data_1.Data.Bidding.get('간접노무비2'));
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 10, 10).SetCellValue(Data_1.Data.Bidding.get('산재보험료2'));
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 11, 10).SetCellValue(Data_1.Data.Bidding.get('고용보험료2'));
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 23, 10).SetCellValue(Data_1.Data.Bidding.get('기타경비2'));
        //적용비율 1, 2 적용 금액 중, 큰 금액 세팅
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 7, 11).SetCellValue(Data_1.Data.Bidding.get('간접노무비max'));
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 10, 11).SetCellValue(Data_1.Data.Bidding.get('산재보험료max'));
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 11, 11).SetCellValue(Data_1.Data.Bidding.get('고용보험료max'));
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 23, 11).SetCellValue(Data_1.Data.Bidding.get('기타경비max'));
        //금액 세팅
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 2, 19).SetCellValue(Data_1.Data.Bidding.get('순공사원가')); //1. 순공사원가
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 3, 19).SetCellValue(Data_1.Data.Bidding.get('직접재료비')); //가. 재료비
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 4, 19).SetCellValue(Data_1.Data.Bidding.get('직접재료비')); //가-1. 직접재료비
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 5, 19).SetCellValue(Data_1.Data.Bidding.get('노무비')); //나. 노무비
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 6, 19).SetCellValue(Data_1.Data.Bidding.get('직접노무비')); //나-1. 직접노무비
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 7, 19).SetCellValue(Data_1.Data.Bidding.get('간접노무비')); //나-2. 간접노무비
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 8, 19).SetCellValue(Data_1.Data.Bidding.get('경비')); //다. 경비
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 9, 19).SetCellValue(Data_1.Data.Bidding.get('산출경비')); //다-1. 산출경비
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 10, 19).SetCellValue(Data_1.Data.Bidding.get('산재보험료')); //다-2. 산재보험료
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 11, 19).SetCellValue(Data_1.Data.Bidding.get('고용보험료')); //다-3. 고용보험료
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 12, 19).SetCellValue(Data_1.Data.Fixed.get('국민건강보험료')); //다-4. 국민건강보험료
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 13, 19).SetCellValue(Data_1.Data.Fixed.get('노인장기요양보험')); //다-5. 노인장기요양보험
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 14, 19).SetCellValue(Data_1.Data.Fixed.get('국민연금보험료')); //다-6. 국민연금보험료
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 15, 19).SetCellValue(Data_1.Data.Fixed.get('퇴직공제부금')); //다-7. 퇴직공제부금
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 16, 19).SetCellValue(Data_1.Data.Fixed.get('산업안전보건관리비')); //다-8. 산업안전보건관리비
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 17, 19).SetCellValue(Data_1.Data.Fixed.get('안전관리비')); //다-9. 안전관리비
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 18, 19).SetCellValue(Data_1.Data.Fixed.get('품질관리비')); //다-10. 품질관리비
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 19, 19).SetCellValue(Data_1.Data.Bidding.get('환경보전비')); //다-11. 환경보전비
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 20, 19).SetCellValue(Data_1.Data.Bidding.get('공사이행보증서발급수수료')); //다-12. 공사이행보증수수료
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 21, 19).SetCellValue(Data_1.Data.Bidding.get('건설하도급보증수수료')); //다-13. 하도급대금지급 보증수수료
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 22, 19).SetCellValue(Data_1.Data.Bidding.get('건설기계대여대금 지급보증서발급금액')); //다-14. 건설기계대여대금 지급보증서 발급금액
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 23, 19).SetCellValue(Data_1.Data.Bidding.get('기타경비')); //다-15. 기타경비
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 24, 19).SetCellValue(Data_1.Data.Bidding.get('일반관리비')); //2. 일반관리비
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 25, 19).SetCellValue(Data_1.Data.Bidding.get('이윤')); //3. 이윤
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 26, 19).SetCellValue(Data_1.Data.Bidding.get('PS')); //3.1 PS
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 27, 19).SetCellValue(Data_1.Data.Bidding.get('제요율적용제외공종')); //3.2 제요율적용제외공종
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 28, 19).SetCellValue(Data_1.Data.Bidding.get('총원가')); //4. 총원가
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 29, 19).SetCellValue(Data_1.Data.Bidding.get('공사손해보험료')); //5. 공사손해보험료
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 30, 19).SetCellValue(Data_1.Data.Bidding.get('소계')); //6. 소계
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 31, 19).SetCellValue(Data_1.Data.Bidding.get('부가가치세')); //7. 부가가치세
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 32, 19).SetCellValue(0); //8. 매입세
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 33, 19).SetCellValue(Data_1.Data.Bidding.get('도급비계')); //9. 도급비계
        //비율 세팅
        //C#: FillCostAccount.GetRate() 결과 (double)로 형식 변환
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 4, 20).SetCellValue(FillCostAccount.GetRate('직접재료비') + '%'); //가-1. 직접재료비
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 6, 20).SetCellValue(FillCostAccount.GetRate('직접노무비') + ' %'); //나-1. 직접노무비
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 7, 20).SetCellValue(FillCostAccount.GetRate('간접노무비') + ' %'); //나-2. 간접노무비
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 9, 20).SetCellValue(FillCostAccount.GetRate('산출경비') + ' %'); //다-1. 산출경비
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 10, 20).SetCellValue(FillCostAccount.GetRate('산재보험료') + ' %'); //다-2. 산재보험료
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 11, 20).SetCellValue(FillCostAccount.GetRate('고용보험료') + ' %'); //다-3. 고용보험료
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 12, 20).SetCellValue(FillCostAccount.GetRate('국민건강보험료') + ' %'); //다-4. 국민건강보험료
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 13, 20).SetCellValue(FillCostAccount.GetRate('노인장기요양보험') + ' %'); //다-5. 노인장기요양보험
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 14, 20).SetCellValue(FillCostAccount.GetRate('국민연금보험료') + ' %'); //다-6. 국민연금보험료
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 15, 20).SetCellValue(FillCostAccount.GetRate('퇴직공제부금') + ' %'); //다-7. 퇴직공제부금
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 16, 20).SetCellValue(FillCostAccount.GetRate('산업안전보건관리비') + ' %'); //다-8. 산업안전보건관리비
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 17, 20).SetCellValue(FillCostAccount.GetRate('안전관리비') + ' %'); //다-9. 안전관리비
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 18, 20).SetCellValue(FillCostAccount.GetRate('품질관리비') + ' %'); //다-10. 품질관리비
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 19, 20).SetCellValue(FillCostAccount.GetRate('환경보전비') + ' %'); //다-11. 환경보전비
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 20, 20).SetCellValue(FillCostAccount.GetRate('공사이행보증서발급수수료') + ' %'); //다-12. 공사이행보증수수료
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 21, 20).SetCellValue(FillCostAccount.GetRate('건설하도급보증수수료') + ' %'); //다-13. 하도급대금지급 보증수수료
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 22, 20).SetCellValue(FillCostAccount.GetRate('건설기계대여대금 지급보증서발급금액') + ' %'); //다-14. 건설기계대여대금 지급보증서 발급금액
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 23, 20).SetCellValue(FillCostAccount.GetRate('기타경비') + ' %'); //다-15. 기타경비
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 24, 20).SetCellValue(FillCostAccount.GetRate('일반관리비') + ' %'); //2. 일반관리비
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 25, 20).SetCellValue('0%'); //3. 이윤
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 26, 20).SetCellValue(FillCostAccount.GetRate('PS') + ' %'); //3.1 PS
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 27, 20).SetCellValue(FillCostAccount.GetRate('제요율적용제외공종') + ' %'); //3.2 제요율적용제외공종
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 29, 20).SetCellValue(FillCostAccount.GetRate('공사손해보험료') + ' %'); //5. 공사손해보험료
        ExcelHandling_1.ExcelHandling.GetCell(sheet, 33, 20).SetCellValue(FillCostAccount.GetRate('도급비계') + ' %'); //9. 도급비계
        costStatementPath = path_1.default.join(Data_1.Data.work_path, '원가계산서_세부결과.xlsx');
        ExcelHandling_1.ExcelHandling.WriteExcel(workbook, costStatementPath);
    };
    //원가계산서 항목별 조사금액 구하여 Dictionary Investigation에 저장
    //보정의 경우, 매개변수로 보정할 항목의 이름(item)과 보정할 금액(price)를 받아 값을 적용
    FillCostAccount.CalculateInvestigationCosts = function (correction //Dictionary<string, long>->Map<string, number>
    ) {
        //직공비
        Data_1.Data.Investigation.set('직공비', FillCostAccount.ToLong(Data_1.Data.RealDirectMaterial + Data_1.Data.RealDirectLabor + Data_1.Data.RealOutputExpense));
        //가-1. 직접재료비
        Data_1.Data.Investigation.set('직접재료비', FillCostAccount.ToLong(Data_1.Data.RealDirectMaterial));
        //나-1. 직접노무비
        Data_1.Data.Investigation.set('직접노무비', FillCostAccount.ToLong(Data_1.Data.RealDirectLabor));
        //나-2.간접노무비
        Data_1.Data.Investigation.set('간접노무비', FillCostAccount.ToLong(Data_1.Data.RealDirectLabor * (Data_1.Data.Rate1.get('간접노무비') * 0.01))); //0.01m->0.01
        //나. 노무비
        Data_1.Data.Investigation.set('노무비', FillCostAccount.ToLong(Data_1.Data.RealDirectLabor) + Data_1.Data.Investigation.get('간접노무비'));
        //다-1. 산출경비
        Data_1.Data.Investigation.set('산출경비', FillCostAccount.ToLong(Data_1.Data.RealOutputExpense));
        //다-2. 산재보험료
        Data_1.Data.Investigation.set('산재보험료', FillCostAccount.ToLong(Data_1.Data.Investigation.get('노무비') * (Data_1.Data.Rate1.get('산재보험료') * 0.01))); //0.01m->0.01
        //다-3. 고용보험료
        Data_1.Data.Investigation.set('고용보험료', FillCostAccount.ToLong(Data_1.Data.Investigation.get('노무비') * (Data_1.Data.Rate1.get('고용보험료') * 0.01))); //0.01m->0.01
        //다-11. 환경보전비
        Data_1.Data.Investigation.set('환경보전비', FillCostAccount.ToLong(Data_1.Data.Investigation.get('직공비') * (Data_1.Data.Rate1.get('환경보전비') * 0.01))); //0.01m->0.01
        //다-12 공사이행보증수수료
        Data_1.Data.Investigation.set('공사이행보증서발급수수료', 0);
        if (Data_1.Data.Rate1.get('공사이행보증서발급수수료') != 0)
            Data_1.Data.Investigation.set('공사이행보증서발급수수료', FillCostAccount.GetConstructionGuaranteeFee(Data_1.Data.Investigation.get('직공비')));
        if (correction.has('공사이행보증서발급수수료'))
            Data_1.Data.Investigation.set('공사이행보증서발급수수료', correction.get('공사이행보증서발급수수료'));
        //다-13. 하도급대금지금보증수수료
        Data_1.Data.Investigation.set('건설하도급보증수수료', FillCostAccount.ToLong(Data_1.Data.Investigation.get('직공비') * (Data_1.Data.Rate1.get('건설하도급보증수수료') * 0.01))); //0.01m->0.01
        //다-14. 건설기계대여대금 지급보증서 발급금액
        Data_1.Data.Investigation.set('건설기계대여대금 지급보증서발급금액', FillCostAccount.ToLong(Data_1.Data.Investigation.get('직공비') *
            (Data_1.Data.Rate1.get('건설기계대여대금 지급보증서발급금액') * 0.01))); //0.01m->0.01
        //다-15. 기타경비
        Data_1.Data.Investigation.set('기타경비', FillCostAccount.ToLong((Data_1.Data.Investigation.get('직접재료비') + Data_1.Data.Investigation.get('노무비')) *
            (Data_1.Data.Rate1.get('기타경비') * 0.01))); //0.01m->0.01
        //다. 경비
        Data_1.Data.Investigation.set('경비', Data_1.Data.Investigation.get('산출경비') +
            Data_1.Data.Investigation.get('산재보험료') +
            Data_1.Data.Investigation.get('고용보험료') +
            Data_1.Data.Fixed.get('국민건강보험료') +
            Data_1.Data.Fixed.get('노인장기요양보험') +
            Data_1.Data.Fixed.get('국민연금보험료') +
            Data_1.Data.Fixed.get('퇴직공제부금') +
            Data_1.Data.Fixed.get('산업안전보건관리비') +
            Data_1.Data.Fixed.get('안전관리비') +
            Data_1.Data.Fixed.get('품질관리비') +
            Data_1.Data.Investigation.get('환경보전비') +
            Data_1.Data.Investigation.get('공사이행보증서발급수수료') +
            Data_1.Data.Investigation.get('건설하도급보증수수료') +
            Data_1.Data.Investigation.get('건설기계대여대금 지급보증서발급금액') +
            Data_1.Data.Investigation.get('기타경비'));
        //1. 순공사원가
        Data_1.Data.Investigation.set('순공사원가', Data_1.Data.Investigation.get('직접재료비') +
            Data_1.Data.Investigation.get('노무비') +
            Data_1.Data.Investigation.get('경비'));
        //2. 일반관리비
        Data_1.Data.Investigation.set('일반관리비', FillCostAccount.ToLong(Data_1.Data.Investigation.get('순공사원가') * (Data_1.Data.Rate1.get('일반관리비') * 0.01))); //0.01m->0.01
        //3. 이윤
        Data_1.Data.Investigation.set('이윤', FillCostAccount.ToLong((Data_1.Data.Investigation.get('노무비') +
            Data_1.Data.Investigation.get('경비') +
            Data_1.Data.Investigation.get('일반관리비')) *
            0.12)); //0.12m->0.12
        if (correction.has('이윤'))
            Data_1.Data.Investigation.set('이윤', correction.get('이윤'));
        //3.1 미확정설계공종(PS)
        Data_1.Data.Investigation.set('PS', FillCostAccount.ToLong(Data_1.Data.PsMaterial + Data_1.Data.PsLabor + Data_1.Data.PsExpense));
        //3.2 제요율적용제외공종
        Data_1.Data.Investigation.set('제요율적용제외공종', FillCostAccount.ToLong(Data_1.Data.ExcludingMaterial + Data_1.Data.ExcludingLabor + Data_1.Data.ExcludingExpense));
        var exSum = Data_1.Data.ExcludingMaterial + Data_1.Data.ExcludingLabor + Data_1.Data.ExcludingExpense;
        var exRate2 = new big_js_1.default(exSum).div(Data_1.Data.Investigation.get('직공비')).round(5).toNumber(); //소수점5자리 남기게 반올림
        Data_1.Data.Rate2.set('제요율적용제외공종', exRate2);
        //4. 총원가
        Data_1.Data.Investigation.set('총원가', Data_1.Data.Investigation.get('순공사원가') +
            Data_1.Data.Investigation.get('일반관리비') +
            Data_1.Data.Investigation.get('이윤') +
            Data_1.Data.Investigation.get('PS') +
            Data_1.Data.Investigation.get('제요율적용제외공종'));
        //5. 공사손해보험료
        Data_1.Data.Investigation.set('공사손해보험료', FillCostAccount.ToLong(Data_1.Data.Investigation.get('직공비') * (Data_1.Data.Rate1.get('공사손해보험료') * 0.01)));
        if (correction.has('공사손해보험료'))
            Data_1.Data.Investigation.set('공사손해보험료', correction.get('공사손해보험료'));
        //작업설 추가 (23.02.06)
        Data_1.Data.Investigation.set('작업설 등', FillCostAccount.ToLong(Data_1.Data.ByProduct));
        //6. 소계
        Data_1.Data.Investigation.set('소계', Data_1.Data.Investigation.get('총원가') +
            Data_1.Data.Investigation.get('공사손해보험료') +
            Data_1.Data.Investigation.get('작업설 등')); //전체 가격 계산에 작업설 추가 (23.02.06)
        //7. 부가가치세
        Data_1.Data.Investigation.set('부가가치세', FillCostAccount.ToLong(Data_1.Data.Investigation.get('소계') * (Data_1.Data.Rate1.get('부가가치세') * 0.01)));
        if (correction.has('부가가치세'))
            Data_1.Data.Investigation.set('부가가치세', correction.get('부가가치세'));
        //8. 매입세(입찰 공사 파일 중, 매입세 있는 공사 없음. 추후 추가할 수 있음)
        //9. 도급비계
        Data_1.Data.Investigation.set('도급비계', Data_1.Data.Investigation.get('소계') + Data_1.Data.Investigation.get('부가가치세'));
    };
    //원가계산서 항목별 입찰금액 구하여 Bidding에 저장
    FillCostAccount.CalculateBiddingCosts = function () {
        //직공비
        Data_1.Data.Bidding.set('직공비', FillCostAccount.ToLong(Data_1.Data.RealDirectMaterial + Data_1.Data.RealDirectLabor + Data_1.Data.RealOutputExpense));
        //적용비율 2를 적용한 금액 계산
        var undirectlabor2 = Data_1.Data.Bidding.get('직공비') * (Data_1.Data.Rate2.get('간접노무비') * 0.01);
        var industrial2 = Data_1.Data.Bidding.get('직공비') * (Data_1.Data.Rate2.get('산재보험료') * 0.01);
        var employ2 = Data_1.Data.Bidding.get('직공비') * (Data_1.Data.Rate2.get('고용보험료') * 0.01);
        var etc2 = Data_1.Data.Bidding.get('직공비') * (Data_1.Data.Rate2.get('기타경비') * 0.01);
        //적용비율 2를 적용한 금액 저장
        Data_1.Data.Bidding.set('간접노무비2', FillCostAccount.ToLong(undirectlabor2));
        Data_1.Data.Bidding.set('산재보험료2', FillCostAccount.ToLong(industrial2));
        Data_1.Data.Bidding.set('고용보험료2', FillCostAccount.ToLong(employ2));
        Data_1.Data.Bidding.set('기타경비2', FillCostAccount.ToLong(etc2));
        //가-1. 직접재료비
        Data_1.Data.Bidding.set('직접재료비', FillCostAccount.ToLong(Data_1.Data.RealDirectMaterial));
        //나-1. 직접노무비
        Data_1.Data.Bidding.set('직접노무비', FillCostAccount.ToLong(Data_1.Data.RealDirectLabor));
        //나-2.간접노무비
        Data_1.Data.Bidding.set('간접노무비1', FillCostAccount.ToLong(Data_1.Data.Bidding.get('직접노무비') * (Data_1.Data.Rate1.get('간접노무비') * 0.01)));
        if (Data_1.Data.Bidding.get('간접노무비1') < Data_1.Data.Bidding.get('간접노무비2')) {
            Data_1.Data.Bidding.set('간접노무비', Data_1.Data.Bidding.get('간접노무비2'));
            Data_1.Data.Bidding.set('간접노무비max', Data_1.Data.Bidding.get('간접노무비2'));
            Data_1.Data.Bidding.set('간접노무비before', Data_1.Data.Bidding.get('간접노무비2'));
            if (Data_1.Data.CostAccountDeduction === '1')
                Data_1.Data.Bidding.set('간접노무비', FillCostAccount.ToLong(Math.ceil(undirectlabor2 * 0.997)));
        }
        else {
            Data_1.Data.Bidding.set('간접노무비', Data_1.Data.Bidding.get('간접노무비1'));
            Data_1.Data.Bidding.set('간접노무비max', Data_1.Data.Bidding.get('간접노무비1'));
            Data_1.Data.Bidding.set('간접노무비before', Data_1.Data.Bidding.get('간접노무비1'));
            if (Data_1.Data.CostAccountDeduction === '1')
                Data_1.Data.Bidding.set('간접노무비', FillCostAccount.ToLong(Math.ceil(Data_1.Data.Bidding.get('직접노무비') *
                    (Data_1.Data.Rate1.get('간접노무비') * 0.01) *
                    0.997)));
        }
        //나. 노무비
        Data_1.Data.Bidding.set('노무비', Data_1.Data.Bidding.get('직접노무비') + Data_1.Data.Bidding.get('간접노무비'));
        //다-1. 산출경비
        Data_1.Data.Bidding.set('산출경비', FillCostAccount.ToLong(Data_1.Data.RealOutputExpense));
        //다-2. 산재보험료
        Data_1.Data.Bidding.set('산재보험료1', FillCostAccount.ToLong(Data_1.Data.Bidding.get('노무비') * (Data_1.Data.Rate1.get('산재보험료') * 0.01)));
        if (Data_1.Data.Bidding.get('산재보험료1') < Data_1.Data.Bidding.get('산재보험료2')) {
            Data_1.Data.Bidding.set('산재보험료', Data_1.Data.Bidding.get('산재보험료2'));
            Data_1.Data.Bidding.set('산재보험료max', Data_1.Data.Bidding.get('산재보험료2'));
        }
        else {
            Data_1.Data.Bidding.set('산재보험료', Data_1.Data.Bidding.get('산재보험료1'));
            Data_1.Data.Bidding.set('산재보험료max', Data_1.Data.Bidding.get('산재보험료1'));
        }
        //다-3. 고용보험료
        Data_1.Data.Bidding.set('고용보험료1', FillCostAccount.ToLong(Data_1.Data.Bidding.get('노무비') * (Data_1.Data.Rate1.get('고용보험료') * 0.01)));
        if (Data_1.Data.Bidding.get('고용보험료1') < Data_1.Data.Bidding.get('고용보험료2')) {
            Data_1.Data.Bidding.set('고용보험료', Data_1.Data.Bidding.get('고용보험료2'));
            Data_1.Data.Bidding.set('고용보험료max', Data_1.Data.Bidding.get('고용보험료2'));
        }
        else {
            Data_1.Data.Bidding.set('고용보험료', Data_1.Data.Bidding.get('고용보험료1'));
            Data_1.Data.Bidding.set('고용보험료max', Data_1.Data.Bidding.get('고용보험료1'));
        }
        //다-11. 환경보전비
        Data_1.Data.Bidding.set('환경보전비', FillCostAccount.ToLong(Data_1.Data.Bidding.get('직공비') * (Data_1.Data.Rate1.get('환경보전비') * 0.01)));
        //다-12 공사이행보증수수료
        Data_1.Data.Bidding.set('공사이행보증서발급수수료', 0);
        if (Data_1.Data.Rate1.get('공사이행보증서발급수수료') != 0)
            Data_1.Data.Bidding.set('공사이행보증서발급수수료', FillCostAccount.GetConstructionGuaranteeFee(Data_1.Data.Bidding.get('직공비')));
        //다-13. 하도급대금지금보증수수료
        Data_1.Data.Bidding.set('건설하도급보증수수료', FillCostAccount.ToLong(Data_1.Data.Bidding.get('직공비') * (Data_1.Data.Rate1.get('건설하도급보증수수료') * 0.01)));
        //다-14. 건설기계대여대금 지급보증서 발급금액
        Data_1.Data.Bidding.set('건설기계대여대금 지급보증서발급금액', FillCostAccount.ToLong(Data_1.Data.Bidding.get('직공비') *
            (Data_1.Data.Rate1.get('건설기계대여대금 지급보증서발급금액') * 0.01)));
        //다-15. 기타경비
        Data_1.Data.Bidding.set('기타경비1', FillCostAccount.ToLong((Data_1.Data.Bidding.get('직접재료비') + Data_1.Data.Bidding.get('노무비')) *
            (Data_1.Data.Rate1.get('기타경비') * 0.01)));
        if (Data_1.Data.Bidding.get('기타경비1') < Data_1.Data.Bidding.get('기타경비2')) {
            Data_1.Data.Bidding.set('기타경비', Data_1.Data.Bidding.get('기타경비2'));
            Data_1.Data.Bidding.set('기타경비max', Data_1.Data.Bidding.get('기타경비2'));
            Data_1.Data.Bidding.set('기타경비before', Data_1.Data.Bidding.get('기타경비2'));
            if (Data_1.Data.CostAccountDeduction === '1')
                Data_1.Data.Bidding.set('기타경비', FillCostAccount.ToLong(Math.ceil(etc2 * 0.997)));
        }
        else {
            Data_1.Data.Bidding.set('기타경비', Data_1.Data.Bidding.get('기타경비1'));
            Data_1.Data.Bidding.set('기타경비max', Data_1.Data.Bidding.get('기타경비1'));
            Data_1.Data.Bidding.set('기타경비before', Data_1.Data.Bidding.get('기타경비1'));
            if (Data_1.Data.CostAccountDeduction === '1')
                Data_1.Data.Bidding.set('기타경비', FillCostAccount.ToLong(Math.ceil((Data_1.Data.Bidding.get('직접재료비') + Data_1.Data.Bidding.get('노무비')) *
                    (Data_1.Data.Rate1.get('기타경비') * 0.01) *
                    0.997))); //decimal 0.01m 0.997m
        }
        //다. 경비
        Data_1.Data.Bidding.set('경비', Data_1.Data.Bidding.get('산출경비') +
            Data_1.Data.Bidding.get('산재보험료') +
            Data_1.Data.Bidding.get('고용보험료') +
            Data_1.Data.Fixed.get('국민건강보험료') +
            Data_1.Data.Fixed.get('노인장기요양보험') +
            Data_1.Data.Fixed.get('국민연금보험료') +
            Data_1.Data.Fixed.get('퇴직공제부금') +
            Data_1.Data.Fixed.get('산업안전보건관리비') +
            Data_1.Data.Fixed.get('안전관리비') +
            Data_1.Data.Fixed.get('품질관리비') +
            Data_1.Data.Bidding.get('환경보전비') +
            Data_1.Data.Bidding.get('공사이행보증서발급수수료') +
            Data_1.Data.Bidding.get('건설하도급보증수수료') +
            Data_1.Data.Bidding.get('건설기계대여대금 지급보증서발급금액') +
            Data_1.Data.Bidding.get('기타경비'));
        //1. 순공사원가
        Data_1.Data.Bidding.set('순공사원가', Data_1.Data.Bidding.get('직접재료비') + Data_1.Data.Bidding.get('노무비') + Data_1.Data.Bidding.get('경비'));
        //2. 일반관리비
        Data_1.Data.Bidding.set('일반관리비', FillCostAccount.ToLong(Data_1.Data.Bidding.get('순공사원가') * (Data_1.Data.Rate1.get('일반관리비') * 0.01))); //decimal 0.01m
        Data_1.Data.Bidding.set('일반관리비before', FillCostAccount.ToLong(Data_1.Data.Bidding.get('순공사원가') * (Data_1.Data.Rate1.get('일반관리비') * 0.01))); //decimal 0.01m
        if (Data_1.Data.CostAccountDeduction === '1') {
            Data_1.Data.Bidding.set('일반관리비before', FillCostAccount.ToLong(Data_1.Data.Bidding.get('순공사원가') * (Data_1.Data.Rate1.get('일반관리비') * 0.01))); //decimal 0.01m
            Data_1.Data.Bidding.set('일반관리비', FillCostAccount.ToLong(Math.ceil(Data_1.Data.Bidding.get('순공사원가') *
                (Data_1.Data.Rate1.get('일반관리비') * 0.01) *
                0.997))); //decimal 0.01m 0.997m
        }
        //3. 이윤(이윤의 입찰 금액은 0원)
        Data_1.Data.Bidding.set('이윤', 0);
        //3.1 미확정설계공종(PS)
        Data_1.Data.Bidding.set('PS', FillCostAccount.ToLong(Data_1.Data.PsMaterial + Data_1.Data.PsLabor + Data_1.Data.PsExpense));
        //3.2 제요율적용제외공종
        Data_1.Data.Bidding.set('제요율적용제외공종', FillCostAccount.ToLong(Data_1.Data.AdjustedExMaterial + Data_1.Data.AdjustedExLabor + Data_1.Data.AdjustedExExpense));
        //4. 총원가
        Data_1.Data.Bidding.set('총원가', Data_1.Data.Bidding.get('순공사원가') +
            Data_1.Data.Bidding.get('일반관리비') +
            Data_1.Data.Bidding.get('이윤') +
            Data_1.Data.Bidding.get('PS') +
            Data_1.Data.Bidding.get('제요율적용제외공종'));
        //5. 공사손해보험료
        Data_1.Data.Bidding.set('공사손해보험료', FillCostAccount.ToLong(Data_1.Data.Bidding.get('직공비') * (Data_1.Data.Rate1.get('공사손해보험료') * 0.01))); //decimal 0.01m
        if (Data_1.Data.CostAccountDeduction === '1') {
            Data_1.Data.Bidding.set('공사손해보험료before', FillCostAccount.ToLong(Data_1.Data.Bidding.get('직공비') * (Data_1.Data.Rate1.get('공사손해보험료') * 0.01))); //decimal 0.01m
            Data_1.Data.Bidding.set('공사손해보험료', FillCostAccount.ToLong(Math.ceil(Data_1.Data.Bidding.get('직공비') * Data_1.Data.Rate1.get('공사손해보험료') * 0.01 * 0.997))); //decimal 0.01m 0.997m
        }
        //작업설 추가 (23.02.06)
        Data_1.Data.Bidding.set('작업설 등', FillCostAccount.ToLong(Data_1.Data.ByProduct));
        //6. 소계
        Data_1.Data.Bidding.set('소계', Data_1.Data.Bidding.get('총원가') +
            Data_1.Data.Bidding.get('공사손해보험료') +
            Data_1.Data.Bidding.get('작업설 등')); //전체 가격 계산에 작업설 추가 (23.02.06)
        //7. 부가가치세
        Data_1.Data.Bidding.set('부가가치세', FillCostAccount.ToLong(Data_1.Data.Bidding.get('소계') * (Data_1.Data.Rate1.get('부가가치세') * 0.01))); //decimal 0.01m
        //8. 매입세(입찰 공사 파일 중, 매입세 있는 공사 없음. 추후 추가할 수 있음)
        //9. 도급비계
        Data_1.Data.Bidding.set('도급비계', Data_1.Data.Bidding.get('소계') + Data_1.Data.Bidding.get('부가가치세'));
        //도급비계 1000원 단위 절상 옵션 적용시
        if (Data_1.Data.BidPriceRaise === '1') {
            var raise = 1000 - (Data_1.Data.Bidding.get('도급비계') % 1000); //1000원 단위 절상 //Convert.ToDecimal(Data.Bidding.get("도급비계")) % 1000)
            var addPrice = raise / 1.1; //decimal 1.1m
            Data_1.Data.Bidding.set('도급비계', FillCostAccount.ToLong(Data_1.Data.Bidding.get('도급비계') + raise));
            Data_1.Data.Bidding.set('일반관리비', FillCostAccount.ToLong(Data_1.Data.Bidding.get('일반관리비') + addPrice)); //절상에 필요한 가격을 일반관리비에 더해 금액을 맞추어줌
            //FillCostAccount.ToLong(Convert.ToDecimal(Data.Bidding.get("일반관리비")) + addPrice)
            //일반관리비 증가에 따른 타 금액 변경
            Data_1.Data.Bidding.set('소계', FillCostAccount.ToLong(Data_1.Data.Bidding.get('소계') + addPrice)); //FillCostAccount.ToLong(Convert.ToDecimal(Data.Bidding.get("소계")) + addPrice)
            Data_1.Data.Bidding.set('부가가치세', FillCostAccount.ToLong(Data_1.Data.Bidding.get('소계') * (Data_1.Data.Rate1.get('부가가치세') * 0.01))); //decimal 0.01m
            //계산된 도급비계 금액이 천원 단위가 아닐 경우, 부가세 조정
            var difference = Data_1.Data.Bidding.get('도급비계') -
                (Data_1.Data.Bidding.get('소계') + Data_1.Data.Bidding.get('부가가치세'));
            Data_1.Data.Bidding.set('부가가치세', FillCostAccount.ToLong(Data_1.Data.Bidding.get('부가가치세') + difference));
            //Console.WriteLine("차이 : " + difference);
        }
    };
    //decimal 금액 원 단위 절사
    FillCostAccount.ToLong = function (price //price: decimal
    ) {
        var bigNum = new big_js_1.default(price).round(0, 0);
        return bigNum.toNumber();
    };
    //공사이행보증서발급수수료 금액 계산 후 반환
    FillCostAccount.GetConstructionGuaranteeFee = function (directSum //long directSum //수정시 자료형 재확인
    ) {
        var guaranteeFee = 0; //long
        var rate = Data_1.Data.Rate1.get('공사이행보증서발급수수료') * 0.01; //decimal 0.01m
        var term = Data_1.Data.ConstructionTerm / 365.0; //decimal  365.0m;
        if (directSum < 7000000000)
            guaranteeFee = FillCostAccount.ToLong(directSum * rate * term);
        else if (directSum < 12000000000)
            guaranteeFee = FillCostAccount.ToLong(((directSum - 5000000000) * rate + 2000000) * term);
        else if (directSum < 25000000000)
            guaranteeFee = FillCostAccount.ToLong(((directSum - 14000000000) * rate + 4000000) * term);
        else if (directSum < 50000000000)
            guaranteeFee = FillCostAccount.ToLong(((directSum - 25000000000) * rate + 6000000) * term);
        else
            guaranteeFee = FillCostAccount.ToLong(((directSum - 50000000000) * rate + 10000000) * term);
        return guaranteeFee;
    };
    //입찰 금액의 조사금액 대 비율 저장
    FillCostAccount.GetRate = function (item) {
        if (Data_1.Data.Fixed.has(item))
            return 100;
        if (Data_1.Data.Investigation.get(item) == 0 && Data_1.Data.Bidding.get(item) == 0)
            return 100;
        //원가계산제경비 옵션 적용 항목은 적용 전, 후의 비율 출력
        if (item === '간접노무비' ||
            item === '기타경비' ||
            item === '일반관리비' ||
            item === '공사손해보험료') {
            var before = item + 'before';
            return new big_js_1.default(Data_1.Data.Bidding[item])
                .div(Data_1.Data.Bidding[before])
                .round(7)
                .times(100)
                .toNumber();
        }
        var rate = new big_js_1.default(Data_1.Data.Bidding[item])
            .div(Data_1.Data.Investigation[item])
            .round(7)
            .times(100)
            .toNumber();
        return rate.toNumber();
    };
    //해당 공사에 특정 원가계산서 항목이 존재하지 않는 경우
    FillCostAccount.CheckKeyNotFound = function () {
        if (!Data_1.Data.Rate1.has('간접노무비'))
            Data_1.Data.Rate1.set('간접노무비', 0);
        if (!Data_1.Data.Rate1.has('산재보험료'))
            Data_1.Data.Rate1.set('산재보험료', 0);
        if (!Data_1.Data.Rate1.has('고용보험료'))
            Data_1.Data.Rate1.set('고용보험료', 0);
        if (!Data_1.Data.Rate1.has('환경보전비'))
            Data_1.Data.Rate1.set('환경보전비', 0);
        if (!Data_1.Data.Rate1.has('공사이행보증서발급수수료'))
            Data_1.Data.Rate1.set('공사이행보증서발급수수료', 0);
        if (!Data_1.Data.Rate1.has('건설하도급보증수수료'))
            Data_1.Data.Rate1.set('건설하도급보증수수료', 0);
        if (!Data_1.Data.Rate1.has('건설기계대여대금 지급보증서발급금액'))
            Data_1.Data.Rate1.set('건설기계대여대금 지급보증서발급금액', 0);
        if (!Data_1.Data.Rate1.has('기타경비'))
            Data_1.Data.Rate1.set('기타경비', 0);
        if (!Data_1.Data.Rate1.has('일반관리비'))
            Data_1.Data.Rate1.set('일반관리비', 0);
        if (!Data_1.Data.Rate1.has('부가가치세'))
            Data_1.Data.Rate1.set('부가가치세', 0);
        if (!Data_1.Data.Rate1.has('공사손해보험료'))
            Data_1.Data.Rate1.set('공사손해보험료', 0);
        if (!Data_1.Data.Fixed.has('국민건강보험료'))
            Data_1.Data.Fixed.set('국민건강보험료', 0);
        if (!Data_1.Data.Fixed.has('노인장기요양보험'))
            Data_1.Data.Fixed.set('노인장기요양보험', 0);
        if (!Data_1.Data.Fixed.has('국민연금보험료'))
            Data_1.Data.Fixed.set('국민연금보험료', 0);
        if (!Data_1.Data.Fixed.has('퇴직공제부금'))
            Data_1.Data.Fixed.set('퇴직공제부금', 0);
        if (!Data_1.Data.Fixed.has('산업안전보건관리비'))
            Data_1.Data.Fixed.set('산업안전보건관리비', 0);
        if (!Data_1.Data.Fixed.has('안전관리비'))
            Data_1.Data.Fixed.set('안전관리비', 0);
        if (!Data_1.Data.Fixed.has('품질관리비'))
            Data_1.Data.Fixed.set('품질관리비', 0);
        if (!Data_1.Data.Rate2.has('간접노무비'))
            Data_1.Data.Rate2.set('간접노무비', 0);
        if (!Data_1.Data.Rate2.has('산재보험료'))
            Data_1.Data.Rate2.set('산재보험료', 0);
        if (!Data_1.Data.Rate2.has('고용보험료'))
            Data_1.Data.Rate2.set('고용보험료', 0);
        if (!Data_1.Data.Rate2.has('기타경비'))
            Data_1.Data.Rate2.set('기타경비', 0);
    };
    return FillCostAccount;
}());
exports.FillCostAccount = FillCostAccount;
