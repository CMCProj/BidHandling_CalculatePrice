//"object is possibly 'undefined'"" -> tsconfig.json 으로 이동하여 "strictNullChecks":false를 추가
//decimal, round사용 부분 수정 필요(라이브러리)

import { ExcelHandling } from './ExcelHandling'
import { Data } from './Data'
import * as fs from 'fs'
import * as path from 'path'

export class FillCostAccount {
    //원가계산서 항목별 조사금액 채움(관리자 보정 후)
    public static async FillInvestigationCosts() {
        let costStatementPath: string = ''
        //원가 계산서 양식 불러오기
        let workbook = await ExcelHandling.GetWorkbook('세부결과_원가계산서.xlsx', '.xlsx')
        let sheet = workbook.getWorksheet(0);
        //적용비율1 작성
        ExcelHandling.GetCell(sheet, 7, 6).value = Data.Rate1.get('간접노무비') + ' %';
        ExcelHandling.GetCell(sheet, 7, 6).value = Data.Rate1.get('간접노무비') + ' %';
        ExcelHandling.GetCell(sheet, 10, 6).value = Data.Rate1.get('산재보험료') + ' %';
        ExcelHandling.GetCell(sheet, 11, 6).value = Data.Rate1.get('고용보험료') + ' %';
        ExcelHandling.GetCell(sheet, 19, 6).value = Data.Rate1.get('환경보전비') + ' %';
        ExcelHandling.GetCell(sheet, 20, 6).value = Data.Rate1.get('공사이행보증서발급수수료') + ' %';
        ExcelHandling.GetCell(sheet, 21, 6).value = Data.Rate1.get('건설하도급보증수수료') + ' %';
        ExcelHandling.GetCell(sheet, 22, 6).value = Data.Rate1.get('건설기계대여대금 지급보증서발급금액') + ' %';
        ExcelHandling.GetCell(sheet, 23, 6).value = Data.Rate1.get('기타경비') + ' %';
        ExcelHandling.GetCell(sheet, 24, 6).value = Data.Rate1.get('일반관리비') + ' %';
        ExcelHandling.GetCell(sheet, 29, 6).value = Data.Rate1.get('공사손해보험료') + ' %';
        ExcelHandling.GetCell(sheet, 31, 6).value = Data.Rate1.get('부가가치세') + ' %';

        //적용비율 2 작성
        ExcelHandling.GetCell(sheet, 7, 7).value = Data.Rate2.get('간접노무비') + ' %'
        ExcelHandling.GetCell(sheet, 10, 7).value = Data.Rate2.get('산재보험료') + ' %'
        ExcelHandling.GetCell(sheet, 11, 7).value = Data.Rate2.get('고용보험료') + ' %'
        ExcelHandling.GetCell(sheet, 23, 7).value = Data.Rate2.get('기타경비') + ' %'

        //금액 세팅
        ExcelHandling.GetCell(sheet, 2, 8).value = Data.Investigation.get('순공사원가') //1. 순공사원가
        ExcelHandling.GetCell(sheet, 3, 8).value = Data.Investigation.get('직접재료비') //가. 재료비
        ExcelHandling.GetCell(sheet, 4, 8).value = Data.Investigation.get('직접재료비') //가-1. 직접재료비
        ExcelHandling.GetCell(sheet, 5, 8).value = Data.Investigation.get('노무비') //나. 노무비
        ExcelHandling.GetCell(sheet, 6, 8).value = Data.Investigation.get('직접노무비') //나-1. 직접노무비
        ExcelHandling.GetCell(sheet, 7, 8).value = Data.Investigation.get('간접노무비') //나-2. 간접노무비
        ExcelHandling.GetCell(sheet, 8, 8).value = Data.Investigation.get('경비') //다. 경비
        ExcelHandling.GetCell(sheet, 9, 8).value = Data.Investigation.get('산출경비') //다-1. 산출경비
        ExcelHandling.GetCell(sheet, 10, 8).value = Data.Investigation.get('산재보험료') //다-2. 산재보험료
        ExcelHandling.GetCell(sheet, 11, 8).value = Data.Investigation.get('고용보험료') //다-3. 고용보험료
        ExcelHandling.GetCell(sheet, 12, 8).value = Data.Fixed.get('국민건강보험료') //다-4. 국민건강보험료
        ExcelHandling.GetCell(sheet, 13, 8).value = Data.Fixed.get('노인장기요양보험') //다-5. 노인장기요양보험
        ExcelHandling.GetCell(sheet, 14, 8).value = Data.Fixed.get('국민연금보험료') //다-6. 국민연금보험료
        ExcelHandling.GetCell(sheet, 15, 8).value = Data.Fixed.get('퇴직공제부금') //다-7. 퇴직공제부금
        ExcelHandling.GetCell(sheet, 16, 8).value = Data.Fixed.get('산업안전보건관리비') //다-8. 산업안전보건관리비
        ExcelHandling.GetCell(sheet, 17, 8).value = Data.Fixed.get('안전관리비') //다-9. 안전관리비
        ExcelHandling.GetCell(sheet, 18, 8).value = Data.Fixed.get('품질관리비') //다-10. 품질관리비
        ExcelHandling.GetCell(sheet, 19, 8).value = Data.Investigation.get('환경보전비') //다-11. 환경보전비
        ExcelHandling.GetCell(sheet, 20, 8).value = Data.Investigation.get('공사이행보증서발급수수료') //다-12. 공사이행보증수수료
        ExcelHandling.GetCell(sheet, 21, 8).value = Data.Investigation.get('건설하도급보증수수료') //다-13. 하도급대금지급 보증수수료
        ExcelHandling.GetCell(sheet, 22, 8).value = Data.Investigation.get('건설기계대여대금 지급보증서발급금액') //다-14. 건설기계대여대금 지급보증서 발급금액
        ExcelHandling.GetCell(sheet, 23, 8).value = Data.Investigation.get('기타경비') //다-15. 기타경비
        ExcelHandling.GetCell(sheet, 24, 8).value = Data.Investigation.get('일반관리비') //2. 일반관리비
        ExcelHandling.GetCell(sheet, 25, 8).value = Data.Investigation.get('이윤') //3. 이윤
        ExcelHandling.GetCell(sheet, 26, 8).value = Data.Investigation.get('PS') //3.1 PS
        ExcelHandling.GetCell(sheet, 27, 8).value = Data.Investigation.get('제요율적용제외공종') //3.2 제요율적용제외공종
        ExcelHandling.GetCell(sheet, 28, 8).value = Data.Investigation.get('총원가') //4. 총원가
        ExcelHandling.GetCell(sheet, 29, 8).value = Data.Investigation.get('공사손해보험료') //5. 공사손해보험료
        ExcelHandling.GetCell(sheet, 30, 8).value = Data.Investigation.get('소계') //6. 소계
        ExcelHandling.GetCell(sheet, 31, 8).value = Data.Investigation.get('부가가치세') //7. 부가가치세
        ExcelHandling.GetCell(sheet, 32, 8).value = 0 //8. 매입세
        ExcelHandling.GetCell(sheet, 33, 8).value = Data.Investigation.get('도급비계') //9. 도급비계

        //원가계산서 조사금액 세팅 시점에 CalculatePrice.cs에서 재계산 시, 초기화를 위한 조사금액 저장
        let FM = Data.FixedPriceDirectMaterial
        let FL = Data.FixedPriceDirectLabor
        let FOE = Data.FixedPriceOutputExpense
        let SM = Data.StandardMaterial
        let SL = Data.StandardLabor
        let SOE = Data.StandardExpense

        Data.InvestigateFixedPriceDirectMaterial = FM
        Data.InvestigateFixedPriceDirectLabor = FL
        Data.InvestigateFixedPriceOutputExpense = FOE
        Data.InvestigateStandardMaterial = SM
        Data.InvestigateStandardLabor = SL
        Data.InvestigateStandardExpense = SOE

        if (fs.existsSync(Data.work_path + '원가계산서.xlsx')) {
            //먼저 기존 원가계산서 파일이 있다면 삭제한다. (23.02.02)
            fs.unlinkSync(Data.work_path + '원가계산서.xlsx')
        }

        costStatementPath = path.join(Data.work_path, '원가계산서.xlsx')
        ExcelHandling.WriteExcel(workbook, costStatementPath)
    }

    //원가계산서 항목별 입찰금액 채움
    public static async FillBiddingCosts() {
        //조사금액을 채운 원가계산서_세부결과.xlsx의 경로
        let costStatementPath: string = path.join(Data.work_path, '원가계산서.xlsx')
        //원가계산서_세부결과 파일 불러오기
        let workbook = await ExcelHandling.GetWorkbook(costStatementPath, '.xlsx')

        let sheet = workbook.getWorksheet(0)

        //적용비율 1, 2 적용금액 원가계산서 반영
        ExcelHandling.GetCell(sheet, 7, 9).value = Data.Bidding.get('간접노무비1')
        ExcelHandling.GetCell(sheet, 10, 9).value = Data.Bidding.get('산재보험료1')
        ExcelHandling.GetCell(sheet, 11, 9).value = Data.Bidding.get('고용보험료1')
        ExcelHandling.GetCell(sheet, 23, 9).value = Data.Bidding.get('기타경비1')
        ExcelHandling.GetCell(sheet, 7, 10).value = Data.Bidding.get('간접노무비2')
        ExcelHandling.GetCell(sheet, 10, 10).value = Data.Bidding.get('산재보험료2')
        ExcelHandling.GetCell(sheet, 11, 10).value = Data.Bidding.get('고용보험료2')
        ExcelHandling.GetCell(sheet, 23, 10).value = Data.Bidding.get('기타경비2')

        //적용비율 1, 2 적용 금액 중, 큰 금액 세팅
        ExcelHandling.GetCell(sheet, 7, 11).value = Data.Bidding.get('간접노무비max')
        ExcelHandling.GetCell(sheet, 10, 11).value = Data.Bidding.get('산재보험료max')
        ExcelHandling.GetCell(sheet, 11, 11).value = Data.Bidding.get('고용보험료max')
        ExcelHandling.GetCell(sheet, 23, 11).value = Data.Bidding.get('기타경비max')

        //금액 세팅
        ExcelHandling.GetCell(sheet, 2, 19).value = Data.Bidding.get('순공사원가') //1. 순공사원가
        ExcelHandling.GetCell(sheet, 3, 19).value = Data.Bidding.get('직접재료비') //가. 재료비
        ExcelHandling.GetCell(sheet, 4, 19).value = Data.Bidding.get('직접재료비') //가-1. 직접재료비
        ExcelHandling.GetCell(sheet, 5, 19).value = Data.Bidding.get('노무비') //나. 노무비
        ExcelHandling.GetCell(sheet, 6, 19).value = Data.Bidding.get('직접노무비') //나-1. 직접노무비
        ExcelHandling.GetCell(sheet, 7, 19).value = Data.Bidding.get('간접노무비') //나-2. 간접노무비
        ExcelHandling.GetCell(sheet, 8, 19).value = Data.Bidding.get('경비') //다. 경비
        ExcelHandling.GetCell(sheet, 9, 19).value = Data.Bidding.get('산출경비') //다-1. 산출경비
        ExcelHandling.GetCell(sheet, 10, 19).value = Data.Bidding.get('산재보험료') //다-2. 산재보험료
        ExcelHandling.GetCell(sheet, 11, 19).value = Data.Bidding.get('고용보험료') //다-3. 고용보험료
        ExcelHandling.GetCell(sheet, 12, 19).value = Data.Fixed.get('국민건강보험료') //다-4. 국민건강보험료
        ExcelHandling.GetCell(sheet, 13, 19).value = Data.Fixed.get('노인장기요양보험') //다-5. 노인장기요양보험
        ExcelHandling.GetCell(sheet, 14, 19).value = Data.Fixed.get('국민연금보험료') //다-6. 국민연금보험료
        ExcelHandling.GetCell(sheet, 15, 19).value = Data.Fixed.get('퇴직공제부금') //다-7. 퇴직공제부금
        ExcelHandling.GetCell(sheet, 16, 19).value = Data.Fixed.get('산업안전보건관리비') //다-8. 산업안전보건관리비
        ExcelHandling.GetCell(sheet, 17, 19).value = Data.Fixed.get('안전관리비') //다-9. 안전관리비
        ExcelHandling.GetCell(sheet, 18, 19).value = Data.Fixed.get('품질관리비') //다-10. 품질관리비
        ExcelHandling.GetCell(sheet, 19, 19).value = Data.Bidding.get('환경보전비') //다-11. 환경보전비
        ExcelHandling.GetCell(sheet, 20, 19).value = Data.Bidding.get('공사이행보증서발급수수료') //다-12. 공사이행보증수수료
        ExcelHandling.GetCell(sheet, 21, 19).value = Data.Bidding.get('건설하도급보증수수료') //다-13. 하도급대금지급 보증수수료
        ExcelHandling.GetCell(sheet, 22, 19).value = Data.Bidding.get('건설기계대여대금 지급보증서발급금액') //다-14. 건설기계대여대금 지급보증서 발급금액
        ExcelHandling.GetCell(sheet, 23, 19).value = Data.Bidding.get('기타경비') //다-15. 기타경비
        ExcelHandling.GetCell(sheet, 24, 19).value = Data.Bidding.get('일반관리비') //2. 일반관리비
        ExcelHandling.GetCell(sheet, 25, 19).value = Data.Bidding.get('이윤') //3. 이윤
        ExcelHandling.GetCell(sheet, 26, 19).value = Data.Bidding.get('PS') //3.1 PS
        ExcelHandling.GetCell(sheet, 27, 19).value = Data.Bidding.get('제요율적용제외공종') //3.2 제요율적용제외공종
        ExcelHandling.GetCell(sheet, 28, 19).value = Data.Bidding.get('총원가') //4. 총원가
        ExcelHandling.GetCell(sheet, 29, 19).value = Data.Bidding.get('공사손해보험료') //5. 공사손해보험료
        ExcelHandling.GetCell(sheet, 30, 19).value = Data.Bidding.get('소계') //6. 소계
        ExcelHandling.GetCell(sheet, 31, 19).value = Data.Bidding.get('부가가치세') //7. 부가가치세
        ExcelHandling.GetCell(sheet, 32, 19).value = 0 //8. 매입세
        ExcelHandling.GetCell(sheet, 33, 19).value = Data.Bidding.get('도급비계') //9. 도급비계

        //비율 세팅
        //C#: FillCostAccount.GetRate() 결과 (double)로 형식 변환
        ExcelHandling.GetCell(sheet, 4, 20).value = FillCostAccount.GetRate('직접재료비') + '%' //가-1. 직접재료비
        ExcelHandling.GetCell(sheet, 6, 20).value = FillCostAccount.GetRate('직접노무비') + ' %' //나-1. 직접노무비
        ExcelHandling.GetCell(sheet, 7, 20).value = FillCostAccount.GetRate('간접노무비') + ' %' //나-2. 간접노무비
        ExcelHandling.GetCell(sheet, 9, 20).value = FillCostAccount.GetRate('산출경비') + ' %' //다-1. 산출경비
        ExcelHandling.GetCell(sheet, 10, 20).value = FillCostAccount.GetRate('산재보험료') + ' %' //다-2. 산재보험료
        ExcelHandling.GetCell(sheet, 11, 20).value = FillCostAccount.GetRate('고용보험료') + ' %' //다-3. 고용보험료
        ExcelHandling.GetCell(sheet, 12, 20).value = FillCostAccount.GetRate('국민건강보험료') + ' %' //다-4. 국민건강보험료
        ExcelHandling.GetCell(sheet, 13, 20).value = FillCostAccount.GetRate('노인장기요양보험') + ' %' //다-5. 노인장기요양보험
        ExcelHandling.GetCell(sheet, 14, 20).value = FillCostAccount.GetRate('국민연금보험료') + ' %' //다-6. 국민연금보험료
        ExcelHandling.GetCell(sheet, 15, 20).value = FillCostAccount.GetRate('퇴직공제부금') + ' %' //다-7. 퇴직공제부금
        ExcelHandling.GetCell(sheet, 16, 20).value = FillCostAccount.GetRate('산업안전보건관리비') + ' %' //다-8. 산업안전보건관리비
        ExcelHandling.GetCell(sheet, 17, 20).value = FillCostAccount.GetRate('안전관리비') + ' %' //다-9. 안전관리비
        ExcelHandling.GetCell(sheet, 18, 20).value = FillCostAccount.GetRate('품질관리비') + ' %' //다-10. 품질관리비
        ExcelHandling.GetCell(sheet, 19, 20).value = FillCostAccount.GetRate('환경보전비') + ' %' //다-11. 환경보전비
        ExcelHandling.GetCell(sheet, 20, 20).value = FillCostAccount.GetRate('공사이행보증서발급수수료') + ' %' //다-12. 공사이행보증수수료
        ExcelHandling.GetCell(sheet, 21, 20).value = FillCostAccount.GetRate('건설하도급보증수수료') + ' %' //다-13. 하도급대금지급 보증수수료
        ExcelHandling.GetCell(sheet, 22, 20).value = FillCostAccount.GetRate('건설기계대여대금 지급보증서발급금액') + ' %' //다-14. 건설기계대여대금 지급보증서 발급금액
        ExcelHandling.GetCell(sheet, 23, 20).value = FillCostAccount.GetRate('기타경비') + ' %' //다-15. 기타경비
        ExcelHandling.GetCell(sheet, 24, 20).value = FillCostAccount.GetRate('일반관리비') + ' %' //2. 일반관리비
        ExcelHandling.GetCell(sheet, 25, 20).value = '0%' //3. 이윤
        ExcelHandling.GetCell(sheet, 26, 20).value = FillCostAccount.GetRate('PS') + ' %' //3.1 PS
        ExcelHandling.GetCell(sheet, 27, 20).value = FillCostAccount.GetRate('제요율적용제외공종') + ' %' //3.2 제요율적용제외공종
        ExcelHandling.GetCell(sheet, 29, 20).value = FillCostAccount.GetRate('공사손해보험료') + ' %' //5. 공사손해보험료
        ExcelHandling.GetCell(sheet, 33, 20).value = FillCostAccount.GetRate('도급비계') + ' %' //9. 도급비계

        costStatementPath = path.join(Data.work_path, '원가계산서_세부결과.xlsx')
        ExcelHandling.WriteExcel(workbook, costStatementPath)
    }

    //원가계산서 항목별 조사금액 구하여 Dictionary Investigation에 저장
    //보정의 경우, 매개변수로 보정할 항목의 이름(item)과 보정할 금액(price)를 받아 값을 적용
    public static CalculateInvestigationCosts(
        correction: Map<string, number> //Dictionary<string, long>->Map<string, number>
    ) {
        //직공비
        Data.Investigation.set(
            '직공비',
            FillCostAccount.ToLong(
                Data.RealDirectMaterial + Data.RealDirectLabor + Data.RealOutputExpense
            )
        )
        //가-1. 직접재료비
        Data.Investigation.set('직접재료비', FillCostAccount.ToLong(Data.RealDirectMaterial))
        //나-1. 직접노무비
        Data.Investigation.set('직접노무비', FillCostAccount.ToLong(Data.RealDirectLabor))
        //나-2.간접노무비
        Data.Investigation.set(
            '간접노무비',
            FillCostAccount.ToLong(Data.RealDirectLabor * (Data.Rate1.get('간접노무비') * 0.01))
        ) //0.01m->0.01
        //나. 노무비
        Data.Investigation.set(
            '노무비',
            FillCostAccount.ToLong(Data.RealDirectLabor) + Data.Investigation.get('간접노무비')
        )
        //다-1. 산출경비
        Data.Investigation.set('산출경비', FillCostAccount.ToLong(Data.RealOutputExpense))
        //다-2. 산재보험료
        Data.Investigation.set(
            '산재보험료',
            FillCostAccount.ToLong(
                Data.Investigation.get('노무비') * (Data.Rate1.get('산재보험료') * 0.01)
            )
        ) //0.01m->0.01
        //다-3. 고용보험료
        Data.Investigation.set(
            '고용보험료',
            FillCostAccount.ToLong(
                Data.Investigation.get('노무비') * (Data.Rate1.get('고용보험료') * 0.01)
            )
        ) //0.01m->0.01
        //다-11. 환경보전비
        Data.Investigation.set(
            '환경보전비',
            FillCostAccount.ToLong(
                Data.Investigation.get('직공비') * (Data.Rate1.get('환경보전비') * 0.01)
            )
        ) //0.01m->0.01
        //다-12 공사이행보증수수료
        Data.Investigation.set('공사이행보증서발급수수료', 0)
        if (Data.Rate1.get('공사이행보증서발급수수료') != 0)
            Data.Investigation.set(
                '공사이행보증서발급수수료',
                FillCostAccount.GetConstructionGuaranteeFee(Data.Investigation.get('직공비'))
            )
        if (correction.has('공사이행보증서발급수수료'))
            Data.Investigation.set(
                '공사이행보증서발급수수료',
                correction.get('공사이행보증서발급수수료')
            )
        //다-13. 하도급대금지금보증수수료
        Data.Investigation.set(
            '건설하도급보증수수료',
            FillCostAccount.ToLong(
                Data.Investigation.get('직공비') * (Data.Rate1.get('건설하도급보증수수료') * 0.01)
            )
        ) //0.01m->0.01
        //다-14. 건설기계대여대금 지급보증서 발급금액
        Data.Investigation.set(
            '건설기계대여대금 지급보증서발급금액',
            FillCostAccount.ToLong(
                Data.Investigation.get('직공비') *
                (Data.Rate1.get('건설기계대여대금 지급보증서발급금액') * 0.01)
            )
        ) //0.01m->0.01
        //다-15. 기타경비
        Data.Investigation.set(
            '기타경비',
            FillCostAccount.ToLong(
                (Data.Investigation.get('직접재료비') + Data.Investigation.get('노무비')) *
                (Data.Rate1.get('기타경비') * 0.01)
            )
        ) //0.01m->0.01
        //다. 경비
        Data.Investigation.set(
            '경비',
            Data.Investigation.get('산출경비') +
            Data.Investigation.get('산재보험료') +
            Data.Investigation.get('고용보험료') +
            Data.Fixed.get('국민건강보험료') +
            Data.Fixed.get('노인장기요양보험') +
            Data.Fixed.get('국민연금보험료') +
            Data.Fixed.get('퇴직공제부금') +
            Data.Fixed.get('산업안전보건관리비') +
            Data.Fixed.get('안전관리비') +
            Data.Fixed.get('품질관리비') +
            Data.Investigation.get('환경보전비') +
            Data.Investigation.get('공사이행보증서발급수수료') +
            Data.Investigation.get('건설하도급보증수수료') +
            Data.Investigation.get('건설기계대여대금 지급보증서발급금액') +
            Data.Investigation.get('기타경비')
        )
        //1. 순공사원가
        Data.Investigation.set(
            '순공사원가',
            Data.Investigation.get('직접재료비') +
            Data.Investigation.get('노무비') +
            Data.Investigation.get('경비')
        )
        //2. 일반관리비
        Data.Investigation.set(
            '일반관리비',
            FillCostAccount.ToLong(
                Data.Investigation.get('순공사원가') * (Data.Rate1.get('일반관리비') * 0.01)
            )
        ) //0.01m->0.01
        //3. 이윤
        Data.Investigation.set(
            '이윤',
            FillCostAccount.ToLong(
                (Data.Investigation.get('노무비') +
                    Data.Investigation.get('경비') +
                    Data.Investigation.get('일반관리비')) *
                0.12
            )
        ) //0.12m->0.12
        if (correction.has('이윤')) Data.Investigation.set('이윤', correction.get('이윤'))
        //3.1 미확정설계공종(PS)
        Data.Investigation.set(
            'PS',
            FillCostAccount.ToLong(Data.PsMaterial + Data.PsLabor + Data.PsExpense)
        )
        //3.2 제요율적용제외공종
        Data.Investigation.set(
            '제요율적용제외공종',
            FillCostAccount.ToLong(
                Data.ExcludingMaterial + Data.ExcludingLabor + Data.ExcludingExpense
            )
        )
        let exSum = Data.ExcludingMaterial + Data.ExcludingLabor + Data.ExcludingExpense
        let exRate2 = Math.trunc((exSum / Data.Investigation.get('직공비')) * 100000) / 100000;
        Data.Rate2.set('제요율적용제외공종', exRate2)
        //4. 총원가
        Data.Investigation.set(
            '총원가',
            Data.Investigation.get('순공사원가') +
            Data.Investigation.get('일반관리비') +
            Data.Investigation.get('이윤') +
            Data.Investigation.get('PS') +
            Data.Investigation.get('제요율적용제외공종')
        )
        //5. 공사손해보험료
        Data.Investigation.set(
            '공사손해보험료',
            FillCostAccount.ToLong(
                Data.Investigation.get('직공비') * (Data.Rate1.get('공사손해보험료') * 0.01)
            )
        )
        if (correction.has('공사손해보험료'))
            Data.Investigation.set('공사손해보험료', correction.get('공사손해보험료'))
        //작업설 추가 (23.02.06)
        Data.Investigation.set('작업설 등', FillCostAccount.ToLong(Data.ByProduct))
        //6. 소계
        Data.Investigation.set(
            '소계',
            Data.Investigation.get('총원가') +
            Data.Investigation.get('공사손해보험료') +
            Data.Investigation.get('작업설 등')
        ) //전체 가격 계산에 작업설 추가 (23.02.06)
        //7. 부가가치세
        Data.Investigation.set(
            '부가가치세',
            FillCostAccount.ToLong(
                Data.Investigation.get('소계') * (Data.Rate1.get('부가가치세') * 0.01)
            )
        )
        if (correction.has('부가가치세'))
            Data.Investigation.set('부가가치세', correction.get('부가가치세'))
        //8. 매입세(입찰 공사 파일 중, 매입세 있는 공사 없음. 추후 추가할 수 있음)
        //9. 도급비계
        Data.Investigation.set(
            '도급비계',
            Data.Investigation.get('소계') + Data.Investigation.get('부가가치세')
        )
    }

    /**원가계산서 항목별 입찰금액 구하여 Bidding에 저장*/
    public static CalculateBiddingCosts() {
        //직공비
        Data.Bidding.set(
            '직공비',
            FillCostAccount.ToLong(
                Data.RealDirectMaterial + Data.RealDirectLabor + Data.RealOutputExpense
            )
        )
        //적용비율 2를 적용한 금액 계산
        let undirectlabor2 = Data.Bidding.get('직공비') * (Data.Rate2.get('간접노무비') * 0.01)
        let industrial2 = Data.Bidding.get('직공비') * (Data.Rate2.get('산재보험료') * 0.01)
        let employ2 = Data.Bidding.get('직공비') * (Data.Rate2.get('고용보험료') * 0.01)
        let etc2 = Data.Bidding.get('직공비') * (Data.Rate2.get('기타경비') * 0.01)

        //적용비율 2를 적용한 금액 저장
        Data.Bidding.set('간접노무비2', FillCostAccount.ToLong(undirectlabor2))
        Data.Bidding.set('산재보험료2', FillCostAccount.ToLong(industrial2))
        Data.Bidding.set('고용보험료2', FillCostAccount.ToLong(employ2))
        Data.Bidding.set('기타경비2', FillCostAccount.ToLong(etc2))

        //가-1. 직접재료비
        Data.Bidding.set('직접재료비', FillCostAccount.ToLong(Data.RealDirectMaterial))
        //나-1. 직접노무비
        Data.Bidding.set('직접노무비', FillCostAccount.ToLong(Data.RealDirectLabor))
        //나-2.간접노무비
        Data.Bidding.set(
            '간접노무비1',
            FillCostAccount.ToLong(
                Data.Bidding.get('직접노무비') * (Data.Rate1.get('간접노무비') * 0.01)
            )
        )
        if (Data.Bidding.get('간접노무비1') < Data.Bidding.get('간접노무비2')) {
            Data.Bidding.set('간접노무비', Data.Bidding.get('간접노무비2'))
            Data.Bidding.set('간접노무비max', Data.Bidding.get('간접노무비2'))
            Data.Bidding.set('간접노무비before', Data.Bidding.get('간접노무비2'))
            if (Data.CostAccountDeduction === '1')
                Data.Bidding.set(
                    '간접노무비',
                    FillCostAccount.ToLong(Math.ceil(undirectlabor2 * 0.997))
                )
        } else {
            Data.Bidding.set('간접노무비', Data.Bidding.get('간접노무비1'))
            Data.Bidding.set('간접노무비max', Data.Bidding.get('간접노무비1'))
            Data.Bidding.set('간접노무비before', Data.Bidding.get('간접노무비1'))
            if (Data.CostAccountDeduction === '1')
                Data.Bidding.set(
                    '간접노무비',
                    FillCostAccount.ToLong(
                        Math.ceil(
                            Data.Bidding.get('직접노무비') *
                            (Data.Rate1.get('간접노무비') * 0.01) *
                            0.997
                        )
                    )
                )
        }
        //나. 노무비
        Data.Bidding.set('노무비', Data.Bidding.get('직접노무비') + Data.Bidding.get('간접노무비'))
        //다-1. 산출경비
        Data.Bidding.set('산출경비', FillCostAccount.ToLong(Data.RealOutputExpense))
        //다-2. 산재보험료
        Data.Bidding.set(
            '산재보험료1',
            FillCostAccount.ToLong(
                Data.Bidding.get('노무비') * (Data.Rate1.get('산재보험료') * 0.01)
            )
        )
        if (Data.Bidding.get('산재보험료1') < Data.Bidding.get('산재보험료2')) {
            Data.Bidding.set('산재보험료', Data.Bidding.get('산재보험료2'))
            Data.Bidding.set('산재보험료max', Data.Bidding.get('산재보험료2'))
        } else {
            Data.Bidding.set('산재보험료', Data.Bidding.get('산재보험료1'))
            Data.Bidding.set('산재보험료max', Data.Bidding.get('산재보험료1'))
        }
        //다-3. 고용보험료
        Data.Bidding.set(
            '고용보험료1',
            FillCostAccount.ToLong(
                Data.Bidding.get('노무비') * (Data.Rate1.get('고용보험료') * 0.01)
            )
        )
        if (Data.Bidding.get('고용보험료1') < Data.Bidding.get('고용보험료2')) {
            Data.Bidding.set('고용보험료', Data.Bidding.get('고용보험료2'))
            Data.Bidding.set('고용보험료max', Data.Bidding.get('고용보험료2'))
        } else {
            Data.Bidding.set('고용보험료', Data.Bidding.get('고용보험료1'))
            Data.Bidding.set('고용보험료max', Data.Bidding.get('고용보험료1'))
        }
        //다-11. 환경보전비
        Data.Bidding.set(
            '환경보전비',
            FillCostAccount.ToLong(
                Data.Bidding.get('직공비') * (Data.Rate1.get('환경보전비') * 0.01)
            )
        )
        //다-12 공사이행보증수수료
        Data.Bidding.set('공사이행보증서발급수수료', 0)
        if (Data.Rate1.get('공사이행보증서발급수수료') != 0)
            Data.Bidding.set(
                '공사이행보증서발급수수료',
                FillCostAccount.GetConstructionGuaranteeFee(Data.Bidding.get('직공비'))
            )
        //다-13. 하도급대금지금보증수수료
        Data.Bidding.set(
            '건설하도급보증수수료',
            FillCostAccount.ToLong(
                Data.Bidding.get('직공비') * (Data.Rate1.get('건설하도급보증수수료') * 0.01)
            )
        )
        //다-14. 건설기계대여대금 지급보증서 발급금액
        Data.Bidding.set(
            '건설기계대여대금 지급보증서발급금액',
            FillCostAccount.ToLong(
                Data.Bidding.get('직공비') *
                (Data.Rate1.get('건설기계대여대금 지급보증서발급금액') * 0.01)
            )
        )
        //다-15. 기타경비
        Data.Bidding.set(
            '기타경비1',
            FillCostAccount.ToLong(
                (Data.Bidding.get('직접재료비') + Data.Bidding.get('노무비')) *
                (Data.Rate1.get('기타경비') * 0.01)
            )
        )
        if (Data.Bidding.get('기타경비1') < Data.Bidding.get('기타경비2')) {
            Data.Bidding.set('기타경비', Data.Bidding.get('기타경비2'))
            Data.Bidding.set('기타경비max', Data.Bidding.get('기타경비2'))
            Data.Bidding.set('기타경비before', Data.Bidding.get('기타경비2'))
            if (Data.CostAccountDeduction === '1')
                Data.Bidding.set('기타경비', FillCostAccount.ToLong(Math.ceil(etc2 * 0.997)))
        } else {
            Data.Bidding.set('기타경비', Data.Bidding.get('기타경비1'))
            Data.Bidding.set('기타경비max', Data.Bidding.get('기타경비1'))
            Data.Bidding.set('기타경비before', Data.Bidding.get('기타경비1'))
            if (Data.CostAccountDeduction === '1')
                Data.Bidding.set(
                    '기타경비',
                    FillCostAccount.ToLong(
                        Math.ceil(
                            (Data.Bidding.get('직접재료비') + Data.Bidding.get('노무비')) *
                            (Data.Rate1.get('기타경비') * 0.01) *
                            0.997
                        )
                    )
                ) //decimal 0.01m 0.997m
        }
        //다. 경비
        Data.Bidding.set(
            '경비',
            Data.Bidding.get('산출경비') +
            Data.Bidding.get('산재보험료') +
            Data.Bidding.get('고용보험료') +
            Data.Fixed.get('국민건강보험료') +
            Data.Fixed.get('노인장기요양보험') +
            Data.Fixed.get('국민연금보험료') +
            Data.Fixed.get('퇴직공제부금') +
            Data.Fixed.get('산업안전보건관리비') +
            Data.Fixed.get('안전관리비') +
            Data.Fixed.get('품질관리비') +
            Data.Bidding.get('환경보전비') +
            Data.Bidding.get('공사이행보증서발급수수료') +
            Data.Bidding.get('건설하도급보증수수료') +
            Data.Bidding.get('건설기계대여대금 지급보증서발급금액') +
            Data.Bidding.get('기타경비')
        )
        //1. 순공사원가
        Data.Bidding.set(
            '순공사원가',
            Data.Bidding.get('직접재료비') + Data.Bidding.get('노무비') + Data.Bidding.get('경비')
        )
        //2. 일반관리비
        Data.Bidding.set(
            '일반관리비',
            FillCostAccount.ToLong(
                Data.Bidding.get('순공사원가') * (Data.Rate1.get('일반관리비') * 0.01)
            )
        ) //decimal 0.01m
        Data.Bidding.set(
            '일반관리비before',
            FillCostAccount.ToLong(
                Data.Bidding.get('순공사원가') * (Data.Rate1.get('일반관리비') * 0.01)
            )
        ) //decimal 0.01m
        if (Data.CostAccountDeduction === '1') {
            Data.Bidding.set(
                '일반관리비before',
                FillCostAccount.ToLong(
                    Data.Bidding.get('순공사원가') * (Data.Rate1.get('일반관리비') * 0.01)
                )
            ) //decimal 0.01m
            Data.Bidding.set(
                '일반관리비',
                FillCostAccount.ToLong(
                    Math.ceil(
                        Data.Bidding.get('순공사원가') *
                        (Data.Rate1.get('일반관리비') * 0.01) *
                        0.997
                    )
                )
            ) //decimal 0.01m 0.997m
        }
        //3. 이윤(이윤의 입찰 금액은 0원)
        Data.Bidding.set('이윤', 0)
        //3.1 미확정설계공종(PS)
        Data.Bidding.set(
            'PS',
            FillCostAccount.ToLong(Data.PsMaterial + Data.PsLabor + Data.PsExpense)
        )
        //3.2 제요율적용제외공종
        Data.Bidding.set(
            '제요율적용제외공종',
            FillCostAccount.ToLong(
                Data.AdjustedExMaterial + Data.AdjustedExLabor + Data.AdjustedExExpense
            )
        )
        //4. 총원가
        Data.Bidding.set(
            '총원가',
            Data.Bidding.get('순공사원가') +
            Data.Bidding.get('일반관리비') +
            Data.Bidding.get('이윤') +
            Data.Bidding.get('PS') +
            Data.Bidding.get('제요율적용제외공종')
        )
        //5. 공사손해보험료
        Data.Bidding.set(
            '공사손해보험료',
            FillCostAccount.ToLong(
                Data.Bidding.get('직공비') * (Data.Rate1.get('공사손해보험료') * 0.01)
            )
        ) //decimal 0.01m
        if (Data.CostAccountDeduction === '1') {
            Data.Bidding.set(
                '공사손해보험료before',
                FillCostAccount.ToLong(
                    Data.Bidding.get('직공비') * (Data.Rate1.get('공사손해보험료') * 0.01)
                )
            ) //decimal 0.01m
            Data.Bidding.set(
                '공사손해보험료',
                FillCostAccount.ToLong(
                    Math.ceil(
                        Data.Bidding.get('직공비') * Data.Rate1.get('공사손해보험료') * 0.01 * 0.997
                    )
                )
            ) //decimal 0.01m 0.997m
        }
        //작업설 추가 (23.02.06)
        Data.Bidding.set('작업설 등', FillCostAccount.ToLong(Data.ByProduct))
        //6. 소계
        Data.Bidding.set(
            '소계',
            Data.Bidding.get('총원가') +
            Data.Bidding.get('공사손해보험료') +
            Data.Bidding.get('작업설 등')
        ) //전체 가격 계산에 작업설 추가 (23.02.06)
        //7. 부가가치세
        Data.Bidding.set(
            '부가가치세',
            FillCostAccount.ToLong(Data.Bidding.get('소계') * (Data.Rate1.get('부가가치세') * 0.01))
        ) //decimal 0.01m
        //8. 매입세(입찰 공사 파일 중, 매입세 있는 공사 없음. 추후 추가할 수 있음)
        //9. 도급비계
        Data.Bidding.set('도급비계', Data.Bidding.get('소계') + Data.Bidding.get('부가가치세'))

        //도급비계 1000원 단위 절상 옵션 적용시
        if (Data.BidPriceRaise === '1') {
            let raise = 1000 - (Data.Bidding.get('도급비계') % 1000) //1000원 단위 절상 //Convert.ToDecimal (Data.Bidding.get("도급비계")) % 1000)
            let addPrice = raise / 1.1 //decimal 1.1m
            Data.Bidding.set(
                '도급비계',
                FillCostAccount.ToLong(Data.Bidding.get('도급비계') + raise)
            )
            Data.Bidding.set(
                '일반관리비',
                FillCostAccount.ToLong(Data.Bidding.get('일반관리비') + addPrice)
            ) //절상에 필요한 가격을 일반관리비에 더해 금액을 맞추어줌
            //FillCostAccount.ToLong(Convert.ToDecimal (Data.Bidding.get("일반관리비")) + addPrice)

            //일반관리비 증가에 따른 타 금액 변경
            Data.Bidding.set('소계', FillCostAccount.ToLong(Data.Bidding.get('소계') + addPrice)) //FillCostAccount.ToLong(Convert.ToDecimal (Data.Bidding.get("소계")) + addPrice)
            Data.Bidding.set(
                '부가가치세',
                FillCostAccount.ToLong(
                    Data.Bidding.get('소계') * (Data.Rate1.get('부가가치세') * 0.01)
                )
            ) //decimal 0.01m
            //계산된 도급비계 금액이 천원 단위가 아닐 경우, 부가세 조정
            let difference =
                Data.Bidding.get('도급비계') -
                (Data.Bidding.get('소계') + Data.Bidding.get('부가가치세'))
            Data.Bidding.set(
                '부가가치세',
                FillCostAccount.ToLong(Data.Bidding.get('부가가치세') + difference)
            )
            //Console.WriteLine("차이 : " + difference);
        }
    }

    //decimal 금액 원 단위 절사
    public static ToLong(price: number) {//price: decimal
        let bigNum = Math.trunc(price);
        return bigNum
    }

    //공사이행보증서발급수수료 금액 계산 후 반환
    static GetConstructionGuaranteeFee(
        directSum: number //long directSum //수정시 자료형 재확인
    ) {
        var guaranteeFee: number = 0 //long
        var rate = Data.Rate1.get('공사이행보증서발급수수료') * 0.01 //decimal 0.01m
        var term = Data.ConstructionTerm / 365.0 //decimal  365.0m;
        if (directSum < 7000000000) guaranteeFee = FillCostAccount.ToLong(directSum * rate * term)
        else if (directSum < 12000000000)
            guaranteeFee = FillCostAccount.ToLong(
                ((directSum - 5000000000) * rate + 2000000) * term
            )
        else if (directSum < 25000000000)
            guaranteeFee = FillCostAccount.ToLong(
                ((directSum - 14000000000) * rate + 4000000) * term
            )
        else if (directSum < 50000000000)
            guaranteeFee = FillCostAccount.ToLong(
                ((directSum - 25000000000) * rate + 6000000) * term
            )
        else
            guaranteeFee = FillCostAccount.ToLong(
                ((directSum - 50000000000) * rate + 10000000) * term
            )
        return guaranteeFee
    }

    //입찰 금액의 조사금액 대 비율 저장
    public static GetRate(item: string) {
        if (Data.Fixed.has(item)) return 100
        if (Data.Investigation.get(item) == 0 && Data.Bidding.get(item) == 0) return 100
        //원가계산제경비 옵션 적용 항목은 적용 전, 후의 비율 출력
        if (
            item === '간접노무비' ||
            item === '기타경비' ||
            item === '일반관리비' ||
            item === '공사손해보험료'
        ) {
            var before: string = item + 'before'
            return Math.round(Data.Bidding[item] / Data.Bidding[before] * 10000000) / 10000000 * 100;
        }
        let rate = Math.round(Data.Bidding[item] / Data.Bidding[before] * 10000000) / 10000000;
        return rate;
    }

    //해당 공사에 특정 원가계산서 항목이 존재하지 않는 경우
    public static CheckKeyNotFound() {
        if (!Data.Rate1.has('간접노무비')) Data.Rate1.set('간접노무비', 0)
        if (!Data.Rate1.has('산재보험료')) Data.Rate1.set('산재보험료', 0)
        if (!Data.Rate1.has('고용보험료')) Data.Rate1.set('고용보험료', 0)
        if (!Data.Rate1.has('환경보전비')) Data.Rate1.set('환경보전비', 0)
        if (!Data.Rate1.has('공사이행보증서발급수수료'))
            Data.Rate1.set('공사이행보증서발급수수료', 0)
        if (!Data.Rate1.has('건설하도급보증수수료')) Data.Rate1.set('건설하도급보증수수료', 0)
        if (!Data.Rate1.has('건설기계대여대금 지급보증서발급금액'))
            Data.Rate1.set('건설기계대여대금 지급보증서발급금액', 0)
        if (!Data.Rate1.has('기타경비')) Data.Rate1.set('기타경비', 0)
        if (!Data.Rate1.has('일반관리비')) Data.Rate1.set('일반관리비', 0)
        if (!Data.Rate1.has('부가가치세')) Data.Rate1.set('부가가치세', 0)
        if (!Data.Rate1.has('공사손해보험료')) Data.Rate1.set('공사손해보험료', 0)

        if (!Data.Fixed.has('국민건강보험료')) Data.Fixed.set('국민건강보험료', 0)
        if (!Data.Fixed.has('노인장기요양보험')) Data.Fixed.set('노인장기요양보험', 0)
        if (!Data.Fixed.has('국민연금보험료')) Data.Fixed.set('국민연금보험료', 0)
        if (!Data.Fixed.has('퇴직공제부금')) Data.Fixed.set('퇴직공제부금', 0)
        if (!Data.Fixed.has('산업안전보건관리비')) Data.Fixed.set('산업안전보건관리비', 0)
        if (!Data.Fixed.has('안전관리비')) Data.Fixed.set('안전관리비', 0)
        if (!Data.Fixed.has('품질관리비')) Data.Fixed.set('품질관리비', 0)

        if (!Data.Rate2.has('간접노무비')) Data.Rate2.set('간접노무비', 0)
        if (!Data.Rate2.has('산재보험료')) Data.Rate2.set('산재보험료', 0)
        if (!Data.Rate2.has('고용보험료')) Data.Rate2.set('고용보험료', 0)
        if (!Data.Rate2.has('기타경비')) Data.Rate2.set('기타경비', 0)
    }
}
