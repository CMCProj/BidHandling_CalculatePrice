import {Data} from "./Data";
import Enumerable from 'linq';
import fs from "fs";

namespace SetUnitPriceByExcel
{
    class Setting
    {
        private static docBID : JSON;
        private static eleBID : JSON;

        public static GetData(): void
        {
            const bidString : string = fs.readFileSync("C:\\Users\\USER\\Documents\\" + "\\OutputDataFromBID.json", 'utf-8');
            this.docBID = JSON.parse(bidString);
            this.eleBID = this.docBID['data'];

            //세부공사별 번호 Data.ConstructionNums 딕셔너리에 저장
            this.GetConstructionNum();

            //세부공사별 리스트 생성(Dic -> key : 세부공사별 번호 / value : 세부공사별 리스트)
            this.AddConstructionList();

            //공내역 xml 파일 읽어들여 데이터 저장
            this.GetDataFromBID();


            if(Data.XlsFiles != null) //실내역으로부터 Data 객체에 단가세팅
            {
                this.SetUnitPrice();
            }
            else
            {
                this.SetUnitPriceNoExcel();//실내역 없이 공내역만으로 Data 객체에 단가 세팅 (23.02.06)
            }

            //고정금액 및 적용비율 저장
            this.GetRate();

            //직공비 제외항목 및 고정금액 계산
            this.GetPrices();

            //표준시장단가 합계(조사금액) 저장
            Data.InvestigateStandardMarket = Data.StandardMaterial + Data.StandardLabor + Data.StandardExpense;
        }

        static GetConstructionNum(): void
        {
            const constNums : object = this.eleBID['T2'];
            let index : string;
            let construction : string;
            for(let num in constNums)
            {
                index = JSON.stringify(constNums[num]['C1']['_text']);
                construction = JSON.stringify(constNums[num]['C3']['_text']);
                if(Data.ConstructionNums.has(construction))
                {
                    construction += "2";
                }
                Data.ConstructionNums.set(index, construction);
            }
        }

        static AddConstructionList(): void
        {
            Data.ConstructionNums.forEach((key, value) => Data.Dic.set(key, new List<Data>()));
        }

        static GetDataFromBID(): void
        {

        }

        static GetItem(bid) : String
        {

        }

        static MatchConstructionNum(): void
        {

        }

        static CopyFile(filePath: String): void
        {

        }

        static SetUnitPrice(): void
        {

        }

        static SetUnitPriceNoExcel(): void
        {

        }

        static GetRate(): void
        {

        }

        public static GetPrices(): void
        {

        }


    }
}