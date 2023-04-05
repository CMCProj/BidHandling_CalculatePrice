import {Data} from "./Data";

namespace SetUnitPriceByExcel
{
    class Setting
    {
        private static docBID : JSON;
        private static eleBID : JSON;

        static GetConstructionNum(): void
        {
            
        }

        static AddConstructionList(): void
        {

        }

        static GetDataFromBID(): void
        {

        }

        static GetItem(bid) : string
        {

        }

        static MatchConstructionNum(): void
        {

        }

        static CopyFile(filePath: string): void
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

        public static GetData(): void
        {
            this.GetConstructionNum();
            this.AddConstructionList();
            this.GetDataFromBID();
            if(Data.XlsFiles != null)
            {
                this.SetUnitPrice();
            }
            else
            {
                this.SetUnitPriceNoExcel();
            }
            this.GetRate();
            this.GetPrices();

            Data.InvestigateStandardMarket = Data.StandardMaterial + Data.StandardLabor + Data.StandardExpense;
        }
    }
}