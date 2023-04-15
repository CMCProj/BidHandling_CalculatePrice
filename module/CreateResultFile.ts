import { Data } from './Data'
import * as  path from 'path'
import * as fs from 'fs'
import * as exceljs from 'exceljs'
import {ExcelHandling} from "./ExcelHandling";
export class CreateResultFile {
    public static Create(): void {
        const directoryPath = Data.folder // 파일 경로
        const extension = '.xls' // 파일 확장자 명
        const xlsList = [] // 파일 경로에 있는 엑셀파일명을 담을 리스트

        fs.readdir(directoryPath, function (err, files) {
            // 파일 경로에 있는 엑셀파일의 절대경로를 모두 저장한다.
            if (err) {
                return console.log('Unable to scan directory: ' + err)
            }
            files.forEach((file) => {
                if (path.extname(file) === extension) {
                    xlsList.push(path.join(directoryPath, file))
                }
            });
        })
        xlsList.forEach((xls) => {
            // 기존의 엑셀 파일을 모두 삭제한다.
            fs.access(xls, fs.constants.F_OK, (err) => {
                if (err) return console.log('error(파일 삭제 실패)')
                fs.unlink(xls, (err) => (err ? console.log(err) : console.log('${xls} 삭제 완료')))
            });
        });
        Data.Dic.forEach(async (value, key) => {
            let workbook = await ExcelHandling.GetWorkbook("입찰내역.xls", "xls");
            let sheet = workbook.getWorksheet(0);
            let resultPath: string;
            let Path: string;

            for (let i = 0; i < value.length; i++) {
                if (value[i].Item === '일반') {
                    sheet.getCell(i, 22).value = Number(value[i].PriceScore);
                    sheet.getCell(i, 23).value = Number(value[i].Score);
                }
                let materialUnit = Number(value[i].MaterialUnit);
                let laborUnit = Number(value[i].LaborUnit);
                let expenseUnit = Number(value[i].ExpenseUnit);
                let material = Number(value[i].Material);
                let labor = Number(value[i].Labor);
                let expense = Number(value[i].Expense);
                let unitpricesum = Number(value[i].UnitPriceSum);
                let pricesum = Number(value[i].PriceSum);

                sheet.getCell(i + 1, 1).value = (i + 1) * 100;
                sheet.getCell(i + 1, 3).value = value[i].Name;
                sheet.getCell(i + 1, 4).value = value[i].Standard;
                sheet.getCell(i + 1, 5).value = value[i].Unit;
                sheet.getCell(i + 1, 6).value = Number(value[i].Quantity);
                sheet.getCell(i + 1, 7).value = materialUnit;
                sheet.getCell(i + 1, 8).value = laborUnit;
                sheet.getCell(i + 1, 9).value = expenseUnit;
                sheet.getCell(i + 1, 10).value = unitpricesum;
                sheet.getCell(i + 1, 11).value = material;
                sheet.getCell(i + 1, 12).value = labor;
                sheet.getCell(i + 1, 13).value = expense;
                sheet.getCell(i + 1, 14).value = pricesum;
                sheet.getCell(i + 1, 18).value = value[i].Code;

                if (value[i].Item === '표준시장단가')
                {
                    sheet.getCell(i+1, 16).value = Number(value[i].Item);
                    sheet.getCell(i+1, 17).value = Number(value[i].Code);
                }
                else
                {
                    sheet.getCell(i + 1, 15).value = value[i].Item;
                }
            }

            resultPath = '입찰내역_' + Data.ConstructionNums[key] + '.xls';
            Path = path.join(Data.folder, resultPath);
            ExcelHandling.WriteExcel(workbook, Path);
        });
    }
}
