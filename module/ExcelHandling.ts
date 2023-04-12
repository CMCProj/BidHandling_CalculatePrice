import * as fs from 'fs'
// import xls from 'exceljs'
// import { Xlsx } from 'exceljs'
import * as exceljs from 'exceljs'
import * as node_xj from 'xls-to-json'
import * as XLSX from 'xlsx'

export class ExcelHandling {
    public static GetRow(sheet: exceljs.Worksheet, rownum: number): exceljs.Row {
        let row = sheet.getRow(rownum); //rownum에 행이 있으면 그 행을 반환하고, 없으면 그 위치에 새로운 빈 행을 만듦
        return row
    }

    public static GetCell2(row: exceljs.Row, cellnum: number): exceljs.Cell {
        let cell = row.getCell(cellnum)
        // if (cell == null) {
        //     cell = row.getCell(cellnum);
        // }
        return cell;
    }

    public static GetCell(sheet: exceljs.Worksheet, rownum: number, cellnum: number): exceljs.Cell {
        let row = this.GetRow(sheet, rownum)
        return this.GetCell2(row, cellnum)
    }

    public static GetWorkbook(filename: string, version: string): Promise<exceljs.Workbook> {
        // 파일을 열고 파일 내용을 읽기/쓰기용 스트림으로 가져옴
        let workbook = new exceljs.Workbook();
        let stream = fs.createReadStream(filename + version, { flags: 'r+' })

        if (version === '.xls') {
            node_xj({ input: filename + version, output: filename + '.json'}, function (err, result) {
                if (err) throw err
                else{
                    //xls -> json으로 변환
                    const readJson = fs.readFileSync(filename + '.json', 'utf-8');
                    let xlsx = JSON.parse(readJson);

                    //json -> xlsx로 변환
                    const sheet = XLSX.utils.json_to_sheet(xlsx);
                    const book = XLSX.utils.book_new();
                    XLSX.utils.book_append_sheet(book, sheet, 'Sheet1');
                    XLSX.writeFile(book, filename + '.xlsx');

                    fs.rmSync(filename + '.json');
                    stream = fs.createReadStream(filename + '.xlsx', { flags: 'r+'})
                }
            })

            return workbook.xlsx.read(stream);
        }
        else if(version === '.xlsx')
            return workbook.xlsx.read(stream);
        
        throw new Error('올바른 Excel파일 형식(.xls/.xlsx)이 아닙니다. 파일을 다시 한 번 확인해주세요.');
    }

    public static WriteExcel(workbook: exceljs.Workbook, filepath: string): void {
        const file = fs.createWriteStream(filepath)
        workbook.xlsx.write(file)
    }
}

//ExcelHandling.GetWorkbook('입찰내역_설비공사', '.xls');
