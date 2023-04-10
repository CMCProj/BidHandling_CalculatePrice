import * as fs from 'fs'
import xls from 'exceljs'
import Xlsx from 'exceljs'

export class ExcelHandling {
    public static GetRow(sheet, rownum: number) {
        let row = sheet.getRow(rownum)
        if (row == null) {
            row = sheet.addRow(rownum)
        }
        return row
    }

    public static GetCell2(row, cellnum: number) {
        let cell = row.GetCell(cellnum)
        if (cell == null) {
            cell = row.addCell(cellnum)
        }
        return cell
    }

    public static GetCell(sheet, rownum: number, cellnum: number) {
        let row = this.GetRow(sheet, rownum)
        return this.GetCell2(row, cellnum)
    }
    public static GetWorkbook(filename: string, version: string) {
        // 파일을 열고 파일 내용을 읽기/쓰기용 스트림으로 가져옴
        const stream = fs.createReadStream(filename, { flags: 'r+' })
        if (version === '.xls') {
            return xls.read(stream, { type: 'buffer' })
        } else if (version === '.xlsx') {
            return Xlsx.read(stream, { type: 'buffer' })
        }
    }

    public static WriteExcel(workbook, filepath: string) {
        const file = fs.createWriteStream(filepath)
        workbook.Write(file)
    }
}
