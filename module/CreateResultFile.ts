import { Data } from './Data'
import path from 'path'
import fs from 'fs'
import exceljs from 'exceljs'

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
            files.forEach(function (file) {
                if (path.extname(file) === extension) {
                    xlsList.push(path.join(directoryPath, file))
                }
            })
        })
        xlsList.forEach((xls) => {
            // 기존의 엑셀 파일을 모두 삭제한다.
            fs.access(xls, fs.constants.F_OK, (err) => {
                if (err) return console.log('error(파일 삭제 실패)')

                fs.unlink(xls, (err) => (err ? console.log(err) : console.log('${xls} 삭제 완료')))
            })
        })
        const worksheet = new exceljs.Workbook()

        Data.Dic.forEach(function (dic) {
            const workbook = worksheet.xlsx.readFile('입찰내역.xls')
            const sheet = workbook.GetSheetAt(0)
            let resultPath: string
            let path: string
            for (let i = 0; i < dic.values.Count; i++) {
                if (dic.values[i].Item === '일반') {
                    //엑셀핸들링
                }
            }
        })
    }
}
