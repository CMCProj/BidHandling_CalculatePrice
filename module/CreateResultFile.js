"use strict";
var __createBinding = (this && this.__createBinding) || (Object.create ? (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    var desc = Object.getOwnPropertyDescriptor(m, k);
    if (!desc || ("get" in desc ? !m.__esModule : desc.writable || desc.configurable)) {
      desc = { enumerable: true, get: function() { return m[k]; } };
    }
    Object.defineProperty(o, k2, desc);
}) : (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    o[k2] = m[k];
}));
var __setModuleDefault = (this && this.__setModuleDefault) || (Object.create ? (function(o, v) {
    Object.defineProperty(o, "default", { enumerable: true, value: v });
}) : function(o, v) {
    o["default"] = v;
});
var __importStar = (this && this.__importStar) || function (mod) {
    if (mod && mod.__esModule) return mod;
    var result = {};
    if (mod != null) for (var k in mod) if (k !== "default" && Object.prototype.hasOwnProperty.call(mod, k)) __createBinding(result, mod, k);
    __setModuleDefault(result, mod);
    return result;
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.CreateResultFile = void 0;
const Data_1 = require("./Data");
const path = __importStar(require("path"));
const fs = __importStar(require("fs"));
const exceljs = __importStar(require("exceljs"));
class CreateResultFile {
    static Create() {
        const directoryPath = Data_1.Data.folder; // 파일 경로
        const extension = '.xls'; // 파일 확장자 명
        const xlsList = []; // 파일 경로에 있는 엑셀파일명을 담을 리스트
        fs.readdir(directoryPath, function (err, files) {
            // 파일 경로에 있는 엑셀파일의 절대경로를 모두 저장한다.
            if (err) {
                return console.log('Unable to scan directory: ' + err);
            }
            files.forEach(function (file) {
                if (path.extname(file) === extension) {
                    xlsList.push(path.join(directoryPath, file));
                }
            });
        });
        xlsList.forEach((xls) => {
            // 기존의 엑셀 파일을 모두 삭제한다.
            fs.access(xls, fs.constants.F_OK, (err) => {
                if (err)
                    return console.log('error(파일 삭제 실패)');
                fs.unlink(xls, (err) => (err ? console.log(err) : console.log('${xls} 삭제 완료')));
            });
        });
        const worksheet = new exceljs.Workbook();
        Data_1.Data.Dic.forEach(function (dic) {
            const workbook = worksheet.xlsx.readFile('입찰내역.xls');
            const sheet = workbook.GetSheetAt(0);
            let resultPath;
            let path;
            for (let i = 0; i < dic.values.length; i++) {
                if (dic.values[i].Item === '일반') {
                    //엑셀핸들링
                }
            }
        });
    }
}
exports.CreateResultFile = CreateResultFile;
