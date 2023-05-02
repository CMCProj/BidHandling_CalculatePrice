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
exports.BidHandling = void 0;
//const { buffer } = require('stream/consumers')
const convert = __importStar(require("xml-js"));
const fs = __importStar(require("fs"));
const AdmZip = require('adm-zip');
const Setting_1 = require("./Setting");
const Data_1 = require("./Data");
//실행 위해 프로그램 내 폴더로 경로 변경
let filename = undefined;
class BidHandling {
    static BidToJson() {
        const copiedFolder = Data_1.Data.folder + '\\EmptyBid'; // EmptyBid폴더 주소 저장 / 폴더 경로 수정 (23.02.02)
        let bidFile = fs.readdirSync(copiedFolder);
        let myFile = bidFile.filter((file) => file.substring(file.length - 4, file.length).toLowerCase() === '.bid')[0]; // 확장자가 .bid인 파일을 찾기
        filename = myFile.substring(0, myFile.length - 4); // 확장자를 뺀 파일의 이름
        fs.copyFileSync(copiedFolder + '\\' + myFile, copiedFolder + '\\' + filename + '.zip'); // 확장자를 .bid에서 .zip으로 교체
        fs.rmSync(copiedFolder + '\\' + myFile); // 기존의 .bid 파일은 삭제
        let zip = new AdmZip(copiedFolder + '\\' + filename + '.zip'); // .zip파일 압축 해제
        zip.extractAllTo(copiedFolder, true);
        bidFile = fs.readdirSync(copiedFolder);
        myFile = bidFile.filter((file) => file.substring(file.length - 4, file.length).toLowerCase() === '.bid')[0]; // 압축 해제되어 나온 .bid파일 찾기
        let text = fs.readFileSync(copiedFolder + '\\' + myFile, 'utf-8'); // 나온 .bid 파일의 텍스트 읽기
        const decodeValue = Buffer.from(text, 'base64'); // base64코드 디코딩
        text = decodeValue.toString('utf-8'); // 디코딩 되어 나온 텍스트 저장
        fs.writeFileSync(Data_1.Data.folder + '\\OutputDataFromBID.xml', text); // xml파일 형식으로 텍스트 쓰기
        const xml = fs.readFileSync(Data_1.Data.folder + '\\OutputDataFromBID.xml', 'utf-8');
        const json = convert.xml2json(xml, { compact: true, spaces: 4 }); // xml파일을 json으로 교체
        fs.writeFileSync(Data_1.Data.folder + '\\OutputDataFromBID.json', json); // json파일 형식으로 다시 쓰기
        //=======이 과정에서 만들어진 파일들은 전부 삭제=======
        fs.rmSync(copiedFolder + '\\' + filename + '.zip');
        fs.rmSync(copiedFolder + '\\' + myFile);
        fs.rmSync(Data_1.Data.folder + '\\OutputDataFromBID.xml');
        Setting_1.Setting.GetData();
    }
    static JsonToBid() {
        const resultFilePath = Data_1.Data.folder + '\\OutputDataFromBID.json';
        const json = fs.readFileSync(resultFilePath, 'utf-8'); // json파일 읽기
        const xml = convert.json2xml(json, { compact: true, ignoreComment: true, spaces: 4 }); // json파일을 xml파일로 교체
        fs.writeFileSync(Data_1.Data.folder + '\\Result_Xml.xml', xml); // xml파일 형식으로 다시 쓰기
        let text = fs.readFileSync(Data_1.Data.folder + '\\Result_Xml.xml', 'utf-8'); // xml파일 내용 읽기
        const encodeValue = Buffer.from(text, 'utf-8');
        text = encodeValue.toString('base64'); // base64로 인코딩
        fs.writeFileSync(Data_1.Data.folder + '\\EmptyBid\\XmlToBID.BID', text); // 인코딩된 텍스트 XmlToBID.BID파일에 쓰기
        let zip = new AdmZip();
        zip.addLocalFile(Data_1.Data.folder + '\\EmptyBid\\XmlToBID.BID');
        zip.writeZip(Data_1.Data.folder + '\\EmptyBid\\' + filename + '.zip'); // XmlToBID.BID파일을 .zip파일로 압축
        fs.copyFileSync(Data_1.Data.folder + '\\EmptyBid\\' + filename + '.zip', Data_1.Data.folder + '\\EmptyBid\\' + filename + '.BID');
        //======이 과정에서 만들어진 파일들은 전부 삭제======
        fs.rmSync(Data_1.Data.folder + '\\EmptyBid\\' + filename + '.zip');
        fs.rmSync(resultFilePath);
        fs.rmSync(Data_1.Data.folder + '\\Result_Xml.xml');
        fs.rmSync(Data_1.Data.folder + '\\EmptyBid\\XmlToBID.BID');
    }
}
exports.BidHandling = BidHandling;
// BidHandling.BidToJson();
// BidHandling.JsonToBid();
