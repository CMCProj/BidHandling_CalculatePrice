"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.BidHandling = void 0;
//const { buffer } = require('stream/consumers')
var convert = require("xml-js");
var fs = require("fs");
var AdmZip = require('adm-zip');
var Setting_1 = require("./Setting");
var Data_1 = require("./Data");
//실행 위해 프로그램 내 폴더로 경로 변경
var filename = undefined;
var BidHandling = /** @class */ (function () {
    function BidHandling() {
    }
    BidHandling.BidToJson = function () {
        var copiedFolder = Data_1.Data.folder + '\\EmptyBid'; // EmptyBid폴더 주소 저장 / 폴더 경로 수정 (23.02.02)
        var bidFile = fs.readdirSync(copiedFolder);
        var myFile = bidFile.filter(function (file) { return file.substring(file.length - 4, file.length).toLowerCase() === '.bid'; })[0]; // 확장자가 .bid인 파일을 찾기
        filename = myFile.substring(0, myFile.length - 4); // 확장자를 뺀 파일의 이름
        fs.copyFileSync(copiedFolder + '\\' + myFile, copiedFolder + '\\' + filename + '.zip'); // 확장자를 .bid에서 .zip으로 교체
        fs.rmSync(copiedFolder + '\\' + myFile); // 기존의 .bid 파일은 삭제
        var zip = new AdmZip(copiedFolder + '\\' + filename + '.zip'); // .zip파일 압축 해제
        zip.extractAllTo(copiedFolder, true);
        bidFile = fs.readdirSync(copiedFolder);
        myFile = bidFile.filter(function (file) { return file.substring(file.length - 4, file.length).toLowerCase() === '.bid'; })[0]; // 압축 해제되어 나온 .bid파일 찾기
        var text = fs.readFileSync(copiedFolder + '\\' + myFile, 'utf-8'); // 나온 .bid 파일의 텍스트 읽기
        var decodeValue = Buffer.from(text, 'base64'); // base64코드 디코딩
        text = decodeValue.toString('utf-8'); // 디코딩 되어 나온 텍스트 저장
        fs.writeFileSync(Data_1.Data.folder + '\\OutputDataFromBID.xml', text); // xml파일 형식으로 텍스트 쓰기
        var xml = fs.readFileSync(Data_1.Data.folder + '\\OutputDataFromBID.xml', 'utf-8');
        var json = convert.xml2json(xml, { compact: true, spaces: 4 }); // xml파일을 json으로 교체
        fs.writeFileSync(Data_1.Data.folder + '\\OutputDataFromBID.json', json); // json파일 형식으로 다시 쓰기
        //=======이 과정에서 만들어진 파일들은 전부 삭제=======
        fs.rmSync(copiedFolder + '\\' + filename + '.zip');
        fs.rmSync(copiedFolder + '\\' + myFile);
        fs.rmSync(Data_1.Data.folder + '\\OutputDataFromBID.xml');
        Setting_1.Setting.GetData();
    };
    BidHandling.JsonToBid = function () {
        var resultFilePath = Data_1.Data.folder + '\\OutputDataFromBID.json';
        var json = fs.readFileSync(resultFilePath, 'utf-8'); // json파일 읽기
        var xml = convert.json2xml(json, { compact: true, ignoreComment: true, spaces: 4 }); // json파일을 xml파일로 교체
        fs.writeFileSync(Data_1.Data.folder + '\\Result_Xml.xml', xml); // xml파일 형식으로 다시 쓰기
        var text = fs.readFileSync(Data_1.Data.folder + '\\Result_Xml.xml', 'utf-8'); // xml파일 내용 읽기
        var encodeValue = Buffer.from(text, 'utf-8');
        text = encodeValue.toString('base64'); // base64로 인코딩
        fs.writeFileSync(Data_1.Data.folder + '\\EmptyBid\\XmlToBID.BID', text); // 인코딩된 텍스트 XmlToBID.BID파일에 쓰기
        var zip = new AdmZip();
        zip.addLocalFile(Data_1.Data.folder + '\\EmptyBid\\XmlToBID.BID');
        zip.writeZip(Data_1.Data.folder + '\\EmptyBid\\' + filename + '.zip'); // XmlToBID.BID파일을 .zip파일로 압축
        fs.copyFileSync(Data_1.Data.folder + '\\EmptyBid\\' + filename + '.zip', Data_1.Data.folder + '\\EmptyBid\\' + filename + '.BID');
        //======이 과정에서 만들어진 파일들은 전부 삭제======
        fs.rmSync(Data_1.Data.folder + '\\EmptyBid\\' + filename + '.zip');
        fs.rmSync(resultFilePath);
        fs.rmSync(Data_1.Data.folder + '\\Result_Xml.xml');
        fs.rmSync(Data_1.Data.folder + '\\EmptyBid\\XmlToBID.BID');
    };
    return BidHandling;
}());
exports.BidHandling = BidHandling;
// BidHandling.BidToJson();
// BidHandling.JsonToBid();
