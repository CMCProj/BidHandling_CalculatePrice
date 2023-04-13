"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.BidHandling = void 0;
//const { buffer } = require('stream/consumers')
var convert = require("xml-js");
var fs = require("fs");
var AdmZip = require('adm-zip');
var Data_1 = require("./Data");
//실행 위해 프로그램 내 폴더로 경로 변경
var filename = undefined;
var BidHandling = /** @class */ (function () {
    function BidHandling() {
    }
    BidHandling.BidToJson = function () {
        var copiedFolder = Data_1.Data.folder + '\\EmptyBid';
        var bidFile = fs.readdirSync(copiedFolder);
        var myFile = bidFile.filter(function (file) { return file.substring(file.length - 4, file.length).toLowerCase() === '.bid'; })[0];
        filename = myFile.substring(0, myFile.length - 4);
        fs.copyFileSync(copiedFolder + '\\' + myFile, copiedFolder + '\\' + filename + '.zip');
        fs.rmSync(copiedFolder + '\\' + myFile);
        var zip = new AdmZip(copiedFolder + '\\' + filename + '.zip');
        zip.extractAllTo(copiedFolder, true);
        bidFile = fs.readdirSync(copiedFolder);
        myFile = bidFile.filter(function (file) { return file.substring(file.length - 4, file.length).toLowerCase() === '.bid'; })[0];
        var text = fs.readFileSync(copiedFolder + '\\' + myFile, 'utf-8');
        var decodeValue = Buffer.from(text, 'base64');
        text = decodeValue.toString('utf-8');
        fs.writeFileSync(Data_1.Data.folder + '\\OutputDataFromBID.xml', text);
        var xml = fs.readFileSync(Data_1.Data.folder + '\\OutputDataFromBID.xml', 'utf-8');
        var json = convert.xml2json(xml, { compact: true, spaces: 4 });
        fs.writeFileSync(Data_1.Data.folder + '\\OutputDataFromBID.json', json);
        fs.rmSync(copiedFolder + '\\' + filename + '.zip');
        fs.rmSync(copiedFolder + '\\' + myFile);
        fs.rmSync(Data_1.Data.folder + '\\OutputDataFromBID.xml');
        // Setting.GetData()
    };
    BidHandling.JsonToBid = function () {
        var resultFilePath = Data_1.Data.folder + '\\OutputDataFromBID.json';
        var json = fs.readFileSync(resultFilePath, 'utf-8');
        var xml = convert.json2xml(json, { compact: true, ignoreComment: true, spaces: 4 });
        fs.writeFileSync(Data_1.Data.folder + '\\Result_Xml.xml', xml);
        var text = fs.readFileSync(Data_1.Data.folder + '\\Result_Xml.xml', 'utf-8');
        var encodeValue = Buffer.from(text, 'utf-8');
        text = encodeValue.toString('base64');
        fs.writeFileSync(Data_1.Data.folder + '\\EmptyBid\\XmlToBID.BID', text);
        var zip = new AdmZip();
        zip.addLocalFile(Data_1.Data.folder + '\\EmptyBid\\XmlToBID.BID');
        zip.writeZip(Data_1.Data.folder + '\\EmptyBid\\' + filename + '.zip');
        fs.copyFileSync(Data_1.Data.folder + '\\EmptyBid\\' + filename + '.zip', Data_1.Data.folder + '\\EmptyBid\\' + filename + '.BID');
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
