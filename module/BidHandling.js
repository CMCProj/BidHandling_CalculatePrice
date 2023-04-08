"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.BidHandling = void 0;
//const { buffer } = require('stream/consumers')
var xml_js_1 = __importDefault(require("xml-js"));
var fs_1 = __importDefault(require("fs"));
var adm_zip_1 = __importDefault(require("adm-zip"));
var Data_1 = require("./Data");
var filename = undefined;
var BidHandling = /** @class */ (function () {
    function BidHandling() {
    }
    BidHandling.BidToJson = function () {
        var copiedFolder = Data_1.Data.folder + '\\EmptyBid';
        var bidFile = fs_1.default.readdirSync(copiedFolder);
        var myFile = bidFile.filter(function (file) { return file.substring(file.length - 4, file.length).toLowerCase() === '.bid'; })[0];
        filename = myFile.substring(0, myFile.length - 4);
        fs_1.default.copyFileSync(copiedFolder + '\\' + myFile, copiedFolder + '\\' + filename + '.zip');
        fs_1.default.rmSync(copiedFolder + '\\' + myFile);
        var zip = new adm_zip_1.default(copiedFolder + '\\' + filename + '.zip');
        zip.extractAllTo(copiedFolder, true);
        bidFile = fs_1.default.readdirSync(copiedFolder);
        myFile = bidFile.filter(function (file) { return file.substring(file.length - 4, file.length).toLowerCase() === '.bid'; })[0];
        var text = fs_1.default.readFileSync(copiedFolder + '\\' + myFile, 'utf-8');
        var decodeValue = Buffer.from(text, 'base64');
        text = decodeValue.toString('utf-8');
        fs_1.default.writeFileSync(Data_1.Data.folder + '\\OutputDataFromBID.xml', text);
        var xml = fs_1.default.readFileSync(Data_1.Data.folder + '\\OutputDataFromBID.xml', 'utf-8');
        var json = xml_js_1.default.xml2json(xml, { compact: true, spaces: 4 });
        fs_1.default.writeFileSync(Data_1.Data.folder + '\\OutputDataFromBID.json', json);
        fs_1.default.rmSync(copiedFolder + '\\' + filename + '.zip');
        fs_1.default.rmSync(copiedFolder + '\\' + myFile);
        fs_1.default.rmSync(Data_1.Data.folder + '\\OutputDataFromBID.xml');
    };
    BidHandling.JsonToBid = function () {
        var resultFilePath = Data_1.Data.folder + '\\OutputDataFromBID.json';
        var json = fs_1.default.readFileSync(resultFilePath, 'utf-8');
        var xml = xml_js_1.default.json2xml(json, { compact: true, ignoreComment: true, spaces: 4 });
        fs_1.default.writeFileSync(Data_1.Data.folder + '\\Result_Xml.xml', xml);
        var text = fs_1.default.readFileSync(Data_1.Data.folder + '\\Result_Xml.xml', 'utf-8');
        var encodeValue = Buffer.from(text, 'utf-8');
        text = encodeValue.toString('base64');
        fs_1.default.writeFileSync(Data_1.Data.folder + '\\EmptyBid\\XmlToBID.BID', text);
        var zip = new adm_zip_1.default();
        zip.addLocalFile(Data_1.Data.folder + '\\EmptyBid\\XmlToBID.BID');
        zip.writeZip(Data_1.Data.folder + '\\EmptyBid\\' + filename + '.zip');
        fs_1.default.copyFileSync(Data_1.Data.folder + '\\EmptyBid\\' + filename + '.zip', Data_1.Data.folder + '\\EmptyBid\\' + filename + '.BID');
        fs_1.default.rmSync(Data_1.Data.folder + '\\EmptyBid\\' + filename + '.zip');
        fs_1.default.rmSync(resultFilePath);
        fs_1.default.rmSync(Data_1.Data.folder + '\\Result_Xml.xml');
        fs_1.default.rmSync(Data_1.Data.folder + '\\EmptyBid\\XmlToBID.BID');
    };
    return BidHandling;
}());
exports.BidHandling = BidHandling;
// BidHandling.BidToJson();
// BidHandling.JsonToBid();
