const Data = require('./Data.js')

var SetUnitPrice
;(function (SetUnitPrice) {
    var fs = require('fs')
    var AdmZip = require('adm-zip')
    var buffer = require('stream/consumers').buffer
    var convert = require('xml-js')
    var filename = undefined
    function BidToJson() {
        //var copiedFolder = "C:\\\\Users\\joung\\OneDrive\\문서\\AutoBID" + "\\EmptyBid";
        console.log(Data.folder, '테스트') //테스트
        var copiedFolder = Data.folder + '\\EmptyBid'
        console.log(copiedFolder, '테스트2') //테스트

        var bidFile = fs.readdirSync(copiedFolder)
        console.log(bidFile, '테스트3') //테스트

        var myFile = bidFile.filter(function (file) {
            return file.substring(file.length - 4, file.length).toLowerCase() === '.bid'
        })[0]
        filename = myFile.substring(0, myFile.length - 4)
        fs.copyFileSync(copiedFolder + '\\' + myFile, copiedFolder + '\\' + filename + '.zip')
        fs.rmSync(copiedFolder + '\\' + myFile)
        var zip = new AdmZip(copiedFolder + '\\' + filename + '.zip')
        zip.extractAllTo(copiedFolder, true)
        bidFile = fs.readdirSync(copiedFolder)
        myFile = bidFile.filter(function (file) {
            return file.substring(file.length - 4, file.length).toLowerCase() === '.bid'
        })[0]
        var text = fs.readFileSync(copiedFolder + '\\' + myFile, 'utf-8')
        var decodeValue = Buffer.from(text, 'base64')
        text = decodeValue.toString('utf-8')
        fs.writeFileSync(Data.folder + '\\OutputDataFromBID.xml', text)
        var xml = fs.readFileSync(Data.folder + '\\OutputDataFromBID.xml', 'utf-8')
        var json = convert.xml2json(xml, { compact: true, spaces: 4 })
        fs.writeFileSync(Data.folder + '\\OutputDataFromBID.json', json)
        fs.rmSync(copiedFolder + '\\' + filename + '.zip')
        fs.rmSync(copiedFolder + '\\' + myFile)
        fs.rmSync(Data.folder + '\\OutputDataFromBID.xml')
    }
    function JsonToBid() {
        var resultFilePath = 'C:\\\\Users\\joung\\OneDrive\\문서\\AutoBID\\OutputDataFromBID.json'
        var json = fs.readFileSync(resultFilePath, 'utf-8')
        var xml = convert.json2xml(json, { compact: true, ignoreComment: true, space: 4 })
        fs.writeFileSync(Data.folder + '\\Result_Xml.xml', xml)
        var text = fs.readFileSync(
            'C:\\\\Users\\joung\\OneDrive\\문서\\AutoBID\\Result_Xml.xml',
            'utf-8'
        )
        var encodeValue = Buffer.from(text, 'utf-8')
        text = encodeValue.toString('base64')
        fs.writeFileSync(Data.folder + '\\EmptyBid\\XmlToBID.BID', text)
        var zip = new AdmZip()
        zip.addLocalFile(Data.folder + '\\EmptyBid\\XmlToBID.BID')
        zip.writeZip(Data.folder + '\\EmptyBid\\' + filename + '.zip')
        fs.copyFileSync(
            Data.folder + '\\EmptyBid\\' + filename + '.zip',
            Data.folder + '\\EmptyBid\\' + filename + '.BID'
        )
        fs.rmSync(Data.folder + '\\EmptyBid\\' + filename + '.zip')
        fs.rmSync(resultFilePath)
        fs.rmSync(Data.folder + '\\Result_Xml.xml')
        fs.rmSync(Data.folder + '\\EmptyBid\\XmlToBID.BID')
    }
    BidToJson()
    //JsonToBid();
})(SetUnitPrice || (SetUnitPrice = {}))
