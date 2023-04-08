//const fs = require('fs')
//const AdmZip = require('adm-zip')
const { buffer } = require('stream/consumers')
const convert = require('xml-js')
//const Data = require('./Data.js')

let filename = undefined
class BidHandling {
    public static BidToJson() {
        const copiedFolder: string = Data.folder + '\\EmptyBid'
        let bidFile = fs.readdirSync(copiedFolder)
        let myFile = bidFile.filter(
            (file) => file.substring(file.length - 4, file.length).toLowerCase() === '.bid'
        )[0]
        filename = myFile.substring(0, myFile.length - 4)

        fs.copyFileSync(copiedFolder + '\\' + myFile, copiedFolder + '\\' + filename + '.zip')
        fs.rmSync(copiedFolder + '\\' + myFile)

        let zip = new AdmZip(copiedFolder + '\\' + filename + '.zip')
        zip.extractAllTo(copiedFolder, true)

        bidFile = fs.readdirSync(copiedFolder)
        myFile = bidFile.filter(
            (file) => file.substring(file.length - 4, file.length).toLowerCase() === '.bid'
        )[0]

        let text = fs.readFileSync(copiedFolder + '\\' + myFile, 'utf-8')
        const decodeValue = Buffer.from(text, 'base64')
        text = decodeValue.toString('utf-8')

        fs.writeFileSync(Data.folder + '\\OutputDataFromBID.xml', text)

        const xml = fs.readFileSync(Data.folder + '\\OutputDataFromBID.xml', 'utf-8')
        const json = convert.xml2json(xml, { compact: true, spaces: 4 })

        fs.writeFileSync(Data.folder + '\\OutputDataFromBID.json', json)

        fs.rmSync(copiedFolder + '\\' + filename + '.zip')
        fs.rmSync(copiedFolder + '\\' + myFile)
        fs.rmSync(Data.folder + '\\OutputDataFromBID.xml')
    }

    public static JsonToBid() {
        const resultFilePath = 'C:\\\\Users\\joung\\OneDrive\\문서\\AutoBID\\OutputDataFromBID.json'
        const json = fs.readFileSync(resultFilePath, 'utf-8')
        const xml = convert.json2xml(json, { compact: true, ignoreComment: true, space: 4 })

        fs.writeFileSync(Data.folder + '\\Result_Xml.xml', xml)

        let text = fs.readFileSync(
            'C:\\\\Users\\joung\\OneDrive\\문서\\AutoBID\\Result_Xml.xml',
            'utf-8'
        )
        const encodeValue = Buffer.from(text, 'utf-8')
        text = encodeValue.toString('base64')

        fs.writeFileSync(Data.folder + '\\EmptyBid\\XmlToBID.BID', text)

        let zip = new AdmZip()
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
}
module.exports = BidHandling

// BidHandling.BidToJson();
// BidHandling.JsonToBid();
