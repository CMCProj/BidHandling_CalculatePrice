//const { buffer } = require('stream/consumers')
import * as convert from 'xml-js'
import * as fs from 'fs'
const AdmZip = require('adm-zip')
import { Setting } from './Setting'
import { Data } from './Data'
//실행 위해 프로그램 내 폴더로 경로 변경

let filename = undefined

export class BidHandling {
    public static BidToJson(): void {
        const copiedFolder: string = Data.folder + '\\EmptyBid' // EmptyBid폴더 주소 저장 / 폴더 경로 수정 (23.02.02)
        let bidFile = fs.readdirSync(copiedFolder)
        let myFile = bidFile.filter(
            (file) => file.substring(file.length - 4, file.length).toLowerCase() === '.bid'
        )[0] // 확장자가 .bid인 파일을 찾기
        filename = myFile.substring(0, myFile.length - 4) // 확장자를 뺀 파일의 이름

        fs.copyFileSync(copiedFolder + '\\' + myFile, copiedFolder + '\\' + filename + '.zip') // 확장자를 .bid에서 .zip으로 교체
        fs.rmSync(copiedFolder + '\\' + myFile) // 기존의 .bid 파일은 삭제

        let zip = new AdmZip(copiedFolder + '\\' + filename + '.zip') // .zip파일 압축 해제
        zip.extractAllTo(copiedFolder, true)

        bidFile = fs.readdirSync(copiedFolder)
        myFile = bidFile.filter(
            (file) => file.substring(file.length - 4, file.length).toLowerCase() === '.bid'
        )[0] // 압축 해제되어 나온 .bid파일 찾기

        let text = fs.readFileSync(copiedFolder + '\\' + myFile, 'utf-8') // 나온 .bid 파일의 텍스트 읽기
        const decodeValue = Buffer.from(text, 'base64') // base64코드 디코딩
        text = decodeValue.toString('utf-8') // 디코딩 되어 나온 텍스트 저장

        fs.writeFileSync(Data.folder + '\\OutputDataFromBID.xml', text) // xml파일 형식으로 텍스트 쓰기

        const xml = fs.readFileSync(Data.folder + '\\OutputDataFromBID.xml', 'utf-8')
        const json = convert.xml2json(xml, { compact: true, spaces: 4 }) // xml파일을 json으로 교체

        fs.writeFileSync(Data.folder + '\\OutputDataFromBID.json', json) // json파일 형식으로 다시 쓰기

        //=======이 과정에서 만들어진 파일들은 전부 삭제=======
        fs.rmSync(copiedFolder + '\\' + filename + '.zip')
        fs.rmSync(copiedFolder + '\\' + myFile)
        fs.rmSync(Data.folder + '\\OutputDataFromBID.xml')

        Setting.GetData()
    }

    public static JsonToBid(): void {
        const resultFilePath = Data.folder + '\\OutputDataFromBID.json'
        const json = fs.readFileSync(resultFilePath, 'utf-8') // json파일 읽기
        const xml = convert.json2xml(json, { compact: true, ignoreComment: true, spaces: 4 }) // json파일을 xml파일로 교체

        fs.writeFileSync(Data.folder + '\\Result_Xml.xml', xml) // xml파일 형식으로 다시 쓰기

        let text = fs.readFileSync(Data.folder + '\\Result_Xml.xml', 'utf-8') // xml파일 내용 읽기
        const encodeValue = Buffer.from(text, 'utf-8')
        text = encodeValue.toString('base64') // base64로 인코딩

        fs.writeFileSync(Data.folder + '\\EmptyBid\\XmlToBID.BID', text) // 인코딩된 텍스트 XmlToBID.BID파일에 쓰기

        let zip = new AdmZip()
        zip.addLocalFile(Data.folder + '\\EmptyBid\\XmlToBID.BID')
        zip.writeZip(Data.folder + '\\EmptyBid\\' + filename + '.zip') // XmlToBID.BID파일을 .zip파일로 압축

        fs.copyFileSync(
            Data.folder + '\\EmptyBid\\' + filename + '.zip',
            Data.folder + '\\EmptyBid\\' + filename + '.BID'
        )

        //======이 과정에서 만들어진 파일들은 전부 삭제======
        fs.rmSync(Data.folder + '\\EmptyBid\\' + filename + '.zip')
        fs.rmSync(resultFilePath)
        fs.rmSync(Data.folder + '\\Result_Xml.xml')
        fs.rmSync(Data.folder + '\\EmptyBid\\XmlToBID.BID')
    }
}
