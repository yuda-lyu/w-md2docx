import fs from 'fs'
import path from 'path'
import zlib from 'zlib'
import { execFileSync } from 'child_process'
import { unzipSync, strFromU8 } from 'fflate'


//長圖與圖名同頁之夾具與量測工具: 長圖PNG、讀圖高、以Word量圖與圖名之頁碼
//note: 沿用 w-html2docx 1.0.34 之 test/tools/imgHeightFixture.mjs 與 figurePages.vbs(其測試資產, 非公開介面, 故複製至本套件)


//crc32, PNG區塊校驗碼
//why: 不依賴zlib.crc32(Node 22.2起才有)
let crcTable = null
function crc32(buf) {
    if (crcTable === null) {
        crcTable = []
        for (let n = 0; n < 256; n++) {
            let c = n
            for (let k = 0; k < 8; k++) {
                c = (c & 1) ? (0xedb88320 ^ (c >>> 1)) : (c >>> 1)
            }
            crcTable.push(c >>> 0)
        }
    }
    let c = 0xffffffff
    for (let i = 0; i < buf.length; i++) {
        c = crcTable[(c ^ buf[i]) & 0xff] ^ (c >>> 8)
    }
    return (c ^ 0xffffffff) >>> 0
}


//genTallPng, 產生單色RGB之PNG, 預設1200×2800px, 模擬整頁高之長圖(網站設計稿、長流程圖)
function genTallPng(w = 1200, h = 2800) {
    let row = Buffer.alloc(1 + w * 3, 220)
    row[0] = 0 //filter none
    let raw = Buffer.concat(Array.from({ length: h }, () => row))
    let chunk = (type, data) => {
        let len = Buffer.alloc(4)
        len.writeUInt32BE(data.length)
        let td = Buffer.concat([Buffer.from(type, 'ascii'), data])
        let crc = Buffer.alloc(4)
        crc.writeUInt32BE(crc32(td))
        return Buffer.concat([len, td, crc])
    }
    let ihdr = Buffer.alloc(13)
    ihdr.writeUInt32BE(w, 0)
    ihdr.writeUInt32BE(h, 4)
    ihdr[8] = 8 //bit depth
    ihdr[9] = 2 //RGB
    let sig = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])
    return Buffer.concat([sig, chunk('IHDR', ihdr), chunk('IDAT', zlib.deflateSync(raw)), chunk('IEND', Buffer.alloc(0))])
}


//getImgHeights, 讀docx內各行內圖片之高度(點), 由word/document.xml之wp:extent cy(EMU)換算
function getImgHeights(fp) {
    let xml = strFromU8(unzipSync(new Uint8Array(fs.readFileSync(fp)))['word/document.xml'])
    return [...xml.matchAll(/<wp:extent cx="(\d+)" cy="(\d+)"/g)].map((m) => Number(m[2]) / 12700)
}


//getFigurePages, 以Word開啟docx(唯讀、不顯示), 回傳各圖片段之頁碼與其後第1個非空段落(圖名)首行、末行之頁碼
function getFigurePages(fp) {
    let fpVbs = path.resolve('./test/tools/figurePages.vbs')
    let out = execFileSync('cscript', ['//nologo', fpVbs, path.resolve(fp)], { encoding: 'utf8', windowsHide: true, timeout: 180000 })
    return out.trim().split(/\r?\n/).filter((v) => v !== '').map((v) => {
        let [img, capStart, capEnd] = v.split(',').map(Number)
        return { img, capStart, capEnd }
    })
}


export {
    genTallPng,
    getImgHeights,
    getFigurePages
}
