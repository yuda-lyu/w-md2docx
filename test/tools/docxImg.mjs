import fs from 'fs'
import os from 'os'
import path from 'path'
import { unzipSync, strFromU8 } from 'fflate'
import { genTallPng } from './imgFixture.mjs'


//圖片相關測試之共用夾具與量測工具: 圖片夾具、讀 docx 內圖片資訊、讀 html 內 img 之 src、本行程暫存中介 html 計數


//mkFixture: 於fd下建圖片夾具(夾名含中文與空白): 圖 資料夾/a.png、圖 資料夾/b.png、c.png、n.png、cocktail.svg; 內容各異(灰階值不同), 大小不同
//  a、c為小於版心且帶96 dpi之pHYs(新舊方案顯示尺寸相同), b為大於版心(兩方案皆縮至版心寬), n為小於版心且無pHYs(新方案以96 dpi顯示, 見 README 之 Breaking changes)
function mkFixture(fd) {
    fs.mkdirSync(path.resolve(fd, '圖 資料夾'), { recursive: true })
    fs.writeFileSync(path.resolve(fd, '圖 資料夾', 'a.png'), genTallPng(40, 20, 60, 96))
    fs.writeFileSync(path.resolve(fd, '圖 資料夾', 'b.png'), genTallPng(800, 450, 120))
    fs.writeFileSync(path.resolve(fd, 'c.png'), genTallPng(300, 150, 200, 96))
    fs.writeFileSync(path.resolve(fd, 'n.png'), genTallPng(40, 20, 90))
    fs.copyFileSync('./test/cocktail.svg', path.resolve(fd, 'cocktail.svg'))
}


//scanDocx: 讀docx之圖片資訊: 各圖(w:drawing)之顯示尺寸(EMU)、r:embed與r:link、內嵌之媒體內容; word/media檔名; 圖片關聯中TargetMode="External"數; 圖之r:link數
function scanDocx(fp) {
    let z = unzipSync(new Uint8Array(fs.readFileSync(fp)))
    let doc = strFromU8(z['word/document.xml'])
    let rels = strFromU8(z['word/_rels/document.xml.rels'])
    let kp = {}
    for (let m of rels.matchAll(/<Relationship\b[^>]*?\/?>/g)) {
        let id = (m[0].match(/\bId="([^"]+)"/) || [])[1]
        kp[id] = {
            target: (m[0].match(/\bTarget="([^"]*)"/) || [])[1],
            image: /relationships\/image"/.test(m[0]),
            external: /TargetMode="External"/.test(m[0]),
        }
    }
    let pics = []
    for (let m of doc.matchAll(/<w:drawing>([\s\S]*?)<\/w:drawing>/g)) {
        let ext = m[1].match(/<wp:extent cx="(\d+)" cy="(\d+)"/)
        let emb = (m[1].match(/<a:blip\b[^>]*\br:embed="([^"]+)"/) || [])[1]
        let lnk = (m[1].match(/<a:blip\b[^>]*\br:link="([^"]+)"/) || [])[1]
        pics.push({
            cx: Number(ext[1]),
            cy: Number(ext[2]),
            embed: emb || '',
            link: lnk || '',
            media: emb && kp[emb] ? z[`word/${kp[emb].target}`] || null : null,
        })
    }
    return {
        pics,
        media: Object.keys(z).filter((v) => v.startsWith('word/media/')).map((v) => v.slice('word/media/'.length)),
        nExternalImg: Object.values(kp).filter((v) => v.image && v.external).length,
        nLinkAttr: (doc.match(/<a:blip\b[^>]*\br:link="/g) || []).length,
    }
}


//imgSrcs: 取html內各img之src屬性值
function imgSrcs(html) {
    return [...html.matchAll(/<img\b[^>]*?\ssrc="([^"]*)"/g)].map((m) => m[1])
}


//nTmpHtmlOfThisProcess: 系統暫存夾內本行程(mocha worker)產生之中介html數(檔名前綴 wmd2docx_<pid>_), 不計並行之他測試檔所產者
function nTmpHtmlOfThisProcess() {
    return fs.readdirSync(os.tmpdir()).filter((v) => v.startsWith(`wmd2docx_${process.pid}_`)).length
}


export {
    mkFixture,
    scanDocx,
    imgSrcs,
    nTmpHtmlOfThisProcess
}
