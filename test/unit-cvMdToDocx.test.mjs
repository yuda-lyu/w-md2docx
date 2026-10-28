import fs from 'fs'
import path from 'path'
import zlib from 'zlib'
import { pathToFileURL } from 'url'
import { execFileSync } from 'child_process'
import assert from 'assert'
import WMd2html from 'w-md2html/src/WMd2html.mjs'
import WHtml2docx from 'w-html2docx/src/WHtml2docx.mjs'
import cvMdToDocx from '../src/cvMdToDocx.mjs'
import cvMdTo from '../src/cvMdTo.mjs'
import { runExclusive } from '../src/utils.mjs'
import hasWord from './tools/hasWord.mjs'
import runWithFakeHtml2docx from './tools/fakeHtml2docx.mjs'
import { parts, buildDocx, readPart, listParas } from './tools/docxFixture.mjs'
import { genTallPng, getImgHeights, getFigurePages } from './tools/imgFixture.mjs'
import { mkFixture, scanDocx, imgSrcs, nTmpHtmlOfThisProcess } from './tools/docxImg.mjs'


//fdTmpRoot: 本測試檔專用暫存夾, 各describe於其下自建子夾並於after刪除
let fdTmpRoot = path.resolve('./test/_tmp/unit-cvMdToDocx')


//readZipEntries: 讀取zip(docx)內各檔案, 回傳{name:Buffer}
//why: 以central directory解析, 不引入額外套件; docx為標準deflate/stored, 無zip64
function readZipEntries(buf) {
    let iE = buf.lastIndexOf(Buffer.from([0x50, 0x4b, 0x05, 0x06]))
    assert.strict.ok(iE >= 0, 'zip end of central directory not found')
    let nEntries = buf.readUInt16LE(iE + 10)
    let p = buf.readUInt32LE(iE + 16)
    let kp = {}
    for (let i = 0; i < nEntries; i++) {
        assert.strict.equal(buf.readUInt32LE(p), 0x02014b50, 'central directory header signature')
        let method = buf.readUInt16LE(p + 10)
        let compSize = buf.readUInt32LE(p + 20)
        let nameLen = buf.readUInt16LE(p + 28)
        let extraLen = buf.readUInt16LE(p + 30)
        let commentLen = buf.readUInt16LE(p + 32)
        let offLocal = buf.readUInt32LE(p + 42)
        let name = buf.toString('utf8', p + 46, p + 46 + nameLen)
        let lNameLen = buf.readUInt16LE(offLocal + 26)
        let lExtraLen = buf.readUInt16LE(offLocal + 28)
        let start = offLocal + 30 + lNameLen + lExtraLen
        let data = buf.subarray(start, start + compSize)
        kp[name] = (method === 8) ? zlib.inflateRawSync(data) : Buffer.from(data)
        p += 46 + nameLen + extraLen + commentLen
    }
    return kp
}


describe('cvMdToDocx 輸入檢查(不需Word)', function() {

    let fpOut = path.resolve(fdTmpRoot, 'input', 'x.docx') //輸入檢查皆於寫檔前拒絕, 不會產生檔案

    it('fpInMd非有效字串', async function() {
        await assert.rejects(cvMdToDocx('', fpOut), (e) => e === 'fpInMd must be a non-empty string')
        await assert.rejects(cvMdToDocx(null, fpOut), (e) => e === 'fpInMd must be a non-empty string')
    })

    it('fpInMd不存在', async function() {
        await assert.rejects(cvMdToDocx('./test/nope.md', fpOut), (e) => /^fpInMd\[.+nope\.md\] does not exist$/.test(e))
    })

    it('fpOutDocx非有效字串', async function() {
        await assert.rejects(cvMdToDocx('./test/report.md', ''), (e) => e === 'fpOutDocx must be a non-empty string')
    })

    it('fpInTemp不存在', async function() {
        await assert.rejects(cvMdToDocx('./test/report.md', fpOut, { fpInTemp: './test/nope.docx' }), (e) => /^fpInTemp\[.+nope\.docx\] does not exist$/.test(e))
    })

})


describe('cvMdToDocx 實轉(需Windows+Microsoft Word)', function() {

    //check
    if (!hasWord()) {
        console.log('[unit-cvMdToDocx] Microsoft Word not available on this machine, skip real conversion tests')
        return
    }

    let fdTmp = path.resolve(fdTmpRoot, 'real')
    let fpIn = './test/report.md'
    let fpOut = path.resolve(fdTmp, 'report.docx')
    let fpOutHtml = path.resolve(fdTmp, 'report.html')
    let fpInTemp = './src/templates/temp_tpc.docx'
    let rt = null
    let zip = null
    let xml = ''

    before(async function() {
        this.timeout(180000)
        fs.mkdirSync(fdTmp, { recursive: true })
        rt = await cvMdToDocx(fpIn, fpOut, {
            fpInTemp,
            fpOutHtml,
            optMd2html: { imgWidthMax: '500px' },
            optHtml2docx: { imgRatioWidthMax: 0.5 },
        })
        zip = readZipEntries(fs.readFileSync(fpOut))
        xml = zip['word/document.xml'].toString('utf8')
    })

    after(function() {
        fs.rmSync(fdTmp, { recursive: true, force: true, maxRetries: 10, retryDelay: 300 })
    })

    it('回傳物件: fpOutDocx/sizeDocx/sizeHtml/ms/fpOutHtml與實體檔一致', function() {
        assert.strict.equal(rt.fpOutDocx, fpOut)
        assert.strict.equal(rt.sizeDocx, fs.statSync(fpOut).size)
        assert.strict.equal(rt.sizeDocx > 0, true)
        assert.strict.equal(rt.fpOutHtml, fpOutHtml)
        assert.strict.equal(rt.sizeHtml, fs.statSync(fpOutHtml).size)
        assert.strict.equal(typeof rt.ms, 'number')
    })

    it('docx為zip且含主文件', function() {
        assert.strict.equal(fs.readFileSync(fpOut).subarray(0, 2).toString('latin1'), 'PK')
        assert.strict.equal(typeof zip['word/document.xml'], 'object')
        assert.strict.equal(typeof zip['[Content_Types].xml'], 'object')
    })

    it('主文件含來源md之標題與清單文字', function() {
        assert.strict.equal(xml.includes('計畫報告'), true) //h1
        assert.strict.equal(xml.includes('資料流程自動化'), true) //li
    })

    it('內文圖片已嵌入(word/media)', function() {
        let medias = Object.keys(zip).filter((v) => v.startsWith('word/media/'))
        assert.strict.equal(medias.length > 0, true)
    })

    it('模板settings之doNotAutoCompressPictures保留', function() {
        let settings = zip['word/settings.xml'].toString('utf8')
        assert.strict.equal(settings.includes('<w:doNotAutoCompressPictures'), true)
    })

    it('未指定fpOutHtml時不留中介檔, 回傳無fpOutHtml', async function() {
        this.timeout(180000)
        let fpOut2 = path.resolve(fdTmp, 'report2.docx')
        let r = await cvMdToDocx(fpIn, fpOut2, { fpInTemp })
        assert.strict.equal(r.fpOutHtml, undefined)
        assert.strict.equal(fs.existsSync(fpOut2), true)
        assert.strict.equal(fs.existsSync(fpOut2.replace(/\.docx$/, '.html')), false)
    })

    it('並發兩份轉檔皆成功(內部佇列序列化)', async function() {
        this.timeout(300000)
        let [ra, rb] = await Promise.all([
            cvMdToDocx(fpIn, path.resolve(fdTmp, 'a.docx'), { fpInTemp }),
            cvMdToDocx(fpIn, path.resolve(fdTmp, 'b.docx'), { fpInTemp }),
        ])
        assert.strict.equal(ra.sizeDocx > 0, true)
        assert.strict.equal(rb.sizeDocx > 0, true)
    })

    it('cvMdTo out非法值(如大小寫不符)退回docx', async function() {
        this.timeout(180000)
        let r = await cvMdTo({ md: '# t', out: 'HTML', fpInTemp, dirWork: fdTmp })
        assert.strict.equal(r.out, 'docx')
        assert.strict.equal(r.html, undefined)
        assert.strict.equal(r.docx.size > 0, true)
    })

    it('toc:true: 添加目錄並重編頁碼, 結果含toc且docx分為封面、目錄、正文3節', async function() {
        this.timeout(300000)
        let fpMd = path.resolve(fdTmp, 'toc.md')
        fs.writeFileSync(fpMd, [
            '<div pretitle style="font-size:14pt; text-align:center;">',
            '    目錄測試',
            '</div>',
            '',
            '# 第一章 緒論',
            '',
            '內容。',
            '',
            '## 1.1 背景',
            '',
            '表n 參數一覽',
            '',
            '| 項目 | 數值 |',
            '|---|---|',
            '| a | 1 |',
            '',
            '# 第二章 方法',
            '',
            '內容。',
            '',
        ].join('\n'), 'utf8')
        let fpOutToc = path.resolve(fdTmp, 'toc.docx')
        let r = await cvMdToDocx(fpMd, fpOutToc, { fpInTemp, toc: true })
        assert.strict.equal(r.toc.skip, '')
        assert.strict.equal(r.toc.cover, true)
        assert.strict.deepEqual(r.toc.lists, ['toc', 'tab'])
        assert.strict.equal(r.toc.toc, 3)
        assert.strict.equal(r.toc.tab, 1)
        assert.strict.equal(r.sizeDocx, fs.statSync(fpOutToc).size)
        let x = readZipEntries(fs.readFileSync(fpOutToc))['word/document.xml'].toString('utf8')
        assert.strict.equal((x.match(/<w:sectPr\b/g) || []).length, 3)
        //目錄欄位之\o自正文最高之標題階層起, 涵蓋md之兩層標題(#、##)
        //note: 階層以docx實際之大綱階層為準, Word匯入html時h1、h2為段落直設之大綱階層(Word 16實測為2、3, 非1、2), 故不寫死
        let pH1 = listParas(x).find((p) => p.text === '第一章 緒論')
        let l0 = Number(pH1.xml.match(/<w:outlineLvl w:val="(\d)"\/>/)[1]) + 1
        assert.strict.deepEqual(r.toc.levels, [l0, l0 + 1])
        assert.strict.equal(x.includes(`> TOC \\o "${l0}-${l0 + 1}" \\h \\z \\u </w:instrText>`), true)
    })

    it('keepCaption:true: 圖片段與表名設與下段同頁, 交付之docx(由fflate寫出)可由Word開啟另存', async function() {
        this.timeout(300000)
        fs.copyFileSync('./test/cocktail.svg', path.resolve(fdTmp, 'cocktail.svg'))
        let fpMd = path.resolve(fdTmp, 'keep.md')
        fs.writeFileSync(fpMd, [
            '# 第一章 圖表',
            '',
            '![p](cocktail.svg)',
            '',
            '圖n 雞尾酒',
            '',
            '表n 參數一覽',
            '',
            '| 項目 | 數值 |',
            '|---|---|',
            '| a | 1 |',
            '',
        ].join('\n'), 'utf8')
        let fpOutKeep = path.resolve(fdTmp, 'keep.docx')
        let r = await cvMdToDocx(fpMd, fpOutKeep, { fpInTemp, keepCaption: true })
        assert.strict.deepEqual([r.keepCaption.figs, r.keepCaption.tabs, r.keepCaption.changed], [1, 1, true])
        let ps = listParas(readPart(new Uint8Array(fs.readFileSync(fpOutKeep)), 'word/document.xml'))
        let iFig = ps.findIndex((p) => p.text.startsWith('圖1'))
        assert.strict.deepEqual([ps[iFig - 1].hasPic, ps[iFig - 1].keepNext, ps[iFig].keepLines], [true, true, true])
        let pTab = ps.find((p) => p.text.startsWith('表1'))
        assert.strict.deepEqual([pTab.keepNext, pTab.keepLines], [true, true])
        //Word 開啟另存: 開檔失敗(檔案損毀)時 updateDocxToc.vbs 以離開碼 1 結束, execFileSync 拋錯
        let out = execFileSync('cscript', ['//nologo', path.resolve('./src/updateDocxToc.vbs'), fpOutKeep, path.resolve(fdTmp, 'keep-resaved.docx')], { encoding: 'utf8', windowsHide: true })
        assert.strict.equal(/(^|\n)ok\s*$/.test(out.trim()), true, out)
    })

    it('整頁高之長圖: 預設預留4格使圖高為(38-4)×18-0.5點, 連續長圖各與其圖名同頁(未開同頁、keepCaption、toc皆然)', async function() {
        this.timeout(600000)
        let fd = path.resolve(fdTmp, 'tall')
        fs.mkdirSync(fd, { recursive: true })
        fs.writeFileSync(path.resolve(fd, 'tall.png'), genTallPng())
        //md: 前言、4個圖區塊(長圖+2行之圖名), 圖區塊間由撰文者寫入w-md2html之換行標記(空段落)
        let cap = (i) => `圖${i} 用過核子燃料最終處置資訊展示網站之首頁設計稿，含導覽列、主視覺、最新消息、計畫沿革、各階段成果展示與頁尾聯絡資訊之完整版面配置示意`
        let lines = ['前言段落。', '']
        for (let i = 1; i <= 4; i++) {
            lines.push('![](tall.png)', '', cap(i), '')
            if (i < 4) {
                lines.push('<div data-for="br"></div>', '')
            }
        }
        lines.push('結尾段落。', '')
        let fpMd = path.resolve(fd, 'tall.md')
        let fpMdToc = path.resolve(fd, 'tallToc.md')
        fs.writeFileSync(fpMd, lines.join('\n'), 'utf8')
        fs.writeFileSync(fpMdToc, ['封面。', '', '# 第一章 長圖', '', ...lines].join('\n'), 'utf8')
        for (let [name, fp, opt] of [['plain', fpMd, {}], ['keep', fpMd, { keepCaption: true }], ['toc', fpMdToc, { toc: true }]]) {
            let fpOut = path.resolve(fd, `${name}.docx`)
            await cvMdToDocx(fp, fpOut, { fpInTemp, ...opt })
            //temp_tpc.docx: A4、上下邊界72點、行距18點之行格線, 版心高697.9點即每頁38格, 預留4格之圖高上限為(38-4)×18-0.5=611.5點
            let hs = getImgHeights(fpOut)
            assert.strict.equal(hs.length, 4, name)
            assert.strict.equal(hs.every((h) => Math.abs(h - 611.5) < 0.05), true, `${name}: ${hs}`)
            //Word排版後各圖片段與其後之圖名(首行、末行)同頁
            let pgs = getFigurePages(fpOut)
            assert.strict.deepEqual(pgs.map((v) => v.img === v.capStart && v.img === v.capEnd), [true, true, true, true], `${name}: ${JSON.stringify(pgs)}`)
        }
    })

    it('cvMdTo out=both: html與docx內容皆回傳且msDocx存在', async function() {
        this.timeout(180000)
        let r = await cvMdTo({ md: '# 內容轉檔\n\n段落', name: '內容', out: 'both', fpInTemp, dirWork: fdTmp })
        assert.strict.equal(r.out, 'both')
        assert.strict.equal(r.html.fileName, '內容.html')
        assert.strict.equal(r.docx.fileName, '內容.docx')
        assert.strict.equal(r.docx.mime, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document')
        assert.strict.equal(r.docx.size, Buffer.from(r.docx.base64, 'base64').length)
        assert.strict.equal(typeof r.msDocx, 'number')
        let z = readZipEntries(Buffer.from(r.docx.base64, 'base64'))
        assert.strict.equal(z['word/document.xml'].toString('utf8').includes('內容轉檔'), true)
        assert.strict.equal(fs.readdirSync(fdTmp).filter((v) => v.startsWith('job_')).length, 0) //作業夾已清除
    })

})


//圖片內嵌於docx: 中介html以本機絕對路徑引用圖, Word讀檔後由w-html2docx斷開連結內嵌(不依賴原圖)
//why: 與上方〈實轉〉同檔依序執行——Word為單一實例, 並行之 mocha 行程同時調用會互相拒絕(com_error -2147418111 接收者已拒絕這個呼叫), 故調用 Word 之測試不另開檔
describe('cvMdToDocx 圖片內嵌於docx(需Windows+Microsoft Word)', function() {

    //check
    if (!hasWord()) {
        console.log('[unit-cvMdToDocx] Microsoft Word not available on this machine, skip image embedding tests')
        return
    }

    let fd = path.resolve(fdTmpRoot, 'img')
    let fpInTemp = path.resolve('./src/templates/temp_tpc.docx')
    let fpMdP = path.resolve(fd, 'png.md') //4張png: md語法圖(角括號路徑)、原始<img>、相對路徑、無pHYs之小圖
    let fpMdS = path.resolve(fd, 'svg.md') //1張svg
    let scanP = null
    let scanS = null
    let scanPb64 = null
    let htmlP = ''

    //convBase64: 對照組, 即過去之路徑: md -> html(w-md2html 預設 base64 內嵌) -> docx, 同 cvMdToDocx 傳予 w-html2docx 之設定
    let convBase64 = async (fpMd, name) => {
        let fpHtml = path.resolve(fd, `${name}.html`)
        let fpDocx = path.resolve(fd, `${name}.docx`)
        await WMd2html(fpMd, fpHtml, {})
        await runExclusive(() => WHtml2docx(fpHtml, fpDocx, { fpInTemp, imgHeightReserveLines: 4 }))
        return fpDocx
    }

    before(async function() {
        this.timeout(600000)
        fs.mkdirSync(fd, { recursive: true })
        mkFixture(fd)
        fs.writeFileSync(fpMdP, [
            '# 四張png',
            '',
            '![甲](<圖 資料夾/a.png>)',
            '',
            '<img src="圖 資料夾/b.png" alt="乙">',
            '',
            '![丙](c.png)',
            '',
            '![戊](n.png)',
            '',
        ].join('\n'), 'utf8')
        fs.writeFileSync(fpMdS, '# 一張svg\n\n![丁](cocktail.svg)\n', 'utf8')
        //optMd2html 給 imgConvertToBase64:true 亦不得使中介 html 內嵌 base64
        let rP = await cvMdToDocx(fpMdP, path.resolve(fd, 'png-path.docx'), { fpInTemp, fpOutHtml: path.resolve(fd, 'png-path.html'), optMd2html: { imgConvertToBase64: true } })
        let rS = await cvMdToDocx(fpMdS, path.resolve(fd, 'svg-path.docx'), { fpInTemp })
        scanP = scanDocx(rP.fpOutDocx)
        scanS = scanDocx(rS.fpOutDocx)
        htmlP = fs.readFileSync(path.resolve(fd, 'png-path.html'), 'utf8')
    })

    after(function() {
        fs.rmSync(fd, { recursive: true, force: true, maxRetries: 10, retryDelay: 300 })
    })

    it('中介html無data:圖(圖以本機絕對路徑引用)', function() {
        assert.strict.equal(htmlP.includes('data:'), false)
        assert.strict.equal(imgSrcs(htmlP).length, 4)
    })

    it('docx之word/media圖數=md之圖數(4), 圖(w:drawing)皆為r:embed內嵌, 且依文件順序為來源圖之原檔內容', function() {
        assert.strict.equal(scanP.media.length, 4)
        assert.strict.equal(scanP.pics.length, 4)
        assert.strict.equal(scanP.pics.every((v) => v.embed !== ''), true)
        let srcs = [
            fs.readFileSync(path.resolve(fd, '圖 資料夾', 'a.png')),
            fs.readFileSync(path.resolve(fd, '圖 資料夾', 'b.png')),
            fs.readFileSync(path.resolve(fd, 'c.png')),
            fs.readFileSync(path.resolve(fd, 'n.png')),
        ]
        //md內順序: a(md語法)、b(原始<img>)、c(md語法)、n(md語法); docx內圖依文件順序
        for (let i = 0; i < 4; i++) {
            assert.strict.equal(scanP.pics[i].media !== null && Buffer.from(scanP.pics[i].media).equals(srcs[i]), true, `pic ${i}`)
        }
    })

    it('docx之圖片關聯無TargetMode="External", 圖無r:link(不依賴原圖, 可移至他處)', function() {
        assert.strict.equal(scanP.nExternalImg, 0)
        assert.strict.equal(scanP.nLinkAttr, 0)
        assert.strict.equal(scanS.nExternalImg, 0)
        assert.strict.equal(scanS.nLinkAttr, 0)
    })

    it('svg: 圖(w:drawing)=md之圖數(1)且為內嵌, media含svg與其點陣預覽', function() {
        assert.strict.equal(scanS.pics.length, 1)
        assert.strict.equal(scanS.pics[0].embed !== '', true)
        assert.strict.equal(scanS.media.filter((v) => /\.svg$/i.test(v)).length, 1)
        assert.strict.equal(scanS.media.filter((v) => /\.png$/i.test(v)).length >= 1, true)
    })

    it('與過去之base64方案同輪雙產比對: 圖數與順序相同, 顯示尺寸差<=0.5點(帶96 dpi之小圖、大於版心之圖、svg)', async function() {
        this.timeout(600000)
        scanPb64 = scanDocx(await convBase64(fpMdP, 'png-b64'))
        let b64S = scanDocx(await convBase64(fpMdS, 'svg-b64'))
        //png 之第4張(n, 無pHYs之小圖)尺寸依 96 dpi 而不同, 見下一測試
        for (let [pth, b64, name, n] of [[scanP, scanPb64, 'png', 3], [scanS, b64S, 'svg', 1]]) {
            assert.strict.equal(pth.pics.length, b64.pics.length, name)
            for (let i = 0; i < n; i++) {
                assert.strict.equal(Math.abs(pth.pics[i].cx - b64.pics[i].cx) <= 0.5 * 12700, true, `${name} ${i} cx ${pth.pics[i].cx} vs ${b64.pics[i].cx}`)
                assert.strict.equal(Math.abs(pth.pics[i].cy - b64.pics[i].cy) <= 0.5 * 12700, true, `${name} ${i} cy ${pth.pics[i].cy} vs ${b64.pics[i].cy}`)
            }
            assert.strict.equal(b64.nExternalImg, 0, name)
        }
    })

    it('無pHYs且小於版心之png: 路徑方案依96 dpi顯示(每像素0.75點), 過去之base64方案為每像素1點(README之Breaking changes 2)', async function() {
        this.timeout(600000)
        if (scanPb64 === null) {
            scanPb64 = scanDocx(await convBase64(fpMdP, 'png-b64'))
        }
        //n.png 為 40×20 像素
        let pt = (emu) => emu / 12700
        assert.strict.equal(Math.abs(pt(scanP.pics[3].cx) - 30) <= 0.5 && Math.abs(pt(scanP.pics[3].cy) - 15) <= 0.5, true, `path ${pt(scanP.pics[3].cx)} x ${pt(scanP.pics[3].cy)}`)
        assert.strict.equal(Math.abs(pt(scanPb64.pics[3].cx) - 40) <= 0.5 && Math.abs(pt(scanPb64.pics[3].cy) - 20) <= 0.5, true, `base64 ${pt(scanPb64.pics[3].cx)} x ${pt(scanPb64.pics[3].cy)}`)
    })

    it('cvMdTo docx: 資產路徑為URL編碼(中文檔名、空白)或含&與%23時, 圖亦解析並內嵌於docx(服務端與用戶端之常見寫法)', async function() {
        this.timeout(300000)
        let b64 = genTallPng(300, 150, 77, 96).toString('base64')
        let r = await cvMdTo({
            md: '# t\n\n![x](pics/%E5%9C%96%20A.png)\n\n<img src="pics/圖 B.png">\n\n![y](pics/a&b.png)\n\n![z](pics/c%23d.png)\n',
            name: 'enc',
            out: 'docx',
            fpInTemp,
            dirWork: path.resolve(fd, 'work-enc'),
            assets: ['pics/%E5%9C%96%20A.png', 'pics/圖 B.png', 'pics/a&b.png', 'pics/c%23d.png'].map((p) => ({ path: p, base64: b64 })),
        })
        let fpEnc = path.resolve(fd, 'enc.docx')
        fs.writeFileSync(fpEnc, Buffer.from(r.docx.base64, 'base64'))
        let sc = scanDocx(fpEnc)
        assert.strict.equal(sc.pics.length, 4)
        assert.strict.equal(sc.pics.every((v) => v.embed !== ''), true)
        assert.strict.equal(sc.nExternalImg, 0)
    })

    it('圖片解析不到(網址不可達): reject並帶w-html2docx列出之來源; 不產出docx、不留中介html', async function() {
        this.timeout(300000)
        let fpMdBad = path.resolve(fd, 'bad.md')
        let fpDocx = path.resolve(fd, 'bad.docx')
        fs.writeFileSync(fpMdBad, '# t\n\n![甲](c.png)\n\n<img src="http://127.0.0.1:1/x.png">\n', 'utf8')
        let n0 = nTmpHtmlOfThisProcess()
        let err = null
        await cvMdToDocx(fpMdBad, fpDocx, { fpInTemp }).catch((e) => {
            err = e
        })
        assert.strict.equal(/^Failed to convert html to docx: code\[1\]:\r?\nunresolved linked picture: http:\/\/127\.0\.0\.1:1\/x\.png\r?\nerror: /.test(String(err)), true, String(err))
        assert.strict.equal(fs.existsSync(fpDocx), false)
        assert.strict.equal(nTmpHtmlOfThisProcess(), n0)
    })

})


//轉檔器 htmlToDocx.exe 之取得由 w-html2docx 負責(經 ensureBins, 快取位置為其套件根之 bin-cache 或 W_BIN_CACHE, 不依 cwd), 上層不得以自身規則事前攔截
//預設模板則由 w-html2docx 以 cwd 查找(WHtml2docx.mjs 之 fdSrv)
describe('cvMdToDocx 轉檔器取得失敗與預設模板查找失敗之訊息(不需Word)', function() {

    //check, htmlToDocx.exe僅供Windows, 非Windows時w-html2docx於取得轉檔器前即拒絕
    if (process.platform !== 'win32') {
        return
    }

    let fpInTemp = path.resolve('./src/templates/temp_tpc.docx')
    let msgHint = 'w-html2docx resolves its default template from cwd, please run from the project root that contains node_modules/w-html2docx or give fpInTemp'
    let msgHintBin = 'the external executable htmlToDocx.exe of w-html2docx is not available, see the section "外部執行檔" of the README of w-html2docx for W_BIN_CACHE, W_BIN_MIRROR, W_BIN_OFFLINE and WHtml2docx.ensureBins()'

    after(function() {
        fs.rmSync(fdTmpRoot, { recursive: true, force: true, maxRetries: 10, retryDelay: 300 })
    })

    it('轉檔器取不到(離線且快取空)時不於上層中斷, 由w-html2docx之[E_OFFLINE_MISSING]原訊息reject並附README指引, 不產出docx且中介html已清除', function() {
        this.timeout(180000)

        //fdCwd: 以子程序於此cwd啟動, 快取指向空夾並設離線, 使 w-html2docx 之 ensureBins 確定失敗(不需連網)
        let fdCwd = path.resolve(fdTmpRoot, 'binfail')
        let fdCache = path.resolve(fdCwd, 'cache')
        let fpMd = path.resolve(fdCwd, 'in.md')
        let fpDocx = path.resolve(fdCwd, 'out.docx')
        fs.rmSync(fdCwd, { recursive: true, force: true })
        fs.mkdirSync(fdCache, { recursive: true })
        fs.writeFileSync(fpMd, '# 轉檔器取不到\n\n段落', 'utf8')

        let script = [
            `import fs from 'fs'`,
            `import os from 'os'`,
            `let m = await import(process.env.T_MOD)`,
            `let err = null`,
            `await m.default(process.env.T_MD, process.env.T_DOCX, { fpInTemp: process.env.T_TEMP }).catch((e) => { err = e })`,
            `let nHtml = fs.readdirSync(os.tmpdir()).filter((v) => v.startsWith('wmd2docx_' + process.pid + '_')).length`,
            `process.stdout.write('@@RESULT@@' + JSON.stringify({ err, nHtml }))`,
            `process.exit(0)`,
        ].join('\n')
        let env = {
            ...process.env,
            W_BIN_CACHE: fdCache,
            W_BIN_OFFLINE: '1',
            T_MOD: pathToFileURL(path.resolve('./src/cvMdToDocx.mjs')).href,
            T_MD: fpMd,
            T_DOCX: fpDocx,
            T_TEMP: fpInTemp,
        }
        delete env.W_BIN_MIRROR
        let out = execFileSync(process.execPath, ['--input-type=module', '-e', script], { cwd: fdCwd, env, windowsHide: true, encoding: 'utf8' })
        let { err, nHtml } = JSON.parse(out.split('@@RESULT@@')[1])

        //由 w-html2docx 之 ensureBins 拒絕(帶錯誤碼), 非上層事前檢查; 原訊息保留並附指引
        assert.strict.equal(/^Failed to convert html to docx: \[E_OFFLINE_MISSING\] /.test(String(err)), true, String(err))
        assert.strict.equal(String(err).endsWith(` (${msgHintBin})`), true, String(err))
        assert.strict.equal(fs.existsSync(fpDocx), false)
        //清理路徑: 未指定fpOutHtml時中介html於失敗時亦刪除
        assert.strict.equal(nHtml, 0)
    })

    it('於不含node_modules/w-html2docx之cwd啟動且給fpInTemp時, 轉檔器之取得與cwd無關: 有Word則轉檔成功; 無Word則由轉檔器以離開碼1回報原因', function() {
        this.timeout(180000)

        //fdCwd: 以子程序於此cwd啟動(其下無 src/ 與 node_modules/), 轉檔器位置依 w-html2docx 套件根(快取), 與 cwd 無關
        let fdCwd = path.resolve(fdTmpRoot, 'nocwdbin')
        let fpMd = path.resolve(fdCwd, 'in.md')
        let fpDocx = path.resolve(fdCwd, 'out.docx')
        fs.rmSync(fdCwd, { recursive: true, force: true })
        fs.mkdirSync(fdCwd, { recursive: true })
        fs.writeFileSync(fpMd, '# 轉檔器與cwd無關\n\n段落', 'utf8')

        let script = [
            `let m = await import(process.env.T_MOD)`,
            `let err = null`,
            `let r = null`,
            `await m.default(process.env.T_MD, process.env.T_DOCX, { fpInTemp: process.env.T_TEMP }).then((v) => { r = v }).catch((e) => { err = e })`,
            `process.stdout.write('@@RESULT@@' + JSON.stringify({ err, size: r ? r.sizeDocx : null }))`,
            `process.exit(0)`,
        ].join('\n')
        let out = execFileSync(process.execPath, ['--input-type=module', '-e', script], {
            cwd: fdCwd,
            env: { ...process.env, T_MOD: pathToFileURL(path.resolve('./src/cvMdToDocx.mjs')).href, T_MD: fpMd, T_DOCX: fpDocx, T_TEMP: fpInTemp },
            windowsHide: true,
            encoding: 'utf8',
            timeout: 170000,
        })
        let { err, size } = JSON.parse(out.split('@@RESULT@@')[1])

        //轉檔器取得失敗(如無法連網下載)時僅能確認錯誤帶ensureBins之錯誤碼, 無法確認後續行為
        if (/^Failed to convert html to docx: \[E_[A-Z_]+\] /.test(String(err))) {
            console.log(`[unit-cvMdToDocx] htmlToDocx.exe could not be obtained, skip: ${err}`)
            this.skip()
        }

        //取得後續行轉檔: 有Word則產出docx; 無Word則 w-html2docx(≥1.0.29) 以離開碼1 reject 並帶轉檔器回報之原因
        //(其 srcPython/說明.txt 第9點: `error: com_error: (-2147221005, …)` 為Word未安裝), 本層加前綴後原樣傳出, 非轉檔器缺檔
        if (err === null) {
            assert.strict.equal(size, fs.statSync(fpDocx).size)
            assert.strict.equal(size > 0, true)
        }
        else {
            assert.strict.equal(hasWord(), false, `conversion failed although Word is available: ${err}`)
            assert.strict.equal(/^Failed to convert html to docx: code\[1\]:\r?\nerror: com_error: /.test(String(err)), true, String(err))
        }
    })

    it('於不含node_modules/w-html2docx之cwd啟動且未給fpInTemp時, 保留w-html2docx原訊息並附cwd提示', function() {
        this.timeout(180000)

        //fdCwd: 以子程序於此cwd啟動, w-html2docx於模組載入時以cwd查找預設模板, 查無即拒絕(WHtml2docx.mjs, 早於取得轉檔器)
        let fdCwd = path.resolve(fdTmpRoot, 'wrongcwd')
        let fpMd = path.resolve(fdCwd, 'in.md')
        let fpDocx = path.resolve(fdCwd, 'out.docx')
        fs.rmSync(fdCwd, { recursive: true, force: true })
        fs.mkdirSync(fdCwd, { recursive: true })
        fs.writeFileSync(fpMd, '# 錯誤工作目錄\n\n段落', 'utf8')

        let script = [
            `import(process.env.T_MOD).then(async (m) => {`,
            `    let err = null`,
            `    await m.default(process.env.T_MD, process.env.T_DOCX).catch((e) => { err = e })`,
            `    process.stdout.write('@@RESULT@@' + JSON.stringify({ err }))`,
            `})`,
        ].join('\n')
        let out = execFileSync(process.execPath, ['--input-type=module', '-e', script], {
            cwd: fdCwd,
            env: {
                ...process.env,
                T_MOD: pathToFileURL(path.resolve('./src/cvMdToDocx.mjs')).href,
                T_MD: fpMd,
                T_DOCX: fpDocx,
            },
            windowsHide: true,
            encoding: 'utf8',
        })
        let { err } = JSON.parse(out.split('@@RESULT@@')[1])

        //原訊息保留, 並附cwd與執行位置提示
        let prefix = 'Failed to convert html to docx: can not find folder for tmp.docx (cwd='
        assert.strict.equal(String(err).startsWith(prefix), true, String(err))
        assert.strict.equal(String(err).toLowerCase().includes(`(cwd=${fdCwd.toLowerCase()}, `), true, String(err))
        assert.strict.equal(String(err).endsWith(`, ${msgHint})`), true, String(err))
    })

})


//轉檔後處理(圖名表名同頁、添加目錄)之接線、略過與失敗清理: 子程序內以假轉檔器取代 w-html2docx(將固定 docx 複製至輸出位置), 不需 Word
describe('cvMdToDocx 轉檔後處理(假轉檔器, 不需Word)', function() {

    //check, 轉檔與目錄添加僅支援 Windows
    if (process.platform !== 'win32') {
        return
    }

    let { heading, body, pic, capFig, capTab, capTabPlain, table, cover } = parts
    let fdTmp = path.resolve(fdTmpRoot, 'post')
    let fpMd = path.resolve(fdTmp, 'in.md')
    let fpNoHead = path.resolve(fdTmp, 'fixture-nohead.docx')
    let fpWord = path.resolve(fdTmp, 'fixture-word.docx')
    let fpFig = path.resolve(fdTmp, 'fixture-fig.docx') //含圖(僅結構標記, 不可交 Word)與表, 無標題
    let fpNotDocx = path.resolve(fdTmp, 'fixture-notdocx.docx')
    let script = [
        `import fs from 'fs'`,
        `let m = await import(process.env.T_MOD)`,
        `let r = null, err = null`,
        `await m.default(process.env.T_MD, process.env.T_DOCX, JSON.parse(process.env.T_OPT)).then((v) => { r = v }).catch((e) => { err = e })`,
        `process.stdout.write('@@RESULT@@' + JSON.stringify({ r, err, exists: fs.existsSync(process.env.T_DOCX) }))`,
    ].join('\n')
    let conv = (fpFixture, name, opt) => {
        let fpDocx = path.resolve(fdTmp, `${name}.docx`)
        let res = runWithFakeHtml2docx(script, {
            fpDocx: fpFixture,
            env: {
                T_MOD: pathToFileURL(path.resolve('./src/cvMdToDocx.mjs')).href,
                T_MD: fpMd,
                T_DOCX: fpDocx,
                T_OPT: JSON.stringify(opt),
            },
        })
        return { ...res, fpDocx }
    }

    before(function() {
        fs.mkdirSync(fdTmp, { recursive: true })
        fs.writeFileSync(fpMd, '# 標題\n\n內容', 'utf8')
        fs.writeFileSync(fpNoHead, buildDocx(cover() + body('沒有標題')))
        fs.writeFileSync(fpWord, buildDocx(cover() + heading('第一章 緒論', 0, { pb: true }) + body('內容。') + capTab('表1 參數一覽') + table() + heading('第二章 方法', 0, { pb: true }) + body('內容。')))
        fs.writeFileSync(fpFig, buildDocx(cover() + body('前文。') + pic() + capFig('圖1 系統架構') + capTabPlain('表1 參數一覽') + table()))
        fs.writeFileSync(fpNotDocx, 'not a docx')
    })

    after(function() {
        fs.rmSync(fdTmpRoot, { recursive: true, force: true, maxRetries: 10, retryDelay: 300 })
    })

    it('toc未給、false或非true與物件之值時不添加目錄, 結果無toc', function() {
        this.timeout(180000)
        for (let [name, opt] of [['none', {}], ['false', { toc: false }], ['str', { toc: 'true' }]]) {
            let res = conv(fpWord, name, opt)
            assert.strict.equal(res.err, null, String(res.err))
            assert.strict.equal(res.r.toc, undefined, name)
            assert.strict.equal(fs.readFileSync(res.fpDocx).equals(fs.readFileSync(fpWord)), true, name) //未經改寫
        }
    })

    it('文件無標題時略過: 結果之toc僅含skip與ms, docx保留', function() {
        this.timeout(180000)
        let res = conv(fpNoHead, 'nohead', { toc: true })
        assert.strict.equal(res.err, null, String(res.err))
        assert.strict.deepEqual(Object.keys(res.r.toc), ['skip', 'ms'])
        assert.strict.equal(res.r.toc.skip, 'no heading (paragraph with an outline level) is found, the TOC is not added')
        assert.strict.equal(res.exists, true)
        assert.strict.equal(res.r.sizeDocx, fs.statSync(res.fpDocx).size)
    })

    it('添加目錄失敗(前提不符)時刪除本次產出之docx並reject', function() {
        this.timeout(180000)
        let res = conv(fpWord, 'badstyle', { toc: { headStyles: ['TPC11報告目錄(標題)'] } })
        assert.strict.equal(res.r, null)
        assert.strict.equal(res.err, 'Failed to add the TOC: Failed to prepare the TOC: the template has a style named "TPC11報告目錄(標題)" that is not bold, please specify another name with headStyles')
        assert.strict.equal(res.exists, false)
    })

    it('無Word時: Word更新失敗之原因可讀, 刪除本次產出之docx並reject', function() {
        this.timeout(180000)
        if (hasWord()) {
            console.log('[unit-cvMdToDocx] Microsoft Word is available, the no-Word failure path is covered by the real conversion tests')
            this.skip()
        }
        let res = conv(fpWord, 'noword', { toc: true })
        assert.strict.equal(res.r, null)
        assert.strict.equal(/^Failed to add the TOC: Word failed to update the fields: /.test(String(res.err)), true, String(res.err))
        assert.strict.equal(String(res.err).includes(String.fromCharCode(0xFFFD)), false, String(res.err)) //以 utf-8 解碼系統字碼頁之原因會含 U+FFFD
        assert.strict.equal(res.exists, false)
    })

    it('keepCaption: 圖名表名與其圖表同頁, 結果含統計, docx經改寫且不留暫存檔', function() {
        this.timeout(180000)
        let res = conv(fpFig, 'keep', { keepCaption: true })
        assert.strict.equal(res.err, null, String(res.err))
        let k = res.r.keepCaption
        assert.strict.equal(typeof k.ms, 'number')
        delete k.ms
        assert.strict.deepEqual(k, { figs: 1, tabs: 1, keepNext: 2, keepLines: 2, changed: true, warns: [] })
        assert.strict.equal(res.r.sizeDocx, fs.statSync(res.fpDocx).size)
        let ps = listParas(readPart(new Uint8Array(fs.readFileSync(res.fpDocx)), 'word/document.xml'))
        let i = ps.findIndex((p) => p.text === '圖1 系統架構')
        assert.strict.deepEqual([ps[i - 1].hasPic, ps[i - 1].keepNext, ps[i].keepLines], [true, true, true])
        assert.strict.deepEqual(fs.readdirSync(fdTmp).filter((v) => v.endsWith('.tmp')), [])
    })

    it('keepCaption未給、false或非true與物件之值時不處理, docx未經改寫且結果無keepCaption', function() {
        this.timeout(180000)
        for (let [name, opt] of [['knone', {}], ['kfalse', { keepCaption: false }], ['kstr', { keepCaption: 'true' }]]) {
            let res = conv(fpFig, name, opt)
            assert.strict.equal(res.err, null, String(res.err))
            assert.strict.equal(res.r.keepCaption, undefined, name)
            assert.strict.equal(fs.readFileSync(res.fpDocx).equals(fs.readFileSync(fpFig)), true, name)
        }
    })

    it('keepCaption而文件無圖名表名時不改寫檔案', function() {
        this.timeout(180000)
        let res = conv(fpNoHead, 'knocap', { keepCaption: true })
        assert.strict.equal(res.err, null, String(res.err))
        assert.strict.equal(res.r.keepCaption.changed, false)
        assert.strict.equal(fs.readFileSync(res.fpDocx).equals(fs.readFileSync(fpNoHead)), true)
    })

    it('keepCaption與toc並用而文件無標題: 目錄略過, 同頁照常套用', function() {
        this.timeout(180000)
        let res = conv(fpFig, 'kboth', { keepCaption: true, toc: true })
        assert.strict.equal(res.err, null, String(res.err))
        assert.strict.equal(res.r.toc.skip, 'no heading (paragraph with an outline level) is found, the TOC is not added')
        assert.strict.deepEqual([res.r.keepCaption.changed, res.r.keepCaption.keepNext], [true, 2])
    })

    it('keepCaption處理失敗(非docx)時刪除本次產出之docx並reject', function() {
        this.timeout(180000)
        let res = conv(fpNotDocx, 'kbad', { keepCaption: true })
        assert.strict.equal(res.r, null)
        assert.strict.equal(/^Failed to keep the captions with their figures and tables: /.test(String(res.err)), true, String(res.err))
        assert.strict.equal(res.exists, false)
    })

    it('imgHeightReserveLines: 未給時以4傳予w-html2docx, 給數字、null或其他值原樣傳下(由w-html2docx判斷), 呼叫端之optHtml2docx不被改動', function() {
        this.timeout(180000)
        //cases: 依序為 opt 未給 optHtml2docx、空物件、僅含他鍵、6、0、null、-1、'x'、optHtml2docx 非物件
        let scriptOpt = [
            `let m = await import(process.env.T_MOD)`,
            `let cases = [undefined, {}, { imgRatioWidthMax: 0.5 }, { imgHeightReserveLines: 6 }, { imgHeightReserveLines: 0 }, { imgHeightReserveLines: null }, { imgHeightReserveLines: -1 }, { imgHeightReserveLines: 'x' }, 'str']`,
            `let snap = JSON.stringify(cases)`,
            `for (let o of cases) { await m.default(process.env.T_MD, process.env.T_DOCX, o === undefined ? {} : { optHtml2docx: o }) }`,
            `let opts = globalThis.fakeHtml2docxOpts`,
            `process.stdout.write('@@RESULT@@' + JSON.stringify({ got: opts.map((v) => v.imgHeightReserveLines), other: opts[2].imgRatioWidthMax, unchanged: JSON.stringify(cases) === snap }))`,
        ].join('\n')
        let res = runWithFakeHtml2docx(scriptOpt, {
            fpDocx: fpNoHead,
            env: {
                T_MOD: pathToFileURL(path.resolve('./src/cvMdToDocx.mjs')).href,
                T_MD: fpMd,
                T_DOCX: path.resolve(fdTmp, 'reserve.docx'),
            },
        })
        assert.strict.deepEqual(res.got, [4, 4, 4, 6, 0, null, -1, 'x', 4])
        assert.strict.equal(res.other, 0.5) //其餘設定照傳
        assert.strict.equal(res.unchanged, true)
    })

    it('w-html2docx失敗訊息之提示: 以[E_xxx]開頭(轉檔器取不到)附README指引; can not find folder for tmp.docx附cwd與預設模板提示; 其餘(含圖片解析不到)原樣傳出', function() {
        this.timeout(180000)
        let hintBin = ' (the external executable htmlToDocx.exe of w-html2docx is not available, see the section "外部執行檔" of the README of w-html2docx for W_BIN_CACHE, W_BIN_MIRROR, W_BIN_OFFLINE and WHtml2docx.ensureBins())'
        let hintTemp = `, w-html2docx resolves its default template from cwd, please run from the project root that contains node_modules/w-html2docx or give fpInTemp)`
        let msgUnres = 'code[1]:\nunresolved linked picture: C:/x/a.png\nerror: LinkedPictureError: 1 linked picture(s) can not be embedded since Word can not resolve their source(see stderr)'
        let cases = [
            ['[E_OFFLINE_MISSING] fsEnsureBins[w-html2docx]: offline, 1 external executable file(s) are not in the cache or the mirror', (m) => `${m}${hintBin}`],
            ['[E_HASH_MISMATCH] fsEnsureBins[w-html2docx] file[htmlToDocx.exe] from github: sha256 mismatch', (m) => `${m}${hintBin}`],
            ['[E_NETWORK] fsEnsureBins[w-html2docx] file[htmlToDocx.exe] from github: fetch failed', (m) => `${m}${hintBin}`],
            ['can not find folder for tmp.docx', (m) => `${m} (cwd=${path.resolve()}${hintTemp}`],
            [msgUnres, (m) => m],
            ['code[1]:\nerror: com_error: (-2147221005, 無效的類別字串)', (m) => m],
            ['fpInTemp[x.docx] does not exist', (m) => m],
            ['x [E_OFFLINE_MISSING] y', (m) => m], //錯誤碼非位於開頭者不附
        ]
        for (let [msg, exp] of cases) {
            let res = runWithFakeHtml2docx(`let m = await import(process.env.T_MOD)\nlet err = null\nawait m.default(process.env.T_MD, process.env.T_DOCX, {}).catch((e) => { err = e })\nprocess.stdout.write('@@RESULT@@' + JSON.stringify({ err }))`, {
                reject: msg,
                env: {
                    T_MOD: pathToFileURL(path.resolve('./src/cvMdToDocx.mjs')).href,
                    T_MD: fpMd,
                    T_DOCX: path.resolve(fdTmp, 'hint.docx'),
                },
            })
            assert.strict.equal(res.err, `Failed to convert html to docx: ${exp(msg)}`, msg)
        }
    })

    it('keepCaption: 轉出之docx超過2 GiB(fs.readFileSync上限)時, 讀檔前以指出原因之訊息reject並刪除該docx, 行程不中止', function() {
        this.timeout(180000)
        let res = runWithFakeHtml2docx(script, {
            size: 2 ** 31, //稀疏檔, 不實寫
            env: {
                T_MOD: pathToFileURL(path.resolve('./src/cvMdToDocx.mjs')).href,
                T_MD: fpMd,
                T_DOCX: path.resolve(fdTmp, 'big-keep.docx'),
                T_OPT: JSON.stringify({ keepCaption: true }),
            },
        })
        assert.strict.equal(res.r, null)
        assert.strict.equal(res.err, 'Failed to keep the captions with their figures and tables: docx is too large to be read into memory in one piece (2147483648 bytes, over the limit of 2147483647 bytes = 2 GiB)')
        assert.strict.equal(res.exists, false)
    })

    it('toc: 轉出之docx超過2 GiB(fs.readFileSync上限)時, 讀檔前以指出原因之訊息reject並刪除該docx, 行程不中止', function() {
        this.timeout(180000)
        let res = runWithFakeHtml2docx(script, {
            size: 2 ** 31, //稀疏檔, 不實寫
            env: {
                T_MOD: pathToFileURL(path.resolve('./src/cvMdToDocx.mjs')).href,
                T_MD: fpMd,
                T_DOCX: path.resolve(fdTmp, 'big-toc.docx'),
                T_OPT: JSON.stringify({ toc: true }),
            },
        })
        assert.strict.equal(res.r, null)
        assert.strict.equal(res.err, 'Failed to add the TOC: Failed to prepare the TOC: docx is too large to be read into memory in one piece (2147483648 bytes, over the limit of 2147483647 bytes = 2 GiB)')
        assert.strict.equal(res.exists, false)
    })

})
