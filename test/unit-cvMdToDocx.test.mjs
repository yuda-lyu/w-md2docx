import fs from 'fs'
import path from 'path'
import zlib from 'zlib'
import { execFileSync } from 'child_process'
import assert from 'assert'
import cvMdToDocx from '../src/cvMdToDocx.mjs'
import cvMdTo from '../src/cvMdTo.mjs'
import { getFpExe } from '../src/utils.mjs'


//hasWord: 本機是否可調用 Microsoft Word(Windows + Word.Application COM 已註冊 + htmlToDocx.exe 存在)
//why: docx 階段須真 Word, 無 Word 之機器僅測輸入檢查, 實轉區塊整段略過(同 w-html2docx 之 isWindows 守門)
function hasWord() {
    if (process.platform !== 'win32') {
        return false
    }
    if (getFpExe() === '') {
        return false
    }
    try {
        execFileSync('reg', ['query', 'HKCR\\Word.Application'], { stdio: 'ignore', windowsHide: true })
        return true
    }
    catch (err) {
        return false
    }
}


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

    it('fpInMd非有效字串', async function() {
        await assert.rejects(cvMdToDocx('', './tmp/x.docx'), (e) => e === 'fpInMd must be a non-empty string')
        await assert.rejects(cvMdToDocx(null, './tmp/x.docx'), (e) => e === 'fpInMd must be a non-empty string')
    })

    it('fpInMd不存在', async function() {
        await assert.rejects(cvMdToDocx('./test/nope.md', './tmp/x.docx'), (e) => /^fpInMd\[.+nope\.md\] does not exist$/.test(e))
    })

    it('fpOutDocx非有效字串', async function() {
        await assert.rejects(cvMdToDocx('./test/report.md', ''), (e) => e === 'fpOutDocx must be a non-empty string')
    })

    it('fpInTemp不存在', async function() {
        await assert.rejects(cvMdToDocx('./test/report.md', './tmp/x.docx', { fpInTemp: './test/nope.docx' }), (e) => /^fpInTemp\[.+nope\.docx\] does not exist$/.test(e))
    })

})


describe('cvMdToDocx 實轉(需Windows+Microsoft Word)', function() {

    //check
    if (!hasWord()) {
        console.log('[unit-cvMdToDocx] Microsoft Word not available on this machine, skip real conversion tests')
        return
    }

    let fdTmp = path.resolve('./tmp/zt_cvMdToDocx')
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
        fs.rmSync(fdTmp, { recursive: true, force: true })
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
