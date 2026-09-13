import fs from 'fs'
import path from 'path'
import zlib from 'zlib'
import { pathToFileURL } from 'url'
import { execFileSync } from 'child_process'
import assert from 'assert'
import cvMdToDocx from '../src/cvMdToDocx.mjs'
import cvMdTo from '../src/cvMdTo.mjs'


//fdTmpRoot: 本測試檔專用暫存夾, 各describe於其下自建子夾並於after刪除
let fdTmpRoot = path.resolve('./test/_tmp/unit-cvMdToDocx')


//hasWord: 本機是否可調用 Microsoft Word(Windows + Word.Application COM 已註冊)
//why: docx 階段須真 Word, 無 Word 之機器實轉區塊整段略過;
//     不以 htmlToDocx.exe 是否存在為條件——缺檔時 w-html2docx 於轉檔時自動下載, 以其為條件會使剛安裝之機器漏測實轉
function hasWord() {
    if (process.platform !== 'win32') {
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


//轉檔器之定位與補齊由 w-html2docx 負責(WHtml2docx.mjs:158 → autoDownloadFiles), 上層不得以自身規則事前攔截
describe('cvMdToDocx 缺轉檔器時交由w-html2docx取得(不得於上層事前攔截)', function() {

    //check, htmlToDocx.exe僅供Windows, 非Windows時w-html2docx於下載前即拒絕
    if (process.platform !== 'win32') {
        return
    }

    let fpInTemp = path.resolve('./src/templates/temp_tpc.docx')
    let msgHint = 'w-html2docx resolves htmlToDocx.exe and its default template from cwd and downloads htmlToDocx.exe automatically when absent, please run from the project root that contains node_modules/w-html2docx'

    after(function() {
        fs.rmSync(fdTmpRoot, { recursive: true, force: true, maxRetries: 10, retryDelay: 300 })
    })

    it('缺檔時不於上層中斷, 由w-html2docx自動下載轉檔器後續行轉檔', async function() {
        this.timeout(180000)

        //fdCwd: 模擬工作目錄, 其下src/與node_modules/w-html2docx/src/皆無htmlToDocx.exe(即安裝時postinstall未執行)
        let fdCwd = path.resolve(fdTmpRoot, 'autodl')
        let fpMd = path.resolve(fdCwd, 'in.md')
        let fpDocx = path.resolve(fdCwd, 'out.docx')
        let fpExeDL = path.resolve(fdCwd, 'src', 'htmlToDocx.exe') //cwd無node_modules/w-html2docx/時之下載落點(autoDownloadFiles.mjs:66)
        fs.rmSync(fdCwd, { recursive: true, force: true })
        fs.mkdirSync(fdCwd, { recursive: true })
        fs.writeFileSync(fpMd, '# 轉檔器自動下載\n\n段落', 'utf8')

        let cwdOri = process.cwd()
        let err = null
        let r = null
        process.chdir(fdCwd)
        try {
            assert.strict.equal(fs.existsSync(fpExeDL), false) //前置: 缺檔
            r = await cvMdToDocx(fpMd, fpDocx, { fpInTemp })
                .catch((e) => {
                    err = e
                })
        }
        finally {
            process.chdir(cwdOri)
        }

        //不得為上層事前檢查之拒絕(未呼叫w-html2docx即中斷)
        assert.strict.equal(/htmlToDocx\.exe not found/.test(String(err)), false, `rejected by upper-layer pre-check: ${err}`)

        //缺檔即自動下載: 轉檔器落地於下載落點
        if (!fs.existsSync(fpExeDL)) {
            //下載失敗(如離線)時僅能確認失敗來自docx階段(已交由w-html2docx), 無法確認落地
            assert.strict.equal(/^Failed to convert html to docx: /.test(String(err)), true, String(err))
            console.log(`[unit-cvMdToDocx] htmlToDocx.exe download failed, skip landing check: ${err}`)
            this.skip()
        }
        assert.strict.equal(fs.statSync(fpExeDL).size > 1024 * 1024, true)

        //下載後續行轉檔: 有Word則產出docx; 無Word則 w-html2docx(≥1.0.29) 以離開碼1 reject 並帶轉檔器回報之原因
        //(其 srcPython/說明.txt 第9點: `error: com_error: (-2147221005, …)` 為Word未安裝), 本層加前綴後原樣傳出, 非轉檔器缺檔
        if (err === null) {
            assert.strict.equal(r.sizeDocx, fs.statSync(fpDocx).size)
            assert.strict.equal(r.sizeDocx > 0, true)
        }
        else {
            assert.strict.equal(hasWord(), false, `conversion failed although Word is available: ${err}`)
            assert.strict.equal(/^Failed to convert html to docx: code\[1\]:\r?\nerror: com_error: /.test(String(err)), true, String(err))
        }
    })

    it('缺檔且下載失敗時reject(不假成功), 不產出docx且中介html已清除', function() {
        this.timeout(180000)

        //fdCwd: 以子程序於此cwd啟動並經不可達之代理下載, 使w-html2docx之自動下載確定失敗(不需連網)
        let fdCwd = path.resolve(fdTmpRoot, 'dlfail')
        let fpMd = path.resolve(fdCwd, 'in.md')
        let fpDocx = path.resolve(fdCwd, 'out.docx')
        let fpExeDL = path.resolve(fdCwd, 'src', 'htmlToDocx.exe')
        fs.rmSync(fdCwd, { recursive: true, force: true })
        fs.mkdirSync(fdCwd, { recursive: true })
        fs.writeFileSync(fpMd, '# 下載失敗\n\n段落', 'utf8')

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
        let out = execFileSync(process.execPath, ['--input-type=module', '-e', script], {
            cwd: fdCwd,
            env: {
                ...process.env,
                NODE_USE_ENV_PROXY: '1',
                HTTPS_PROXY: 'http://127.0.0.1:1',
                HTTP_PROXY: 'http://127.0.0.1:1',
                T_MOD: pathToFileURL(path.resolve('./src/cvMdToDocx.mjs')).href,
                T_MD: fpMd,
                T_DOCX: fpDocx,
                T_TEMP: fpInTemp,
            },
            windowsHide: true,
            encoding: 'utf8',
        })
        let { err, nHtml } = JSON.parse(out.split('@@RESULT@@')[1])

        //Node未支援NODE_USE_ENV_PROXY時下載會實際成功, 無法構成下載失敗之前置, 標為pending
        if (fs.existsSync(fpExeDL)) {
            console.log('[unit-cvMdToDocx] proxy env not honored by this Node version, skip download failure check')
            this.skip()
        }

        //下載失敗才失敗: 由docx階段reject, 不假成功
        assert.strict.equal(/^Failed to convert html to docx: /.test(String(err)), true, String(err))
        assert.strict.equal(fs.existsSync(fpDocx), false)
        //清理路徑: 未指定fpOutHtml時中介html於失敗時亦刪除
        assert.strict.equal(nHtml, 0)
    })

    it('於不含node_modules/w-html2docx之cwd啟動且未給fpInTemp時, 保留w-html2docx原訊息並附cwd提示', function() {
        this.timeout(180000)

        //fdCwd: 以子程序於此cwd啟動, w-html2docx於模組載入時以cwd查找預設模板, 查無即拒絕(WHtml2docx.mjs:136, 早於下載)
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
