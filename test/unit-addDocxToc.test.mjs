import fs from 'fs'
import path from 'path'
import assert from 'assert'
import addDocxToc, { psScript, execPs, fromClixml } from '../src/addDocxToc.mjs'
import { parts, buildDocx, getSectFinal } from './tools/docxFixture.mjs'
import { sha, nTmpToc, bodyWord, snapPs1 } from './tools/tocFixture.mjs'


let { run, para, body, cover } = parts


//調用 Word 之〈Word更新〉案在 unit-cvMdToDocx.test.mjs〈addDocxToc Word更新〉(規則帳本 R-TEST-1: Word 不可由 mocha --parallel 之不同檔同時調用), 本檔不啟動 Word
describe('addDocxToc', function() {

    //check, Word 只存在於 Windows
    if (process.platform !== 'win32') {
        it('非Windows時reject', async function() {
            await assert.rejects(addDocxToc('./x.docx'), (e) => e === 'addDocxToc requires Microsoft Word on Windows')
        })
        return
    }

    let fdTmp = path.resolve('./test/_tmp/unit-addDocxToc')

    before(function() {
        fs.mkdirSync(fdTmp, { recursive: true })
    })

    after(function() {
        fs.rmSync(fdTmp, { recursive: true, force: true, maxRetries: 10, retryDelay: 300 })
    })

    //note: 共用的探測腳本, 讀入帶入之環境變數並以 UTF-8 輸出
    let psProbe = [
        '[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)',
        '[Console]::Out.WriteLine(\'[\' + $env:W_MD2DOCX_SRC + \'][\' + $env:W_MD2DOCX_DST + \']\')',
    ].join('\n')

    it('psScript: 僅含ASCII, 路徑不在腳本內(由環境變數帶入), 編碼後之命令列遠低於32767字元上限', function() {
        assert.strict.equal([...psScript].every((c) => c.charCodeAt(0) < 0x80), true)
        assert.strict.equal(psScript.includes('$env:W_MD2DOCX_SRC'), true)
        assert.strict.equal(psScript.includes('$env:W_MD2DOCX_DST'), true)
        assert.strict.equal(/[A-Za-z]:\\/.test(psScript), false) //不含磁碟機路徑
        let enc = Buffer.from(psScript, 'utf16le').toString('base64')
        assert.strict.equal(enc.length < 16383, true, `encoded script ${enc.length} chars`) //CreateProcess 命令列上限 32767 字元, 留一半餘裕
    })

    it('套件不含獨立腳本檔: src無.vbs與.ps1, 原始碼不呼叫cscript與wscript, package.json不提.vbs', function() {
        let names = fs.readdirSync('./src')
        assert.strict.deepEqual(names.filter((v) => /\.(vbs|ps1)$/i.test(v)), [])
        for (let name of names.filter((v) => v.endsWith('.mjs'))) {
            //去除註解後檢查(註解可說明緣由而提及 .vbs)
            let code = fs.readFileSync(path.resolve('./src', name), 'utf8').replace(/\/\*[\s\S]*?\*\//g, '').replace(/^\s*\/\/.*$/gm, '')
            assert.strict.equal(/\b(cscript|wscript)\b|\.vbs\b/i.test(code), false, name)
        }
        assert.strict.equal(/\.vbs\b/i.test(fs.readFileSync('./package.json', 'utf8')), false)
    })

    it('fromClixml: 還原CLIXML(丟棄進度、錯誤取文字、換行與跳脫還原), 非CLIXML原樣', function() {
        assert.strict.equal(fromClixml('plain\r\ntext'), 'plain\r\ntext')
        assert.strict.equal(fromClixml(''), '')
        let prog = '<Objs Version="1.1.0.1" xmlns="http://schemas.microsoft.com/powershell/2004/04"><Obj S="progress" RefId="0"><TN RefId="0"><T>X</T></TN><MS><I64 N="SourceId">1</I64></MS></Obj></Objs>'
        assert.strict.equal(fromClixml(`#< CLIXML\r\nerror: a\r\n${prog}`), 'error: a\r\n')
        let err = '<Objs Version="1.1.0.1" xmlns="http://schemas.microsoft.com/powershell/2004/04"><S S="Error">x &lt;y&gt; &amp; 中文_x000D__x000A_</S><S S="Error">z_x000D__x000A_</S></Objs>'
        assert.strict.equal(fromClixml(`code[1]:\n#< CLIXML\r\n${err}`), 'code[1]:\nx <y> & 中文\r\nz\r\n')
    })

    it('execPs: 路徑經環境變數原樣帶入(含空白、中文、引號與&^%!;,=\'`$), 執行後環境變數還原, 無.ps1落地', async function() {
        this.timeout(120000)
        let bq = String.fromCharCode(96)
        let pairs = [
            ['D:\\opensrc\\a b\\plain.docx', 'D:\\opensrc\\a b\\plain-out.docx'],
            ['D:\\開源 資料夾\\報告 測試.docx', 'D:\\開源 資料夾\\輸出 測試.docx'],
            ['D:\\x\\q"uote.docx', `D:\\x\\a&b^c%PATH%!d;e,f=g'h${bq}i$j.docx`],
        ]
        let s0 = snapPs1()
        process.env.W_MD2DOCX_DST = 'keep-me' //執行前已存在者, 結束後須還原為原值
        delete process.env.W_MD2DOCX_SRC
        try {
            for (let [a, b] of pairs) {
                let out = await execPs(psProbe, { W_MD2DOCX_SRC: a, W_MD2DOCX_DST: b }, 60000)
                assert.strict.equal(out.trim(), `[${a}][${b}]`)
                assert.strict.equal(process.env.W_MD2DOCX_SRC, undefined)
                assert.strict.equal(process.env.W_MD2DOCX_DST, 'keep-me')
            }
        }
        finally {
            delete process.env.W_MD2DOCX_DST
        }
        assert.strict.equal(snapPs1(), s0)
    })

    it('execPs: 離開碼非0為`code[N]:`、未捕捉之例外取其文字(無CLIXML)、逾時為`timeout[Nms]:`且結束子程序', async function() {
        this.timeout(120000)
        let s0 = snapPs1()
        let rej = async (script, env, timeout) => {
            let err = null
            await execPs(script, env, timeout).catch((e) => {
                err = e
            })
            assert.strict.equal(typeof err, 'string')
            return err
        }
        let head = '[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)\n'
        assert.strict.equal((await rej(head + 'exit 3', {}, 60000)).startsWith('code[3]:'), true)
        let eThrow = await rej(head + '$ErrorActionPreference = \'Stop\'\nthrow \'中文例外\'', {}, 60000)
        assert.strict.equal(eThrow.startsWith('code[1]:'), true, eThrow)
        assert.strict.equal(eThrow.includes('中文例外'), true, eThrow)
        assert.strict.equal(/CLIXML|<Objs/.test(eThrow), false, eThrow)
        let t0 = Date.now()
        let eTimeout = await rej('Start-Sleep -Seconds 60', {}, 1500)
        assert.strict.equal(eTimeout.startsWith('timeout[1500ms]:'), true, eTimeout)
        assert.strict.equal(Date.now() - t0 < 30000, true)
        assert.strict.equal(snapPs1(), s0)
    })

    it('psScript: 缺環境變數為用法錯誤(離開碼2, 訊息指出變數名), 無Word亦同; Word無法啟動時字頭為`error: cannot start Microsoft Word: `(離開碼1, 原因可讀)', async function() {
        this.timeout(120000)
        let rej = async (script, env) => {
            let err = null
            await execPs(script, env, 60000).catch((e) => {
                err = e
            })
            assert.strict.equal(typeof err, 'string')
            return err
        }
        let usage = 'code[2]:\nerror: usage: set the environment variables W_MD2DOCX_SRC and W_MD2DOCX_DST (absolute paths of the source and the destination docx)'
        assert.strict.equal((await rej(psScript, {})).trim(), usage)
        assert.strict.equal((await rej(psScript, { W_MD2DOCX_SRC: 'D:\\x\\a.docx' })).trim(), usage)
        //以未註冊之 COM 類別模擬 Word 無法啟動(不需 Word)
        let eStart = await rej(psScript.replace('Word.Application', 'Md2docxNoSuchApp.Application'), { W_MD2DOCX_SRC: 'D:\\x\\a.docx', W_MD2DOCX_DST: 'D:\\x\\b.docx' })
        assert.strict.equal(eStart.startsWith('code[1]:\nerror: cannot start Microsoft Word: '), true, eStart)
        assert.strict.equal(eStart.includes(String.fromCharCode(0xFFFD)), false, eStart)
        assert.strict.equal(/CLIXML|<Objs/.test(eStart), false, eStart)
    })

    it('輸入檢查: fpDocx非有效字串、不存在、非docx', async function() {
        await assert.rejects(addDocxToc(''), (e) => e === 'fpDocx must be a non-empty string')
        await assert.rejects(addDocxToc(null), (e) => e === 'fpDocx must be a non-empty string')
        await assert.rejects(addDocxToc('./test/nope.docx'), (e) => /^fpDocx\[.+nope\.docx\] does not exist$/.test(e))
        await assert.rejects(addDocxToc('./test/report.md'), (e) => /^Failed to prepare the TOC: /.test(e))
    })

    it('docx超過2 GiB(fs.readFileSync上限): 讀檔前以指出原因之訊息reject, 原檔不變且不啟動Word(不留暫存檔)', async function() {
        let fp = path.resolve(fdTmp, 'big.docx')
        let h = fs.openSync(fp, 'w')
        fs.ftruncateSync(h, 2 ** 31) //稀疏檔, 不實寫
        fs.closeSync(h)
        let n0 = nTmpToc()
        try {
            await assert.rejects(addDocxToc(fp), (e) => {
                assert.strict.equal(e, 'Failed to prepare the TOC: docx is too large to be read into memory in one piece (2147483648 bytes, over the limit of 2147483647 bytes = 2 GiB)')
                return true
            })
            assert.strict.equal(fs.statSync(fp).size, 2 ** 31)
            assert.strict.equal(nTmpToc(), n0)
        }
        finally {
            fs.unlinkSync(fp)
        }
    })

    it('文件無標題時不處理: 僅回傳skip與ms且原檔不變', async function() {
        let fp = path.resolve(fdTmp, 'nohead.docx')
        fs.writeFileSync(fp, buildDocx(cover() + body('沒有標題')))
        let h = sha(fp)
        let r = await addDocxToc(fp)
        assert.strict.deepEqual(Object.keys(r), ['skip', 'ms'])
        assert.strict.equal(r.skip, 'no heading (paragraph with an outline level) is found, the TOC is not added')
        assert.strict.equal(sha(fp), h)
    })

    it('前提不符時reject且原檔不變', async function() {
        let fp = path.resolve(fdTmp, 'twosect.docx')
        fs.writeFileSync(fp, buildDocx(para(run('封面'), `<w:jc w:val="center"/>${getSectFinal()}`) + bodyWord))
        let h = sha(fp)
        await assert.rejects(addDocxToc(fp), (e) => e === 'Failed to prepare the TOC: the document has 2 sections, renumbering pages requires a document of 1 section (set pageNumbers to false to keep the sections)')
        assert.strict.equal(sha(fp), h)
    })

})
