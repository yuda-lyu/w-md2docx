import fs from 'fs'
import path from 'path'
import { pathToFileURL } from 'url'
import assert from 'assert'
import cvMdTo from '../src/cvMdTo.mjs'
import runWithFakeHtml2docx from './tools/fakeHtml2docx.mjs'
import { parts, buildDocx } from './tools/docxFixture.mjs'
import { genTallPng } from './tools/imgFixture.mjs'


//cvMdTo之html路徑(不需Microsoft Word); docx路徑於unit-cvMdToDocx以Word守門測試
describe('cvMdTo', function() {

    let fdTmp = path.resolve('./test/_tmp/unit-cvMdTo')
    let svgB64 = fs.readFileSync('./test/cocktail.svg').toString('base64')
    let toHtml = (r) => Buffer.from(r.html.base64, 'base64').toString('utf8')

    before(function() {
        fs.mkdirSync(fdTmp, { recursive: true })
    })

    after(function() {
        fs.rmSync(fdTmp, { recursive: true, force: true, maxRetries: 10, retryDelay: 300 })
    })

    it('out=html: 回傳物件形狀與檔名', async function() {
        let r = await cvMdTo({ md: '# 標題A\n\n內文B', name: '報告R00.01', out: 'html' })
        assert.strict.equal(r.name, '報告R00.01')
        assert.strict.equal(r.out, 'html')
        assert.strict.equal(r.nAssets, 0)
        assert.strict.equal(typeof r.ms, 'number')
        assert.strict.equal(r.docx, undefined)
        assert.strict.equal(r.html.fileName, '報告R00.01.html')
        assert.strict.equal(r.html.mime, 'text/html; charset=utf-8')
        assert.strict.equal(r.html.size, Buffer.from(r.html.base64, 'base64').length)
        let h = toHtml(r)
        assert.strict.equal(h.includes('標題A'), true)
        assert.strict.equal(h.includes('內文B'), true)
        assert.strict.equal(h.includes('<title>報告R00.01</title>'), true) //title取自md檔名(=name)
    })

    it('name含路徑與非法字元時被清理', async function() {
        let r = await cvMdTo({ md: '# t', name: '../a/b:c.md', out: 'html' })
        assert.strict.equal(r.name, 'b_c')
        assert.strict.equal(r.html.fileName, 'b_c.html')
    })

    it('name未給則為output', async function() {
        let r = await cvMdTo({ md: '# t', out: 'html' })
        assert.strict.equal(r.name, 'output')
    })

    it('mdBase64可替代md', async function() {
        let r = await cvMdTo({ mdBase64: Buffer.from('# 由b64', 'utf8').toString('base64'), out: 'html' })
        assert.strict.equal(toHtml(r).includes('由b64'), true)
    })

    it('md優先於mdBase64', async function() {
        let r = await cvMdTo({ md: '# 主', mdBase64: Buffer.from('# 副', 'utf8').toString('base64'), out: 'html' })
        let h = toHtml(r)
        assert.strict.equal(h.includes('主'), true)
        assert.strict.equal(h.includes('副'), false)
    })

    it('md為空(未給/空白/非字串)則reject', async function() {
        let msg = 'md is empty: md (string) or mdBase64 is required'
        await assert.rejects(cvMdTo({ out: 'html' }), (e) => e === msg)
        await assert.rejects(cvMdTo({ md: '   \n', out: 'html' }), (e) => e === msg)
        await assert.rejects(cvMdTo({ md: 123, out: 'html' }), (e) => e === msg)
        await assert.rejects(cvMdTo(), (e) => e === msg)
    })

    it('資產: 以md內原字串為path, 同時支援markdown圖與<img>之URL編碼引用, 圖片內嵌為base64', async function() {
        let r = await cvMdTo({
            md: '# t\n\n![p](pics/圖.svg)\n\n<img src="pics/%E5%9C%96.svg" />',
            out: 'html',
            assets: [{ path: 'pics/%E5%9C%96.svg', base64: svgB64 }],
        })
        assert.strict.equal(r.nAssets, 1)
        let h = toHtml(r)
        assert.strict.equal((h.match(/src="data:image\/png;base64,/g) || []).length, 2)
        assert.strict.equal(h.includes('pics/'), false) //相對路徵已全數替換
    })

    it('資產: assets非陣列視為無資產', async function() {
        let r = await cvMdTo({ md: '# t', out: 'html', assets: 'x' })
        assert.strict.equal(r.nAssets, 0)
    })

    it('資產缺漏(<img>引用): 英文可行動訊息且不含工作夾路徑', async function() {
        await assert.rejects(cvMdTo({ md: '# t\n\n<img src="pics/no.png" />', out: 'html' }), (e) => {
            assert.strict.equal(e, 'Failed to convert md to html: the referenced file (pics/no.png) does not exist, please make sure it is included in assets')
            return true
        })
    })

    it('資產缺漏(md語法圖與<img>併用, 多個): 全數列出且不含工作夾路徑, 單數與複數形各依個數', async function() {
        await assert.rejects(cvMdTo({ md: '# t\n\n![x](pics/a.png)\n\n<img src="b.png" />', out: 'html' }), (e) => {
            //列出之順序依w-md2html, 不作為契約, 故排序後比對
            let mt = e.match(/^Failed to convert md to html: the referenced files \((.+)\) do not exist, please make sure they are included in assets$/)
            assert.strict.notEqual(mt, null, e)
            assert.strict.deepEqual(mt[1].split(', ').sort(), ['b.png', 'pics/a.png'])
            return true
        })
        await assert.rejects(cvMdTo({ md: '# t\n\n![x](pics/a.png)', out: 'html' }), (e) => {
            assert.strict.equal(e, 'Failed to convert md to html: the referenced file (pics/a.png) does not exist, please make sure it is included in assets')
            return true
        })
    })

    it('資產穿越/絕對路徑/空base64: 以asset開頭之訊息reject', async function() {
        await assert.rejects(cvMdTo({ md: '# t', out: 'html', assets: [{ path: '../x.png', base64: 'QUI=' }] }), (e) => e === 'asset.path is outside the working folder: ../x.png')
        await assert.rejects(cvMdTo({ md: '# t', out: 'html', assets: [{ path: 'C:/x.png', base64: 'QUI=' }] }), (e) => e === 'asset.path must be a relative path: C:/x.png')
        await assert.rejects(cvMdTo({ md: '# t', out: 'html', assets: [{ path: 'x.png' }] }), (e) => e === 'asset.base64 is empty: x.png')
    })

    it('optMd2html傳遞至w-md2html(imgWidthMax)', async function() {
        let r = await cvMdTo({
            md: '![p](pics/a.svg)',
            out: 'html',
            assets: [{ path: 'pics/a.svg', base64: svgB64 }],
            optMd2html: { imgWidthMax: '123px' },
        })
        assert.strict.equal(/max-width:\s*123px/.test(toHtml(r)), true)
    })

    it('dirWork指定則作業夾建於其下; keepWork=true保留, false清除', async function() {
        let dirWork = path.resolve(fdTmp, 'work')
        await cvMdTo({ md: '# t', out: 'html', dirWork, keepWork: true })
        let jobs = fs.readdirSync(dirWork).filter((v) => v.startsWith('job_'))
        assert.strict.equal(jobs.length, 1)
        assert.strict.equal(fs.existsSync(path.resolve(dirWork, jobs[0], 'output.md')), true)
        assert.strict.equal(fs.existsSync(path.resolve(dirWork, jobs[0], 'output.html')), true)
        await cvMdTo({ md: '# t', out: 'html', dirWork })
        assert.strict.equal(fs.readdirSync(dirWork).filter((v) => v.startsWith('job_')).length, 1) //僅前一次保留者
    })

    it('keepWork非布林值視為false', async function() {
        let dirWork = path.resolve(fdTmp, 'work2')
        await cvMdTo({ md: '# t', out: 'html', dirWork, keepWork: 'true' })
        assert.strict.equal(fs.readdirSync(dirWork).filter((v) => v.startsWith('job_')).length, 0)
    })

    it('失敗時作業夾仍被清除', async function() {
        let dirWork = path.resolve(fdTmp, 'work3')
        await assert.rejects(cvMdTo({ md: '# t\n\n<img src="no.png" />', out: 'html', dirWork }))
        assert.strict.equal(fs.readdirSync(dirWork).filter((v) => v.startsWith('job_')).length, 0)
    })

    it('out=html時toc不適用(無docx), 結果無toc', async function() {
        let r = await cvMdTo({ md: '# t', out: 'html', toc: true })
        assert.strict.equal(r.toc, undefined)
    })

    it('產出docx且toc時結果帶回cvMdToDocx之toc(假轉檔器取代w-html2docx, 文件無標題時為skip)', function() {
        this.timeout(180000)
        if (process.platform !== 'win32') {
            this.skip()
        }
        let fpFixture = path.resolve(fdTmp, 'fixture-nohead.docx')
        fs.writeFileSync(fpFixture, buildDocx(parts.cover() + parts.body('沒有標題')))
        let script = [
            `let m = await import(process.env.T_MOD)`,
            `let r = await m.default({ md: '# t', out: 'docx', toc: true, dirWork: process.env.T_WORK })`,
            `process.stdout.write('@@RESULT@@' + JSON.stringify({ toc: r.toc, size: r.docx.size }))`,
        ].join('\n')
        let res = runWithFakeHtml2docx(script, {
            fpDocx: fpFixture,
            env: {
                T_MOD: pathToFileURL(path.resolve('./src/cvMdTo.mjs')).href,
                T_WORK: path.resolve(fdTmp, 'work-toc'),
            },
        })
        assert.strict.equal(res.toc.skip, 'no heading (paragraph with an outline level) is found, the TOC is not added')
        assert.strict.equal(res.size, fs.statSync(fpFixture).size)
    })

    it('產出docx且keepCaption時結果帶回cvMdToDocx之keepCaption(假轉檔器取代w-html2docx)', function() {
        this.timeout(180000)
        if (process.platform !== 'win32') {
            this.skip()
        }
        let fpFixture = path.resolve(fdTmp, 'fixture-fig.docx')
        fs.writeFileSync(fpFixture, buildDocx(parts.cover() + parts.pic() + parts.capFig('圖1 系統架構')))
        let script = [
            `let m = await import(process.env.T_MOD)`,
            `let r = await m.default({ md: '# t', out: 'docx', keepCaption: true, dirWork: process.env.T_WORK })`,
            `process.stdout.write('@@RESULT@@' + JSON.stringify({ keepCaption: r.keepCaption }))`,
        ].join('\n')
        let res = runWithFakeHtml2docx(script, {
            fpDocx: fpFixture,
            env: {
                T_MOD: pathToFileURL(path.resolve('./src/cvMdTo.mjs')).href,
                T_WORK: path.resolve(fdTmp, 'work-keep'),
            },
        })
        assert.strict.deepEqual([res.keepCaption.figs, res.keepCaption.keepNext, res.keepCaption.keepLines, res.keepCaption.changed], [1, 1, 1, true])
    })

    it('out=html時keepCaption不適用(無docx), 結果無keepCaption', async function() {
        let r = await cvMdTo({ md: '# t', out: 'html', keepCaption: true })
        assert.strict.equal(r.keepCaption, undefined)
    })

    it('out=html: 回傳之html恆自足, optMd2html給imgConvertToBase64:false亦不得使圖以(將被刪除之)工作夾路徑引用', async function() {
        let r = await cvMdTo({
            md: '# t\n\n![p](pics/a.svg)\n\n<img src="pics/b.svg" />',
            out: 'html',
            assets: [{ path: 'pics/a.svg', base64: svgB64 }, { path: 'pics/b.svg', base64: svgB64 }],
            optMd2html: { imgConvertToBase64: false },
        })
        let h = toHtml(r)
        let srcs = [...h.matchAll(/<img\b[^>]*?\ssrc="([^"]*)"/g)].map((m) => m[1])
        assert.strict.equal(srcs.length, 2)
        assert.strict.equal(srcs.every((v) => v.startsWith('data:image/png;base64,')), true, srcs.map((v) => v.slice(0, 40)).join(' | '))
        assert.strict.equal(/job_\d+_/.test(h), false) //不含工作夾路徑
    })

})


//cvMdTo之docx路徑與回傳上限: 子程序內以假轉檔器取代w-html2docx(不需Word), 稀疏檔模擬超大docx(不實寫)
describe('cvMdTo 回傳內容(假轉檔器, 不需Word)', function() {

    //check, 假轉檔器掛鉤與docx階段僅支援Windows
    if (process.platform !== 'win32') {
        return
    }

    let fdTmp = path.resolve('./test/_tmp/unit-cvMdTo-fake')
    let fpFixture = path.resolve(fdTmp, 'fixture-nohead.docx')
    let png = (gray) => genTallPng(8, 8, gray).toString('base64')
    let nJob = (dir) => fs.readdirSync(dir).filter((v) => v.startsWith('job_')).length

    //runCvMdTo: 子程序內呼叫 cvMdTo, 回傳 { err, nFake, htmlSeen, docxSize, htmlSize, htmlText, keys }
    //  optJs 為傳予 cvMdTo 之設定物件之 JS 運算式文字(大資產於子程序內產生, 不經環境變數傳遞)
    let runCvMdTo = (optJs, fakeOpt, dirWork) => {
        let script = [
            `let m = await import(process.env.T_MOD)`,
            `let opt = (${optJs})`,
            `opt.dirWork = process.env.T_WORK`,
            `let r = null, err = null`,
            `await m.default(opt).then((v) => { r = v }).catch((e) => { err = typeof e === 'string' ? e : String(e && e.message) })`,
            `let o = { err, nFake: (globalThis.fakeHtml2docxOpts || []).length, htmlSeen: globalThis.fakeHtml2docxHtml || [] }`,
            `if (r) { o.keys = Object.keys(r); o.docxSize = r.docx ? r.docx.size : null; o.htmlSize = r.html ? r.html.size : null; o.htmlText = r.html ? Buffer.from(r.html.base64, 'base64').toString('utf8') : null }`,
            `process.stdout.write('@@RESULT@@' + JSON.stringify(o))`,
        ].join('\n')
        return runWithFakeHtml2docx(script, {
            ...fakeOpt,
            env: { T_MOD: pathToFileURL(path.resolve('./src/cvMdTo.mjs')).href, T_WORK: dirWork },
        })
    }

    before(function() {
        fs.mkdirSync(fdTmp, { recursive: true })
        fs.writeFileSync(fpFixture, buildDocx(parts.cover() + parts.body('沒有標題')))
    })

    after(function() {
        fs.rmSync(fdTmp, { recursive: true, force: true, maxRetries: 10, retryDelay: 300 })
    })

    it('out=both: html為自足之base64內嵌(與out=html同輸入之結果逐字相同), 不含工作夾路徑; docx階段之中介html則以本機絕對路徑引用圖', function() {
        this.timeout(180000)
        let dirWork = path.resolve(fdTmp, 'work-both')
        let assets = [{ path: 'pics/a.png', base64: png(60) }, { path: 'pics/b.png', base64: png(180) }]
        let md = '# 圖\n\n![甲](pics/a.png)\n\n<img src="pics/b.png" alt="乙">'
        let rBoth = runCvMdTo(JSON.stringify({ md, name: '手冊', out: 'both', assets }), { fpDocx: fpFixture }, dirWork)
        let rHtml = runCvMdTo(JSON.stringify({ md, name: '手冊', out: 'html', assets }), { fpDocx: fpFixture }, dirWork)
        assert.strict.equal(rBoth.err, null)
        assert.strict.equal(rBoth.docxSize, fs.statSync(fpFixture).size)
        let srcs = [...rBoth.htmlText.matchAll(/<img\b[^>]*?\ssrc="([^"]*)"/g)].map((m) => m[1])
        assert.strict.equal(srcs.length, 2)
        assert.strict.equal(srcs.every((v) => v.startsWith('data:image/png;base64,')), true, srcs.map((v) => v.slice(0, 40)).join(' | '))
        assert.strict.equal(/job_\d+_/.test(rBoth.htmlText), false)
        assert.strict.equal(rBoth.htmlText, rHtml.htmlText)
        //Word 階段(假轉檔器所收)之中介 html: 無 data URI, 圖以反斜線絕對路徑引用
        assert.strict.equal(rBoth.htmlSeen.length, 1)
        let seen = [...rBoth.htmlSeen[0].matchAll(/<img\b[^>]*?\ssrc="([^"]*)"/g)].map((m) => m[1])
        assert.strict.equal(seen.length, 2)
        assert.strict.equal(seen.every((v) => /^[A-Za-z]:\\.+pics\\[ab]\.png$/.test(v)), true, seen.join(' | '))
        assert.strict.equal(rBoth.htmlSeen[0].includes('data:image'), false)
        assert.strict.equal(nJob(dirWork), 0) //作業夾已清除
    })

    it('out=both: 圖片base64總量超過w-md2html上限時於轉docx前reject(不呼叫轉檔器), 訊息指出原因, 作業夾已清除', function() {
        this.timeout(180000)
        let dirWork = path.resolve(fdTmp, 'work-both-big')
        //100 MiB 之內容(w-md2html 以檔案大小預檢, 不讀內容): base64 約 1.4 億字元, 超過其 2^27 字元上限
        let optJs = `{ md: '# t\\n\\n![](a.png)', out: 'both', assets: [{ path: 'a.png', base64: Buffer.alloc(100 * 1024 * 1024).toString('base64') }] }`
        let r = runCvMdTo(optJs, { fpDocx: fpFixture }, dirWork)
        assert.strict.equal(/^Failed to convert md to html: images are too large to embed as base64 \(imgConvertToBase64=true\): /.test(r.err), true, String(r.err))
        assert.strict.equal(r.nFake, 0)
        assert.strict.equal(nJob(dirWork), 0)
    })

    it('缺圖: docx路徑與html路徑皆於轉docx前以相同之可行動訊息reject(不呼叫轉檔器)', function() {
        this.timeout(180000)
        let dirWork = path.resolve(fdTmp, 'work-miss')
        for (let out of ['docx', 'both', 'html']) {
            let r = runCvMdTo(JSON.stringify({ md: '# t\n\n![x](pics/no.png)', out }), { fpDocx: fpFixture }, dirWork)
            assert.strict.equal(r.err, 'Failed to convert md to html: the referenced file (pics/no.png) does not exist, please make sure it is included in assets', out)
            assert.strict.equal(r.nFake, 0, out)
        }
        assert.strict.equal(nJob(dirWork), 0)
    })

    it('Word解析不到工作夾內之圖(w-html2docx reject): 訊息經cvMdToDocx加前綴後隱去工作夾路徑, 作業夾已清除', function() {
        this.timeout(180000)
        let dirWork = path.resolve(fdTmp, 'work-unres')
        let reject = 'code[1]:\nunresolved linked picture: C:\\Users\\x\\AppData\\Local\\Temp\\w-md2docx\\job_123_abc\\pics\\a.png\nerror: LinkedPictureError: 1 linked picture(s) can not be embedded'
        let r = runCvMdTo(JSON.stringify({ md: '# t', out: 'docx' }), { reject }, dirWork)
        assert.strict.equal(r.err, 'Failed to convert html to docx: code[1]:\nunresolved linked picture: pics/a.png\nerror: LinkedPictureError: 1 linked picture(s) can not be embedded')
        assert.strict.equal(nJob(dirWork), 0)
    })

    it('docx超過可置於JSON字串之base64上限(約401,866,734位元組): 讀檔與轉base64前以指出原因之訊息reject, 作業夾已清除', function() {
        this.timeout(180000)
        let dirWork = path.resolve(fdTmp, 'work-big')
        let r = runCvMdTo(JSON.stringify({ md: '# t', out: 'docx' }), { size: 410000000 }, dirWork)
        assert.strict.equal(r.err, 'the result is too large to be returned as base64 in one JSON string: docx 410000000 bytes (546666668 chars of base64 in total), over the limit of 535822312 chars (about 401866734 bytes of files in total)')
        assert.strict.equal(nJob(dirWork), 0)
    })

    it('out=both: docx與html合計超過上限(docx單獨未超過)亦reject, 訊息列出兩者', function() {
        this.timeout(180000)
        let dirWork = path.resolve(fdTmp, 'work-big-both')
        //docx 401,866,000 B 單獨為 535,821,336 字元(未超過上限 535,822,312, 餘 976 字元), 加上 html(數 KB)後超過
        let r = runCvMdTo(JSON.stringify({ md: '# t', out: 'both' }), { size: 401866000 }, dirWork)
        assert.strict.equal(/^the result is too large to be returned as base64 in one JSON string: docx 401866000 bytes, html \d+ bytes \(\d+ chars of base64 in total\), over the limit of 535822312 chars \(about 401866734 bytes of files in total\)$/.test(r.err), true, String(r.err))
        assert.strict.equal(nJob(dirWork), 0)
    })

})
