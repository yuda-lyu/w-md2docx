import fs from 'fs'
import path from 'path'
import { pathToFileURL } from 'url'
import assert from 'assert'
import cvMdToDocx from '../src/cvMdToDocx.mjs'
import runWithFakeHtml2docx from './tools/fakeHtml2docx.mjs'
import { mkFixture, imgSrcs, nTmpHtmlOfThisProcess } from './tools/docxImg.mjs'


//cvMdToDocx之圖片處理(不需Word): 中介html一律以本機絕對路徑引用圖(不內嵌base64), 缺圖於md轉html階段reject
//  (md2html 1.1.12 imgConvertToBase64:false 寫反斜線絕對路徑並於缺圖時reject; 圖內嵌於docx之實轉測試見 unit-cvMdToDocx.test.mjs
//  之〈圖片內嵌於docx〉, 因 Word 不可跨 mocha 平行行程同時調用, 凡調用 Word 之測試併於該檔依序執行)
let fdTmpRoot = path.resolve('./test/_tmp/unit-cvMdToDocxImg')


describe('cvMdToDocx 中介html以本機絕對路徑引用圖(假轉檔器, 不需Word)', function() {

    //check, 假轉檔器掛鉤僅支援Windows
    if (process.platform !== 'win32') {
        return
    }

    let fd = path.resolve(fdTmpRoot, 'fake')
    let fpMd = path.resolve(fd, 'in.md')
    let fpFixture = path.resolve(fd, 'fixture.docx')

    //run: 子程序內(假轉檔器取代w-html2docx)呼叫cvMdToDocx, 回傳 { r, err, seen(假轉檔器所收之中介html) }
    let run = (opt) => {
        let script = [
            `let m = await import(process.env.T_MOD)`,
            `let r = null, err = null`,
            `await m.default(process.env.T_MD, process.env.T_DOCX, JSON.parse(process.env.T_OPT)).then((v) => { r = v }).catch((e) => { err = e })`,
            `process.stdout.write('@@RESULT@@' + JSON.stringify({ r, err, seen: globalThis.fakeHtml2docxHtml || [] }))`,
        ].join('\n')
        return runWithFakeHtml2docx(script, {
            fpDocx: fpFixture,
            env: {
                T_MOD: pathToFileURL(path.resolve('./src/cvMdToDocx.mjs')).href,
                T_MD: fpMd,
                T_DOCX: path.resolve(fd, 'out.docx'),
                T_OPT: JSON.stringify(opt),
            },
        })
    }

    before(function() {
        fs.mkdirSync(fd, { recursive: true })
        mkFixture(fd)
        fs.writeFileSync(fpFixture, 'fake docx')
        fs.writeFileSync(fpMd, [
            '# 圖片路徑',
            '',
            '![甲](<圖 資料夾/a.png>)',
            '',
            '<img src="圖 資料夾/b.png" alt="乙">',
            '',
            '![丙](c.png)',
            '',
            '![丁](cocktail.svg)',
            '',
            '<img src="https://example.com/x.png" alt="網址圖">',
            '',
        ].join('\n'), 'utf8')
    })

    after(function() {
        fs.rmSync(fdTmpRoot, { recursive: true, force: true, maxRetries: 10, retryDelay: 300 })
    })

    //本機圖之預期絕對路徑(html內順序依w-md2html, 故以排序後集合比對)
    let expected = () => [
        path.resolve(fd, '圖 資料夾', 'a.png'),
        path.resolve(fd, '圖 資料夾', 'b.png'),
        path.resolve(fd, 'c.png'),
        path.resolve(fd, 'cocktail.svg'),
    ].sort()

    it('中介html不內嵌base64: 本機圖(md語法圖與原始<img>, 路徑含中文與空白)皆為存在之反斜線絕對路徑, 網址圖原樣', function() {
        this.timeout(180000)
        let res = run({})
        assert.strict.equal(res.err, null)
        assert.strict.equal(res.seen.length, 1)
        let html = res.seen[0]
        assert.strict.equal(html.includes('data:'), false)
        let srcs = imgSrcs(html)
        assert.strict.equal(srcs.length, 5)
        let locals = srcs.filter((v) => !v.startsWith('https://'))
        assert.strict.equal(locals.every((v) => /^[A-Za-z]:\\/.test(v) && fs.existsSync(v)), true, locals.join(' | '))
        assert.strict.deepEqual([...locals].sort(), expected())
        assert.strict.deepEqual(srcs.filter((v) => v.startsWith('https://')), ['https://example.com/x.png'])
    })

    it('optMd2html給imgConvertToBase64:true亦不得覆寫(中介html仍無base64), 其餘md2html設定照傳', function() {
        this.timeout(180000)
        let res = run({ optMd2html: { imgConvertToBase64: true, imgWidthMax: '123px' } })
        assert.strict.equal(res.err, null)
        let html = res.seen[0]
        assert.strict.equal(html.includes('data:'), false)
        assert.strict.deepEqual(imgSrcs(html).filter((v) => !v.startsWith('https://')).sort(), expected())
        assert.strict.equal(/max-width:\s*123px/.test(html), true)
    })

    it('fpOutHtml另存之中介html與交予轉檔器者相同(供Word之中介檔, 以本機絕對路徑引用圖, 不自足)', function() {
        this.timeout(180000)
        let fpOutHtml = path.resolve(fd, 'keep.html')
        let res = run({ fpOutHtml })
        assert.strict.equal(res.err, null)
        assert.strict.equal(res.r.fpOutHtml, fpOutHtml)
        let kept = fs.readFileSync(fpOutHtml, 'utf8')
        assert.strict.equal(kept, res.seen[0])
        assert.strict.equal(kept.includes('data:'), false)
        assert.strict.equal(res.r.sizeHtml, fs.statSync(fpOutHtml).size)
    })

    it('缺圖(md語法圖與原始<img>併用): 於md轉html階段reject, 訊息列出全部缺圖之絕對路徑; 不產出docx、不留中介html、不寫出另存之html', async function() {
        let fpMdMiss = path.resolve(fd, 'miss.md')
        let fpDocx = path.resolve(fd, 'miss.docx')
        let fpOutHtml = path.resolve(fd, 'miss.html')
        fs.writeFileSync(fpMdMiss, '# t\n\n![甲](<圖 資料夾/a.png>)\n\n![乙](nope1.png)\n\n<img src="沒有 的圖/nope2.png">\n', 'utf8') //甲存在, 不列入
        let n0 = nTmpHtmlOfThisProcess()
        let err = null
        await cvMdToDocx(fpMdMiss, fpDocx, { fpOutHtml }).catch((e) => {
            err = e
        })
        let lines = String(err).split('\n')
        assert.strict.equal(lines[0], 'Failed to convert md to html: 2 image file(s) do not exist:')
        assert.strict.deepEqual(lines.slice(1).sort(), [path.resolve(fd, 'nope1.png'), path.resolve(fd, '沒有 的圖', 'nope2.png')].sort())
        assert.strict.equal(fs.existsSync(fpDocx), false)
        assert.strict.equal(fs.existsSync(fpOutHtml), false)
        assert.strict.equal(nTmpHtmlOfThisProcess(), n0)
    })

    it('缺圖: 單一缺圖之訊息', async function() {
        let fpMdMiss = path.resolve(fd, 'miss1.md')
        fs.writeFileSync(fpMdMiss, '# t\n\n![乙](nope1.png)\n', 'utf8')
        await assert.rejects(cvMdToDocx(fpMdMiss, path.resolve(fd, 'miss1.docx')), (e) => {
            assert.strict.equal(e, `Failed to convert md to html: 1 image file(s) do not exist:\n${path.resolve(fd, 'nope1.png')}`)
            return true
        })
    })

})
