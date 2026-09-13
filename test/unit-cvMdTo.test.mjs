import fs from 'fs'
import path from 'path'
import assert from 'assert'
import cvMdTo from '../src/cvMdTo.mjs'


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

})
