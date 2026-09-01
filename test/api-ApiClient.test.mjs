import fs from 'fs'
import path from 'path'
import assert from 'assert'
import ApiServer from '../src/ApiServer.mjs'
import ApiClient, { cvMdTo, health, scanAssetPaths, readAssets } from '../src/ApiClient.mjs'


describe('ApiClient', function() {

    describe('scanAssetPaths(純函數)', function() {

        it('涵蓋<img src>與![]() , 去重, 保留原字串', function() {
            let md = [
                '![a](pics/a.png)',
                '<img src="pics/b.png" />',
                `<IMG class='x' SRC='pics/c.png'>`,
                '![a2](pics/a.png)',
                '![e](<pics/e.png>)',
                '![enc](pics/%E5%9C%96.png)',
            ].join('\n')
            assert.strict.deepEqual(scanAssetPaths(md), ['pics/b.png', 'pics/c.png', 'pics/a.png', 'pics/e.png', 'pics/%E5%9C%96.png'])
        })

        it('排除網路位址/data/mailto/錨點/絕對路徑', function() {
            let md = [
                '![1](https://x.com/1.png)',
                '![2](//x.com/2.png)',
                '![3](data:image/png;base64,AAAA)',
                '[m](mailto:a@b.c)',
                '![4](#top)',
                '![5](C:/x/5.png)',
                '![6](/abs/6.png)',
                '![ok](rel/ok.png)',
            ].join('\n')
            assert.strict.deepEqual(scanAssetPaths(md), ['rel/ok.png'])
        })

        it('去除錨點與查詢字串', function() {
            assert.strict.deepEqual(scanAssetPaths('![a](pics/a.png?v=1#x)'), ['pics/a.png'])
        })

        it('非字串回空陣列', function() {
            assert.strict.deepEqual(scanAssetPaths(null), [])
            assert.strict.deepEqual(scanAssetPaths(123), [])
        })

    })

    describe('readAssets(讀本機檔)', function() {

        let fdTmp = path.resolve('./tmp/zt_ApiClient_ra')

        before(function() {
            fs.mkdirSync(path.resolve(fdTmp, 'pics'), { recursive: true })
            fs.writeFileSync(path.resolve(fdTmp, 'pics/圖.png'), 'IMG')
            fs.writeFileSync(path.resolve(fdTmp, 'pics/a b.png'), 'IMG2')
        })

        after(function() {
            fs.rmSync(fdTmp, { recursive: true, force: true })
        })

        it('URL編碼引用可對應解碼後實體檔, path保留原字串', function() {
            let { assets, missing } = readAssets('![p](pics/%E5%9C%96.png)\n<img src="pics/a%20b.png">', fdTmp)
            assert.strict.deepEqual(missing, [])
            assert.strict.deepEqual(assets.map((v) => v.path), ['pics/a%20b.png', 'pics/%E5%9C%96.png'])
            assert.strict.equal(Buffer.from(assets[1].base64, 'base64').toString('utf8'), 'IMG')
        })

        it('找不到者列入missing', function() {
            let { assets, missing } = readAssets('![p](pics/圖.png)\n![n](pics/no.png)', fdTmp)
            assert.strict.equal(assets.length, 1)
            assert.strict.deepEqual(missing, ['pics/no.png'])
        })

        it('dirMd未給以cwd為基準', function() {
            let { assets, missing } = readAssets('![c](test/cocktail.svg)')
            assert.strict.equal(assets.length, 1)
            assert.strict.deepEqual(missing, [])
        })

    })

    describe('cvMdTo/health(對接ApiServer, html路徑不需Word)', function() {

        let fdTmp = path.resolve('./tmp/zt_ApiClient')
        let dirWork = path.resolve(fdTmp, 'work')
        let token = 'tk-client'
        let srv = null
        let host = '127.0.0.1'
        let port = 0
        let url = ''
        let fpOutHtml = path.resolve(fdTmp, 'out/report.html')

        before(async function() {
            fs.mkdirSync(fdTmp, { recursive: true })
            srv = await ApiServer({ port: 0, host, token, dirWork })
            port = srv.server.info.port
            url = `http://${host}:${port}`
        })

        after(async function() {
            if (srv) {
                await srv.stop()
            }
            fs.rmSync(fdTmp, { recursive: true, force: true })
        })

        it('default export匯整四函數', function() {
            assert.strict.deepEqual(Object.keys(ApiClient), ['cvMdTo', 'health', 'scanAssetPaths', 'readAssets'])
            assert.strict.equal(ApiClient.cvMdTo, cvMdTo)
        })

        it('health', async function() {
            let r = await health({ host, port, token })
            assert.strict.equal(r.success, true)
            assert.strict.equal(r.tokenRequired, true)
            assert.strict.deepEqual(r.templates, ['temp_tpc.docx'])
        })

        it('host/port未給時取環境變數WMD2DOCX_HOST/WMD2DOCX_PORT', async function() {
            process.env.WMD2DOCX_HOST = host
            process.env.WMD2DOCX_PORT = String(port)
            try {
                let r = await health({ token })
                assert.strict.equal(r.success, true)
            }
            finally {
                delete process.env.WMD2DOCX_HOST
                delete process.env.WMD2DOCX_PORT
            }
        })

        it('host/port與環境變數皆未給時預設127.0.0.1:22000(與ApiServer預設埠一致)', async function() {
            await assert.rejects(health({ token, timeoutMs: 3000 }), (e) => /^Unable to connect to the conversion service http:\/\/127\.0\.0\.1:22000: /.test(e) || /^Invalid response from the conversion service http:\/\/127\.0\.0\.1:22000/.test(e))
        })

        it('port非正整數視為未給', async function() {
            await assert.rejects(health({ host, port: 'x', token, timeoutMs: 3000 }), (e) => /http:\/\/127\.0\.0\.1:22000/.test(e))
        })

        it('url給予時覆寫host/port', async function() {
            let r = await health({ url, host: '10.255.255.1', port: 1, token })
            assert.strict.equal(r.success, true)
        })

        it('health 連不上則reject', async function() {
            await assert.rejects(health({ host: '127.0.0.1', port: 1, token, timeoutMs: 3000 }), (e) => /^Unable to connect to the conversion service http:\/\/127\.0\.0\.1:1: /.test(e))
        })

        it('cvMdTo fpOutHtml: 寫出檔案並回傳大小(自動夾帶md引用之圖)', async function() {
            let r = await cvMdTo({ host, port, token, fpInMd: './test/report.md', fpOutHtml })
            assert.strict.equal(r.nAssets, 1) //report.md 引用 ./cocktail.svg
            assert.strict.equal(r.template, 'default')
            assert.strict.equal(typeof r.ms, 'number')
            assert.strict.equal(r.docx, undefined)
            assert.strict.equal(r.html.fp, fpOutHtml)
            assert.strict.equal(r.html.size, fs.statSync(fpOutHtml).size)
            let h = fs.readFileSync(fpOutHtml, 'utf8')
            assert.strict.equal(h.includes('計畫報告'), true)
            assert.strict.equal(h.includes('<title>report</title>'), true) //name預設取md檔名
            assert.strict.equal(h.includes('src="data:image/png;base64,'), true)
        })

        it('cvMdTo name/url尾斜線/環境變數token', async function() {
            let fpOut2 = path.resolve(fdTmp, 'out/n.html')
            process.env.WMD2DOCX_TOKEN = token
            try {
                let r = await cvMdTo({ url: `${url}///`, fpInMd: './test/report.md', fpOutHtml: fpOut2, name: '自訂名' })
                assert.strict.equal(r.html.fp, fpOut2)
                assert.strict.equal(fs.readFileSync(fpOut2, 'utf8').includes('<title>自訂名</title>'), true)
            }
            finally {
                delete process.env.WMD2DOCX_TOKEN
            }
        })

        it('cvMdTo 輸入檢查', async function() {
            await assert.rejects(cvMdTo({ host, port, token, fpOutHtml }), (e) => e === 'fpInMd must be a non-empty string')
            await assert.rejects(cvMdTo({ host, port, token, fpInMd: './test/nope.md', fpOutHtml }), (e) => /^fpInMd\[.+nope\.md\] does not exist$/.test(e))
            await assert.rejects(cvMdTo({ host, port, token, fpInMd: './test/report.md', fpOutHtml, fpInTemp: './nope.docx' }), (e) => /^fpInTemp\[.+nope\.docx\] does not exist$/.test(e))
        })

        it('cvMdTo 資產缺漏預設reject, allowMissingAssets=true放行', async function() {
            let fpMd = path.resolve(fdTmp, 'missing.md')
            fs.writeFileSync(fpMd, '# m\n\n![x](nope.png)\n![y](nope2.png)\n', 'utf8')
            await assert.rejects(cvMdTo({ host, port, token, fpInMd: fpMd, fpOutHtml }), (e) => e === '2 asset file(s) referenced in md do not exist: nope.png, nope2.png. Set allowMissingAssets:true to skip this check')
            let r = await cvMdTo({ host, port, token, fpInMd: fpMd, fpOutHtml, allowMissingAssets: true })
            assert.strict.equal(r.nAssets, 0)
            assert.strict.equal(r.html.size > 0, true)
        })

        it('cvMdTo fpInTemp 以base64夾帶模板(template=request)', async function() {
            let r = await cvMdTo({ host, port, token, fpInMd: './test/report.md', fpOutHtml, fpInTemp: './src/templates/temp_tpc.docx' })
            assert.strict.equal(r.template, 'request')
        })

        it('cvMdTo templateName 指定服務端模板(template=server), 不存在則HTTP錯誤不重試', async function() {
            let r = await cvMdTo({ host, port, token, fpInMd: './test/report.md', fpOutHtml, templateName: 'temp_tpc.docx' })
            assert.strict.equal(r.template, 'server')
            await assert.rejects(cvMdTo({ host, port, token, fpInMd: './test/report.md', fpOutHtml, templateName: 'nope.docx' }), (e) => e === 'Conversion failed: templateName[nope.docx] does not exist (available templates: temp_tpc.docx)')
        })

        it('cvMdTo token錯誤', async function() {
            await assert.rejects(cvMdTo({ host, port, token: 'bad', fpInMd: './test/report.md', fpOutHtml }), (e) => e === 'Conversion failed: unauthorized: a valid x-api-token header is required')
        })

        it('cvMdTo 連不上: retries=0不重試', async function() {
            let t0 = Date.now()
            await assert.rejects(cvMdTo({ host: '127.0.0.1', port: 1, token, fpInMd: './test/report.md', fpOutHtml, retries: 0 }), (e) => /^Unable to connect to the conversion service http:\/\/127\.0\.0\.1:1 \(retried 0 times\): /.test(e))
            assert.strict.equal(Date.now() - t0 < 3000, true) //無重試等待
        })

        it('cvMdTo 逾時: 不重試, 訊息含timeoutMs', async function() {
            //以極短timeout使fetch逾時(轉檔需時>1ms)
            await assert.rejects(cvMdTo({ host, port, token, fpInMd: './test/report.md', fpOutHtml, timeoutMs: 1 }), (e) => e === `Conversion timed out (no response within 1 ms): ${url}`)
        })

    })

})
