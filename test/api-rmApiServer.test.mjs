import fs from 'fs'
import path from 'path'
import assert from 'assert'
import rmApiServer from '../src/rmApiServer.mjs'


//rmApiServer: 直打HTTP端點(不經UI); docx相關以html路徵替代, 不需Word
describe('rmApiServer', function() {

    let fdTmp = path.resolve('./test/_tmp/api-rmApiServer')
    let dirWork = path.resolve(fdTmp, 'work')
    let token = 'tk-test'
    let svgB64 = fs.readFileSync('./test/cocktail.svg').toString('base64')
    let portMain = 8201 //固定埠(8000以上), 與api-rmApiClient(8211)錯開以免並行撞埠
    let portStop = 8202
    let srv = null
    let url = ''

    let post = (body, q = '', tk = token) => {
        let headers = { 'content-type': 'application/json' }
        if (tk !== '') {
            headers['x-api-token'] = tk
        }
        return fetch(`${url}/api/convert${q}`, { method: 'POST', headers, body: JSON.stringify(body) })
    }

    before(async function() {
        fs.mkdirSync(fdTmp, { recursive: true })
        srv = await rmApiServer({ port: portMain, host: '127.0.0.1', token, dirWork })
        url = `http://127.0.0.1:${portMain}`
    })

    after(async function() {
        if (srv) {
            await srv.stop()
        }
        fs.rmSync(fdTmp, { recursive: true, force: true, maxRetries: 10, retryDelay: 300 })
    })

    it('回傳物件與settings', function() {
        assert.strict.equal(srv.server.info.port, portMain)
        assert.strict.equal(srv.settings.port, portMain)
        assert.strict.equal(typeof srv.getState, 'function')
        assert.strict.equal(typeof srv.stop, 'function')
        assert.strict.equal(srv.settings.host, '127.0.0.1')
        assert.strict.equal(srv.settings.token, token)
        assert.strict.equal(srv.settings.maxMb, 256)
        assert.strict.equal(srv.settings.maxBytes, 256 * 1024 * 1024)
        assert.strict.equal(srv.settings.dirWork, dirWork)
        assert.strict.equal(srv.settings.templateDef, 'temp_tpc.docx')
        assert.strict.equal(/src[\\/]templates$/.test(srv.settings.dirTemplates), true) //預設模板夾為套件 src/templates
        assert.strict.equal(fs.existsSync(dirWork), true) //啟動時建立工作夾
    })

    it('GET / 不需token, 回服務說明', async function() {
        let res = await fetch(`${url}/`)
        assert.strict.equal(res.status, 200)
        let r = await res.json()
        assert.strict.equal(r.service, 'w-md2docx')
        assert.strict.deepEqual(Object.keys(r.endpoints), ['GET /api/health', 'GET /api/selftest', 'POST /api/convert'])
    })

    it('/api/* 無token或錯token回401', async function() {
        let res = await fetch(`${url}/api/health`)
        assert.strict.equal(res.status, 401)
        let r = await res.json()
        assert.strict.deepEqual(r, { success: false, error: 'unauthorized: a valid x-api-token header is required' })
        let res2 = await post({ md: '# t', out: 'html' }, '', 'bad')
        assert.strict.equal(res2.status, 401)
    })

    it('GET /api/health 回環境狀態', async function() {
        let res = await fetch(`${url}/api/health`, { headers: { 'x-api-token': token } })
        assert.strict.equal(res.status, 200)
        let r = await res.json()
        assert.strict.equal(r.success, true)
        assert.strict.equal(r.platform, process.platform)
        assert.strict.equal(r.node, process.version)
        assert.strict.equal(r.tokenRequired, true)
        assert.strict.equal(r.dirWork, dirWork)
        assert.strict.deepEqual(r.templates, ['temp_tpc.docx'])
        assert.strict.equal(/temp_tpc\.docx$/.test(r.templateDefault), true)
        assert.strict.equal(r.queue, 0)
        assert.strict.equal(typeof r.winwordProcesses, 'number')
        assert.strict.equal(r.maxBytes, 256 * 1024 * 1024)
        assert.strict.equal(typeof r.docxReady, 'boolean')
        let st = await srv.getState()
        assert.strict.equal(st.docxReady, r.docxReady)
    })

    it('POST /api/convert out=html: 回傳結構與資產內嵌', async function() {
        let res = await post({
            md: '# 標題A\n\n![p](pics/圖.svg)',
            name: '報告',
            out: 'html',
            assets: [{ path: 'pics/圖.svg', base64: svgB64 }],
        })
        assert.strict.equal(res.status, 200)
        let r = await res.json()
        assert.strict.equal(r.success, true)
        assert.strict.equal(r.name, '報告')
        assert.strict.equal(r.out, 'html')
        assert.strict.equal(r.nAssets, 1)
        assert.strict.equal(r.template, 'default') //服務端預設模板(temp_tpc.docx)存在
        assert.strict.equal(typeof r.ms, 'number')
        assert.strict.equal(r.docx, undefined)
        assert.strict.equal(r.html.fileName, '報告.html')
        assert.strict.equal(r.html.mime, 'text/html; charset=utf-8')
        let h = Buffer.from(r.html.base64, 'base64').toString('utf8')
        assert.strict.equal(h.includes('標題A'), true)
        assert.strict.equal(h.includes('src="data:image/png;base64,'), true)
    })

    it('POST /api/convert mdBase64 與 templateName 指定服務端模板', async function() {
        let res = await post({ mdBase64: Buffer.from('# b64', 'utf8').toString('base64'), out: 'html', templateName: 'temp_tpc.docx' })
        assert.strict.equal(res.status, 200)
        let r = await res.json()
        assert.strict.equal(r.template, 'server')
    })

    it('POST /api/convert templateBase64 夾帶模板, 用畢刪除暫存模板', async function() {
        let tplB64 = fs.readFileSync('./src/templates/temp_tpc.docx').toString('base64')
        let res = await post({ md: '# t', out: 'html', templateBase64: tplB64 })
        assert.strict.equal(res.status, 200)
        let r = await res.json()
        assert.strict.equal(r.template, 'request')
        assert.strict.equal(fs.readdirSync(dirWork).filter((v) => v.startsWith('tmpl_')).length, 0)
    })

    it('POST /api/convert 非JSON物件回400', async function() {
        let res = await fetch(`${url}/api/convert`, { method: 'POST', headers: { 'content-type': 'application/json', 'x-api-token': token }, body: '"str"' })
        assert.strict.equal(res.status, 400)
        let r = await res.json()
        assert.strict.deepEqual(r, { success: false, error: 'parameters must be provided as a JSON object' })
    })

    it('POST /api/convert 輸入錯誤回400(md為空/templateName不存在/資產穿越)', async function() {
        let res1 = await post({ out: 'html' })
        assert.strict.equal(res1.status, 400)
        assert.strict.deepEqual(await res1.json(), { success: false, error: 'md is empty: md (string) or mdBase64 is required' })
        let res2 = await post({ md: '# t', out: 'html', templateName: 'nope.docx' })
        assert.strict.equal(res2.status, 400)
        assert.strict.deepEqual(await res2.json(), { success: false, error: 'templateName[nope.docx] does not exist (available templates: temp_tpc.docx)' })
        let res3 = await post({ md: '# t', out: 'html', assets: [{ path: '../x.png', base64: 'QUI=' }] })
        assert.strict.equal(res3.status, 400)
        assert.strict.deepEqual(await res3.json(), { success: false, error: 'asset.path is outside the working folder: ../x.png' })
    })

    it('POST /api/convert 轉檔錯誤(資產缺漏)回500', async function() {
        let res = await post({ md: '# t\n\n<img src="no.png" />', out: 'html' })
        assert.strict.equal(res.status, 500)
        assert.strict.deepEqual(await res.json(), { success: false, error: 'Failed to convert md to html: the referenced file (no.png) does not exist, please make sure it is included in assets' })
    })

    it('POST /api/convert?download=1 單一格式直接回二進位, 中文檔名走filename*', async function() {
        let res = await post({ md: '# hi', out: 'html', name: '報告' }, '?download=1')
        assert.strict.equal(res.status, 200)
        assert.strict.equal(res.headers.get('content-type').startsWith('text/html'), true)
        assert.strict.equal(res.headers.get('content-disposition'), `attachment; filename="__.html"; filename*=UTF-8''%E5%A0%B1%E5%91%8A.html`)
        let h = await res.text()
        assert.strict.equal(h.includes('hi'), true)
    })

    it('POST /api/convert?download=true out=both 轉檔前即回400', async function() {
        let res = await post({ md: '# hi', out: 'both' }, '?download=true')
        assert.strict.equal(res.status, 400)
        assert.strict.deepEqual(await res.json(), { success: false, error: 'download only supports out of html or docx (single format)' })
    })

    it('工作夾於請求後不殘留job_', async function() {
        assert.strict.equal(fs.readdirSync(dirWork).filter((v) => v.startsWith('job_')).length, 0)
    })

    it('opt非法值採預設; autoStart=false不啟動', async function() {
        let s2 = await rmApiServer({ port: 'x', host: '', token: 5, maxMb: -1, templateDef: '', autoStart: false, dirWork: path.resolve(fdTmp, 'work2') })
        assert.strict.equal(s2.settings.port, 22000)
        assert.strict.equal(s2.settings.host, '0.0.0.0')
        assert.strict.equal(s2.settings.token, '')
        assert.strict.equal(s2.settings.maxMb, 256)
        assert.strict.equal(s2.settings.templateDef, 'temp_tpc.docx')
        assert.strict.equal(s2.server.info.started, 0) //未啟動
        await s2.stop()
    })

    it('stop後不再回應', async function() {
        let s3 = await rmApiServer({ port: portStop, host: '127.0.0.1', dirWork: path.resolve(fdTmp, 'work3') })
        let u3 = `http://127.0.0.1:${portStop}`
        let res = await fetch(`${u3}/api/health`)
        assert.strict.equal(res.status, 200) //無token時不需標頭
        await s3.stop()
        await assert.rejects(fetch(`${u3}/api/health`, { signal: AbortSignal.timeout(3000) }))
    })

})
