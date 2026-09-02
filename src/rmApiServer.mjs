import fs from 'fs'
import os from 'os'
import path from 'path'
import { fileURLToPath } from 'url'
import { execFile } from 'child_process'
import get from 'lodash-es/get.js'
import isstr from 'wsemi/src/isstr.mjs'
import isestr from 'wsemi/src/isestr.mjs'
import isbol from 'wsemi/src/isbol.mjs'
import isobj from 'wsemi/src/isobj.mjs'
import isp0int from 'wsemi/src/isp0int.mjs'
import ispnum from 'wsemi/src/ispnum.mjs'
import cint from 'wsemi/src/cint.mjs'
import cdbl from 'wsemi/src/cdbl.mjs'
import fsIsFile from 'wsemi/src/fsIsFile.mjs'
import fsIsFolder from 'wsemi/src/fsIsFolder.mjs'
import Hapi from '@hapi/hapi'
import cvMdTo from './cvMdTo.mjs'
import { toErrText, checkDocxReady, getQueueSize, cleanWorkDir } from './utils.mjs'


//fdSelf (場景B: 內建模板隨模組所在資料夾, 不依啟動時之相對路徑漂移)
let fdSelf = path.dirname(fileURLToPath(import.meta.url))


//getWinwordCount: 取當前 WINWORD 行程數(診斷用)
//why: 殘留之 WINWORD 行程會鎖檔並使後續轉檔失敗, 健檢時一併呈現供判斷
function getWinwordCount() {
    return new Promise((resolve) => {
        if (process.platform !== 'win32') {
            resolve(-1)
            return
        }
        execFile('tasklist', ['/FI', 'IMAGENAME eq WINWORD.EXE', '/NH'], { windowsHide: true, timeout: 8000 }, (err, stdout) => {
            if (err) {
                resolve(-1)
                return
            }
            let n = (String(stdout).match(/WINWORD\.EXE/gi) || []).length
            resolve(n)
        })
    })
}


//toContentDisposition: 組附檔標頭(中文檔名以 RFC5987 之 filename* 提供)
function toContentDisposition(fileName) {
    let ascii = fileName.replace(/[^\x20-\x7E]/g, '_').replace(/"/g, '')
    return `attachment; filename="${ascii}"; filename*=UTF-8''${encodeURIComponent(fileName)}`
}


//isInputError: 判斷是否為輸入面錯誤(回 400), 其餘視為轉檔或環境錯誤(回 500)
//note: 前綴須與 cvMdTo/utils 之錯誤訊息一致
function isInputError(msg) {
    return /^(md is empty|asset|templateName)/.test(msg)
}


/**
 * 建立並啟動Markdown轉Html/Docx之API服務
 *
 * 以hapi提供三個端點：GET /api/health(健康檢查與環境狀態)、GET /api/selftest(實跑一次最小轉檔驗證本機Word可用)、POST /api/convert(轉檔主端點，JSON參數同cvMdTo，另可用templateName指定服務端模板或templateBase64夾帶模板；query帶download=1時單一格式直接回傳二進位檔)。
 *
 * @param {Object} [opt={}] 輸入設定物件，預設{}
 * @param {Integer} [opt.port=22000] 輸入服務埠號整數，給0則由系統自動配置(實際埠號見回傳之server.info.port)，預設22000
 * @param {String} [opt.host='0.0.0.0'] 輸入監聽位址字串，預設'0.0.0.0'
 * @param {String} [opt.token=''] 輸入權杖字串，非空時所有/api/*須帶x-api-token標頭，預設''
 * @param {Number} [opt.maxMb=256] 輸入請求大小上限數字，單位MB，預設256
 * @param {String} [opt.dirTemplates=''] 輸入docx模板資料夾位置字串，未給則用本套件src/templates，預設''
 * @param {String} [opt.dirWork=''] 輸入工作資料夾位置字串，未給則用系統暫存夾下之w-md2docx，預設''
 * @param {String} [opt.templateDef='temp_tpc.docx'] 輸入預設模板檔名字串，預設'temp_tpc.docx'
 * @param {Boolean} [opt.autoStart=true] 輸入是否立即啟動服務布林值，預設true
 * @returns {Promise} 回傳Promise，resolve回傳{server,settings,getState,stop}，其中server為hapi實例，getState為取環境狀態之async函數，stop為停止服務之async函數，reject回傳錯誤訊息
 * @example
 *
 * import rmApiServer from 'w-md2docx/src/rmApiServer.mjs'
 *
 * let { server, settings } = await rmApiServer({
 *     port: 22000,
 *     token: '',
 * })
 * console.log(`listening on ${server.info.uri}`, settings)
 *
 */
async function rmApiServer(opt = {}) {

    //port (0 為由系統自動配置, 供測試等場景避免撞埠; 實際埠號由 server.info.port 取得)
    let port = get(opt, 'port', null)
    if (!isp0int(port)) {
        port = 22000
    }
    port = cint(port)

    //host
    let host = get(opt, 'host', '')
    if (!isestr(host)) {
        host = '0.0.0.0'
    }
    host = host.trim()

    //token
    let token = get(opt, 'token', '')
    if (!isstr(token)) {
        token = ''
    }
    token = token.trim()

    //maxMb
    let maxMb = get(opt, 'maxMb', null)
    if (!ispnum(maxMb)) {
        maxMb = 256
    }
    maxMb = cdbl(maxMb)
    let maxBytes = Math.round(maxMb * 1024 * 1024)

    //dirTemplates
    let dirTemplates = get(opt, 'dirTemplates', '')
    if (!isestr(dirTemplates)) {
        dirTemplates = path.resolve(fdSelf, 'templates')
    }
    dirTemplates = path.resolve(dirTemplates)

    //dirWork
    let dirWork = get(opt, 'dirWork', '')
    if (!isestr(dirWork)) {
        dirWork = path.join(os.tmpdir(), 'w-md2docx')
    }
    dirWork = path.resolve(dirWork)

    //templateDef
    let templateDef = get(opt, 'templateDef', '')
    if (!isestr(templateDef)) {
        templateDef = 'temp_tpc.docx'
    }
    templateDef = templateDef.trim()

    //autoStart
    let autoStart = get(opt, 'autoStart', true)
    if (!isbol(autoStart)) {
        autoStart = true
    }

    //listTemplates: 列出可用之 docx 模板檔名
    let listTemplates = () => {
        if (!fsIsFolder(dirTemplates)) {
            return []
        }
        return fs.readdirSync(dirTemplates).filter((v) => /\.docx$/i.test(v))
    }

    //getFpTemplateDef: 取預設模板位置(不存在則回空字串)
    let getFpTemplateDef = () => {
        let fp = path.resolve(dirTemplates, path.basename(templateDef))
        return fsIsFile(fp) ? fp : ''
    }

    //pickTemplate: 決定本次使用之模板
    //順序: 請求端夾帶(base64) > 請求指定之服務端模板名 > 服務端預設模板 > 空(交由底層自行決定)
    //回傳 { fpInTemp, from, fpTmpDel } ; fpTmpDel 為須於用畢刪除之暫存模板檔
    let pickTemplate = (inp) => {

        //請求端夾帶
        let b64 = get(inp, 'templateBase64', '')
        if (!isestr(b64)) {
            b64 = get(inp, 'template.base64', '')
        }
        if (isestr(b64)) {
            b64 = b64.replace(/^data:[^;]+;base64,/, '')
            fs.mkdirSync(dirWork, { recursive: true })
            let fp = path.resolve(dirWork, `tmpl_${Date.now()}_${Math.random().toString(36).slice(2, 8)}.docx`)
            fs.writeFileSync(fp, Buffer.from(b64, 'base64'))
            return { fpInTemp: fp, from: 'request', fpTmpDel: fp }
        }

        //指定服務端模板名
        let fn = get(inp, 'templateName', '')
        if (isestr(fn)) {
            fn = fn.trim()
        }
        else {
            fn = ''
        }
        if (fn !== '') {
            let fp = path.resolve(dirTemplates, path.basename(fn)) //basename 化, 限定取用 templates 內之檔案
            if (!fsIsFile(fp)) {
                throw new Error(`templateName[${fn}] does not exist (available templates: ${listTemplates().join(', ') || 'none'})`)
            }
            return { fpInTemp: fp, from: 'server', fpTmpDel: '' }
        }

        //服務端預設模板
        let fpDef = getFpTemplateDef()
        if (fpDef !== '') {
            return { fpInTemp: fpDef, from: 'default', fpTmpDel: '' }
        }

        return { fpInTemp: '', from: 'none', fpTmpDel: '' }
    }

    //getState: 取服務環境狀態
    let getState = async () => {
        let rd = checkDocxReady()
        let nWinword = await getWinwordCount()
        return {
            platform: rd.platform,
            node: process.version,
            cwd: rd.cwd,
            exeFound: rd.exeFound,
            exePath: rd.exePath,
            docxReady: rd.ready, //僅代表環境具備條件, 實際 Word 可用性須以 /api/selftest 驗證
            templateDefault: getFpTemplateDef(),
            templates: listTemplates(),
            dirTemplates,
            dirWork,
            queue: getQueueSize(),
            winwordProcesses: nWinword,
            maxBytes,
            tokenRequired: token !== '',
        }
    }

    //convertCore: 依請求內容轉檔(模板由服務端決定後交由核心處理)
    let convertCore = async (inp) => {
        let rTemp = pickTemplate(inp)
        try {
            let keepTmp = get(inp, 'keepTmp', false)
            let keepWork = get(inp, 'keepWork', false)
            let r = await cvMdTo({
                md: get(inp, 'md', ''),
                mdBase64: get(inp, 'mdBase64', ''),
                name: get(inp, 'name', ''),
                out: get(inp, 'out', ''),
                assets: get(inp, 'assets', []),
                dirWork,
                fpInTemp: rTemp.fpInTemp,
                optMd2html: get(inp, 'optMd2html', {}),
                optHtml2docx: get(inp, 'optHtml2docx', {}),
                keepWork: (keepTmp === true || keepTmp === 'true' || keepWork === true || keepWork === 'true'),
            })
            return {
                success: true,
                ...r,
                template: rTemp.from,
            }
        }
        finally {
            //刪除本次夾帶之暫存模板
            if (rTemp.fpTmpDel !== '' && fsIsFile(rTemp.fpTmpDel)) {
                try {
                    fs.unlinkSync(rTemp.fpTmpDel)
                }
                catch (err) {
                    //殘檔被鎖, 留待下次清理
                }
            }
        }
    }

    //清除逾時殘留之作業資料夾
    let nClean = cleanWorkDir(dirWork)

    //server
    let server = Hapi.server({
        port,
        host,
        routes: {
            cors: true, //內網工具服務, 便於自瀏覽器或他機測試
            payload: {
                maxBytes,
            },
        },
    })

    //token 驗證(未設定時不啟用)
    server.ext('onRequest', (req, h) => {
        if (token === '') {
            return h.continue
        }
        if (!req.path.startsWith('/api/')) {
            return h.continue
        }
        let t = get(req, 'headers.x-api-token', '')
        if (t !== token) {
            return h.response({ success: false, error: 'unauthorized: a valid x-api-token header is required' }).code(401).takeover()
        }
        return h.continue
    })

    //apis
    server.route([

        {
            //服務說明(供瀏覽器直接開啟確認服務存活)
            method: 'GET',
            path: '/',
            handler: () => {
                return {
                    service: 'w-md2docx',
                    description: 'API service for converting Markdown to HTML/DOCX',
                    endpoints: {
                        'GET /api/health': 'health check and environment state',
                        'GET /api/selftest': 'run a minimal conversion to verify that Microsoft Word is available',
                        'POST /api/convert': 'main conversion endpoint (JSON)',
                    },
                }
            },
        },

        {
            method: 'GET',
            path: '/api/health',
            handler: async () => {
                let st = await getState()
                return {
                    success: true,
                    ...st,
                }
            },
        },

        {
            //實跑一次最小 md -> docx, 驗證本機 Word 確實可用(健檢僅檢查靜態條件, 無法確認 Word 能否被調用)
            method: 'GET',
            path: '/api/selftest',
            handler: async (req, h) => {
                let msStart = Date.now()
                try {
                    let r = await convertCore({
                        md: '# selftest\n\nThis file is generated by the self test of the conversion service.\n',
                        name: 'selftest',
                        out: 'both',
                    })
                    return {
                        success: true,
                        ms: Date.now() - msStart,
                        htmlSize: get(r, 'html.size', 0),
                        docxSize: get(r, 'docx.size', 0),
                        template: r.template,
                    }
                }
                catch (err) {
                    let st = await getState()
                    return h.response({
                        success: false,
                        error: toErrText(err),
                        ms: Date.now() - msStart,
                        env: st,
                    }).code(500)
                }
            },
        },

        {
            method: 'POST',
            path: '/api/convert',
            options: {
                payload: {
                    maxBytes,
                    parse: true,
                    allow: 'application/json',
                    timeout: false, //允許大檔慢速上傳
                },
                timeout: {
                    socket: false, //轉檔耗時較長, 不以 socket 逾時中斷
                },
            },
            handler: async (req, h) => {

                let inp = req.payload
                if (!isobj(inp)) {
                    return h.response({ success: false, error: 'parameters must be provided as a JSON object' }).code(400)
                }

                //download: 單一格式時可直接回傳二進位檔(便於 curl -o 與瀏覽器下載)
                //why: 不合法組合(out='both')於轉檔前即擋下, 避免先跑完耗時之 Word 轉檔才回 400
                let qd = get(req, 'query.download', '')
                let download = (qd === '1' || qd === 'true')
                if (download && get(inp, 'out', '') === 'both') {
                    return h.response({ success: false, error: 'download only supports out of html or docx (single format)' }).code(400)
                }

                try {

                    let r = await convertCore(inp)

                    if (download) {
                        let one = null
                        if (r.out === 'docx' && isobj(r.docx)) {
                            one = r.docx
                        }
                        else if (r.out === 'html' && isobj(r.html)) {
                            one = r.html
                        }
                        if (one === null) {
                            return h.response({ success: false, error: 'download only supports out of html or docx (single format)' }).code(400)
                        }
                        return h.response(Buffer.from(one.base64, 'base64'))
                            .type(one.mime)
                            .header('content-disposition', toContentDisposition(one.fileName))
                    }

                    return r

                }
                catch (err) {
                    let msg = toErrText(err)
                    let code = isInputError(msg) ? 400 : 500
                    return h.response({ success: false, error: msg }).code(code)
                }

            },
        },

    ])

    //autoStart
    if (autoStart) {
        await server.start()
    }

    //定期清除逾時殘留之作業資料夾
    let idInterval = setInterval(() => {
        cleanWorkDir(dirWork)
    }, 3600 * 1000)
    idInterval.unref()

    //stop
    let stop = async () => {
        clearInterval(idInterval)
        await server.stop()
    }

    return {
        server,
        getState,
        stop,
        settings: { port, host, token, maxMb, maxBytes, dirTemplates, dirWork, templateDef, nClean },
    }

}


export default rmApiServer
