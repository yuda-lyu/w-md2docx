import fs from 'fs'
import path from 'path'
import { createRequire } from 'module'
import WHtml2docx from 'w-html2docx/src/WHtml2docx.mjs'


//binEnsure: 測試端取得轉檔器 htmlToDocx.exe 之唯一入口(規則帳本 R-TEST-2)
//
//why(集中取得, 只下載一次): mocha --parallel 下多個 worker 若各自於轉檔時隱式取得, 快取為空(剛安裝或 w-html2docx 升版後)時各自下載(約 9 MB),
//  且下載時間落在各測試之 mocha 時限(--timeout 180000)內互搶.
//  wsemi 之 fsEnsureBins 自 1.9.75 起帶跨程序鎖(見其〈併發〉說明, 規則 R-BIN-ENSURE-01): 同一檔只下載一次, 他方等候後讀快取.
//  需要轉檔器之測試檔(含起子程序真轉檔者)於 before 呼叫 ensureHtmlToDocxBin(), 其後(含子程序)之隱式取得皆命中快取.
//
//取得函式(實際執行者): w-html2docx 之公開入口 WHtml2docx.ensureBins, 與產品碼(src/cvMdToDocx.mjs)引用同一模組 'w-html2docx/src/WHtml2docx.mjs'.
//  其 src/ensureBins.mjs 引用 'wsemi/src/fsEnsureBins.mjs', 由本套件之 node_modules/wsemi 解析(wsemi ≥ 1.9.75 帶跨程序鎖).
//  注意: w-html2docx 之 dist(UMD)內嵌其建置時之 wsemi(1.0.72 之 dist 不含跨程序鎖), 升級本套件之 wsemi 對經 dist 之取得無效;
//  本套件之產品碼與測試皆引 src, 故有效; 此一前提由 test/unit-binEnsure.test.mjs 以版本、同一份 wsemi、import 陳述、鎖逾時與多行程只取得一次確認(測試全綠不證明只下載一次).
//
//不另寫下載、快取、鏡像、校驗: 全交 w-html2docx 之 ensureBins(即 wsemi 之 fsEnsureBins); 本檔只給時限、觀察事件與一行取得紀錄.


//BIN_NAME: 轉檔器名稱(w-html2docx 之 bin/manifest.json 之 name)
let BIN_NAME = 'htmlToDocx.exe'


//BIN_TOTAL_MS: 交 opt.totalMs 之值(每次下載之整體時間上限, 毫秒), 取 wsemi 之預設 600000.
//  與 w-html2docx 轉檔時之隱式取得(只給 names, 用預設)同值: 持鎖者之持有上限依其 totalMs 推得, 等候者之 totalMs 不得小於持鎖者,
//  否則持鎖者尚未逾上限時, 等候者已先 E_LOCK_TIMEOUT
let BIN_TOTAL_MS = 600000


//binBudgetMs: wsemi 對單一檔(非壓縮檔)之取得預算(毫秒), 即鎖之持有上限與等候上限(fsEnsureBinsCore.mjs 之 lockBudget)
//  預算 = n × totalMs + min(totalMs, 60000)
//  n 為可能之下載次數: 無鏡像 1; 有鏡像(opt.mirror 或環境變數 W_BIN_MIRROR)2(先試鏡像再試來源); min(totalMs, 60000) 為校驗、更名、刪暫存之餘裕
function binBudgetMs(mirror = false, totalMs = BIN_TOTAL_MS) {
    let n = mirror ? 2 : 1
    return n * totalMs + Math.min(totalMs, 60000)
}


//BIN_ENSURE_TIMEOUT_MS: before 之 mocha 時限(毫秒) = 取得預算 + 60 秒(事件回調與一行紀錄之餘裕);
//  預設 mocha --timeout 180000 短於下載預算, 故呼叫 ensureHtmlToDocxBin() 之 before 須自給 this.timeout(BIN_ENSURE_TIMEOUT_MS);
//  隨環境有無 W_BIN_MIRROR 取有鏡像或無鏡像之預算(env 為行程啟動時之值, 本檔載入時讀取)
//  最壞情形(持鎖者近其上限才失敗, 本行程接手後再下載一次)之總時間可達約 2 倍預算, 屬極端情形, 不另加長時限
let BIN_ENSURE_TIMEOUT_MS = binBudgetMs(typeof process.env.W_BIN_MIRROR === 'string' && process.env.W_BIN_MIRROR.trim() !== '') + 60000


//本行程之取得(記憶): 同一行程內多個 describe 呼叫只取得一次; 失敗不記憶, 下次呼叫重取
let _pm = null


/**
 * 取得 htmlToDocx.exe 之單一綁定(本檔所引用之 WHtml2docx.ensureBins), 預設 totalMs 為 BIN_TOTAL_MS
 *
 * ensureHtmlToDocxBin() 與 test/unit-binEnsure.test.mjs 及其子行程(test/tools/binEnsureChild.mjs)皆經此函數取得, 故該測試走過的即本檔實際用於 before 之綁定
 *
 * @param {Object} [opt={}] 輸入設定物件, 交 WHtml2docx.ensureBins(cacheDir、mirror、offline、totalMs、onProgress、onTrace 等), names 固定為 BIN_NAME
 * @returns {Promise} 回傳 Promise, resolve 回傳 { [檔名]: 絕對路徑 }, reject 回傳帶 code 之 Error
 */
function getBin(opt = {}) {
    return WHtml2docx.ensureBins({ totalMs: BIN_TOTAL_MS, ...opt, names: BIN_NAME })
}


//ensure, 取得並回傳 { fp, events, traces, nDownload, nMirror, nWait, nTakeover, ms }
//  events 為 onProgress 之事件(step: cache、mirror、download、wait、ready 等), traces 為 onTrace 之事件(type: lock-wait、lock-takeover 等)
async function ensure() {
    let t0 = Date.now()
    let events = []
    let traces = []
    let r = await getBin({
        onProgress: (v) => {
            events.push(v)
        },
        onTrace: (v) => {
            traces.push(v)
        },
    })
    let fp = r[BIN_NAME]
    let nDownload = events.filter((v) => v.step === 'download').length
    let nMirror = events.filter((v) => v.step === 'mirror' && v.result === 'ok').length
    let nWait = events.filter((v) => v.step === 'wait').length
    let nTakeover = traces.filter((v) => v.type === 'lock-takeover').length
    let ms = Date.now() - t0
    console.log(`[binEnsure] pid=${process.pid} ${BIN_NAME}: download=${nDownload} mirror=${nMirror} wait=${nWait} takeover=${nTakeover} ${ms}ms`)
    return { fp, events, traces, nDownload, nMirror, nWait, nTakeover, ms }
}


/**
 * 取得 htmlToDocx.exe(快取命中亦校驗; 未命中自鏡像或 GitHub 取得並校驗; 跨程序只取得一次), 本行程內只執行一次
 *
 * 呼叫處須於 before 內且自給 this.timeout(BIN_ENSURE_TIMEOUT_MS)
 *
 * @returns {Promise} 回傳 Promise, resolve 回傳 { fp(轉檔器之絕對路徑), events, traces, nDownload, nMirror, nWait, nTakeover, ms }, reject 回傳 w-html2docx 之 ensureBins 之錯誤(帶 code)
 */
function ensureHtmlToDocxBin() {
    if (_pm === null) {
        _pm = ensure().catch((err) => {
            _pm = null
            throw err
        })
    }
    return _pm
}


//getWsemiOfHtml2docx: w-html2docx 之 src(實際執行取得函式者)解析到之 wsemi, 回傳 { version, fd(wsemi 之資料夾), fdWh(w-html2docx 之資料夾) }
function getWsemiOfHtml2docx() {
    let req = createRequire(import.meta.url)
    let fdWh = path.dirname(req.resolve('w-html2docx/package.json'))
    let reqWh = createRequire(path.join(fdWh, 'src', 'ensureBins.mjs'))
    let fpPkg = reqWh.resolve('wsemi/package.json')
    let version = JSON.parse(fs.readFileSync(fpPkg, 'utf8')).version
    return { version, fd: path.dirname(fpPkg), fdWh }
}


export { BIN_NAME, BIN_TOTAL_MS, BIN_ENSURE_TIMEOUT_MS, binBudgetMs, getBin, ensureHtmlToDocxBin, getWsemiOfHtml2docx }
