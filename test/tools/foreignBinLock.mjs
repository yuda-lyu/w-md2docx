import fs from 'fs'
import path from 'path'
import crypto from 'crypto'
import WHtml2docx from 'w-html2docx/src/WHtml2docx.mjs'
import { BIN_NAME, getWsemiOfHtml2docx } from './binEnsure.mjs'


//foreignBinLock: 以替身模擬「他行程正取得轉檔器而持有跨程序鎖」(規則帳本 R-DIAG-2、R-TEST-2)
//
//why(替身持鎖): 測試要重現「他行程持有取得鎖」而不真的下載(9 MB, 慢且依賴網路): 於測試自建之快取夾內放一把他方之鎖.
//  鎖之位置與 owner.json 之格式為 wsemi 之 fsEnsureBins 〈併發〉(規則 R-BIN-ENSURE-01 規則 19)所載之公開規格:
//  <快取>/.lock-<鍵>/owner.json, 鍵為 'f/<sha256>/<name>'(Windows 轉小寫)之 SHA-256 前 32 碼; 持有者為本行程(存活)、持有上限一小時, 故不會被接手.
//  只用於測試自建之快取夾; 產品碼不得推算鎖位置(R-DIAG-2), 不得對真實快取放替身鎖.


/**
 * 於快取夾內放他方之鎖(替身), 回傳鎖資料夾
 *
 * 釋放: 呼叫 rmForeignLock(僅限本函式所建者); 等候中之取得於其輪詢間隔(至多約 1.5 秒)內察覺鎖已不存在
 *
 * @param {String} fdCacheRoot 輸入快取夾(其下為 <sha256>/<name> 與鎖); 用 opt.cacheDir 者即該值, 用環境變數 W_BIN_CACHE 者為 <W_BIN_CACHE>/w-html2docx
 * @returns {String} 回傳鎖資料夾之絕對路徑
 */
function mkForeignLock(fdCacheRoot) {
    let fdWh = getWsemiOfHtml2docx().fdWh
    let manifest = JSON.parse(fs.readFileSync(path.resolve(fdWh, 'bin', 'manifest.json'), 'utf8'))
    let ent = manifest.files.find((v) => v.name === BIN_NAME)
    let id = `f/${ent.sha256}/${BIN_NAME}`
    if (process.platform === 'win32') {
        id = id.toLowerCase()
    }
    let key = crypto.createHash('sha256').update(id).digest('hex').slice(0, 32)
    let fdLock = path.resolve(fdCacheRoot, `.lock-${key}`)
    fs.mkdirSync(fdLock, { recursive: true })
    let rec = { v: 1, pid: process.pid, host: process.env.COMPUTERNAME || process.env.HOSTNAME || '', at: Date.now(), limitMs: 3600000, token: `tmd-${process.pid}` }
    fs.writeFileSync(path.resolve(fdLock, 'owner.json'), JSON.stringify(rec), 'utf8')
    return fdLock
}


/**
 * 移除 mkForeignLock 所放之替身鎖(已不存在者不報錯)
 *
 * @param {String} fdLock 輸入 mkForeignLock 回傳之鎖資料夾
 */
function rmForeignLock(fdLock) {
    fs.rmSync(fdLock, { recursive: true, force: true, maxRetries: 10, retryDelay: 100 })
}


/**
 * 在「他行程持有轉檔器取得鎖」之情境下執行 fun, 結束時必還原
 *
 * 做的事: 建快取夾 <fdCache>/w-html2docx 並放替身鎖; 環境變數 W_BIN_CACHE 指向 fdCache(並暫刪 W_BIN_OFFLINE、W_BIN_MIRROR);
 *   以記錄用之轉呼叫包覆 WHtml2docx.ensureBins(轉呼叫真實函式, 不換成替身; 產品碼與本檔引用同一模組), 記下每次呼叫之選項與回傳之 Promise;
 *   防懸置上限 guardMs: 逾時由本檔移除替身鎖(使「不該等候而等候」之被測物落定, 不使測試與行程懸置), 並令回傳之 guarded 為 true, 呼叫端須斷言 guarded 為 false.
 * 結束時: 移除替身鎖、等記錄之取得全部落定、還原 ensureBins 與環境變數.
 *
 * @param {String} fdCache 輸入測試自建之快取根夾(其下建 w-html2docx 夾)
 * @param {Function} fun 輸入被執行之 async 函數, 收到 { fdCache, fdPkg, fdLock, calls, release, settle }:
 *   fdPkg 為實際之快取夾(<fdCache>/w-html2docx); calls 為 [{ opt, pm }](本情境內 WHtml2docx.ensureBins 之呼叫); release() 移除替身鎖(可重複呼叫); settle() 等 calls 內之 pm 全部落定
 * @param {Number} [guardMs=8000] 輸入防懸置上限毫秒數
 * @returns {Promise} 回傳 Promise, resolve 回傳 { value(fun 之回傳值), guarded(是否由防懸置上限移除了替身鎖) }
 */
async function withForeignBinLock(fdCache, fun, guardMs = 8000) {
    let fdPkg = path.resolve(fdCache, 'w-html2docx')
    fs.mkdirSync(fdPkg, { recursive: true })
    let fdLock = mkForeignLock(fdPkg)

    let calls = []
    let orgEnsure = WHtml2docx.ensureBins
    WHtml2docx.ensureBins = (...args) => {
        let pm = orgEnsure(...args)
        calls.push({ opt: args[0], pm })
        return pm
    }

    let kv = { W_BIN_CACHE: fdCache, W_BIN_OFFLINE: undefined, W_BIN_MIRROR: undefined }
    let ori = {}
    for (let k of Object.keys(kv)) {
        ori[k] = process.env[k]
        if (kv[k] === undefined) {
            delete process.env[k]
        }
        else {
            process.env[k] = kv[k]
        }
    }

    let guarded = false
    let guard = setTimeout(() => {
        guarded = true
        rmForeignLock(fdLock)
    }, guardMs)

    let release = () => {
        rmForeignLock(fdLock)
    }
    let settle = async () => {
        await Promise.allSettled(calls.map((v) => v.pm))
        //產品碼之後續處理(清除進行中之探測)為數個微任務之接續, 登記先於本處; setImmediate 為巨任務邊界, 到達時微任務佇列已清空(同步點, 非延時)
        await new Promise((resolve) => {
            setImmediate(resolve)
        })
    }

    try {
        let value = await fun({ fdCache, fdPkg, fdLock, calls, release, settle })
        return { value, guarded }
    }
    finally {
        clearTimeout(guard)
        release()
        await settle()
        WHtml2docx.ensureBins = orgEnsure
        for (let k of Object.keys(ori)) {
            if (ori[k] === undefined) {
                delete process.env[k]
            }
            else {
                process.env[k] = ori[k]
            }
        }
    }
}


export { mkForeignLock, rmForeignLock, withForeignBinLock }
