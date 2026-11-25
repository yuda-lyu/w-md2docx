import fs from 'fs'
import path from 'path'
import http from 'http'
import crypto from 'crypto'
import { createRequire } from 'module'
import { spawn } from 'child_process'
import assert from 'assert'
import WHtml2docx from 'w-html2docx/src/WHtml2docx.mjs'
import { BIN_NAME, BIN_TOTAL_MS, BIN_ENSURE_TIMEOUT_MS, binBudgetMs, getBin, ensureHtmlToDocxBin, getWsemiOfHtml2docx } from './tools/binEnsure.mjs'


//轉檔器 htmlToDocx.exe 之集中取得(test/tools/binEnsure.mjs, 規則帳本 R-TEST-2)
//對應 wsemi fsEnsureBins 之〈併發〉(規則 R-BIN-ENSURE-01): 取得帶跨程序鎖, 同一檔只下載一次, 他方等候後讀快取
//下列各案不連外網: 鎖與多行程案之快取夾為本測試自建之空夾, 來源為本機鏡像(內容為已取得之轉檔器, 以其計數)
//鎖與多行程案(第4、5案)皆經 binEnsure.mjs 之 getBin(其子行程亦然), 即 before 實際使用之綁定: 該檔之 import 若改走無鎖之處(如 w-html2docx 之 dist), 此二案轉紅


//cmpVer: 比較 'a.b.c' 版本, 回傳 -1、0、1
let cmpVer = (a, b) => {
    let pa = a.split('.').map((v) => Number(v))
    let pb = b.split('.').map((v) => Number(v))
    for (let i = 0; i < 3; i++) {
        if (pa[i] !== pb[i]) {
            return pa[i] < pb[i] ? -1 : 1
        }
    }
    return 0
}


describe('binEnsure 轉檔器之集中取得與跨程序只取得一次', function() {

    //check, htmlToDocx.exe 僅供 Windows
    if (process.platform !== 'win32') {
        return
    }

    let fdTmp = path.resolve('./test/_tmp/unit-binEnsure')
    let fpChild = path.resolve('./test/tools/binEnsureChild.mjs')
    let manifest = JSON.parse(fs.readFileSync('./node_modules/w-html2docx/bin/manifest.json', 'utf8'))
    let entManifest = manifest.files.find((v) => v.name === BIN_NAME)
    let sha = entManifest.sha256
    let size = entManifest.size
    let ent = null

    //mkForeignLock: 在 fdCache 內放他方之鎖 <快取>/.lock-<鍵>/owner.json, 回傳鎖資料夾
    //  鍵為 'f/<sha256>/<name>' 轉小寫(Windows)之 SHA-256 前 32 碼(見 wsemi fsEnsureBins 之〈併發〉, 規則 R-BIN-ENSURE-01); 持有者為本行程(存活), 持有上限一小時, 故不會被接手
    let mkForeignLock = (fdCache) => {
        let key = crypto.createHash('sha256').update(`f/${sha}/${BIN_NAME}`.toLowerCase()).digest('hex').slice(0, 32)
        let fdLock = path.resolve(fdCache, `.lock-${key}`)
        fs.mkdirSync(fdLock, { recursive: true })
        let rec = { v: 1, pid: process.pid, host: process.env.COMPUTERNAME || process.env.HOSTNAME || '', at: Date.now(), limitMs: 3600000, token: `tmd-${process.pid}` }
        fs.writeFileSync(path.resolve(fdLock, 'owner.json'), JSON.stringify(rec), 'utf8')
        return fdLock
    }

    before(async function() {
        this.timeout(BIN_ENSURE_TIMEOUT_MS)
        fs.mkdirSync(fdTmp, { recursive: true })
        ent = await ensureHtmlToDocxBin()
    })

    after(function() {
        fs.rmSync(fdTmp, { recursive: true, force: true, maxRetries: 10, retryDelay: 300 })
    })

    it('取得之轉檔器: 為清單所列之檔(大小與SHA-256相符), 本行程重複呼叫不重取; 其後之隱式取得(w-html2docx 預設選項)命中快取, 不下載、不等候', async function() {
        let st = fs.statSync(ent.fp)
        assert.strict.equal(st.isFile(), true)
        assert.strict.equal(st.size, size)
        assert.strict.equal(crypto.createHash('sha256').update(fs.readFileSync(ent.fp)).digest('hex'), sha)
        assert.strict.equal(await ensureHtmlToDocxBin(), ent) //本行程記憶, 同一物件

        //隱式取得之寫法同 w-html2docx 轉檔時(只給 names)
        let events = []
        let r = await WHtml2docx.ensureBins({
            names: BIN_NAME,
            onProgress: (v) => {
                events.push(v)
            },
        })
        assert.strict.equal(r[BIN_NAME], ent.fp)
        assert.strict.equal(events.some((v) => v.step === 'download' || v.step === 'mirror' || v.step === 'wait'), false, JSON.stringify(events.map((v) => v.step)))
        assert.strict.equal(events.some((v) => v.step === 'cache' && v.hit === true), true)
    })

    it('時限對齊下載預算: 預算=n×totalMs+min(totalMs,60000)(無鏡像n=1, 有鏡像n=2), before 時限=預算+60秒且長於 mocha 預設之180秒', function() {
        //預算算法之前提: 轉檔器為單一檔(無壓縮檔 extract、無分割 parts), 否則 wsemi 之預算不是 n×totalMs+min(totalMs,60000)(wsemi 逐檔計算, 只取 names 所列者)
        assert.strict.notEqual(entManifest, undefined)
        assert.strict.equal(entManifest.extract, undefined)
        assert.strict.equal(entManifest.parts, undefined)

        assert.strict.equal(BIN_TOTAL_MS, 600000) //wsemi 之預設, 與 w-html2docx 隱式取得同值
        assert.strict.equal(binBudgetMs(false), 660000)
        assert.strict.equal(binBudgetMs(true), 1260000)
        assert.strict.equal(binBudgetMs(false, 1000), 2000)
        let mirror = typeof process.env.W_BIN_MIRROR === 'string' && process.env.W_BIN_MIRROR.trim() !== ''
        assert.strict.equal(BIN_ENSURE_TIMEOUT_MS, binBudgetMs(mirror) + 60000)
        assert.strict.equal(BIN_ENSURE_TIMEOUT_MS > 180000, true)
    })

    it('執行取得函式者為帶跨程序鎖之 wsemi: w-html2docx 之 src 解析到與本套件相同之 wsemi 且版本≥1.9.75, 本套件宣告之 wsemi 亦≥1.9.75, 產品碼與本工具皆引用 w-html2docx 之 src(不經 dist)', function() {
        let w = getWsemiOfHtml2docx()
        let fdOwn = path.dirname(createRequire(import.meta.url).resolve('wsemi/package.json'))
        assert.strict.equal(fs.realpathSync(w.fd), fs.realpathSync(fdOwn), `w-html2docx resolves wsemi from ${w.fd}, this package resolves it from ${fdOwn}`)
        assert.strict.equal(cmpVer(w.version, '1.9.75') >= 0, true, `wsemi ${w.version} has no cross-process lock (needs >= 1.9.75)`)
        let m = /(\d+\.\d+\.\d+)/.exec(JSON.parse(fs.readFileSync('./package.json', 'utf8')).dependencies.wsemi)
        assert.strict.equal(m !== null && cmpVer(m[1], '1.9.75') >= 0, true, 'package.json should declare wsemi 1.9.75 or later')

        //w-html2docx 之 dist(UMD)內嵌其建置時之 wsemi, 升級 wsemi 對經 dist 之取得無效: 產品碼與本工具之 import 皆須為 src
        //(以行首錨定之 import 陳述判定, 註解中之字樣不算; 此二處亦由案4、5以行為守住)
        for (let fp of ['./src/cvMdToDocx.mjs', './test/tools/binEnsure.mjs']) {
            assert.strict.equal(/^import\s+\S+\s+from\s+'w-html2docx\/src\/WHtml2docx\.mjs'/m.test(fs.readFileSync(fp, 'utf8')), true, `${fp} should import w-html2docx/src/WHtml2docx.mjs`)
        }
    })

    it('他方持有鎖時等候: 逾預算(此案 totalMs=1000 即 2000 毫秒)以 E_LOCK_TIMEOUT 失敗, 不下載、不誤報缺檔, 他方之鎖原樣保留', async function() {
        this.timeout(60000)
        let fdCache = path.resolve(fdTmp, 'lock-probe')
        fs.rmSync(fdCache, { recursive: true, force: true })

        let fdLock = mkForeignLock(fdCache)

        let events = []
        let t0 = Date.now()
        let err = await getBin({
            cacheDir: fdCache,
            mirror: '',
            offline: true,
            totalMs: 1000,
            onProgress: (v) => {
                events.push(v)
            },
        })
            .then(() => null, (e) => e)
        let ms = Date.now() - t0
        assert.strict.notEqual(err, null, 'should not get the file while another process holds the lock')
        assert.strict.equal(err.code, 'E_LOCK_TIMEOUT', `${err.code} ${err.message}`)
        assert.strict.equal(err.holderPid, process.pid)

        //等候時間: 下限為預算(2000), 上限寬至預算之 2 倍加 1.5 秒(輪詢間隔至多約 1.5 秒之誤差與排程延遲), 預算被放大約 3 倍以上之變異會超過上限
        let budget = binBudgetMs(false, 1000)
        assert.strict.equal(ms >= budget - 100 && ms < budget * 2 + 1500, true, `waited ${ms} ms, the budget is ${budget} ms`)
        assert.strict.equal(events.some((v) => v.step === 'wait'), true)
        assert.strict.equal(events.some((v) => v.step === 'download'), false)
        assert.strict.equal(fs.existsSync(path.resolve(fdLock, 'owner.json')), true)
    })

    it('有鏡像時預算較大: 他方持有鎖且給鏡像(空夾), 逾預算(此案 totalMs=2000 即 2×2000+2000=6000 毫秒, 無鏡像者為 4000)以 E_LOCK_TIMEOUT 失敗', async function() {
        this.timeout(60000)
        let fdCache = path.resolve(fdTmp, 'lock-probe-mirror')
        let fdMirror = path.resolve(fdTmp, 'empty-mirror')
        fs.rmSync(fdCache, { recursive: true, force: true })
        fs.rmSync(fdMirror, { recursive: true, force: true })
        fs.mkdirSync(fdMirror, { recursive: true })
        mkForeignLock(fdCache)

        let t0 = Date.now()
        let err = await getBin({
            cacheDir: fdCache,
            mirror: fdMirror,
            offline: true,
            totalMs: 2000,
        })
            .then(() => null, (e) => e)
        let ms = Date.now() - t0
        assert.strict.notEqual(err, null, 'should not get the file while another process holds the lock')
        assert.strict.equal(err.code, 'E_LOCK_TIMEOUT', `${err.code} ${err.message}`)

        //下限為有鏡像之預算(6000); 無鏡像之預算 4000 加輪詢間隔之誤差(至多約 1.5 秒)仍低於 5900, 故「有鏡像時 n=2」被放大或縮回 1 倍皆可由此分辨
        let budget = binBudgetMs(true, 2000)
        assert.strict.equal(budget, 6000)
        assert.strict.equal(ms >= budget - 100 && ms < budget * 2, true, `waited ${ms} ms, the budget with a mirror is ${budget} ms (without a mirror ${binBudgetMs(false, 2000)} ms)`)
    })

    it('3個行程同時取得(快取夾空, 來源為本機鏡像): 鏡像之轉檔器只被取走1次, 其餘等候後讀快取, 3個行程皆得到校驗相符之檔', async function() {
        this.timeout(120000)
        let fdCache = path.resolve(fdTmp, 'multi-cache')
        fs.rmSync(fdCache, { recursive: true, force: true })
        fs.mkdirSync(fdCache, { recursive: true })

        //本機鏡像: <鏡像>/<pkg>/<sha256>/<name>, 內容為已取得之轉檔器; 第1個請求之回應扣住, 至「2個行程已進入等候」(事件同步, 不用延時)或「出現第2個請求」(無鎖時)才放行, 25 秒為防懸置上限
        let bin = fs.readFileSync(ent.fp)
        let pathExe = `/${manifest.pkg}/${sha}/${BIN_NAME}`
        let nGet = 0
        let nWait = 0
        let others = []
        let release = null
        let why = ''
        let gate = new Promise((resolve) => {
            release = (v) => {
                if (why === '') {
                    why = v
                }
                resolve()
            }
        })
        let timer = setTimeout(() => release('cap'), 25000)
        let server = http.createServer((req, res) => {
            if (req.method === 'GET' && req.url === pathExe) {
                nGet++
                if (nGet >= 2) {
                    release('second-get')
                }
                gate.then(() => {
                    res.writeHead(200, { 'content-type': 'application/octet-stream', 'content-length': String(bin.length) })
                    res.end(bin)
                })
                return
            }
            others.push(`${req.method} ${req.url}`)
            res.writeHead(404)
            res.end()
        })
        await new Promise((resolve) => {
            server.listen(0, '127.0.0.1', resolve)
        })
        let mirror = `http://127.0.0.1:${server.address().port}`

        let env = { ...process.env, T_CACHE: fdCache, T_MIRROR: mirror }
        for (let k of ['W_BIN_CACHE', 'W_BIN_MIRROR', 'W_BIN_OFFLINE']) {
            delete env[k]
        }
        let procs = []
        let run = () => {
            return new Promise((resolve) => {
                let p = spawn(process.execPath, [fpChild], { env, windowsHide: true, stdio: ['ignore', 'pipe', 'pipe'] })
                procs.push(p)
                let buf = ''
                let result = null
                let errText = ''
                p.stdout.on('data', (d) => {
                    buf += String(d)
                    let ls = buf.split('\n')
                    buf = ls.pop()
                    for (let l of ls) {
                        let o = null
                        try {
                            o = JSON.parse(l)
                        }
                        catch (err) {
                            continue
                        }
                        if (o.ev === 'wait') {
                            nWait++
                            if (nWait >= 2) {
                                release('waits')
                            }
                        }
                        if (o.result) {
                            result = o.result
                        }
                    }
                })
                p.stderr.on('data', (d) => {
                    errText += String(d)
                })
                p.on('error', (err) => {
                    resolve({ code: -1, result, errText: `spawn failed: ${err}` })
                })
                p.on('close', (code) => {
                    resolve({ code, result, errText })
                })
            })
        }
        let rs = []
        try {
            rs = await Promise.all([run(), run(), run()])
        }
        finally {
            clearTimeout(timer)
            for (let p of procs) {
                if (p.exitCode === null && p.signalCode === null) {
                    p.kill() //失敗或逾時時不留子行程(免其占用 fdTmp 而使 after 之刪除失敗)
                }
            }
            await new Promise((resolve) => {
                server.close(resolve)
                if (typeof server.closeAllConnections === 'function') {
                    server.closeAllConnections()
                }
            })
        }

        let fpExpect = path.resolve(fdCache, sha, BIN_NAME)
        for (let r of rs) {
            assert.strict.equal(r.code, 0, r.errText)
            assert.strict.notEqual(r.result, null, r.errText)
            assert.strict.equal(r.result.err, null, JSON.stringify(r.result.err))
            assert.strict.equal(r.result.fp, fpExpect)
        }
        assert.strict.equal(crypto.createHash('sha256').update(fs.readFileSync(fpExpect)).digest('hex'), sha)
        assert.strict.equal(nGet, 1, `the mirror was asked ${nGet} times for ${BIN_NAME} (release by ${why}), others: ${others.join(', ')}`)
        assert.strict.equal(nWait >= 1, true, `no process waited for the lock (release by ${why}): the case did not establish contention`)
        assert.strict.equal(rs.filter((r) => r.result.steps.includes('mirror:ok')).length, 1, JSON.stringify(rs.map((r) => r.result.steps)))
    })

    it('取得失敗不被記憶: 離線且快取空時 reject, 恢復環境後同一模組再呼叫即取得(不重播舊失敗)', async function() {
        this.timeout(60000)
        let fdEmpty = path.resolve(fdTmp, 'memo-fail-cache')
        fs.mkdirSync(fdEmpty, { recursive: true })

        //另載一份模組實例(查詢字串使其不共用本檔已載入者之記憶), 以環境變數令取得確定失敗(wsemi 於呼叫時讀取 process.env)
        let m = await import(`./tools/binEnsure.mjs?memo-fail=${Date.now()}`)
        let keys = ['W_BIN_CACHE', 'W_BIN_OFFLINE', 'W_BIN_MIRROR']
        let ori = {}
        for (let k of keys) {
            ori[k] = process.env[k]
        }
        let err = null
        try {
            process.env.W_BIN_CACHE = fdEmpty
            process.env.W_BIN_OFFLINE = '1'
            delete process.env.W_BIN_MIRROR
            err = await m.ensureHtmlToDocxBin().then(() => null, (e) => e)
        }
        finally {
            for (let k of keys) {
                if (ori[k] === undefined) {
                    delete process.env[k]
                }
                else {
                    process.env[k] = ori[k]
                }
            }
        }
        assert.strict.notEqual(err, null, 'should fail while offline with an empty cache')
        assert.strict.equal(err.code, 'E_OFFLINE_MISSING', `${err.code} ${err.message}`)
        let r = await m.ensureHtmlToDocxBin()
        assert.strict.equal(r.fp, ent.fp)
    })

})
