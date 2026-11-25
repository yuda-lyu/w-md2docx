import fs from 'fs'
import path from 'path'
import http from 'http'
import crypto from 'crypto'
import { createRequire } from 'module'
import { spawn } from 'child_process'
import assert from 'assert'
import WHtml2docx from 'w-html2docx/src/WHtml2docx.mjs'
import fsEnsureBinsBudget from 'wsemi/src/fsEnsureBinsBudget.mjs'
import { BIN_NAME, BIN_TOTAL_MS, BIN_ENSURE_TIMEOUT_MS, binBudgetMs, getBin, ensureHtmlToDocxBin, getWsemiOfHtml2docx } from './tools/binEnsure.mjs'


//轉檔器 htmlToDocx.exe 之集中取得(test/tools/binEnsure.mjs, 規則帳本 R-TEST-2)
//對應 wsemi fsEnsureBins 之〈併發〉(規則 R-BIN-ENSURE-01): 取得帶跨程序鎖, 同一檔只下載一次, 他方等候後讀快取
//首次取得之時限取自 wsemi 之 fsEnsureBinsBudget(與鎖之持有與等候上限同一算式; 制式規格〈外部執行檔之存放下載校驗與離線〉§6 E5-3), 並行時取約兩倍; 本檔與工具皆不自寫算式
//下列各案不連外網: 鎖與多行程案之快取夾為本測試自建之空夾, 來源為本機鏡像(內容為已取得之轉檔器, 以其計數)
//鎖與多行程案(第4、5、6案)皆經 binEnsure.mjs 之 getBin(其子行程亦然), 即 before 實際使用之綁定: 該檔之 import 若改走無鎖之處(如 w-html2docx 之 dist), 此三案轉紅


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

    it('時限取自 wsemi 之 fsEnsureBinsBudget 且含並行倍數: binBudgetMs 與獨立求得者逐項相同(無鏡像、有鏡像、offline、totalMs 1000/2000/5000/預設), 舊簽章(非物件)被拒, before 時限=2×取得預算+60秒(有 W_BIN_MIRROR 時隨之加大)且長於 mocha 預設之180秒, 工具不自寫算式', async function() {
        //獨立求預算: 本案直接呼叫 fsEnsureBinsBudget, 清單給物件(工具給清單路徑), pkgRoot 為 w-html2docx 之套件根(清單所屬套件)
        let fdWh = getWsemiOfHtml2docx().fdWh
        let budgetOf = (opt = {}) => fsEnsureBinsBudget(manifest, { pkgRoot: fdWh, names: BIN_NAME, ...opt })
        let mirrorPath = path.resolve(fdTmp, 'budget-mirror') //只作為「有鏡像」之值, 預算之計算不讀取

        assert.strict.equal(BIN_TOTAL_MS, 600000) //wsemi 之預設, 與 w-html2docx 隱式取得同值
        for (let opt of [
            { mirror: '' },
            { mirror: '', offline: false },
            { mirror: mirrorPath },
            { mirror: '', totalMs: 1000 },
            { mirror: mirrorPath, totalMs: 2000 },
            { mirror: mirrorPath, offline: true, totalMs: 5000 },
        ]) {
            let b = binBudgetMs(opt)
            let e = budgetOf({ totalMs: BIN_TOTAL_MS, ...opt })
            assert.strict.equal(b, e, `${JSON.stringify(opt)}: binBudgetMs=${b}, fsEnsureBinsBudget=${e}`)
        }
        //前提: 有鏡像之預算大於無鏡像者(否則上列有、無鏡像之比對無鑑別力)
        assert.strict.equal(budgetOf({ mirror: mirrorPath }) > budgetOf({ mirror: '' }), true)

        //舊簽章 binBudgetMs(mirror布林, totalMs) 之呼叫(非物件)明確拒絕, 不靜默回預設預算(否則舊呼叫者拿到錯數而不自知)
        for (let v of [true, false, null, 1000, 'x']) {
            assert.throws(() => binBudgetMs(v), TypeError, `binBudgetMs(${JSON.stringify(v)}) should throw TypeError`)
        }

        //before 時限=並行倍數(2)×取得預算+60秒: 並行時等候者最壞先等他方一個預算, 他方失敗才自行取得再一個預算(規格 E5-3)
        let budget = budgetOf({ totalMs: BIN_TOTAL_MS })
        assert.strict.equal(BIN_ENSURE_TIMEOUT_MS, 2 * budget + 60000)
        assert.strict.equal(BIN_ENSURE_TIMEOUT_MS >= 2 * budget, true)
        assert.strict.equal(BIN_ENSURE_TIMEOUT_MS > 180000, true)

        //環境變數 W_BIN_MIRROR(行程啟動時之值, 工具載入時取預算): 另載兩份工具(查詢字串使其不共用已載入者), 無與有 W_BIN_MIRROR 之時限各等於獨立求得者, 且有者較大
        let ori = process.env.W_BIN_MIRROR
        let tNo = null
        let tMirror = null
        try {
            delete process.env.W_BIN_MIRROR
            tNo = (await import(`./tools/binEnsure.mjs?budget-no=${Date.now()}`)).BIN_ENSURE_TIMEOUT_MS
            process.env.W_BIN_MIRROR = mirrorPath
            tMirror = (await import(`./tools/binEnsure.mjs?budget-mirror=${Date.now()}`)).BIN_ENSURE_TIMEOUT_MS
        }
        finally {
            if (ori === undefined) {
                delete process.env.W_BIN_MIRROR
            }
            else {
                process.env.W_BIN_MIRROR = ori
            }
        }
        assert.strict.equal(tNo, 2 * budgetOf({ mirror: '', totalMs: BIN_TOTAL_MS }) + 60000)
        assert.strict.equal(tMirror, 2 * budgetOf({ mirror: mirrorPath, totalMs: BIN_TOTAL_MS }) + 60000)
        assert.strict.equal(tMirror > tNo, true)

        //工具不自寫算式: 取得預算只經 wsemi 之 fsEnsureBinsBudget(行首錨定之 import 陳述, 註解不算), 且無 Math.min 之算式(舊算式 n×totalMs+min(totalMs,60000))
        let src = fs.readFileSync('./test/tools/binEnsure.mjs', 'utf8')
        assert.strict.equal(/^import\s+fsEnsureBinsBudget\s+from\s+'wsemi\/src\/fsEnsureBinsBudget\.mjs'/m.test(src), true, 'binEnsure.mjs should import wsemi/src/fsEnsureBinsBudget.mjs')
        assert.strict.equal(/Math\.min\s*\(/.test(src), false, 'binEnsure.mjs should not compute the budget itself')
    })

    it('執行取得函式者為帶跨程序鎖之 wsemi: w-html2docx 之 src 解析到與本套件相同之 wsemi 且版本≥1.9.76(1.9.75 起帶鎖, 1.9.76 起有 fsEnsureBinsBudget), 本套件宣告之 wsemi 亦≥1.9.76, 產品碼與本工具皆引用 w-html2docx 之 src(不經 dist)', function() {
        let w = getWsemiOfHtml2docx()
        let fdOwn = path.dirname(createRequire(import.meta.url).resolve('wsemi/package.json'))
        assert.strict.equal(fs.realpathSync(w.fd), fs.realpathSync(fdOwn), `w-html2docx resolves wsemi from ${w.fd}, this package resolves it from ${fdOwn}`)
        assert.strict.equal(cmpVer(w.version, '1.9.76') >= 0, true, `wsemi ${w.version} has no cross-process lock or no fsEnsureBinsBudget (needs >= 1.9.76)`)
        let m = /(\d+\.\d+\.\d+)/.exec(JSON.parse(fs.readFileSync('./package.json', 'utf8')).dependencies.wsemi)
        assert.strict.equal(m !== null && cmpVer(m[1], '1.9.76') >= 0, true, 'package.json should declare wsemi 1.9.76 or later')

        //w-html2docx 之 dist(UMD)內嵌其建置時之 wsemi(1.0.73 起之 dist 已帶鎖, 更早者不帶): 本套件仍經 src, 使取得所用之 wsemi 為本套件所裝者; 產品碼與本工具之 import 皆須為 src
        //(以行首錨定之 import 陳述判定, 註解中之字樣不算; 此二處亦由案4、5、6以行為守住)
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
        let opt = { cacheDir: fdCache, mirror: '', offline: true, totalMs: 1000 } //取得與求預算共用同一組選項
        let t0 = Date.now()
        let err = await getBin({
            ...opt,
            onProgress: (v) => {
                events.push(v)
            },
        })
            .then(() => null, (e) => e)
        let ms = Date.now() - t0
        assert.strict.notEqual(err, null, 'should not get the file while another process holds the lock')
        assert.strict.equal(err.code, 'E_LOCK_TIMEOUT', `${err.code} ${err.message}`)
        assert.strict.equal(err.holderPid, process.pid)

        //等候時間: 下限為預算(此案 2000), 上限寬至預算之 2 倍加 1.5 秒(輪詢間隔至多約 1.5 秒之誤差與排程延遲), 預算被放大約 3 倍以上之變異會超過上限
        //預算取自 fsEnsureBinsBudget(經 binBudgetMs, 與上之取得同一組選項): 它與鎖之等候上限同一算式, 兩者不一致即此處轉紅(規格 E5-3)
        let budget = binBudgetMs(opt)
        assert.strict.equal(ms >= budget - 100 && ms < budget * 2 + 1500, true, `waited ${ms} ms, the budget is ${budget} ms`)
        assert.strict.equal(events.some((v) => v.step === 'wait'), true)
        assert.strict.equal(events.some((v) => v.step === 'download'), false)
        assert.strict.equal(fs.existsSync(path.resolve(fdLock, 'owner.json')), true)
    })

    it('有鏡像時預算較大: 他方持有鎖且給鏡像(空夾), 逾預算(取自 fsEnsureBinsBudget, 此案 totalMs=2000 之有鏡像者大於無鏡像者)以 E_LOCK_TIMEOUT 失敗', async function() {
        this.timeout(60000)
        let fdCache = path.resolve(fdTmp, 'lock-probe-mirror')
        let fdMirror = path.resolve(fdTmp, 'empty-mirror')
        fs.rmSync(fdCache, { recursive: true, force: true })
        fs.rmSync(fdMirror, { recursive: true, force: true })
        fs.mkdirSync(fdMirror, { recursive: true })
        mkForeignLock(fdCache)

        let opt = { cacheDir: fdCache, mirror: fdMirror, offline: true, totalMs: 2000 } //取得與求預算共用同一組選項
        let t0 = Date.now()
        let err = await getBin(opt)
            .then(() => null, (e) => e)
        let ms = Date.now() - t0
        assert.strict.notEqual(err, null, 'should not get the file while another process holds the lock')
        assert.strict.equal(err.code, 'E_LOCK_TIMEOUT', `${err.code} ${err.message}`)

        //下限為有鏡像之預算; 前提(使本案有鑑別力): 無鏡像之預算加輪詢間隔之誤差(至多約 1.5 秒)仍低於有鏡像預算之下限, 故「有鏡像時預算較大」被縮回無鏡像者可由此分辨
        let budget = binBudgetMs(opt)
        let budgetNoMirror = binBudgetMs({ ...opt, mirror: '' })
        assert.strict.equal(budgetNoMirror + 1500 < budget - 100, true, `budget with a mirror ${budget} ms, without a mirror ${budgetNoMirror} ms: the case cannot tell them apart`)
        assert.strict.equal(ms >= budget - 100 && ms < budget * 2, true, `waited ${ms} ms, the budget with a mirror is ${budget} ms (without a mirror ${budgetNoMirror} ms)`)
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
        let lockLimit = null //第1個請求到達時(持有者此時必已持鎖)讀鎖之 owner.json 之 limitMs, 即實際之持有上限
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
                if (nGet === 1) {
                    try {
                        for (let n of fs.readdirSync(fdCache)) {
                            if (n.startsWith('.lock-')) {
                                lockLimit = JSON.parse(fs.readFileSync(path.resolve(fdCache, n, 'owner.json'), 'utf8')).limitMs
                            }
                        }
                    }
                    catch (err) {
                        lockLimit = `read failed: ${err}`
                    }
                }
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

        //實際之持鎖上限(子行程以 getBin 之預設選項取得, 有鏡像)等於 binBudgetMs 對同一組選項之值: 時限取自預算函式而預算與鎖之上限同一算式(規格 E5-3), 亦守住 getBin 預設之 totalMs 與 BIN_TOTAL_MS 同步
        let budget = binBudgetMs({ cacheDir: fdCache, mirror, offline: false })
        assert.strict.equal(lockLimit, budget, `the holder recorded limitMs=${lockLimit}, binBudgetMs=${budget}`)
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
