import fs from 'fs'
import path from 'path'
import { constants } from 'buffer'
import assert from 'assert'
import {
    mimeHtml,
    mimeDocx,
    toErrText,
    retryBusy,
    runExclusive,
    getQueueSize,
    checkDocxReady,
    assertReadable,
    assertB64Fit,
    toSafeName,
    resolveAssetPaths,
    writeAssets,
    cleanWorkDir,
    transErrAsset,
    normReqOpt,
    normTocReq,
    normKeepReq
} from '../src/utils.mjs'


let sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms))


describe('utils', function() {

    let fdTmp = path.resolve('./test/_tmp/unit-utils')

    before(function() {
        fs.mkdirSync(fdTmp, { recursive: true })
    })

    after(function() {
        fs.rmSync(fdTmp, { recursive: true, force: true, maxRetries: 10, retryDelay: 300 })
    })

    describe('mime', function() {

        it('mimeHtml/mimeDocx', function() {
            assert.strict.equal(mimeHtml, 'text/html; charset=utf-8')
            assert.strict.equal(mimeDocx, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document')
        })

    })

    describe('toErrText', function() {

        it('字串原樣回傳', function() {
            assert.strict.equal(toErrText('abc'), 'abc')
            assert.strict.equal(toErrText(''), '')
        })

        it('Error取message', function() {
            assert.strict.equal(toErrText(new Error('boom')), 'boom')
        })

        it('物件無message則JSON化', function() {
            assert.strict.equal(toErrText({ code: 'E1' }), '{"code":"E1"}')
        })

        it('null/undefined轉為文字', function() {
            assert.strict.equal(toErrText(null), 'null')
            assert.strict.equal(toErrText(undefined), 'undefined')
        })

    })

    describe('retryBusy', function() {

        it('fun非函數則reject', async function() {
            await assert.rejects(retryBusy('x'), (e) => e === 'fun must be a function')
        })

        it('EBUSY/EPERM重試至成功', async function() {
            let n = 0
            let r = await retryBusy(async () => {
                n++
                if (n <= 2) {
                    let e = new Error('locked')
                    e.code = (n === 1) ? 'EBUSY' : 'EPERM'
                    throw e
                }
                return 'ok'
            }, { delayMs: 5 })
            assert.strict.equal(r, 'ok')
            assert.strict.equal(n, 3)
        })

        it('非busy錯誤不重試, 直接拋出', async function() {
            let n = 0
            await assert.rejects(retryBusy(async () => {
                n++
                throw new Error('other')
            }, { delayMs: 5 }), /other/)
            assert.strict.equal(n, 1)
        })

        it('字串reject(底層轉檔器慣例)不重試, 原樣拋出', async function() {
            let n = 0
            await assert.rejects(retryBusy(() => {
                n++
                return Promise.reject('fpIn[x] does not exist')
            }, { delayMs: 5 }), (e) => e === 'fpIn[x] does not exist')
            assert.strict.equal(n, 1)
        })

        it('tries耗盡則拋出最後錯誤', async function() {
            let n = 0
            await assert.rejects(retryBusy(async () => {
                n++
                let e = new Error('locked')
                e.code = 'EBUSY'
                throw e
            }, { tries: 3, delayMs: 5 }), /locked/)
            assert.strict.equal(n, 3)
        })

        it('tries/delayMs非法值採預設(tries=6)', async function() {
            let n = 0
            await assert.rejects(retryBusy(async () => {
                n++
                let e = new Error('locked')
                e.code = 'EBUSY'
                throw e
            }, { tries: -1, delayMs: 'x' }), /locked/)
            assert.strict.equal(n, 6)
        })

    })

    describe('runExclusive/getQueueSize', function() {

        it('fun非函數則reject', async function() {
            await assert.rejects(runExclusive(null), (e) => e === 'fun must be a function')
        })

        it('依序執行, 佇列數隨進出增減', async function() {
            let order = []
            assert.strict.equal(getQueueSize(), 0)
            let p1 = runExclusive(async () => {
                await sleep(60)
                order.push(1)
                return 'a'
            })
            let p2 = runExclusive(async () => {
                order.push(2)
                return 'b'
            })
            assert.strict.equal(getQueueSize(), 2)
            let rs = await Promise.all([p1, p2])
            await sleep(5) //佇列減量於pm.then之後, 給予一個tick
            assert.strict.deepEqual(rs, ['a', 'b'])
            assert.strict.deepEqual(order, [1, 2]) //後進者須等前者完成
            assert.strict.equal(getQueueSize(), 0)
        })

        it('前一件失敗不阻斷後續', async function() {
            let p1 = runExclusive(() => Promise.reject('fail1'))
            let p2 = runExclusive(async () => 'ok2')
            await assert.rejects(p1, (e) => e === 'fail1')
            assert.strict.equal(await p2, 'ok2')
            await sleep(5)
            assert.strict.equal(getQueueSize(), 0)
        })

    })

    describe('checkDocxReady', function() {

        //manifest: w-html2docx 的外部執行檔清單, 快取位置為 <W_BIN_CACHE>/w-html2docx/<sha256>/htmlToDocx.exe
        let manifest = JSON.parse(fs.readFileSync('./node_modules/w-html2docx/bin/manifest.json', 'utf8'))
        let sha = manifest.files[0].sha256
        let fnExe = manifest.files[0].name

        //withEnv: 暫時設定環境變數(值 undefined 為刪除), 結束後還原
        let withEnv = async (kv, fun) => {
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
            try {
                return await fun()
            }
            finally {
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

        //fpExeReal: 本機預設快取內已備妥之轉檔器(實跑過 Word 轉檔或 ensureBins 者才有), 無則為空字串
        let fpExeReal = path.resolve('./node_modules/w-html2docx/bin-cache', sha, fnExe)
        if (!fs.existsSync(fpExeReal)) {
            fpExeReal = ''
        }

        it('欄位與一致性: exePath與exeFound同進退, ready僅以作業系統為Windows判定', async function() {
            let r = await checkDocxReady()
            assert.strict.equal(r.platform, process.platform)
            assert.strict.equal(r.isWindows, process.platform === 'win32')
            assert.strict.equal(r.cwd, path.resolve())
            assert.strict.equal(r.exeFound, r.exePath !== '')
            if (r.exeFound) {
                assert.strict.equal(fs.existsSync(r.exePath), true)
                assert.strict.equal(path.basename(r.exePath), fnExe)
            }
            assert.strict.equal(r.ready, r.isWindows) //ready不以轉檔器存在為條件(缺檔於首次轉檔時由w-html2docx取得)
        })

        it('快取空: exeFound為false且exePath為空字串, 但ready不因此為false; 不下載(快取夾仍空)', async function() {
            let fdCache = path.resolve(fdTmp, 'cache-empty')
            fs.mkdirSync(fdCache, { recursive: true })
            let r = await withEnv({ W_BIN_CACHE: fdCache, W_BIN_OFFLINE: undefined, W_BIN_MIRROR: undefined }, () => checkDocxReady())
            assert.strict.equal(r.exeFound, false)
            assert.strict.equal(r.exePath, '')
            assert.strict.equal(r.ready, process.platform === 'win32')
            assert.strict.deepEqual(fs.readdirSync(fdCache), [])
        })

        it('快取內有檔但內容與清單之SHA-256不符: 視為未備妥(命中亦校驗)', async function() {
            let fdCache = path.resolve(fdTmp, 'cache-corrupt')
            fs.mkdirSync(path.resolve(fdCache, 'w-html2docx', sha), { recursive: true })
            fs.writeFileSync(path.resolve(fdCache, 'w-html2docx', sha, fnExe), 'not an executable')
            let r = await withEnv({ W_BIN_CACHE: fdCache, W_BIN_OFFLINE: undefined, W_BIN_MIRROR: undefined }, () => checkDocxReady())
            assert.strict.equal(r.exeFound, false)
            assert.strict.equal(r.exePath, '')
        })

        it('快取內有校驗相符之轉檔器: exeFound為true且exePath為該快取內之檔', async function() {
            if (fpExeReal === '') {
                console.log('[unit-utils] htmlToDocx.exe is not in the default cache of w-html2docx on this machine, skip')
                this.skip()
            }
            let fdCache = path.resolve(fdTmp, 'cache-hit')
            let fpHit = path.resolve(fdCache, 'w-html2docx', sha, fnExe)
            fs.mkdirSync(path.dirname(fpHit), { recursive: true })
            fs.copyFileSync(fpExeReal, fpHit)
            let r = await withEnv({ W_BIN_CACHE: fdCache, W_BIN_OFFLINE: undefined, W_BIN_MIRROR: undefined }, () => checkDocxReady())
            assert.strict.equal(r.exeFound, true)
            assert.strict.equal(r.exePath, fpHit)
        })

        it('只看快取: 環境變數之鏡像內有檔也不取用(不複製入快取)', async function() {
            if (fpExeReal === '') {
                console.log('[unit-utils] htmlToDocx.exe is not in the default cache of w-html2docx on this machine, skip')
                this.skip()
            }
            let fdCache = path.resolve(fdTmp, 'cache-mirror')
            let fdMirror = path.resolve(fdTmp, 'mirror')
            let fpMirror = path.resolve(fdMirror, 'w-html2docx', sha, fnExe) //鏡像結構: <鏡像>/<pkg>/<sha256>/<name>
            fs.mkdirSync(fdCache, { recursive: true })
            fs.mkdirSync(path.dirname(fpMirror), { recursive: true })
            fs.copyFileSync(fpExeReal, fpMirror)
            let r = await withEnv({ W_BIN_CACHE: fdCache, W_BIN_MIRROR: fdMirror, W_BIN_OFFLINE: undefined }, () => checkDocxReady())
            assert.strict.equal(r.exeFound, false)
            assert.strict.equal(r.exePath, '')
            assert.strict.deepEqual(fs.readdirSync(fdCache), [])
        })

    })

    describe('assertReadable', function() {

        it('單檔大小以fs.readFileSync上限(2 GiB - 1位元組)為界: 以內不拋錯, 超過則於讀檔前拋出指出原因之錯誤', function() {
            let fp = path.resolve(fdTmp, 'sparse.docx')
            let mk = (n) => {
                let h = fs.openSync(fp, 'w')
                fs.ftruncateSync(h, n) //稀疏檔, 不實寫
                fs.closeSync(h)
            }
            mk(0)
            assertReadable(fp, 'docx')
            mk(2 ** 31 - 1)
            assertReadable(fp, 'docx')
            mk(2 ** 31)
            assert.throws(() => assertReadable(fp, 'docx'), (e) => {
                assert.strict.equal(e instanceof Error, true)
                assert.strict.equal(e.message, 'docx is too large to be read into memory in one piece (2147483648 bytes, over the limit of 2147483647 bytes = 2 GiB)')
                return true
            })
            fs.unlinkSync(fp)
        })

    })

    describe('assertB64Fit', function() {

        //maxChars: 回傳內容以 base64 置於單一 JSON 字串, 受 V8 字串長度上限限制, 另預留 1 MiB 給 JSON 外殼與其他欄位
        let maxChars = constants.MAX_STRING_LENGTH - 1024 * 1024
        let maxBytes = Math.floor(maxChars / 4) * 3 //base64 每 3 位元組成 4 字元

        it('各檔 base64 字元數(含進位)合計以上限為界: 以內不拋錯', function() {
            assertB64Fit({ docx: 0, html: 0 })
            assertB64Fit({ docx: maxBytes })
            assertB64Fit({ docx: maxBytes - 3, html: 3 })
        })

        it('超過上限時於轉 base64 前拋出指出原因之錯誤(列出各檔位元組數與上限)', function() {
            assert.throws(() => assertB64Fit({ docx: maxBytes + 1 }), (e) => {
                assert.strict.equal(e instanceof Error, true)
                assert.strict.equal(e.message, `the result is too large to be returned as base64 in one JSON string: docx ${maxBytes + 1} bytes (${4 * Math.ceil((maxBytes + 1) / 3)} chars of base64 in total), over the limit of ${maxChars} chars (about ${maxBytes} bytes of files in total)`)
                return true
            })
        })

        it('docx與html合計超過上限(各自未超過)亦拋錯, 訊息列出兩者', function() {
            let half = Math.floor(maxBytes / 2)
            assertB64Fit({ docx: half })
            assert.throws(() => assertB64Fit({ docx: half + 3, html: half + 3 }), (e) => {
                assert.strict.equal(e.message.startsWith(`the result is too large to be returned as base64 in one JSON string: docx ${half + 3} bytes, html ${half + 3} bytes (`), true, e.message)
                return true
            })
        })

    })

    describe('toSafeName', function() {

        it('去除路徑成分與副檔名', function() {
            assert.strict.equal(toSafeName('../report/報告R00.01.md'), '報告R00.01')
            assert.strict.equal(toSafeName('a\\b\\c.docx'), 'c')
            assert.strict.equal(toSafeName('x.HTML'), 'x')
        })

        it('非法字元替換為底線', function() {
            //note: 冒號不置於第2字元, 否則 Windows 之 path.basename 視 'a:' 為磁碟機前綴(路徑成分)而剝除
            assert.strict.equal(toSafeName('a*b?c"d<e>f|g:h'), 'a_b_c_d_e_f_g_h')
            assert.strict.equal(toSafeName('a\tb\rc\nd'), 'a_b_c_d')
        })

        it('去除開頭點號, 空值回output', function() {
            assert.strict.equal(toSafeName('..hidden'), 'hidden')
            assert.strict.equal(toSafeName(''), 'output')
            assert.strict.equal(toSafeName('   '), 'output')
            assert.strict.equal(toSafeName(null), 'output')
            assert.strict.equal(toSafeName(123), 'output')
            assert.strict.equal(toSafeName('...'), 'output')
        })

        it('長度截至120', function() {
            assert.strict.equal(toSafeName('x'.repeat(200)).length, 120)
        })

    })

    describe('resolveAssetPaths', function() {

        it('一般相對路徑回傳單一絕對路徑', function() {
            let rs = resolveAssetPaths(fdTmp, 'pics/a.png')
            assert.strict.deepEqual(rs, [path.resolve(fdTmp, 'pics/a.png')])
        })

        it('反斜線正規化', function() {
            let rs = resolveAssetPaths(fdTmp, 'pics\\a.png')
            assert.strict.deepEqual(rs, [path.resolve(fdTmp, 'pics/a.png')])
        })

        it('URL編碼路徑回傳原字串與解碼後兩者', function() {
            let rs = resolveAssetPaths(fdTmp, 'pics/%E5%9C%96.png')
            assert.strict.deepEqual(rs, [path.resolve(fdTmp, 'pics/%E5%9C%96.png'), path.resolve(fdTmp, 'pics/圖.png')])
        })

        it('非法percent序列視為未編碼', function() {
            let rs = resolveAssetPaths(fdTmp, 'pics/100%.png')
            assert.strict.deepEqual(rs, [path.resolve(fdTmp, 'pics/100%.png')])
        })

        it('空路徑/非字串拋錯', function() {
            assert.throws(() => resolveAssetPaths(fdTmp, ''), /^Error: asset\.path is empty$/)
            assert.throws(() => resolveAssetPaths(fdTmp, null), /^Error: asset\.path is empty$/)
        })

        it('絕對路徑(磁碟機/根/UNC)拋錯', function() {
            assert.throws(() => resolveAssetPaths(fdTmp, 'C:/x.png'), /^Error: asset\.path must be a relative path: C:\/x\.png$/)
            assert.throws(() => resolveAssetPaths(fdTmp, '/x.png'), /must be a relative path/)
            assert.throws(() => resolveAssetPaths(fdTmp, '//srv/x.png'), /must be a relative path/)
        })

        it('路徑穿越(含編碼形式)拋錯', function() {
            assert.throws(() => resolveAssetPaths(fdTmp, '../x.png'), /^Error: asset\.path is outside the working folder: \.\.\/x\.png$/)
            assert.throws(() => resolveAssetPaths(fdTmp, 'a/../../x.png'), /is outside the working folder/)
            assert.throws(() => resolveAssetPaths(fdTmp, '%2e%2e/x.png'), /is outside the working folder/)
        })

        it('同前綴兄弟資料夾不放行', function() {
            let dirJob = path.resolve(fdTmp, 'job_1')
            assert.throws(() => resolveAssetPaths(dirJob, '../job_1x/x.png'), /is outside the working folder/)
        })

    })

    describe('writeAssets', function() {

        it('非陣列回0', function() {
            assert.strict.equal(writeAssets(fdTmp, null), 0)
            assert.strict.equal(writeAssets(fdTmp, 'x'), 0)
        })

        it('寫入檔案(含子資料夾), 編碼路徵同時寫出兩檔, 容許data URI與contentBase64', function() {
            let dirJob = path.resolve(fdTmp, 'job_w')
            let n = writeAssets(dirJob, [
                { path: 'pics/%E5%9C%96.png', base64: Buffer.from('AB').toString('base64') },
                { path: 'b.txt', base64: `data:text/plain;base64,${Buffer.from('CD').toString('base64')}` },
                { path: 'c.txt', contentBase64: Buffer.from('EF').toString('base64') },
            ])
            assert.strict.equal(n, 3)
            assert.strict.equal(fs.readFileSync(path.resolve(dirJob, 'pics/%E5%9C%96.png'), 'utf8'), 'AB')
            assert.strict.equal(fs.readFileSync(path.resolve(dirJob, 'pics/圖.png'), 'utf8'), 'AB')
            assert.strict.equal(fs.readFileSync(path.resolve(dirJob, 'b.txt'), 'utf8'), 'CD')
            assert.strict.equal(fs.readFileSync(path.resolve(dirJob, 'c.txt'), 'utf8'), 'EF')
        })

        it('base64為空拋錯', function() {
            assert.throws(() => writeAssets(fdTmp, [{ path: 'x.png' }]), /^Error: asset\.base64 is empty: x\.png$/)
            assert.throws(() => writeAssets(fdTmp, [{ path: 'x.png', base64: 123 }]), /asset\.base64 is empty/)
        })

        it('路徑穿越拋錯且不寫檔', function() {
            assert.throws(() => writeAssets(fdTmp, [{ path: '../zt_evil.txt', base64: 'QUI=' }]), /is outside the working folder/)
            assert.strict.equal(fs.existsSync(path.resolve(fdTmp, '../zt_evil.txt')), false)
        })

    })

    describe('cleanWorkDir', function() {

        it('非有效字串回0', function() {
            assert.strict.equal(cleanWorkDir(''), 0)
            assert.strict.equal(cleanWorkDir(null), 0)
        })

        it('資料夾不存在則建立並回0', function() {
            let d = path.resolve(fdTmp, 'work_new')
            assert.strict.equal(fs.existsSync(d), false)
            assert.strict.equal(cleanWorkDir(d), 0)
            assert.strict.equal(fs.existsSync(d), true)
        })

        it('僅刪逾時之job_夾, 非job_與未逾時者保留', function() {
            let d = path.resolve(fdTmp, 'work_clean')
            let dOld = path.resolve(d, 'job_old')
            let dNew = path.resolve(d, 'job_new')
            let dOther = path.resolve(d, 'other_old')
            fs.mkdirSync(dOld, { recursive: true })
            fs.mkdirSync(dNew, { recursive: true })
            fs.mkdirSync(dOther, { recursive: true })
            let tOld = (Date.now() - 2 * 3600 * 1000) / 1000
            fs.utimesSync(dOld, tOld, tOld)
            fs.utimesSync(dOther, tOld, tOld)
            let n = cleanWorkDir(d, { msAge: 3600 * 1000 })
            assert.strict.equal(n, 1)
            assert.strict.equal(fs.existsSync(dOld), false)
            assert.strict.equal(fs.existsSync(dNew), true)
            assert.strict.equal(fs.existsSync(dOther), true)
        })

        it('msAge非法值採預設6小時', function() {
            let d = path.resolve(fdTmp, 'work_def')
            let d5h = path.resolve(d, 'job_5h')
            let d7h = path.resolve(d, 'job_7h')
            fs.mkdirSync(d5h, { recursive: true })
            fs.mkdirSync(d7h, { recursive: true })
            let t5 = (Date.now() - 5 * 3600 * 1000) / 1000
            let t7 = (Date.now() - 7 * 3600 * 1000) / 1000
            fs.utimesSync(d5h, t5, t5)
            fs.utimesSync(d7h, t7, t7)
            assert.strict.equal(cleanWorkDir(d, { msAge: 'x' }), 1)
            assert.strict.equal(fs.existsSync(d5h), true)
            assert.strict.equal(fs.existsSync(d7h), false)
        })

    })

    describe('transErrAsset', function() {

        //w-md2html(≥1.1.12)缺圖之 reject 訊息: `N image file(s) do not exist:` 後每行一個絕對路徑
        it('缺1個圖: 轉為可行動之英文提示並隱去工作夾路徑(正斜線路徑)', function() {
            let msg = 'Failed to convert md to html: 1 image file(s) do not exist:\nC:/tmp/w-md2docx/job_123_abc/pics/圖.png'
            assert.strict.equal(transErrAsset(msg), 'Failed to convert md to html: the referenced file (pics/圖.png) does not exist, please make sure it is included in assets')
        })

        it('缺1個圖: 反斜線路徑與無前綴之訊息亦可解析', function() {
            let msg = '1 image file(s) do not exist:\r\nC:\\tmp\\job_1_a\\pics\\a.png'
            assert.strict.equal(transErrAsset(msg), 'Failed to convert md to html: the referenced file (pics/a.png) does not exist, please make sure it is included in assets')
        })

        it('缺多個圖: 全數列出(依訊息順序), 複數形', function() {
            let msg = 'Failed to convert md to html: 2 image file(s) do not exist:\nC:\\tmp\\job_1_a\\pics\\a.png\nC:\\tmp\\job_1_a\\b.png\n'
            assert.strict.equal(transErrAsset(msg), 'Failed to convert md to html: the referenced files (pics/a.png, b.png) do not exist, please make sure they are included in assets')
        })

        it('缺圖之路徑不全在工作夾內(如md內寫絕對路徑)則原樣回傳', function() {
            let msg = 'Failed to convert md to html: 2 image file(s) do not exist:\nC:\\tmp\\job_1_a\\a.png\nD:\\other\\b.png'
            assert.strict.equal(transErrAsset(msg), msg)
            let msg2 = 'Failed to convert md to html: 1 image file(s) do not exist:\nD:\\other\\b.png'
            assert.strict.equal(transErrAsset(msg2), msg2)
        })

        it('Word解析不到之圖(unresolved linked picture): 來源在工作夾內者隱去工作夾路徑, 網址與工作夾外之路徑原樣', function() {
            let msg = 'Failed to convert html to docx: code[1]:\r\nunresolved linked picture: C:\\tmp\\w-md2docx\\job_1_a\\pics\\圖 a.png\r\nunresolved linked picture: http://127.0.0.1:1/job_9/x.png\r\nunresolved linked picture: D:\\other\\b.png\r\nunresolved linked picture: C:/tmp/job_2_b/c.png\r\nerror: LinkedPictureError: 4 linked picture(s) can not be embedded'
            let exp = 'Failed to convert html to docx: code[1]:\r\nunresolved linked picture: pics/圖 a.png\r\nunresolved linked picture: http://127.0.0.1:1/job_9/x.png\r\nunresolved linked picture: D:\\other\\b.png\r\nunresolved linked picture: c.png\r\nerror: LinkedPictureError: 4 linked picture(s) can not be embedded'
            assert.strict.equal(transErrAsset(msg), exp)
        })

        it('非資產缺漏訊息原樣回傳', function() {
            assert.strict.equal(transErrAsset('other error'), 'other error')
            assert.strict.equal(transErrAsset('fpInMd[x] does not exist'), 'fpInMd[x] does not exist')
            assert.strict.equal(transErrAsset('Failed to convert md to html: images are too large to embed as base64 (imgConvertToBase64=true): 3 image file(s)'), 'Failed to convert md to html: images are too large to embed as base64 (imgConvertToBase64=true): 3 image file(s)')
        })

        it('非字串轉為文字', function() {
            assert.strict.equal(transErrAsset(new Error('e1')), 'e1')
        })

    })

    describe('normReqOpt/normKeepReq', function() {

        it('normReqOpt: true與字串true回true; 物件回移除omit之鍵後之複本且不改原物件; 其餘回false', function() {
            assert.strict.equal(normReqOpt(true), true)
            assert.strict.equal(normReqOpt('true'), true)
            let o = { a: 1, b: 2 }
            assert.strict.deepEqual(normReqOpt(o, ['b']), { a: 1 })
            assert.strict.deepEqual(o, { a: 1, b: 2 })
            assert.strict.deepEqual(normReqOpt(o, 'b'), { a: 1, b: 2 }) //omit非陣列視為無
            for (let v of [false, undefined, null, 'false', '', 1, []]) {
                assert.strict.equal(normReqOpt(v), false, JSON.stringify(v))
            }
        })

        it('normKeepReq: true、字串true與設定物件(複本)原樣通過, 其餘回false', function() {
            assert.strict.equal(normKeepReq(true), true)
            assert.strict.equal(normKeepReq('true'), true)
            let o = { labels: { fig: 'Figure', tab: 'Table' } }
            assert.strict.deepEqual(normKeepReq(o), o)
            assert.strict.notEqual(normKeepReq(o), o)
            assert.strict.equal(normKeepReq('x'), false)
        })

    })

    describe('normTocReq', function() {

        it('true與字串true回傳true', function() {
            assert.strict.equal(normTocReq(true), true)
            assert.strict.equal(normTocReq('true'), true)
        })

        it('設定物件回傳移除timeoutMs之複本(逾時由服務端決定), 不改動請求物件', function() {
            let o = { maxLevels: 2, labels: { fig: 'Figure' }, timeoutMs: 1 }
            assert.strict.deepEqual(normTocReq(o), { maxLevels: 2, labels: { fig: 'Figure' } })
            assert.strict.deepEqual(o, { maxLevels: 2, labels: { fig: 'Figure' }, timeoutMs: 1 })
        })

        it('其餘值回傳false', function() {
            for (let v of [false, undefined, null, 'false', '', 1, []]) {
                assert.strict.equal(normTocReq(v), false, JSON.stringify(v))
            }
        })

    })

})
