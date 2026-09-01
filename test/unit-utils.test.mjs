import fs from 'fs'
import path from 'path'
import assert from 'assert'
import {
    mimeHtml,
    mimeDocx,
    toErrText,
    retryBusy,
    runExclusive,
    getQueueSize,
    getFpExe,
    checkDocxReady,
    toSafeName,
    resolveAssetPaths,
    writeAssets,
    cleanWorkDir,
    transErrAsset
} from '../src/utils.mjs'


let sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms))


describe('utils', function() {

    let fdTmp = path.resolve('./tmp/zt_utils')

    before(function() {
        fs.mkdirSync(fdTmp, { recursive: true })
    })

    after(function() {
        fs.rmSync(fdTmp, { recursive: true, force: true })
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

    describe('getFpExe/checkDocxReady', function() {

        it('getFpExe回傳字串, 存在時為實體檔', function() {
            let fp = getFpExe()
            assert.strict.equal(typeof fp, 'string')
            if (fp !== '') {
                assert.strict.equal(fs.existsSync(fp), true)
                assert.strict.equal(/htmlToDocx\.exe$/.test(fp), true)
            }
        })

        it('checkDocxReady欄位與一致性', function() {
            let r = checkDocxReady()
            assert.strict.equal(r.platform, process.platform)
            assert.strict.equal(r.isWindows, process.platform === 'win32')
            assert.strict.equal(r.cwd, path.resolve())
            assert.strict.equal(r.exePath, getFpExe())
            assert.strict.equal(r.exeFound, r.exePath !== '')
            assert.strict.equal(r.ready, r.isWindows && r.exeFound)
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

        it('資產缺漏訊息轉為可行動之英文提示並隱去工作夾路徑', function() {
            let msg = 'fp[C:/tmp/w-md2docx/job_123_abc/pics/圖.png] does not exist'
            assert.strict.equal(transErrAsset(msg), 'Failed to convert md to html: the referenced file (pics/圖.png) does not exist, please make sure it is included in assets')
        })

        it('反斜線路徑亦可解析', function() {
            let msg = 'Failed to convert md to html: fp[C:\\tmp\\job_1_a\\pics\\a.png] does not exist'
            assert.strict.equal(transErrAsset(msg), 'Failed to convert md to html: the referenced file (pics/a.png) does not exist, please make sure it is included in assets')
        })

        it('非資產缺漏訊息原樣回傳', function() {
            assert.strict.equal(transErrAsset('other error'), 'other error')
            assert.strict.equal(transErrAsset('fp[x] does not exist'), 'fp[x] does not exist') //無 job_ 片段
        })

        it('非字串轉為文字', function() {
            assert.strict.equal(transErrAsset(new Error('e1')), 'e1')
        })

    })

})
