import fs from 'fs'
import os from 'os'
import path from 'path'
import crypto from 'crypto'
import assert from 'assert'
import addDocxToc, { getFpVbs } from '../src/addDocxToc.mjs'
import hasWord from './tools/hasWord.mjs'
import { parts, buildDocx, getSectFinal } from './tools/docxFixture.mjs'


let { run, para, heading, body, capTab, table, cover } = parts


describe('addDocxToc', function() {

    //check, Word 只存在於 Windows
    if (process.platform !== 'win32') {
        it('非Windows時reject', async function() {
            await assert.rejects(addDocxToc('./x.docx'), (e) => e === 'addDocxToc requires Microsoft Word on Windows')
        })
        return
    }

    let fdTmp = path.resolve('./test/_tmp/unit-addDocxToc')
    let sha = (fp) => crypto.createHash('sha256').update(fs.readFileSync(fp)).digest('hex')
    let nTmpToc = () => fs.readdirSync(os.tmpdir()).filter((v) => v.startsWith('wmd2docx_toc_')).length

    //bodyWord: 交 Word 開啟之正文(不含僅供結構判斷之 pic 標記): 章 2、節 1、表 2
    let bodyWord = [
        heading('第一章 緒論', 0, { pb: true }),
        body('內容。'),
        heading('1.1 背景', 1),
        capTab('表1 參數一覽'),
        table(),
        heading('第二章 方法', 0, { pb: true }),
        capTab('表2 結果'),
        table('表2 表格內之文字'),
    ].join('')

    before(function() {
        fs.mkdirSync(fdTmp, { recursive: true })
    })

    after(function() {
        fs.rmSync(fdTmp, { recursive: true, force: true, maxRetries: 10, retryDelay: 300 })
    })

    it('getFpVbs: 以src執行時取同層, 以dist執行時取../src, 皆無則回空字串', function() {
        let fpSrc = path.resolve('./src/updateDocxToc.vbs')
        assert.strict.equal(getFpVbs(), fpSrc)
        assert.strict.equal(getFpVbs(path.resolve('./src')), fpSrc)
        assert.strict.equal(getFpVbs(path.resolve('./dist')), fpSrc)
        assert.strict.equal(getFpVbs(path.resolve(fdTmp, 'nowhere')), '')
    })

    it('updateDocxToc.vbs僅含ASCII(cscript以系統字碼頁讀檔, 非ASCII字串會使編譯失敗)', function() {
        let buf = fs.readFileSync('./src/updateDocxToc.vbs')
        assert.strict.equal(buf.every((b) => b < 0x80), true)
    })

    it('輸入檢查: fpDocx非有效字串、不存在、非docx', async function() {
        await assert.rejects(addDocxToc(''), (e) => e === 'fpDocx must be a non-empty string')
        await assert.rejects(addDocxToc(null), (e) => e === 'fpDocx must be a non-empty string')
        await assert.rejects(addDocxToc('./test/nope.docx'), (e) => /^fpDocx\[.+nope\.docx\] does not exist$/.test(e))
        await assert.rejects(addDocxToc('./test/report.md'), (e) => /^Failed to prepare the TOC: /.test(e))
    })

    it('文件無標題時不處理: 僅回傳skip與ms且原檔不變', async function() {
        let fp = path.resolve(fdTmp, 'nohead.docx')
        fs.writeFileSync(fp, buildDocx(cover() + body('沒有標題')))
        let h = sha(fp)
        let r = await addDocxToc(fp)
        assert.strict.deepEqual(Object.keys(r), ['skip', 'ms'])
        assert.strict.equal(r.skip, 'no heading (paragraph with an outline level) is found, the TOC is not added')
        assert.strict.equal(sha(fp), h)
    })

    it('前提不符時reject且原檔不變', async function() {
        let fp = path.resolve(fdTmp, 'twosect.docx')
        fs.writeFileSync(fp, buildDocx(para(run('封面'), `<w:jc w:val="center"/>${getSectFinal()}`) + bodyWord))
        let h = sha(fp)
        await assert.rejects(addDocxToc(fp), (e) => e === 'Failed to prepare the TOC: the document has 2 sections, renumbering pages requires a document of 1 section (set pageNumbers to false to keep the sections)')
        assert.strict.equal(sha(fp), h)
    })

    it('Word更新: 有Word時添加目錄並通過核對; 無Word時reject之原因以系統字碼頁正確解碼、原檔不變; 暫存檔皆清除', async function() {
        this.timeout(180000)
        let fp = path.resolve(fdTmp, 'word.docx')
        fs.writeFileSync(fp, buildDocx(cover() + bodyWord))
        let h = sha(fp)
        let n0 = nTmpToc()
        let res = null
        let err = null
        await addDocxToc(fp)
            .then((r) => {
                res = r
            })
            .catch((e) => {
                err = e
            })
        assert.strict.equal(nTmpToc(), n0)
        if (hasWord()) {
            assert.strict.equal(err, null, String(err))
            assert.strict.equal(res.skip, '')
            assert.strict.equal(res.cover, true)
            assert.strict.deepEqual(res.lists, ['toc', 'tab'])
            assert.strict.equal(res.toc, 3)
            assert.strict.equal(res.tab, 2)
            assert.strict.notEqual(sha(fp), h)
        }
        else {
            assert.strict.equal(res, null)
            assert.strict.equal(typeof err, 'string')
            assert.strict.equal(err.startsWith('Word failed to update the fields: '), true, err)
            assert.strict.equal(err.includes('�'), false, err) //Word 之原因為系統字碼頁, 以 utf-8 解碼會成亂碼
            assert.strict.equal(sha(fp), h)
        }
    })

})
