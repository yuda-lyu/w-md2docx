import assert from 'assert'
import { normLabels, prepDocxKeep } from '../src/docxKeep.mjs'
import { prepDocxToc } from '../src/docxToc.mjs'
import { parts, buildDocx, getSectFinal, readPart, listParas, checkXml } from './tools/docxFixture.mjs'


let { run, para, heading, body, pic, capFig, capTab, capTabPlain, blank, table, cover } = parts


//keep: 合成 docx 並執行 prepDocxKeep, 回傳結果與處理後之段落
let keep = (bodyXml, opt, fixtureOpt) => {
    let u8In = buildDocx(bodyXml, fixtureOpt)
    let r = prepDocxKeep(u8In, opt)
    let doc = readPart(r.u8, 'word/document.xml')
    return { ...r, u8In, doc, ps: listParas(doc) }
}
let byText = (ps, t) => ps.find((p) => p.text === t)
let keepNextOf = (r) => r.ps.map((p) => p.keepNext)


describe('docxKeep', function() {

    describe('normLabels', function() {

        it('未給時為圖、表; 可部分覆寫; 非字串、空字串、含空白引號反斜線者該項退回預設; 兩者相同時皆退回預設', function() {
            assert.strict.deepEqual(normLabels(), { fig: '圖', tab: '表' })
            assert.strict.deepEqual(normLabels({ fig: 'Figure', tab: 'Table' }), { fig: 'Figure', tab: 'Table' })
            assert.strict.deepEqual(normLabels({ tab: 'Table' }), { fig: '圖', tab: 'Table' })
            for (let v of ['', 1, null, 'Fig ure', 'F"g', 'F\\g']) {
                assert.strict.deepEqual(normLabels({ fig: v }), { fig: '圖', tab: '表' }, JSON.stringify(v))
            }
            assert.strict.deepEqual(normLabels({ fig: 'X', tab: 'X' }), { fig: '圖', tab: '表' })
            assert.strict.deepEqual(normLabels('x'), { fig: '圖', tab: '表' })
        })

    })

    describe('prepDocxKeep', function() {

        it('圖在上圖名在下: 圖片段設與下段同頁, 圖名設段落內不分頁', function() {
            let r = keep(cover() + body('前文。') + pic() + capFig('圖1 系統架構') + body('後文。'))
            assert.strict.deepEqual(r.info, { figs: 1, tabs: 0, keepNext: 1, keepLines: 1, changed: true, warns: [] })
            let i = r.ps.findIndex((p) => p.text === '圖1 系統架構')
            assert.strict.equal(r.ps[i - 1].hasPic && r.ps[i - 1].keepNext, true)
            assert.strict.equal(r.ps[i].keepLines, true)
            assert.strict.equal(r.ps[i].keepNext, false)
            assert.strict.equal(byText(r.ps, '前文。').keepNext, false)
        })

        it('圖名在上圖在下: 圖名設與下段同頁及段落內不分頁', function() {
            let r = keep(cover() + capFig('圖1 系統架構') + pic() + body('後文。'))
            let p = byText(r.ps, '圖1 系統架構')
            assert.strict.equal(p.keepNext, true)
            assert.strict.equal(p.keepLines, true)
            assert.strict.deepEqual([r.info.keepNext, r.info.keepLines], [1, 1])
        })

        it('表名在上表格在下: 表名設與下段同頁及段落內不分頁; 已有與下段同頁者不重複添加亦不計數', function() {
            let r = keep(cover() + capTabPlain('表1 參數一覽') + table())
            assert.strict.deepEqual(r.info, { figs: 0, tabs: 1, keepNext: 1, keepLines: 1, changed: true, warns: [] })
            assert.strict.equal(byText(r.ps, '表1 參數一覽').keepNext, true)
            let r2 = keep(cover() + capTab('表1 參數一覽') + table())
            assert.strict.deepEqual([r2.info.keepNext, r2.info.keepLines], [0, 1])
            assert.strict.equal((byText(r2.ps, '表1 參數一覽').xml.match(/<w:keepNext\/>/g) || []).length, 1)
        })

        it('表名在表格下方者不處理(與下段同頁之設定只用於表名在上)', function() {
            let r = keep(cover() + table('格內') + capTabPlain('表1 參數一覽') + body('後文。'))
            assert.strict.equal(r.info.tabs, 1)
            assert.strict.equal(r.info.keepNext, 0)
            assert.strict.equal(byText(r.ps, '表1 參數一覽').keepNext, false)
        })

        it('圖名之下段為圖片時優先視為該圖之圖名: 連續之圖(圖、圖名、圖、圖名)前後串連', function() {
            let r = keep(cover() + pic() + capFig('圖1 前') + pic() + capFig('圖2 後'))
            assert.strict.deepEqual(keepNextOf(r).slice(-4), [false, true, true, false]) //「圖1 前」之下段為圖片故設於圖名; 「圖2 後」之上段為圖片
            assert.strict.deepEqual([r.info.keepNext, r.info.keepLines], [2, 2])
        })

        it('不相鄰時不設與下段同頁: 圖片與圖名之間有表格或文字段落', function() {
            let r = keep(cover() + pic() + table('格內') + capFig('圖1 系統架構'))
            assert.strict.equal(r.info.keepNext, 0)
            let r2 = keep(cover() + pic() + body('說明。') + capFig('圖1 系統架構'))
            assert.strict.equal(r2.info.keepNext, 0)
            assert.strict.equal(r2.info.keepLines, 1)
        })

        it('越過空段落: 圖片與圖名之間、表名與表格之間、圖名與其下之圖片之間的空段落一併設與下段同頁', function() {
            let r = keep(cover() + pic() + blank() + blank() + capFig('圖1 系統架構'))
            assert.strict.deepEqual(keepNextOf(r).slice(-4), [true, true, true, false])
            assert.strict.equal(r.info.keepNext, 3)
            let r2 = keep(cover() + capTabPlain('表1 參數一覽') + blank() + table())
            let i2 = r2.ps.findIndex((p) => p.text === '表1 參數一覽')
            assert.strict.deepEqual([r2.ps[i2].keepNext, r2.ps[i2 + 1].keepNext], [true, true])
            let r3 = keep(cover() + capFig('圖1 系統架構') + blank() + pic())
            let i3 = r3.ps.findIndex((p) => p.text === '圖1 系統架構')
            assert.strict.deepEqual([r3.ps[i3].keepNext, r3.ps[i3 + 1].keepNext, r3.ps[i3 + 2].keepNext], [true, true, false])
        })

        it('含分頁符號、段落前分頁或分節設定之空段落不越過', function() {
            let brk = [
                para('<w:r><w:br w:type="page"/></w:r>'),
                para('', '<w:pageBreakBefore/>'),
                para('', `<w:jc w:val="center"/>${getSectFinal()}`),
            ]
            for (let b of brk) {
                let r = keep(cover() + pic() + b + capFig('圖1 系統架構'))
                assert.strict.equal(r.info.keepNext, 0, b)
            }
        })

        it('表格內段首「表1 」者不視為表名', function() {
            let r = keep(cover() + table('表1 表格內之文字'))
            assert.strict.deepEqual(r.info, { figs: 0, tabs: 0, keepNext: 0, keepLines: 0, changed: false, warns: [] })
        })

        it('段首為標籤與編號而其後非空白者(如「圖1：名稱」)不處理, 列入warns', function() {
            let r = keep(cover() + pic() + capFig('圖1：系統架構'))
            assert.strict.equal(r.info.figs, 0)
            assert.strict.equal(r.info.changed, false)
            assert.strict.deepEqual(r.info.warns, ['1 paragraph(s) start with a figure or table label and a number that is not followed by a space, so they are not treated as captions: 圖1：系統架構'])
        })

        it('labels: 英文報告給Figure與Table', function() {
            let r = keep(cover() + pic() + capFig('Figure 1 Architecture') + capTabPlain('Table 1 Parameters') + table(), { labels: { fig: 'Figure', tab: 'Table' } })
            assert.strict.deepEqual([r.info.figs, r.info.tabs, r.info.keepNext, r.info.keepLines], [1, 1, 2, 2])
            let r2 = keep(cover() + pic() + capFig('Figure 1 Architecture'))
            assert.strict.equal(r2.info.figs, 0) //未給labels時標籤為圖、表
        })

        it('不沿用添加目錄之前提: 編號不連續、分節不只1個、無標題、已有目錄欄位皆照常處理', function() {
            let tocField = para('<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>')
            let r = keep(para(run('封面'), `<w:jc w:val="center"/>${getSectFinal()}`) + tocField + pic() + capFig('圖3 後') + pic() + capFig('圖1 前'))
            assert.strict.deepEqual([r.info.figs, r.info.keepNext, r.info.keepLines], [2, 2, 2])
        })

        it('文字方塊: 框內與含框之段落不作圖名亦不作圖片段, 其餘照常處理並列入warns', function() {
            let tb = (inner) => `<w:r><w:pict><v:shape><v:textbox><w:txbxContent>${inner}</w:txbxContent></v:textbox></v:shape></w:pict></w:r>`
            let host = para(tb(para(run('圖1 框內說明')) + para(run('圖2 框內第二段'))) + run('框後文字'))
            let r = keep(cover() + pic() + host + capFig('圖1 系統架構') + body('說明。') + pic() + capFig('圖2 流程'))
            assert.strict.deepEqual(r.info.warns, ['the document has text boxes, the paragraphs in or holding text boxes are not treated as captions or figures'])
            assert.strict.deepEqual([r.info.figs, r.info.keepNext, r.info.keepLines], [2, 1, 2]) //「圖1 系統架構」之上為含框段落, 不相鄰; 「圖2 流程」之上為圖片
            assert.strict.equal(r.doc.includes('<w:p><w:r><w:t xml:space="preserve">圖1 框內說明</w:t>'), false) //框內段落未加任何設定
            assert.strict.deepEqual(checkXml(r.doc), [])
        })

        it('無圖名表名時不改動: changed為false且回傳原內容', function() {
            let u8 = buildDocx(cover() + heading('第一章', 0) + body('內容。'))
            let r = prepDocxKeep(u8)
            assert.strict.equal(r.u8, u8)
            assert.strict.deepEqual(r.info, { figs: 0, tabs: 0, keepNext: 0, keepLines: 0, changed: false, warns: [] })
        })

        it('冪等: 處理兩次與處理一次相同', function() {
            let r1 = prepDocxKeep(buildDocx(cover() + pic() + blank() + capFig('圖1 系統架構') + capTabPlain('表1 參數') + table()))
            let r2 = prepDocxKeep(r1.u8)
            assert.strict.equal(r2.info.changed, false)
            assert.strict.equal(r2.u8, r1.u8)
        })

        it('非docx時拋錯', function() {
            assert.throws(() => prepDocxKeep(new Uint8Array(Buffer.from('not a zip'))))
        })

        it('輸出之document.xml標籤成對且子元素順序合規', function() {
            let rs = [
                keep(cover() + pic() + capFig('圖1 系統架構') + capTabPlain('表1 參數') + table()),
                keep(cover() + capFig('圖1 系統架構') + blank() + pic()),
                keep(body('x') + pic() + blank() + capFig('圖1 系統架構')),
            ]
            for (let r of rs) {
                assert.strict.deepEqual(checkXml(r.doc), [])
            }
        })

    })

    describe('與添加目錄並用', function() {

        let bodyMix = [
            heading('第一章 緒論', 0, { pb: true }),
            body('內容。'),
            pic(),
            blank(),
            capFig('圖1 系統架構'),
            capTabPlain('表1 參數一覽'),
            table(),
            heading('第二章 方法', 0, { pb: true }),
            capFig('圖2 流程'),
            pic(),
        ].join('')

        it('先處理同頁再添加目錄, 與只添加目錄之document.xml、styles.xml相同(同一規則, 已有之設定原樣回傳)', function() {
            let u8 = buildDocx(cover() + bodyMix)
            let rTocOnly = prepDocxToc(u8)
            let rBoth = prepDocxToc(prepDocxKeep(u8).u8)
            assert.strict.equal(readPart(rBoth.u8, 'word/document.xml'), readPart(rTocOnly.u8, 'word/document.xml'))
            assert.strict.equal(readPart(rBoth.u8, 'word/styles.xml'), readPart(rTocOnly.u8, 'word/styles.xml'))
        })

        it('目錄之keepWithObject為false時, 先前之同頁設定保留(任一要求即套用)', function() {
            let u8 = buildDocx(cover() + bodyMix)
            let r = prepDocxToc(prepDocxKeep(u8).u8, { keepWithObject: false })
            let ps = listParas(readPart(r.u8, 'word/document.xml'))
            let i = ps.findIndex((p) => p.text === '圖1 系統架構')
            assert.strict.deepEqual([ps[i - 2].keepNext, ps[i - 1].keepNext], [true, true])
            let r0 = prepDocxToc(u8, { keepWithObject: false })
            let ps0 = listParas(readPart(r0.u8, 'word/document.xml'))
            let i0 = ps0.findIndex((p) => p.text === '圖1 系統架構')
            assert.strict.deepEqual([ps0[i0 - 2].keepNext, ps0[i0 - 1].keepNext], [false, false])
        })

    })

})
