import assert from 'assert'
import { normLabels, isPicPara, isBrkBefore, isBrkAfter, prepDocxKeep } from '../src/docxKeep.mjs'
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

        it('連續之圖、圖名在下(圖、圖名、圖、圖名): 圖名上下皆為圖而只有上方有圖之圖名居多, 各圖片段設與下段同頁, 圖名不設', function() {
            let r = keep(cover() + pic() + capFig('圖1 前') + pic() + capFig('圖2 後'))
            assert.strict.deepEqual(keepNextOf(r).slice(-4), [true, false, true, false]) //「圖2 後」只有上方有圖(圖名在下 1 : 在上 0), 故「圖1 前」與其上之圖同頁
            assert.strict.deepEqual([r.info.keepNext, r.info.keepLines], [2, 2])
        })

        it('連續之圖、圖名在下且圖區塊間有空段落(圖、圖名、空段落、圖、圖名…): 各圖片段設與下段同頁, 圖名與其後之空段落不設', function() {
            let r = keep(cover() + [1, 2, 3].map((i) => pic() + capFig(`圖${i} 長圖`)).join(blank()) + body('後文。'))
            assert.strict.deepEqual(keepNextOf(r).slice(2), [true, false, false, true, false, false, true, false, false])
            assert.strict.deepEqual([r.info.keepNext, r.info.keepLines], [3, 3])
        })

        it('連續之圖、圖名在上(圖名、圖、圖名、圖): 只有下方有圖之圖名居多, 上下皆為圖之圖名與其下之圖同頁', function() {
            let r = keep(cover() + capFig('圖1 前') + pic() + capFig('圖2 後') + pic() + body('後文。'))
            assert.strict.deepEqual(keepNextOf(r).slice(2), [true, false, true, false, false]) //「圖1 前」只有下方有圖(圖名在上 1 : 在下 0)
            assert.strict.deepEqual([r.info.keepNext, r.info.keepLines], [2, 2])
        })

        it('圖名皆為上下皆有圖者(平手): 依圖名在圖下方之慣例, 與其上之圖同頁', function() {
            let r = keep(cover() + pic() + capFig('圖1 前') + pic() + capFig('圖2 中') + pic() + body('後文。'))
            assert.strict.deepEqual(keepNextOf(r).slice(2), [true, false, true, false, false, false])
        })

        it('圖名上下皆為圖而取向之一方之圖已歸屬其他圖名時改取另一方, 兩方皆已歸屬時不設(同一張圖不串連前後兩個圖名)', function() {
            //「圖1」「圖2」只有上方有圖(在下 2), 「圖3」只有下方有圖(在上 1); 「圖4」上下皆為圖, 多數方(上方)之圖已歸屬「圖3」, 改與其下之圖同頁
            let r = keep(cover() + pic() + capFig('圖1 甲') + body('說明。') + pic() + capFig('圖2 乙') + body('說明。') + capFig('圖3 丙') + pic() + capFig('圖4 丁') + pic() + body('後文。'))
            let ps = r.ps.slice(2)
            assert.strict.deepEqual(ps.map((p) => p.keepNext), [true, false, false, true, false, false, true, false, true, false, false])
            //「圖1」「圖4」只有下方有圖(在上 2), 「圖3」只有上方有圖(在下 1); 「圖2」上方之圖歸「圖1」、下方之圖歸「圖3」, 兩方皆已歸屬, 不設
            let r2 = keep(cover() + capFig('圖1 甲') + pic() + capFig('圖2 乙') + pic() + capFig('圖3 丙') + body('後文。') + capFig('圖4 丁') + pic() + body('後文。'))
            assert.strict.deepEqual(r2.ps.slice(2).map((p) => p.keepNext), [true, false, false, true, false, false, true, false, false])
        })

        it('取下段時依文件逆序歸屬: 圖名在上居多之文件中, 以只有上方有圖之圖名結尾之連續圖, 各圖與其下之圖名同頁', function() {
            //「圖1」「圖2」只有下方有圖(在上 2), 「圖5」只有上方有圖(在下 1), 取向為下段; 自「圖5」往前: 「圖4」下方之圖歸「圖5」改取上方, 「圖3」同理
            let r = keep(cover() + capFig('圖1 甲') + pic() + body('說明。') + capFig('圖2 乙') + pic() + body('說明。') + pic() + capFig('圖3 丙') + pic() + capFig('圖4 丁') + pic() + capFig('圖5 戊') + body('後文。'))
            assert.strict.deepEqual(r.ps.slice(2).map((p) => p.keepNext), [true, false, false, true, false, false, true, false, true, false, true, false, false])
        })

        it('表名在表格下方而表格連續(表格、表名、表格、表名): 表名不與下一個表格同頁', function() {
            //「表2」只有上方有表格(在下 1), 取向為上段; 「表1」取上方之表格, 表格在上不設, 亦不改與下一個表格同頁
            let r = keep(cover() + table('格1') + capTabPlain('表1 甲') + table('格2') + capTabPlain('表2 乙') + body('後文。'))
            assert.strict.deepEqual([r.info.tabs, r.info.keepNext, r.info.keepLines], [2, 0, 2])
        })

        it('圖在版面表格內而圖名在表格下方: 不與下一個表格同頁, 且計為圖名在下, 不使取向偏向下段', function() {
            //「圖3」只有上方有表格(在下 1); 「圖1」「圖2」上下皆為表格, 取上段即不設; 「圖4」上下皆為圖, 依在下居多取上段
            let r = keep(cover() + table('圖A') + capFig('圖1 甲') + table('圖B') + capFig('圖2 乙') + table('圖C') + capFig('圖3 丙') + body('說明。') + pic() + capFig('圖4 丁') + pic() + body('後文。'))
            assert.strict.deepEqual(r.ps.slice(2).map((p) => p.keepNext), [false, false, false, false, true, false, false, false]) //listParas 不列表格內之段落
        })

        it('水平線(md之---經Word匯入之VML橫線)與含文字之物件段落不作圖片段落', function() {
            let hr = para('<w:r><w:pict><v:rect id="_x0000_i1025" style="width:0;height:1.5pt" o:hralign="center" o:hrstd="t" o:hr="t" fillcolor="#a0a0a0" stroked="f"/></w:pict></w:r>') //2026-10-10 Word 16 匯入 <hr> 之段落
            let icon = para(run('如下') + '<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="300000" cy="300000"/><wp:docPr id="2" name="icon"/></wp:inline></w:drawing></w:r>')
            for (let before of [hr, icon]) {
                let r = keep(cover() + before + capFig('圖1 甲') + pic() + body('後文。'))
                assert.strict.deepEqual(r.ps.slice(2).map((p) => p.keepNext), [false, true, false, false]) //「圖1」只有下方有圖
            }
            let r2 = keep(cover() + pic() + capFig('圖1 甲') + hr + body('後文。'))
            assert.strict.deepEqual(r2.ps.slice(2).map((p) => p.keepNext), [true, false, false, false]) //下方為水平線, 「圖1」只有上方有圖
        })

        it('圖名與其圖之間分頁或分節時不相鄰: 圖名或圖片段落設段落前分頁、上方之圖片段落含分節設定', function() {
            let picPb = pic().replace('<w:jc w:val="center"/>', '<w:pageBreakBefore/><w:jc w:val="center"/>')
            let picSect = pic().replace('<w:jc w:val="center"/>', `<w:jc w:val="center"/>${getSectFinal()}`)
            let capPb = para(run('圖1 甲'), '<w:pageBreakBefore/><w:jc w:val="center"/>')
            for (let x of [pic() + capPb, capFig('圖1 甲') + picPb, picSect + capFig('圖1 甲')]) {
                let r = keep(cover() + x + body('後文。'))
                assert.strict.deepEqual([r.info.keepNext, r.info.keepLines], [0, 1], x)
            }
            //分頁在圖名之前而非其與圖之間者照常: 段落前分頁之圖名與其下之圖同頁
            let r2 = keep(cover() + body('前文。') + capPb + pic() + body('後文。'))
            assert.strict.deepEqual(r2.ps.slice(2).map((p) => p.keepNext), [false, true, false, false])
        })

        it('isPicPara、isBrkBefore、isBrkAfter', function() {
            let p = (xml) => ({ xml })
            assert.strict.equal(isPicPara(p(pic())), true)
            assert.strict.equal(isPicPara(p(para('<w:r><w:pict><v:rect o:hr="t"/></w:pict></w:r>'))), false)
            assert.strict.equal(isPicPara(p(para(run('如下') + '<w:r><w:drawing/></w:r>'))), false)
            assert.strict.equal(isPicPara(p(body('文字。'))), false)
            assert.strict.deepEqual([isBrkBefore(p(para(run('a'), '<w:pageBreakBefore/>'))), isBrkAfter(p(para(run('a'), '<w:pageBreakBefore/>')))], [true, false])
            assert.strict.deepEqual([isBrkBefore(p(para(run('a'), '<w:pageBreakBefore w:val="0"/>'))), isBrkBefore(p(para(run('a'), '<w:pageBreakBefore w:val="false"/>')))], [false, false])
            assert.strict.deepEqual([isBrkBefore(p(para(run('a'), getSectFinal()))), isBrkAfter(p(para(run('a'), getSectFinal())))], [false, true])
            assert.strict.deepEqual([isBrkBefore(p(para('<w:r><w:br w:type="page"/></w:r>'))), isBrkAfter(p(para('<w:r><w:br w:type="column"/></w:r>')))], [true, true])
        })

        it('表名不計入圖名之多數, 上下皆有候選時與下段之表格同頁', function() {
            //「圖1」上下皆為圖, 全文無只有一方有圖之圖名(平手), 與其上之圖同頁; 「表1」上方為圖、下方為表格, 與表格同頁
            let r = keep(cover() + pic() + capFig('圖1 前') + pic() + capTabPlain('表1 參數') + table())
            let ps = r.ps.slice(2)
            assert.strict.deepEqual(ps.map((p) => p.keepNext), [true, false, false, true]) //listParas 不列表格內之段落
            assert.strict.deepEqual([r.info.figs, r.info.tabs, r.info.keepNext], [1, 1, 2])
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
            assert.strict.deepEqual(keepNextOf(r).slice(-4), [true, false, true, false]) //「圖1 前」只有上方有圖, 「圖3 後」依此取上方之圖
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

        it('連續之圖(圖名在下): 目錄之keepWithObject與keepCaption為同一歸屬, 先處理同頁再添加目錄與只添加目錄之document.xml相同', function() {
            let u8 = buildDocx(cover() + heading('第一章 長圖', 0, { pb: true }) + body('前言。') + [1, 2, 3].map((i) => pic() + capFig(`圖${i} 長圖`)).join(blank()) + body('後文。'))
            let docToc = readPart(prepDocxToc(u8).u8, 'word/document.xml')
            let ps = listParas(docToc)
            let i = ps.findIndex((p) => p.hasPic)
            assert.strict.deepEqual(ps.slice(i, i + 9).map((p) => p.keepNext), [true, false, false, true, false, false, true, false, false]) //圖、圖名、空段落 ×3
            assert.strict.equal(readPart(prepDocxToc(prepDocxKeep(u8).u8).u8, 'word/document.xml'), docToc)
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
