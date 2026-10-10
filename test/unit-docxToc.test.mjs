import fs from 'fs'
import assert from 'assert'
import { normTocOpt, prepDocxToc, checkDocxToc } from '../src/docxToc.mjs'
import { parts, buildDocx, getSectFinal, readPart, listParas, checkXml, simulateWordToc } from './tools/docxFixture.mjs'


let { run, para, heading, body, pic, capFig, capTab, blank, table, cover } = parts
let fpTplWh = './node_modules/w-html2docx/src/tmp.docx' //w-html2docx 內建模板: 無 TPC 樣式與目錄樣式
let escRe = (s) => String(s).replace(/[.*+?^${}()|[\]\\]/g, '\\$&')


//bodyStd: 正文(章、節、小節 3 層與第 4 層標題, 圖 2、表 2), 第 1 個標題與第二章開頭帶分頁符號
let bodyStd = [
    heading('第一章 緒論', 0, { pb: true }),
    body('本章說明計畫緣起。'),
    heading('1.1 背景', 1),
    body('背景說明。'),
    pic(),
    capFig('圖1 系統架構'),
    capTab('表1 參數一覽'),
    table(),
    heading('1.1.1 細節', 2),
    heading('1.1.1.1 更細之標題', 3),
    body('細節說明。'),
    heading('第二章 方法', 0, { pb: true }),
    pic(),
    capFig('圖2 流程'),
    capTab('表2 結果'),
    table('表2 表格內之文字'),
].join('')


//prep: 合成 docx 並執行 prepDocxToc, 回傳結果與解析後之段落、欄位指令、分節設定
let prep = (bodyXml, opt, fixtureOpt) => {
    let u8In = buildDocx(bodyXml, fixtureOpt)
    let r = prepDocxToc(u8In, opt)
    if (r.info.skip) {
        return { ...r, u8In }
    }
    let doc = readPart(r.u8, 'word/document.xml')
    let sty = readPart(r.u8, 'word/styles.xml')
    let instr = (doc.match(/<w:instrText[^>]*>[^<]*<\/w:instrText>/g) || []).map((v) => v.replace(/<[^>]+>/g, '').replace(/&quot;/g, '"').trim())
    return { ...r, u8In, doc, sty, ps: listParas(doc), instr, tocs: instr.filter((v) => v.startsWith('TOC')), sects: doc.match(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/g) || [] }
}
let byText = (ps, t) => ps.find((p) => p.text === t)
let styleOf = (sty, name) => (sty.match(new RegExp(`<w:style\\b[^>]*>(?:(?!</w:style>)[\\s\\S])*?<w:name w:val="${escRe(name)}"/>[\\s\\S]*?</w:style>`)) || [''])[0]
let titlesOf = (ps) => ps.filter((p) => ['目錄', '圖目錄', '表目錄'].includes(p.text))


describe('docxToc', function() {

    describe('normTocOpt', function() {

        it('未給、true或非物件時採預設', function() {
            for (let v of [undefined, true, 'x', null]) {
                let o = normTocOpt(v)
                assert.strict.deepEqual(o, {
                    labels: { fig: '圖', tab: '表' },
                    titles: { toc: '目錄', fig: '圖目錄', tab: '表目錄' },
                    frontHeadings: ['摘要', 'ABSTRACT', 'Abstract'],
                    maxLevels: 3,
                    titleStyle: 'TPC11報告目錄(標題)',
                    titleStyleGiven: false,
                    headStyles: ['章標題', '節標題', '小節標題'],
                    pageNumbers: true,
                    keepLines: true,
                    keepWithObject: true,
                    replaceTocStyles: true,
                })
            }
        })

        it('labels: 可部分覆寫; 空字串、非字串、含空白引號反斜線者退回預設; 兩者相同時皆退回預設', function() {
            assert.strict.deepEqual(normTocOpt({ labels: { fig: 'Figure', tab: 'Table' } }).labels, { fig: 'Figure', tab: 'Table' })
            assert.strict.deepEqual(normTocOpt({ labels: { fig: 'Figure' } }).labels, { fig: 'Figure', tab: '表' })
            for (let v of ['', 123, null, 'Fig ure', 'Fi"g', 'F\\g']) {
                assert.strict.deepEqual(normTocOpt({ labels: { fig: v } }).labels, { fig: '圖', tab: '表' }, JSON.stringify(v))
            }
            assert.strict.deepEqual(normTocOpt({ labels: { fig: 'X', tab: 'X' } }).labels, { fig: '圖', tab: '表' })
        })

        it('titles: 可部分覆寫, 非有效字串退回預設', function() {
            assert.strict.deepEqual(normTocOpt({ titles: { toc: 'Contents', fig: '', tab: null } }).titles, { toc: 'Contents', fig: '圖目錄', tab: '表目錄' })
        })

        it('frontHeadings 僅保留非空字串, 非陣列退回預設; maxLevels 為正整數且上限9', function() {
            assert.strict.deepEqual(normTocOpt({ frontHeadings: ['Summary', 1, ''] }).frontHeadings, ['Summary'])
            assert.strict.deepEqual(normTocOpt({ frontHeadings: 'Abstract' }).frontHeadings, ['摘要', 'ABSTRACT', 'Abstract'])
            let kp = [[2, 2], ['4', 4], [20, 9], [0, 3], [-1, 3], [2.5, 3], ['x', 3]]
            for (let [v, e] of kp) {
                assert.strict.equal(normTocOpt({ maxLevels: v }).maxLevels, e, JSON.stringify(v))
            }
        })

        it('titleStyle 指定時記錄titleStyleGiven; headStyles 須為非空字串組成之非空陣列; 布林設定非布林值退回預設', function() {
            let o = normTocOpt({ titleStyle: '目錄標題', headStyles: ['A', 'B'], pageNumbers: false, keepLines: 'false', replaceTocStyles: false })
            assert.strict.equal(o.titleStyle, '目錄標題')
            assert.strict.equal(o.titleStyleGiven, true)
            assert.strict.deepEqual(o.headStyles, ['A', 'B'])
            assert.strict.equal(o.pageNumbers, false)
            assert.strict.equal(o.keepLines, true)
            assert.strict.equal(o.replaceTocStyles, false)
            assert.strict.deepEqual(normTocOpt({ headStyles: [] }).headStyles, ['章標題', '節標題', '小節標題'])
            assert.strict.deepEqual(normTocOpt({ headStyles: ['A', 1] }).headStyles, ['章標題', '節標題', '小節標題'])
        })

    })

    describe('prepDocxToc', function() {

        let rStd = prep(cover() + bodyStd)

        it('有封面: 分3節(封面無頁尾與頁碼、目錄節沿用原頁尾且大寫羅馬數字自I起、正文節不指定頁尾而自1起), 封面最後一段帶封面節設定', function() {
            assert.strict.equal(rStd.info.skip, '')
            assert.strict.equal(rStd.info.cover, true)
            assert.strict.equal(rStd.sects.length, 3)
            let [s1, s2, s3] = rStd.sects
            assert.strict.equal(/<w:footerReference\b|<w:headerReference\b|<w:pgNumType\b/.test(s1), false)
            assert.strict.equal(s2.includes('<w:footerReference w:type="default" r:id="rId8"/>'), true)
            assert.strict.equal(s2.includes('<w:pgNumType w:fmt="upperRoman" w:start="1"/>'), true)
            assert.strict.equal(/<w:footerReference\b|<w:headerReference\b/.test(s3), false)
            assert.strict.equal(/<w:pgNumType\b[^>]*w:start="1"/.test(s3), true)
            assert.strict.equal(byText(rStd.ps, '2026年10月').sectPr, s1)
        })

        it('目錄欄位: 層級依正文標題與maxLevels(第4層不列), 有圖名表名才列圖目錄表目錄, 無前置章節不加書籤', function() {
            assert.strict.deepEqual(rStd.tocs, ['TOC \\o "1-3" \\h \\z \\u', 'TOC \\h \\z \\c "圖"', 'TOC \\h \\z \\c "表"'])
            assert.strict.deepEqual(rStd.info.levels, [1, 3])
            assert.strict.deepEqual(rStd.info.lists, ['toc', 'fig', 'tab'])
            assert.strict.deepEqual(rStd.info.expect, {
                toc: ['第一章 緒論', '1.1 背景', '1.1.1 細節', '第二章 方法'],
                fig: ['圖1 系統架構', '圖2 流程'],
                tab: ['表1 參數一覽', '表2 結果'],
            })
            assert.strict.equal(rStd.info.bookmark, '')
            assert.strict.equal(rStd.doc.includes('_WMdTocBody'), false)
            let r2 = prep(cover() + bodyStd, { maxLevels: 2 })
            assert.strict.equal(r2.tocs[0], 'TOC \\o "1-2" \\h \\z \\u')
            assert.strict.deepEqual(r2.info.expect.toc, ['第一章 緒論', '1.1 背景', '第二章 方法'])
        })

        it('圖名表名之編號改為SEQ欄位且段落文字不變; 表格內之「表1」不改', function() {
            assert.strict.deepEqual(rStd.instr.filter((v) => v.startsWith('SEQ')), ['SEQ 圖 \\* ARABIC', 'SEQ 表 \\* ARABIC', 'SEQ 圖 \\* ARABIC', 'SEQ 表 \\* ARABIC'])
            for (let t of ['圖1 系統架構', '圖2 流程', '表1 參數一覽', '表2 結果']) {
                assert.strict.equal(byText(rStd.ps, t).nSeq, 1, t)
            }
            assert.strict.equal(rStd.doc.includes('<w:t xml:space="preserve">表1 表格內之文字</w:t>'), true)
        })

        it('列入目錄之標題改用承接樣式且文字段不帶直接粗體; 標題與圖名表名設段落內不分頁; 未列入之第4層標題不改', function() {
            let kp = [['第一章 緒論', 'WMdTocHead1'], ['1.1 背景', 'WMdTocHead2'], ['1.1.1 細節', 'WMdTocHead3'], ['第二章 方法', 'WMdTocHead1']]
            for (let [t, id] of kp) {
                let p = byText(rStd.ps, t)
                assert.strict.equal(p.pStyle, id, t)
                assert.strict.equal(p.directBold, false, t)
                assert.strict.equal(p.keepLines, true, t)
            }
            for (let t of ['圖1 系統架構', '圖2 流程', '表1 參數一覽', '表2 結果']) {
                assert.strict.equal(byText(rStd.ps, t).keepLines, true, t)
            }
            let p4 = byText(rStd.ps, '1.1.1.1 更細之標題')
            assert.strict.equal(p4.pStyle, '')
            assert.strict.equal(p4.directBold, true)
            assert.strict.equal(p4.keepLines, false)
            let s1 = styleOf(rStd.sty, '章標題')
            assert.strict.equal(s1.includes('w:styleId="WMdTocHead1"'), true)
            assert.strict.equal(s1.includes('<w:b/>'), true)
        })

        it('圖片段落設與下段(圖名)同頁; 表名維持與表格同頁', function() {
            for (let t of ['圖1 系統架構', '圖2 流程']) {
                let i = rStd.ps.findIndex((p) => p.text === t)
                assert.strict.equal(rStd.ps[i - 1].hasPic, true)
                assert.strict.equal(rStd.ps[i - 1].keepNext, true, t)
            }
            assert.strict.equal(byText(rStd.ps, '表1 參數一覽').keepNext, true)
        })

        it('圖片與圖名之間夾空段落時, 空段落一併設與下段同頁(與keepCaption共用之規則)', function() {
            let r = prep(cover() + bodyStd.replace(pic() + capFig('圖1 系統架構'), pic() + blank() + capFig('圖1 系統架構')))
            let i = r.ps.findIndex((p) => p.text === '圖1 系統架構')
            assert.strict.deepEqual([r.ps[i - 2].hasPic, r.ps[i - 2].keepNext, r.ps[i - 1].text, r.ps[i - 1].keepNext], [true, true, '', true])
        })

        it('正文第1個標題移除開頭之分頁符號(新節自新頁開始), 其後各章保留', function() {
            assert.strict.equal(byText(rStd.ps, '第一章 緒論').pageBreakRun, false)
            assert.strict.equal(byText(rStd.ps, '第二章 方法').pageBreakRun, true)
        })

        it('目錄標題使用模板之樣式; 分節時第1個目錄標題不加段前分頁, 圖目錄表目錄加; 無警告', function() {
            let ts = titlesOf(rStd.ps)
            assert.strict.deepEqual(ts.map((p) => p.pStyle), ['TPC11', 'TPC11', 'TPC11'])
            assert.strict.deepEqual(ts.map((p) => p.pageBreakBefore), [false, true, true])
            assert.strict.deepEqual(rStd.info.warns, [])
        })

        it('目錄項目樣式取代模板既有之toc與table of figures(第1層粗體、右縮排、懸掛縮排), 字型沿用正文第1個標題', function() {
            let toc1 = styleOf(rStd.sty, 'toc 1')
            assert.strict.equal(toc1.includes('<w:rFonts w:ascii="Times New Roman" w:eastAsia="標楷體" w:hAnsi="Times New Roman"/><w:b/>'), true)
            assert.strict.equal(toc1.includes('<w:ind w:left="0" w:right="480"/>'), true)
            assert.strict.equal(toc1.includes('w:autoRedefine'), false) //模板原有者已被取代
            assert.strict.equal(styleOf(rStd.sty, 'toc 2').includes('<w:ind w:left="960" w:right="480" w:hanging="480"/>'), true)
            assert.strict.equal(styleOf(rStd.sty, 'table of figures').includes('<w:ind w:left="720" w:right="480" w:hanging="720"/>'), true)
            assert.strict.equal((rStd.sty.match(/<w:name w:val="toc 1"\/>/g) || []).length, 1)
        })

        it('無封面(第1段即標題, 或其前僅空段落): 不建封面節而分2節, 目錄節為第1節', function() {
            for (let pre of ['', para('')]) {
                let r = prep(pre + bodyStd)
                assert.strict.equal(r.info.cover, false)
                assert.strict.equal(r.sects.length, 2)
                assert.strict.equal(r.sects[0].includes('<w:pgNumType w:fmt="upperRoman" w:start="1"/>'), true)
                assert.strict.equal(r.sects[0].includes('<w:footerReference '), true)
                assert.strict.equal(/<w:footerReference\b/.test(r.sects[1]), false)
                let iTitle = r.ps.findIndex((p) => p.text === '目錄')
                assert.strict.equal(r.ps.slice(0, iTitle).every((p) => p.sectPr === ''), true)
                assert.strict.equal(r.ps[iTitle].pageBreakBefore, false)
                assert.strict.equal(byText(r.ps, '第一章 緒論').pageBreakRun, false)
            }
        })

        it('第1個標題前有表格即視為有封面', function() {
            let r = prep(table('封面表格') + bodyStd)
            assert.strict.equal(r.info.cover, true)
            assert.strict.equal(r.sects.length, 3)
        })

        it('單一#文件標題: 文件標題列為目錄第1層且正文自其起算(已知限制, 欲置於封面須以非標題段落撰寫)', function() {
            let r = prep(para(run('測試範例')) + heading('計畫報告', 0) + heading('一、摘要', 1) + body('摘要內容。') + heading('二、背景與目的', 1) + heading('2.1 背景說明', 2) + body('內容。'))
            assert.strict.equal(r.info.cover, true)
            assert.strict.deepEqual(r.info.expect.toc, ['計畫報告', '一、摘要', '二、背景與目的', '2.1 背景說明'])
        })

        it('前置章節: 排於封面與目錄之間且不列入目錄, 三目錄以書籤限定正文, 前置章節內之圖名照樣編號但不列入', function() {
            let r = prep(cover() + heading('摘要', 0, { pb: true }) + body('摘要內容。') + pic() + capFig('圖1 摘要附圖') + bodyStd.replace('圖1 系統架構', '圖2 系統架構').replace('圖2 流程', '圖3 流程'))
            assert.strict.deepEqual(r.info.front, ['摘要'])
            assert.strict.equal(r.info.bookmark, '_WMdTocBody')
            assert.strict.deepEqual(r.tocs, ['TOC \\o "1-3" \\h \\z \\u \\b _WMdTocBody', 'TOC \\h \\z \\c "圖" \\b _WMdTocBody', 'TOC \\h \\z \\c "表" \\b _WMdTocBody'])
            assert.strict.deepEqual(r.info.expect.fig, ['圖2 系統架構', '圖3 流程'])
            assert.strict.equal(r.info.expect.toc.includes('摘要'), false)
            assert.strict.equal(byText(r.ps, '圖1 摘要附圖').nSeq, 1)
            assert.strict.equal(byText(r.ps, '目錄').pageBreakBefore, true)
            assert.strict.equal(byText(r.ps, '摘要').pageBreakRun, false)
            assert.strict.equal(byText(r.ps, '第一章 緒論').pageBreakRun, false)
            let iStart = r.doc.indexOf('<w:bookmarkStart w:id="0" w:name="_WMdTocBody"/>')
            assert.strict.equal(iStart > r.doc.lastIndexOf('Update field to build the table'), true)
            assert.strict.equal(iStart < r.doc.indexOf('第一章 緒論'), true)
            assert.strict.equal(/<w:bookmarkEnd w:id="0"\/><w:sectPr\b[^>]*>(?:(?!<\/w:sectPr>)[\s\S])*<\/w:sectPr>\s*<\/w:body>/.test(r.doc), true)
            assert.strict.equal(r.sects.length, 3)
        })

        it('pageNumbers:false: 不動分節與頁尾(原有2節亦可), 有封面時三目錄標題皆段前分頁, 無封面時第1個不加, 第1個標題保留分頁符號', function() {
            let r = prep(cover() + bodyStd, { pageNumbers: false })
            assert.strict.deepEqual(r.sects, [getSectFinal()])
            assert.strict.deepEqual(titlesOf(r.ps).map((p) => p.pageBreakBefore), [true, true, true])
            assert.strict.equal(byText(r.ps, '第一章 緒論').pageBreakRun, true)
            let r2 = prep(bodyStd, { pageNumbers: false })
            assert.strict.deepEqual(titlesOf(r2.ps).map((p) => p.pageBreakBefore), [false, true, true])
            let r3 = prep(para(run('封面'), `<w:jc w:val="center"/>${getSectFinal()}`) + bodyStd, { pageNumbers: false })
            assert.strict.equal(r3.sects.length, 2)
        })

        it('keepLines:false與keepWithObject:false: 不加段落內不分頁與圖片同頁設定', function() {
            let r = prep(cover() + bodyStd, { keepLines: false, keepWithObject: false })
            assert.strict.equal(r.ps.some((p) => p.keepLines), false)
            let i = r.ps.findIndex((p) => p.text === '圖1 系統架構')
            assert.strict.equal(r.ps[i - 1].keepNext, false)
            assert.strict.equal(byText(r.ps, '表1 參數一覽').keepNext, true) //原有之設定不動
        })

        it('圖名之編號跨文字段時亦改為SEQ欄位且文字不變', function() {
            let r = prep(cover() + bodyStd.replace(run('圖1 系統架構'), run('圖') + run('1') + run(' 系統架構')))
            assert.strict.equal(byText(r.ps, '圖1 系統架構').nSeq, 1)
            assert.strict.deepEqual(r.info.expect.fig, ['圖1 系統架構', '圖2 流程'])
        })

        it('段首為標籤與編號但其後非空白者(如「圖1：名稱」)不視為圖名, 以warns提示筆數與前3例', function() {
            let r = prep(cover() + bodyStd.replace('圖1 系統架構', '圖1：系統架構').replace('圖2 流程', '圖2：流程'))
            assert.strict.deepEqual(r.info.lists, ['toc', 'tab'])
            assert.strict.equal(r.instr.some((v) => v.startsWith('SEQ 圖')), false)
            assert.strict.deepEqual(r.info.warns, ['2 paragraph(s) start with a figure or table label and a number that is not followed by a space, so they are not treated as captions: 圖1：系統架構; 圖2：流程'])
        })

        it('英文報告: labels與titles改英文', function() {
            let b = bodyStd.replace('圖1 系統架構', 'Figure 1 Architecture').replace('圖2 流程', 'Figure 2 Flow').replace('表1 參數一覽', 'Table 1 Parameters').replace('表2 結果', 'Table 2 Results')
            let r = prep(cover() + b, { labels: { fig: 'Figure', tab: 'Table' }, titles: { toc: 'Contents', fig: 'List of Figures', tab: 'List of Tables' } })
            assert.strict.deepEqual(r.tocs, ['TOC \\o "1-3" \\h \\z \\u', 'TOC \\h \\z \\c "Figure"', 'TOC \\h \\z \\c "Table"'])
            assert.strict.deepEqual(r.instr.filter((v) => v.startsWith('SEQ')), ['SEQ Figure \\* ARABIC', 'SEQ Table \\* ARABIC', 'SEQ Figure \\* ARABIC', 'SEQ Table \\* ARABIC'])
            assert.strict.equal(['Contents', 'List of Figures', 'List of Tables'].every((t) => byText(r.ps, t)), true)
        })

        it('標題已有段落樣式者不改其粗體並列入warns', function() {
            let r = prep(cover() + bodyStd.replace(heading('1.1 背景', 1), heading('1.1 背景', 1, { pStyle: 'a' })))
            assert.strict.deepEqual(r.info.warns, ['the heading already has a paragraph style, its TOC entry may carry its direct formatting: 1.1 背景'])
            let p = byText(r.ps, '1.1 背景')
            assert.strict.equal(p.pStyle, 'a')
            assert.strict.equal(p.directBold, true)
        })

        it('replaceTocStyles:false: 沿用模板既有之目錄樣式, 僅補建缺少者', function() {
            let styTpl = readPart(new Uint8Array(fs.readFileSync('./src/templates/temp_tpc.docx')), 'word/styles.xml')
            let r = prep(cover() + bodyStd, { replaceTocStyles: false })
            for (let n of ['toc 1', 'toc 2', 'toc 3', 'table of figures']) {
                assert.strict.equal(styleOf(r.sty, n), styleOf(styTpl, n), n)
            }
            let r2 = prep(cover() + bodyStd, { replaceTocStyles: false }, { fpTemplate: fpTplWh })
            assert.strict.equal(styleOf(r2.sty, 'toc 1').includes('w:styleId="WMdToc1"'), true)
            assert.strict.equal(styleOf(r2.sty, 'table of figures').includes('w:styleId="WMdTocFigs"'), true)
        })

        it('titleStyle: 模板無預設樣式時另建置中粗體樣式而不提示; 呼叫端指定之樣式不存在時另建並提示', function() {
            let r = prep(cover() + bodyStd, {}, { fpTemplate: fpTplWh })
            assert.strict.deepEqual(r.info.warns, [])
            assert.strict.deepEqual(titlesOf(r.ps).map((p) => p.pStyle), ['WMdTocTitle', 'WMdTocTitle', 'WMdTocTitle'])
            assert.strict.equal(styleOf(r.sty, 'TPC11報告目錄(標題)').includes('<w:jc w:val="center"/>'), true)
            let r2 = prep(cover() + bodyStd, { titleStyle: '目錄標題' }, { fpTemplate: fpTplWh })
            assert.strict.deepEqual(r2.info.warns, ['the template has no paragraph style named "目錄標題", a centered bold style is created for the TOC titles'])
            assert.strict.equal(styleOf(r2.sty, '目錄標題').includes('w:styleId="WMdTocTitle"'), true)
        })

        it('目錄項目字型沿用正文第1個標題之字型, 標題無字型時不指定', function() {
            let fGeorgia = '<w:rFonts w:ascii="Georgia" w:eastAsia="新細明體" w:hAnsi="Georgia"/>'
            let h1 = heading('第一章 緒論', 0, { pb: true })
            let r = prep(cover() + bodyStd.replace(h1, heading('第一章 緒論', 0, { pb: true, fonts: fGeorgia })))
            assert.strict.equal(styleOf(r.sty, 'toc 1').includes(fGeorgia), true)
            assert.strict.equal(styleOf(r.sty, 'table of figures').includes(fGeorgia), true)
            let r2 = prep(cover() + bodyStd.replace(h1, heading('第一章 緒論', 0, { pb: true, fonts: '' })))
            assert.strict.equal(styleOf(r2.sty, 'toc 1').includes('<w:rFonts'), false)
        })

        it('前提不符時拋錯', function() {
            let ex = (bodyXml, opt, re) => assert.throws(() => prepDocxToc(buildDocx(bodyXml), opt), (e) => re.test(e.message), String(re))
            ex(para(run('封面'), `<w:jc w:val="center"/>${getSectFinal()}`) + bodyStd, {}, /^the document has 2 sections, renumbering pages requires a document of 1 section \(set pageNumbers to false to keep the sections\)$/)
            ex(cover() + para('<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>') + bodyStd, {}, /^the document already has a table of contents field/)
            ex(cover() + para('<w:r><w:pict><v:shape><v:textbox><w:txbxContent><w:p><w:r><w:t>框</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></w:r>') + bodyStd, {}, /^text boxes are not supported$/)
            ex(cover() + bodyStd + body('表 1 之參數如下所述。'), {}, /^the numbers of the captions labeled 表 are not consecutive from 1 in document order \(1,2,1\.\.\.\)/)
            ex(cover() + bodyStd, { headStyles: ['TPC11報告目錄(標題)'] }, /^the template has a style named "TPC11報告目錄\(標題\)" that is not bold, please specify another name with headStyles$/)
            ex(cover() + bodyStd.replace(capFig('圖2 流程'), para('<w:r><w:t>圖2</w:t><w:tab/><w:t xml:space="preserve"> 流程</w:t></w:r>')), {}, /^the text run of the caption number contains elements other than text: /)
        })

        it('無標題或只有前置章節時不處理, skip說明原因且內容不變', function() {
            let u8 = buildDocx(cover() + body('沒有標題'))
            let r = prepDocxToc(u8)
            assert.strict.deepEqual(r.info, { skip: 'no heading (paragraph with an outline level) is found, the TOC is not added' })
            assert.strict.equal(r.u8, u8)
            assert.strict.deepEqual(prepDocxToc(buildDocx(cover() + heading('摘要', 0) + body('x'))).info, { skip: 'only front headings are found, the TOC is not added' })
        })

        it('各情境輸出之document.xml與styles.xml標籤成對且子元素順序合規', function() {
            let rs = [
                rStd,
                prep(bodyStd),
                prep(cover() + heading('摘要', 0, { pb: true }) + body('x') + bodyStd),
                prep(cover() + bodyStd, { pageNumbers: false }),
                prep(cover() + bodyStd, { replaceTocStyles: false }, { fpTemplate: fpTplWh }),
                prep(cover() + bodyStd.replace(run('圖1 系統架構'), run('圖') + run('1') + run(' 系統架構'))),
            ]
            for (let r of rs) {
                assert.strict.deepEqual(checkXml(r.doc), [])
                assert.strict.deepEqual(checkXml(r.sty), [])
            }
        })

    })

    describe('checkDocxToc', function() {

        let r = prepDocxToc(buildDocx(cover() + bodyStd))
        let chk = (opt) => checkDocxToc(simulateWordToc(r.u8, r.info, opt), r.info)

        it('Word更新後之目錄與預期相符時ok, 回傳各目錄項目數與目錄第1項之頁碼', function() {
            assert.strict.deepEqual(chk(), { ok: true, errs: [], stats: { toc: 4, fig: 2, tab: 2, pagesFirst: '1' } })
        })

        it('欄位結束位於獨立段落時亦相符', function() {
            assert.strict.equal(chk({ endSeparate: true }).ok, true)
        })

        it('項目文字不符', function() {
            assert.strict.deepEqual(chk({ tamper: 'text' }).errs, ['toc entry 2 "1.1 背景X" differs from the expected "1.1 背景"', 'toc entry 2 links to "1.1 背景" instead of the paragraph of the entry'])
        })

        it('頁碼非依序', function() {
            assert.strict.deepEqual(chk({ tamper: 'page' }).errs, ['fig entry 2 has a page number "0" that is not a number in non-decreasing order'])
        })

        it('連結落點不符', function() {
            assert.strict.deepEqual(chk({ tamper: 'anchor' }).errs, ['tab entry 1 links to "表2 結果" instead of the paragraph of the entry'])
        })

        it('目錄無項目', function() {
            assert.strict.deepEqual(chk({ emptyToc: true }).errs, ['toc has 0 entries but 4 are expected'])
        })

        it('缺少應有之目錄欄位', function() {
            assert.strict.deepEqual(chk({ tamper: 'dropFig' }).errs, ['the fig field is missing'])
        })

        it('正文節指定頁尾', function() {
            assert.strict.deepEqual(chk({ tamper: 'bodyFooter' }).errs, ['the body section should inherit the footer of the previous section and be numbered from 1'])
        })

        it('無封面之文件以2節核對', function() {
            let r2 = prepDocxToc(buildDocx(bodyStd))
            assert.strict.equal(r2.info.cover, false)
            assert.strict.equal(checkDocxToc(simulateWordToc(r2.u8, r2.info), r2.info).ok, true)
            assert.strict.deepEqual(checkDocxToc(simulateWordToc(r2.u8, r2.info, { tamper: 'bodyFooter' }), r2.info).errs, ['the body section should inherit the footer of the previous section and be numbered from 1'])
            assert.strict.deepEqual(checkDocxToc(simulateWordToc(r.u8, r.info), { ...r.info, cover: false }).errs, ['the document has 3 sections but 2 are expected'])
        })

    })

})
