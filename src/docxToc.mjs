import get from 'lodash-es/get.js'
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate'
import isobj from 'wsemi/src/isobj.mjs'
import isarr from 'wsemi/src/isarr.mjs'
import isestr from 'wsemi/src/isestr.mjs'
import isbol from 'wsemi/src/isbol.mjs'
import ispint from 'wsemi/src/ispint.mjs'
import cint from 'wsemi/src/cint.mjs'


//目錄之預設設定(中文報告)
let tocDef = {
    labels: { fig: '圖', tab: '表' }, //圖名、表名之標籤(段首「圖N」「表N」; 亦為 SEQ 欄位之識別名稱)
    titles: { toc: '目錄', fig: '圖目錄', tab: '表目錄' }, //三目錄之標題
    frontHeadings: ['摘要', 'ABSTRACT', 'Abstract'], //位於封面與目錄之間之前置章節(文件開頭之標題文字與之相同者)
    maxLevels: 3, //目錄列出之標題層數(自正文最高之大綱階層起算)
    titleStyle: 'TPC11報告目錄(標題)', //三目錄標題之段落樣式名稱(本套件所附 temp_tpc.docx 之樣式), 模板無此樣式時另建置中粗體之樣式
    headStyles: ['章標題', '節標題', '小節標題'], //承接標題粗體之段落樣式名稱(依層)
    pageNumbers: true, //是否分節: 封面無頁碼、前置章節與三目錄之頁碼為大寫羅馬數字(I 起)、正文之頁碼自 1 起
    keepLines: true, //目錄來源段落(列入目錄之標題、圖名、表名)設為段落內不分頁, 使目錄之頁碼即該段落所在之頁(多行之表名跨頁時, 首行留在前頁而表格在次頁, 表目錄之頁碼會指向前頁)
    keepWithObject: true, //圖名、表名與其圖、表同頁: 下段為表格或圖片時, 圖名表名設與下段同頁; 上段為圖片時, 該段設與下段同頁(圖在頁底而圖名被擠到次頁時, 圖目錄之頁碼會指向次頁)
    replaceTocStyles: true, //是否以本套件之目錄項目樣式取代模板既有之同名樣式(toc N、table of figures)
}


/**
 * 整理添加目錄之設定
 *
 * 各設定非有效值時採預設：labels之fig與tab須為不含空白、引號與反斜線之非空字串且兩者不同(其為SEQ欄位之識別名稱，並置於TOC欄位之\c參數內)，否則兩者皆退回預設；titles之各值須為非空字串；frontHeadings須為陣列(僅保留非空字串)；maxLevels須為正整數，上限9；titleStyle須為非空字串；headStyles須為非空字串組成之非空陣列；pageNumbers、keepLines、keepWithObject、replaceTocStyles須為布林值。
 *
 * @param {Object|Boolean} [opt={}] 輸入設定物件，true或非物件時皆用預設，預設{}
 * @returns {Object} 回傳設定物件{labels,titles,frontHeadings,maxLevels,titleStyle,titleStyleGiven,headStyles,pageNumbers,keepLines,keepWithObject,replaceTocStyles}，titleStyleGiven表示titleStyle是否由呼叫端指定
 */
function normTocOpt(opt = {}) {

    //opt
    if (!isobj(opt)) {
        opt = {}
    }

    //labels, 須為單一字詞(SEQ 識別名稱不可含空白; 引號與反斜線會截斷 TOC 欄位之 \c 參數), 兩者相同時無法區分圖名與表名
    let isLabel = (v) => isestr(v) && !/[\s"\\]/.test(v)
    let labels = {
        fig: get(opt, 'labels.fig', ''),
        tab: get(opt, 'labels.tab', ''),
    }
    for (let k of ['fig', 'tab']) {
        if (!isLabel(labels[k])) {
            labels[k] = tocDef.labels[k]
        }
    }
    if (labels.fig === labels.tab) {
        labels = { ...tocDef.labels }
    }

    //titles
    let titles = {}
    for (let k of ['toc', 'fig', 'tab']) {
        let v = get(opt, `titles.${k}`, '')
        titles[k] = isestr(v) ? v : tocDef.titles[k]
    }

    //frontHeadings
    let frontHeadings = get(opt, 'frontHeadings', null)
    if (isarr(frontHeadings)) {
        frontHeadings = frontHeadings.filter(isestr)
    }
    else {
        frontHeadings = [...tocDef.frontHeadings]
    }

    //maxLevels
    let maxLevels = get(opt, 'maxLevels', null)
    if (ispint(maxLevels)) {
        maxLevels = Math.min(cint(maxLevels), 9)
    }
    else {
        maxLevels = tocDef.maxLevels
    }

    //titleStyle
    let titleStyle = get(opt, 'titleStyle', '')
    let titleStyleGiven = isestr(titleStyle)
    if (!titleStyleGiven) {
        titleStyle = tocDef.titleStyle
    }

    //headStyles
    let headStyles = get(opt, 'headStyles', null)
    if (!isarr(headStyles) || headStyles.length === 0 || !headStyles.every(isestr)) {
        headStyles = [...tocDef.headStyles]
    }

    //pageNumbers, keepLines, keepWithObject, replaceTocStyles
    let bols = {}
    for (let k of ['pageNumbers', 'keepLines', 'keepWithObject', 'replaceTocStyles']) {
        let v = get(opt, k, null)
        bols[k] = isbol(v) ? v : tocDef[k]
    }

    return {
        labels,
        titles,
        frontHeadings,
        maxLevels,
        titleStyle,
        titleStyleGiven,
        headStyles,
        pageNumbers: bols.pageNumbers,
        keepLines: bols.keepLines,
        keepWithObject: bols.keepWithObject,
        replaceTocStyles: bols.replaceTocStyles,
    }
}


//xml 之文字與跳脫
let unesc = (s) => s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, '\'').replace(/&amp;/g, '&')
let esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')
let escRe = (s) => String(s).replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
let textOf = (x) => unesc((x.match(/<w:t(?: [^>]*)?>[^<]*<\/w:t>/g) || []).map((t) => t.replace(/<[^>]+>/g, '')).join(''))
let tElem = (s) => `<w:t${/^\s|\s$/.test(s) ? ' xml:space="preserve"' : ''}>${esc(s)}</w:t>`
let pPrOf = (p) => (p.match(/^<w:p\b[^>]*>\s*(<w:pPr>[\s\S]*?<\/w:pPr>)/) || [])[1] || ''


//scanParas: 依序列出段落與其是否位於表格內(表格內之段落不作為標題、圖名與表名)
function scanParas(doc) {
    let out = []
    let depth = 0
    let re = /<w:tbl\b[^>]*>|<\/w:tbl>|<w:p\b[^>]*\/>|<w:p\b[^>]*>[\s\S]*?<\/w:p>/g
    let m
    while ((m = re.exec(doc)) !== null) {
        let s = m[0]
        if (s.startsWith('<w:tbl')) {
            depth++
            continue
        }
        if (s === '</w:tbl>') {
            depth--
            continue
        }
        out.push({ start: m.index, end: m.index + s.length, xml: s, inTbl: depth > 0 })
    }
    return out
}


//styleInfo: 樣式表之查詢(依名稱取編號、段落樣式之大綱階層含 basedOn 鏈、預設段落樣式)
function styleInfo(sty) {
    let list = []
    for (let m of sty.matchAll(/<w:style\b([^>]*)>([\s\S]*?)<\/w:style>/g)) {
        let at = m[1]
        let body = m[2]
        list.push({
            id: (at.match(/w:styleId="([^"]+)"/) || [])[1],
            type: (at.match(/w:type="([^"]+)"/) || [])[1],
            isDef: /w:default="(1|true)"/.test(at),
            name: unesc((body.match(/<w:name w:val="([^"]+)"\/>/) || [])[1] || ''),
            ol: (body.match(/<w:outlineLvl w:val="(\d)"\/>/) || [])[1],
            based: (body.match(/<w:basedOn w:val="([^"]+)"\/>/) || [])[1],
        })
    }
    let byId = {}
    for (let s of list) {
        byId[s.id] = s
    }
    let olOf = (id, n = 0) => {
        let s = byId[id]
        if (!s || n > 20) {
            return undefined
        }
        return s.ol !== undefined ? s.ol : olOf(s.based, n + 1)
    }
    let idOf = (name) => (list.find((s) => s.type === 'paragraph' && s.name === name) || {}).id
    let defId = (list.find((s) => s.type === 'paragraph' && s.isDef) || {}).id || ''
    return { olOf, idOf, defId }
}


//toSeq: 段落文字中 [s, e) 之編號改為 SEQ 欄位, 顯示之編號不變(編號所在之文字段須只含文字)
function toSeq(p, s, e, ident) {
    let runs = []
    let re = /<w:r\b[^>]*>([\s\S]*?)<\/w:r>/g
    let m
    let off = 0
    while ((m = re.exec(p)) !== null) {
        let t = textOf(m[0])
        runs.push({ i0: m.index, i1: m.index + m[0].length, inner: m[1], t, a: off, b: off + t.length })
        off += t.length
    }
    let hit = runs.filter((r) => r.b > s && r.a < e)
    if (hit.length === 0) {
        throw new Error(`the caption number is not inside a text run: ${textOf(p).slice(0, 30)}`)
    }
    for (let r of hit) {
        //lastRenderedPageBreak 為 Word 記錄上次分頁位置之提示, 存檔時重算, 可略去
        if (!/^(<w:rPr>[\s\S]*?<\/w:rPr>)?(<w:t(?: xml:space="preserve")?>[^<]*<\/w:t>)+$/.test(r.inner.replace(/<w:lastRenderedPageBreak\/>/g, ''))) {
            throw new Error(`the text run of the caption number contains elements other than text: ${textOf(p).slice(0, 30)}`)
        }
    }
    let r0 = hit[0]
    let r1 = hit[hit.length - 1]
    let rPr0 = (r0.inner.match(/^<w:rPr>[\s\S]*?<\/w:rPr>/) || [''])[0]
    let rPr1 = (r1.inner.match(/^<w:rPr>[\s\S]*?<\/w:rPr>/) || [''])[0]
    let run = (rPr, inner) => `<w:r>${rPr}${inner}</w:r>`
    let all = runs.map((r) => r.t).join('')
    let pre = r0.t.slice(0, s - r0.a)
    let post = r1.t.slice(e - r1.a)
    let x = ''
    if (pre) {
        x += run(rPr0, tElem(pre))
    }
    x += run(rPr0, '<w:fldChar w:fldCharType="begin"/>')
    x += run(rPr0, `<w:instrText xml:space="preserve"> SEQ ${esc(ident)} \\* ARABIC </w:instrText>`)
    x += run(rPr0, '<w:fldChar w:fldCharType="separate"/>')
    x += run(rPr0, tElem(all.slice(s, e)))
    x += run(rPr0, '<w:fldChar w:fldCharType="end"/>')
    if (post) {
        x += run(rPr1, tElem(post))
    }
    return p.slice(0, r0.i0) + x + p.slice(r1.i1)
}


//moveBold: 標題段落之粗體改由段落樣式提供(Word 產生目錄時會把標題文字之直接格式帶進目錄項目, 改由樣式提供後目錄項目不帶粗體, 標題外觀不變)
function moveBold(p, styleId) {
    if (/^<w:p\b[^>]*>\s*<w:pPr>/.test(p)) {
        p = p.replace(/^(<w:p\b[^>]*>\s*)<w:pPr>/, `$1<w:pPr><w:pStyle w:val="${styleId}"/>`)
    }
    else {
        p = p.replace(/^(<w:p\b[^>]*>)/, `$1<w:pPr><w:pStyle w:val="${styleId}"/></w:pPr>`)
    }
    let reOn = /<w:b\/>|<w:b w:val="(?:1|true|on)"\/>/
    let off = '<w:b w:val="0"/><w:bCs w:val="0"/>'
    return p.replace(/<w:r\b[^>]*>[\s\S]*?<\/w:r>/g, (r) => {
        if (!/<w:t[ >]/.test(r)) {
            return r
        }
        if (reOn.test(r)) {
            return r.replace(/<w:b\/>|<w:b w:val="(?:1|true|on)"\/>/g, '').replace(/<w:bCs\/>|<w:bCs w:val="(?:1|true|on)"\/>/g, '')
        }
        if (/<w:b w:val="(?:0|false|off)"\/>/.test(r)) {
            return r
        }
        //原非粗體之文字段明確設為不粗體, 維持原外觀(rPr 子元素順序: rStyle, rFonts, b, bCs, …)
        if (/<w:rPr>/.test(r)) {
            if (/<w:rFonts\b[^>]*\/>/.test(r)) {
                return r.replace(/(<w:rFonts\b[^>]*\/>)/, `$1${off}`)
            }
            if (/<w:rStyle\b[^>]*\/>/.test(r)) {
                return r.replace(/(<w:rStyle\b[^>]*\/>)/, `$1${off}`)
            }
            return r.replace('<w:rPr>', `<w:rPr>${off}`)
        }
        return r.replace(/^(<w:r\b[^>]*>)/, `$1<w:rPr>${off}</w:rPr>`)
    })
}


//addKeepLines: 段落設為段落內不分頁(pPr 子元素順序: pStyle, keepNext, keepLines, …)
function addKeepLines(p) {
    let pPr = pPrOf(p)
    if (/<w:keepLines\/>|<w:keepLines w:val="(?:1|true|on)"\/>/.test(pPr)) {
        return p
    }
    if (/<w:keepLines\b/.test(pPr)) {
        return p.replace(/<w:keepLines\b[^>]*\/>/, '<w:keepLines/>')
    }
    if (pPr) {
        return p.replace(/^(<w:p\b[^>]*>\s*<w:pPr>(?:<w:pStyle\b[^>]*\/>)?(?:<w:keepNext\b[^>]*\/>)?)/, '$1<w:keepLines/>')
    }
    return p.replace(/^(<w:p\b[^>]*>)/, '$1<w:pPr><w:keepLines/></w:pPr>')
}


//addKeepNext: 段落設為與下段同頁(pPr 子元素順序: pStyle, keepNext, keepLines, …)
function addKeepNext(p) {
    let pPr = pPrOf(p)
    if (/<w:keepNext\/>|<w:keepNext w:val="(?:1|true|on)"\/>/.test(pPr)) {
        return p
    }
    if (/<w:keepNext\b/.test(pPr)) {
        return p.replace(/<w:keepNext\b[^>]*\/>/, '<w:keepNext/>')
    }
    if (pPr) {
        return p.replace(/^(<w:p\b[^>]*>\s*<w:pPr>(?:<w:pStyle\b[^>]*\/>)?)/, '$1<w:keepNext/>')
    }
    return p.replace(/^(<w:p\b[^>]*>)/, '$1<w:pPr><w:keepNext/></w:pPr>')
}


//dropLeadBreak: 移除段落開頭之分頁(段落前分頁、開頭只含分頁符號之文字段); 新的一節已自新頁開始, 留著會多出空白頁
function dropLeadBreak(p) {
    let p2 = p.replace(/<w:pageBreakBefore\/>/, '')
    p2 = p2.replace(/^(<w:p\b[^>]*>\s*(?:<w:pPr>[\s\S]*?<\/w:pPr>)?)<w:r\b[^>]*>(?:<w:rPr>[\s\S]*?<\/w:rPr>)?<w:br w:type="page"\/><\/w:r>/, '$1')
    return p2
}


/**
 * 準備添加目錄之docx內容(純結構處理, 不含Word)
 *
 * 處理內容：圖名與表名之編號改為SEQ欄位、標題粗體改由段落樣式提供、目錄來源段落(列入目錄之標題、圖名、表名)設為段落內不分頁(keepLines)、圖名表名與其圖表同頁(keepWithObject)、於正文第1個標題前插入目錄與圖目錄、表目錄之欄位(無圖名或表名時不列該目錄)、分節與頁碼、目錄項目之樣式。目錄之內容與頁碼須再交Word更新。
 *
 * 分節(pageNumbers為true)：第1個標題前有內容(含文字或圖片之段落、表格)者視為有封面，分為封面(無頁碼)、前置章節與目錄(大寫羅馬數字自I起)、正文(自1起)共3節；無封面者不建封面節，分為2節，以免多出空白首頁。
 *
 * 圖名、表名以段首「標籤＋編號＋空白」判斷；段首為標籤與編號但其後非空白者(如「圖1：名稱」)不視為圖名表名，以warns提示筆數與前3例。目錄項目之字型沿用正文第1個標題之文字段字型，查無則沿用模板預設。
 *
 * 前提不符時拋錯(不產出錯誤之目錄)：含文字方塊、已有目錄欄位、分節不只1個(pageNumbers為true時)、圖名或表名之編號未依文件順序自1連續、編號所在之文字段含文字以外之元素、headStyles指定之既有樣式非粗體。
 *
 * @param {Uint8Array} u8 輸入docx檔之內容
 * @param {Object|Boolean} [opt={}] 輸入設定物件，見normTocOpt
 * @returns {Object} 回傳{u8,info}，u8為處理後之docx內容，info為{skip,cover,levels,front,expect,labels,lists,bookmark,pageNumbers,warns}；無標題時info僅含skip(原因字串)且u8為原內容
 */
function prepDocxToc(u8, opt = {}) {
    let o = normTocOpt(opt)
    let z = unzipSync(u8)
    if (!z['word/document.xml'] || !z['word/styles.xml']) {
        throw new Error('the docx lacks word/document.xml or word/styles.xml')
    }
    let doc = strFromU8(z['word/document.xml'])
    let sty = strFromU8(z['word/styles.xml'])
    let warns = []

    //前提
    if (/<w:txbxContent/.test(doc)) {
        throw new Error('text boxes are not supported')
    }
    if (/<w:instrText[^>]*>\s*TOC\b/.test(doc) || /<w:fldSimple w:instr="\s*TOC\b/.test(doc)) {
        throw new Error('the document already has a table of contents field, the TOC must be added to a docx just converted')
    }
    let nSect = (doc.match(/<w:sectPr\b/g) || []).length
    if (o.pageNumbers && nSect !== 1) {
        throw new Error(`the document has ${nSect} sections, renumbering pages requires a document of 1 section (set pageNumbers to false to keep the sections)`)
    }

    let si = styleInfo(sty)
    let paras = scanParas(doc)
    let tops = paras.filter((p) => !p.inTbl)
    let hasPic = (p) => /<w:drawing\b|<w:pict\b|<w:object\b/.test(p.xml)
    let levelOf = (p) => {
        let pPr = pPrOf(p)
        let m = pPr.match(/<w:outlineLvl w:val="(\d)"\/>/)
        if (m) {
            return Number(m[1])
        }
        let ps = pPr.match(/<w:pStyle w:val="([^"]+)"\/>/)
        if (ps) {
            let ol = si.olOf(ps[1])
            if (ol !== undefined) {
                return Number(ol)
            }
        }
        return 9
    }
    let heads = tops.map((p) => ({ ...p, lv: levelOf(p.xml), text: textOf(p.xml).trim() })).filter((p) => p.lv < 9 && p.text !== '')
    if (heads.length === 0) {
        return { u8, info: { skip: 'no heading (paragraph with an outline level) is found, the TOC is not added' } }
    }
    let iFront = 0
    while (iFront < heads.length && o.frontHeadings.includes(heads[iFront].text)) {
        iFront++
    }
    let front = heads.slice(0, iFront)
    let body = heads.slice(iFront)
    if (body.length === 0) {
        return { u8, info: { skip: 'only front headings are found, the TOC is not added' } }
    }
    let h0 = heads[0]
    let bodyFirst = body[0]
    let l0 = Math.min(...body.map((h) => h.lv))
    let l1 = Math.min(Math.max(...body.map((h) => h.lv)), l0 + o.maxLevels - 1)
    let headsIn = body.filter((h) => h.lv <= l1)

    //hasCover: 第 1 個標題之前有內容(含文字或圖片之段落、表格)才視為有封面; 無封面時不建封面節, 免得多出空白首頁
    let hasCover = /<w:tbl\b/.test(doc.slice(0, h0.start)) || tops.some((p) => p.end <= h0.start && (textOf(p.xml).trim() !== '' || hasPic(p)))

    //圖名與表名: 段首「圖N」「表N」(標籤與編號間可有空白, 編號後須接空白), 編號須依文件順序自 1 連續
    let reCap = new RegExp(`^(${escRe(o.labels.fig)}|${escRe(o.labels.tab)})(\\s*)(\\d+)\\s`)
    let reCapLike = new RegExp(`^(?:${escRe(o.labels.fig)}|${escRe(o.labels.tab)})\\s*\\d+`)
    let caps = []
    let capsLike = []
    for (let p of tops) {
        let t = textOf(p.xml)
        let m = t.match(reCap)
        if (m) {
            let kind = m[1] === o.labels.fig ? 'fig' : 'tab'
            let s = m[1].length + m[2].length
            caps.push({ ...p, kind, num: Number(m[3]), s, e: s + m[3].length, text: t.trim() })
        }
        else if (reCapLike.test(t)) {
            capsLike.push(t.trim())
        }
    }
    if (capsLike.length > 0) {
        warns.push(`${capsLike.length} paragraph(s) start with a figure or table label and a number that is not followed by a space, so they are not treated as captions nor listed: ${capsLike.slice(0, 3).map((t) => t.slice(0, 30)).join('; ')}`)
    }
    for (let kind of ['fig', 'tab']) {
        let nums = caps.filter((c) => c.kind === kind).map((c) => c.num)
        if (!nums.every((v, i) => v === i + 1)) {
            throw new Error(`the numbers of the captions labeled ${o.labels[kind]} are not consecutive from 1 in document order (${nums.slice(0, 10).join(',')}...), Word would renumber them after they become SEQ fields`)
        }
    }
    let capsBody = caps.filter((c) => c.start > bodyFirst.start)
    let nFig = capsBody.filter((c) => c.kind === 'fig').length
    let nTab = capsBody.filter((c) => c.kind === 'tab').length

    //樣式: 標題粗體之承接樣式(依層)
    let idHead = {}
    for (let lv = l0; lv <= l1; lv++) {
        let name = o.headStyles[Math.min(lv - l0, o.headStyles.length - 1)]
        let id = si.idOf(name)
        if (id) {
            //既有同名樣式須為粗體, 否則標題移除直接粗體後外觀會改變
            let mSty = sty.match(new RegExp(`<w:style\\b[^>]*w:styleId="${escRe(id)}"[^>]*>[\\s\\S]*?</w:style>`))
            if (!mSty || !/<w:b\/>|<w:b w:val="(?:1|true|on)"\/>/.test(mSty[0])) {
                throw new Error(`the template has a style named "${name}" that is not bold, please specify another name with headStyles`)
            }
        }
        else {
            id = `WMdTocHead${lv - l0 + 1}`
            sty = sty.replace('</w:styles>', `<w:style w:type="paragraph" w:customStyle="1" w:styleId="${id}"><w:name w:val="${esc(name)}"/>${si.defId ? `<w:basedOn w:val="${si.defId}"/><w:next w:val="${si.defId}"/>` : ''}<w:qFormat/><w:rPr><w:b/><w:bCs/></w:rPr></w:style></w:styles>`)
        }
        idHead[lv] = id
    }

    //各段落之改寫(以段落起點為鍵): fns 依序改寫該段落, pre 為依序插於該段落之前之內容
    let edits = new Map()
    let editOf = (p) => {
        if (!edits.has(p.start)) {
            edits.set(p.start, { p, fns: [], pre: [] })
        }
        return edits.get(p.start)
    }
    let addEdit = (p, fn) => {
        editOf(p).fns.push(fn)
    }
    let addPre = (p, s) => {
        editOf(p).pre.push(s)
    }
    for (let c of caps) {
        addEdit(c, (x) => toSeq(x, c.s, c.e, o.labels[c.kind]))
    }
    for (let h of headsIn) {
        if (/<w:pStyle /.test(pPrOf(h.xml))) {
            warns.push(`the heading already has a paragraph style, its TOC entry may carry its direct formatting: ${h.text}`)
            continue
        }
        addEdit(h, (x) => moveBold(x, idHead[h.lv]))
    }
    if (o.keepLines) {
        for (let p of [...headsIn, ...caps]) {
            addEdit(p, addKeepLines)
        }
    }
    if (o.keepWithObject) {
        //相鄰判斷: 兩段落之間只有空白或書籤(中間有表格即不相鄰)
        let between = (a, b) => doc.slice(a, b).replace(/<w:bookmark(?:Start|End)\b[^>]*\/>/g, '').trim()
        for (let c of caps) {
            let i = paras.findIndex((p) => p.start === c.start)
            let prev = i > 0 ? paras[i - 1] : null
            let next = i >= 0 && i < paras.length - 1 ? paras[i + 1] : null
            let tblNext = /^<w:tbl\b/.test(doc.slice(c.end).replace(/^(?:\s|<w:bookmark(?:Start|End)\b[^>]*\/>)+/, ''))
            let picNext = next && !next.inTbl && between(c.end, next.start) === '' && hasPic(next)
            let picPrev = prev && !prev.inTbl && between(prev.end, c.start) === '' && hasPic(prev)
            if (tblNext || picNext) {
                addEdit(c, addKeepNext)
            }
            else if (picPrev) {
                addEdit(prev, addKeepNext)
            }
        }
    }

    //三目錄區塊
    let idTitle = si.idOf(o.titleStyle)
    if (!idTitle) {
        idTitle = 'WMdTocTitle'
        sty = sty.replace('</w:styles>', `<w:style w:type="paragraph" w:customStyle="1" w:styleId="${idTitle}"><w:name w:val="${esc(o.titleStyle)}"/>${si.defId ? `<w:basedOn w:val="${si.defId}"/>` : ''}<w:qFormat/><w:pPr><w:jc w:val="center"/></w:pPr><w:rPr><w:b/><w:bCs/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr></w:style></w:styles>`)
        if (o.titleStyleGiven) {
            //使用預設名稱(本套件所附模板之樣式)而模板無此樣式屬常態, 僅於呼叫端指定時提示
            warns.push(`the template has no paragraph style named "${o.titleStyle}", a centered bold style is created for the TOC titles`)
        }
    }
    let bookmark = ''
    let bmId = 0
    if (front.length > 0) {
        //書籤名不以 _Toc 開頭(Word 更新目錄時會自行管理 _Toc 開頭之書籤)
        bookmark = '_WMdTocBody'
        let ids = [...doc.matchAll(/<w:bookmarkStart\b[^>]*w:id="(\d+)"/g)].map((m) => Number(m[1]))
        bmId = ids.length ? Math.max(...ids) + 1 : 0
    }
    let swB = bookmark ? ` \\b ${bookmark}` : ''
    let title = (t, pb) => `<w:p><w:pPr><w:pStyle w:val="${idTitle}"/>${pb ? '<w:pageBreakBefore/>' : ''}<w:spacing w:after="240"/></w:pPr><w:r>${tElem(t)}</w:r></w:p>`
    let field = (instr) => `<w:p><w:r><w:fldChar w:fldCharType="begin" w:dirty="true"/></w:r><w:r><w:instrText xml:space="preserve"> ${esc(instr)} </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>Update field to build the table</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>`
    let lists = ['toc']
    //分節: 封面節(未指定頁尾即無頁尾, 無封面者不建)、前置章節與三目錄節(沿用原頁尾, 大寫羅馬數字自 I 起)、正文節(頁尾沿用前節, 自 1 起)
    let pageNumbers = o.pageNumbers
    //目錄標題之段落前分頁: 前有前置章節時須換頁; 不分節而有封面時亦須換頁; 分節時新的一節已自新頁開始
    let firstPb = front.length > 0 || (!pageNumbers && hasCover)
    let block = title(o.titles.toc, firstPb) + field(`TOC \\o "${l0 + 1}-${l1 + 1}" \\h \\z \\u${swB}`)
    if (nFig > 0) {
        block += title(o.titles.fig, true) + field(`TOC \\h \\z \\c "${o.labels.fig}"${swB}`)
        lists.push('fig')
    }
    if (nTab > 0) {
        block += title(o.titles.tab, true) + field(`TOC \\h \\z \\c "${o.labels.tab}"${swB}`)
        lists.push('tab')
    }
    let sectFinalNew = ''
    let iSect = -1
    let sectFinalOld = ''
    if (pageNumbers) {
        let mSect = doc.match(/<w:sectPr\b[^>]*>([\s\S]*?)<\/w:sectPr>(?=\s*<\/w:body>)/)
        if (!mSect) {
            throw new Error('the final section properties of the document are not found')
        }
        iSect = doc.lastIndexOf(mSect[0])
        sectFinalOld = mSect[0]
        let inner = mSect[1]
        let part = (tag) => (inner.match(new RegExp(`<w:${tag}\\b[^>]*/>|<w:${tag}\\b[^>]*>[\\s\\S]*?</w:${tag}>`)) || [''])[0]
        let refs = (inner.match(/<w:headerReference\b[^>]*\/>/g) || []).join('') + (inner.match(/<w:footerReference\b[^>]*\/>/g) || []).join('')
        if (!/<w:footerReference\b/.test(refs)) {
            warns.push('the document has no footer, so the page numbers are not shown')
        }
        //sectPr 子元素順序: headerReference, footerReference, …, pgSz, pgMar, …, pgNumType, cols, …, docGrid
        let sect1 = `<w:sectPr>${part('pgSz')}${part('pgMar')}${part('cols')}${part('docGrid')}</w:sectPr>`
        let sect2 = `<w:sectPr>${refs}${part('pgSz')}${part('pgMar')}<w:pgNumType w:fmt="upperRoman" w:start="1"/>${part('cols')}${part('docGrid')}</w:sectPr>`
        let inner3 = inner.replace(/<w:headerReference\b[^>]*\/>|<w:footerReference\b[^>]*\/>/g, '')
        if (/<w:pgNumType\b/.test(inner3)) {
            inner3 = inner3.replace(/<w:pgNumType\b([^>]*?)\/>/, (x, a) => `<w:pgNumType${/w:start=/.test(a) ? a.replace(/w:start="\d+"/, 'w:start="1"') : `${a} w:start="1"`}/>`)
        }
        else {
            inner3 = inner3.replace(/(<w:pgMar\b[^>]*\/>)/, '$1<w:pgNumType w:start="1"/>')
        }
        sectFinalNew = sectFinalOld.replace(inner, inner3)
        //目錄節之結束段落: 行高與字級 1 點, 免得最後一頁恰好滿頁時被擠到下一頁成為空白頁
        block += `<w:p><w:pPr><w:snapToGrid w:val="0"/><w:spacing w:before="0" w:after="0" w:line="20" w:lineRule="exact"/><w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr>${sect2}</w:pPr></w:p>`
        if (hasCover) {
            //封面之結束: 第 1 個標題前之段落帶封面節之分節設定(不另加段落, 免得封面滿頁時多出空白頁); 其前不是段落時才另加 1 點高之段落
            let prev = tops.filter((p) => p.end <= h0.start).pop()
            if (prev && doc.slice(prev.end, h0.start).trim() === '') {
                addEdit(prev, (x) => (/^<w:p\b[^>]*\/>$/.test(x) ? x.replace(/\/>$/, `><w:pPr>${sect1}</w:pPr></w:p>`) : /^<w:p\b[^>]*>\s*<w:pPr>/.test(x) ? x.replace('</w:pPr>', `${sect1}</w:pPr>`) : x.replace(/^(<w:p\b[^>]*>)/, `$1<w:pPr>${sect1}</w:pPr>`)))
            }
            else {
                addPre(h0, `<w:p><w:pPr><w:snapToGrid w:val="0"/><w:spacing w:before="0" w:after="0" w:line="20" w:lineRule="exact"/><w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr>${sect1}</w:pPr></w:p>`)
            }
        }
        //各節之第 1 個標題移除開頭之分頁
        addEdit(h0, dropLeadBreak)
        if (front.length > 0) {
            addEdit(bodyFirst, dropLeadBreak)
        }
    }
    //正文第 1 個標題之前插入三目錄(有前置章節時另以書籤限定目錄只收正文)
    addPre(bodyFirst, block + (bookmark ? `<w:bookmarkStart w:id="${bmId}" w:name="${bookmark}"/>` : ''))

    //套用(由後往前, 位置不受前段改寫影響)
    if (bookmark) {
        let iEnd = iSect >= 0 ? iSect : doc.lastIndexOf('<w:sectPr')
        doc = doc.slice(0, iEnd) + `<w:bookmarkEnd w:id="${bmId}"/>` + doc.slice(iEnd)
    }
    if (pageNumbers) {
        let iS = doc.lastIndexOf(sectFinalOld)
        doc = doc.slice(0, iS) + sectFinalNew + doc.slice(iS + sectFinalOld.length)
    }
    for (let { p, fns, pre } of [...edits.values()].sort((a, b) => b.p.start - a.p.start)) {
        let x = p.xml
        for (let fn of fns) {
            x = fn(x)
        }
        doc = doc.slice(0, p.start) + pre.join('') + x + doc.slice(p.end)
    }

    //目錄項目之樣式(依樣式名稱; 行距: 第 1 層 1.5 倍、其餘單行; 章粗體、段前 6 點; 下層每層縮排 2 字並懸掛縮排使換行對齊; 右縮排 2 字; 點線引導至靠右之頁碼)
    //  定位點位置取版心寬(頁寬減左右邊界)
    let pgW = Number((doc.match(/<w:pgSz\b[^>]*w:w="(\d+)"/) || [])[1] || 11906)
    let mar = doc.match(/<w:pgMar\b[^>]*\/>/)
    let marL = Number(((mar && mar[0].match(/w:left="(\d+)"/)) || [])[1] || 1800)
    let marR = Number(((mar && mar[0].match(/w:right="(\d+)"/)) || [])[1] || 1800)
    let tabPos = pgW - marL - marR - 10
    //字型沿用正文第 1 個標題之文字段字型(w-html2docx 以 fontFamilies 套用於全文, 故與正文一致), 查無則不指定而沿用模板預設
    let fonts = ''
    for (let r of bodyFirst.xml.match(/<w:r\b[^>]*>[\s\S]*?<\/w:r>/g) || []) {
        if (/<w:t[ >]/.test(r)) {
            fonts = (r.match(/<w:rFonts\b[^>]*\/>/) || [''])[0]
            break
        }
    }
    let tab = `<w:tabs><w:tab w:val="right" w:leader="dot" w:pos="${tabPos}"/></w:tabs>`
    let setStyle = (name, idNew, pPr, rPr, ui) => {
        let id = si.idOf(name)
        if (id && !o.replaceTocStyles) {
            return //沿用模板既有之樣式
        }
        id = id || idNew
        let xml = `<w:style w:type="paragraph" w:styleId="${id}"><w:name w:val="${esc(name)}"/>${si.defId ? `<w:basedOn w:val="${si.defId}"/><w:next w:val="${si.defId}"/>` : ''}<w:uiPriority w:val="${ui}"/><w:unhideWhenUsed/><w:pPr>${pPr}</w:pPr><w:rPr>${rPr}</w:rPr></w:style>`
        let re = new RegExp(`<w:style\\b[^>]*w:styleId="${escRe(id)}"[^>]*>[\\s\\S]*?</w:style>`)
        sty = re.test(sty) ? sty.replace(re, xml) : sty.replace('</w:styles>', `${xml}</w:styles>`)
    }
    for (let lv = l0; lv <= l1; lv++) {
        let k = lv - l0
        if (k === 0) {
            setStyle(`toc ${lv + 1}`, `WMdToc${lv + 1}`, `${tab}<w:spacing w:before="120" w:line="360" w:lineRule="auto"/><w:ind w:left="0" w:right="480"/>`, `${fonts}<w:b/><w:bCs/><w:noProof/><w:kern w:val="0"/>`, 39)
        }
        else {
            setStyle(`toc ${lv + 1}`, `WMdToc${lv + 1}`, `${tab}<w:ind w:left="${k * 480 + 480}" w:right="480" w:hanging="480"/>`, `${fonts}<w:noProof/><w:kern w:val="0"/>`, 39)
        }
    }
    setStyle('table of figures', 'WMdTocFigs', `${tab}<w:ind w:left="720" w:right="480" w:hanging="720"/>`, `${fonts}<w:noProof/><w:kern w:val="0"/>`, 99)

    //寫出(圖檔已壓縮, 不再壓)
    let out = {}
    for (let [k, v] of Object.entries(z)) {
        let d = k === 'word/document.xml' ? strToU8(doc) : k === 'word/styles.xml' ? strToU8(sty) : v
        out[k] = /^word\/media\//.test(k) ? [d, { level: 0 }] : d
    }
    let u8Out = zipSync(out, { level: 6 })
    let info = {
        skip: '',
        cover: hasCover,
        levels: [l0 + 1, l1 + 1],
        front: front.map((h) => h.text),
        expect: {
            toc: headsIn.map((h) => h.text),
            fig: capsBody.filter((c) => c.kind === 'fig').map((c) => c.text),
            tab: capsBody.filter((c) => c.kind === 'tab').map((c) => c.text),
        },
        labels: o.labels,
        lists,
        bookmark,
        pageNumbers,
        warns,
    }
    return { u8: u8Out, info }
}


/**
 * 核對Word更新後之目錄
 *
 * 逐項取出目錄、圖目錄、表目錄之項目，檢查：項目與預期之標題、圖名、表名逐一相同且順序相同、各項為超連結且其書籤落在同文字之段落、頁碼為依序不遞減之數字、應有之目錄欄位皆存在；分節時另查各節之頁尾與頁碼設定(有封面3節，無封面2節)。
 *
 * @param {Uint8Array} u8 輸入Word更新後之docx內容
 * @param {Object} info 輸入prepDocxToc回傳之info
 * @returns {Object} 回傳{ok,errs,stats}，errs為不符之說明字串陣列，stats為{toc,fig,tab,pagesFirst}，分別為各目錄之項目數與目錄第1項之頁碼
 */
function checkDocxToc(u8, info) {
    let z = unzipSync(u8, { filter: (f) => f.name === 'word/document.xml' })
    let doc = strFromU8(z['word/document.xml'])
    let errs = []

    //依序走訪段落、欄位(含巢狀)、文字、定位字元、超連結、書籤
    let re = /<w:p\b[^>]*>|<\/w:p>|<w:fldChar\b[^>]*w:fldCharType="(begin|separate|end)"[^>]*\/>|<w:instrText\b[^>]*>([^<]*)<\/w:instrText>|<w:t(?: [^>]*)?>([^<]*)<\/w:t>|<w:tab\/>|<w:hyperlink\b[^>]*w:anchor="([^"]+)"[^>]*>|<w:bookmarkStart\b[^>]*w:name="([^"]+)"[^>]*\/>/g
    let paras = []
    let cur = null
    let stack = []
    let tocs = []
    let bm = {}
    let openToc = () => {
        for (let k = stack.length - 1; k >= 0; k--) {
            if (stack[k].toc && stack[k].sep) {
                return stack[k].toc
            }
        }
        return null
    }
    for (let m of doc.matchAll(re)) {
        let s = m[0]
        if (s.startsWith('<w:p')) {
            cur = { text: '', segs: [''], anchor: null, toc: openToc() }
            paras.push(cur)
        }
        else if (s === '</w:p>') {
            if (cur) {
                //目錄之首段含欄位之開始、末段含欄位之結束, 故於段首與段尾各判斷一次
                cur.toc = cur.toc || openToc()
                if (cur.toc && cur.anchor) {
                    cur.toc.entries.push({ anchor: cur.anchor, text: cur.segs.slice(0, -1).join('').trim(), page: cur.segs[cur.segs.length - 1].trim() })
                }
            }
            cur = null
        }
        else if (m[1] === 'begin') {
            stack.push({ instr: '', sep: false, toc: null })
        }
        else if (m[1] === 'separate') {
            let f = stack[stack.length - 1]
            if (f) {
                f.sep = true
                if (/^\s*TOC\b/.test(f.instr)) {
                    f.toc = { instr: f.instr.trim(), entries: [] }
                    tocs.push(f.toc)
                }
            }
        }
        else if (m[1] === 'end') {
            stack.pop()
        }
        else if (m[2] !== undefined) {
            if (stack.length) {
                stack[stack.length - 1].instr += unesc(m[2])
            }
        }
        else if (m[3] !== undefined) {
            if (cur) {
                cur.text += unesc(m[3])
                cur.segs[cur.segs.length - 1] += unesc(m[3])
            }
        }
        else if (s === '<w:tab/>') {
            if (cur) {
                cur.segs.push('')
            }
        }
        else if (m[4]) {
            if (cur && !cur.anchor) {
                cur.anchor = m[4]
            }
        }
        else if (m[5]) {
            bm[m[5]] = paras.length - 1
        }
    }
    let kindOf = (instr) => {
        let mc = instr.match(/\\c\s+"([^"]+)"/)
        if (!mc) {
            return 'toc'
        }
        return mc[1] === info.labels.fig ? 'fig' : mc[1] === info.labels.tab ? 'tab' : 'unknown'
    }
    let stats = { toc: 0, fig: 0, tab: 0, pagesFirst: '' }
    let seen = []
    for (let t of tocs) {
        let kind = kindOf(t.instr)
        seen.push(kind)
        let exp = info.expect[kind] || []
        stats[kind] = t.entries.length
        if (t.entries.length !== exp.length) {
            errs.push(`${kind} has ${t.entries.length} entries but ${exp.length} are expected`)
        }
        let last = 0
        t.entries.forEach((e, i) => {
            let ip = bm[e.anchor]
            let target = ip === undefined ? null : paras[ip].text.trim()
            if (e.text !== exp[i]) {
                errs.push(`${kind} entry ${i + 1} "${e.text}" differs from the expected "${exp[i]}"`)
            }
            if (target !== e.text) {
                errs.push(`${kind} entry ${i + 1} links to "${target}" instead of the paragraph of the entry`)
            }
            let pg = /^\d+$/.test(e.page) ? Number(e.page) : NaN
            if (Number.isNaN(pg) || pg < last) {
                errs.push(`${kind} entry ${i + 1} has a page number "${e.page}" that is not a number in non-decreasing order`)
            }
            else {
                last = pg
            }
        })
        if (kind === 'toc' && t.entries.length) {
            stats.pagesFirst = t.entries[0].page
        }
    }
    for (let kind of info.lists) {
        if (!seen.includes(kind)) {
            errs.push(`the ${kind} field is missing`)
        }
    }

    //分節: 有封面 3 節(封面、前置章節與目錄、正文), 無封面 2 節(前置章節與目錄、正文)
    if (info.pageNumbers) {
        let hasCover = info.cover !== false
        let nExp = hasCover ? 3 : 2
        let secs = [...doc.matchAll(/<w:sectPr\b[^>]*>([\s\S]*?)<\/w:sectPr>/g)].map((mm) => mm[1])
        if (secs.length !== nExp) {
            errs.push(`the document has ${secs.length} sections but ${nExp} are expected`)
        }
        else {
            let secToc = secs[hasCover ? 1 : 0]
            let secBody = secs[secs.length - 1]
            if (hasCover && /<w:footerReference\b/.test(secs[0])) {
                errs.push('the cover section should have no footer')
            }
            if (!/<w:pgNumType\b[^>]*w:fmt="upperRoman"[^>]*w:start="1"/.test(secToc) && !/<w:pgNumType\b[^>]*w:start="1"[^>]*w:fmt="upperRoman"/.test(secToc)) {
                errs.push('the section of the front headings and the TOC should be numbered in upper Roman numerals from I')
            }
            if (/<w:footerReference\b/.test(secBody) || !/<w:pgNumType\b[^>]*w:start="1"/.test(secBody)) {
                errs.push('the body section should inherit the footer of the previous section and be numbered from 1')
            }
        }
    }
    return { ok: errs.length === 0, errs, stats }
}


export { normTocOpt, prepDocxToc, checkDocxToc }
