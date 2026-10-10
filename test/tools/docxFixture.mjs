import fs from 'fs'
import path from 'path'
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate'


//docxFixture: 以 docx 模板之樣式、頁尾與分節為底, 合成仿 w-html2docx 產出之 document.xml, 供不需 Word 之目錄測試
//  標題為段落直接設定之大綱階層與直接粗體、圖在上圖名在下、表名在上帶 keepNext(HTML 之 page-break-after:avoid)、表格內文字不計為圖表名
//  pic() 僅含 w:drawing 標記供純結構處理判斷, 非完整圖片, 含 pic() 之 fixture 不可交 Word 開啟


let esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
let unesc = (s) => s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, '\'').replace(/&amp;/g, '&')
let textOf = (x) => unesc((x.match(/<w:t(?: [^>]*)?>[^<]*<\/w:t>/g) || []).map((t) => t.replace(/<[^>]+>/g, '')).join(''))
let fontsDef = '<w:rFonts w:ascii="Times New Roman" w:eastAsia="標楷體" w:hAnsi="Times New Roman"/>'

let run = (t, { b = false, sz = 24, fonts = fontsDef } = {}) => `<w:r><w:rPr>${fonts}${b ? '<w:b/><w:bCs/>' : ''}<w:sz w:val="${sz}"/></w:rPr><w:t xml:space="preserve">${esc(t)}</w:t></w:r>`
let para = (inner, pPr = '') => `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}${inner}</w:p>`
let heading = (t, lv, { pb = false, pStyle = '', fonts = fontsDef } = {}) => para(`${pb ? '<w:r><w:br w:type="page"/></w:r>' : ''}${run(t, { b: true, sz: 32 - lv * 4, fonts })}`, `${pStyle ? `<w:pStyle w:val="${pStyle}"/>` : ''}<w:keepNext/><w:spacing w:before="240" w:after="120"/><w:outlineLvl w:val="${lv}"/>`)
let body = (t) => para(run(t), '<w:spacing w:line="360" w:lineRule="auto"/><w:ind w:firstLine="480"/><w:jc w:val="both"/>')
let pic = () => para('<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="3000000" cy="2000000"/><wp:docPr id="1" name="pic"/></wp:inline></w:drawing></w:r>', '<w:jc w:val="center"/>')
let capFig = (t) => para(run(t), '<w:jc w:val="center"/>')
let capTab = (t) => para(run(t), '<w:keepNext/><w:jc w:val="center"/>')
let table = (t = '表1 表格內之文字') => `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>${para(run(t))}</w:tc></w:tr></w:tbl>`
let cover = () => para(run('測試計畫報告', { b: true, sz: 40 }), '<w:jc w:val="center"/>') + para(run('2026年10月'), '<w:jc w:val="center"/>')


//buildDocx: 以模板之 document.xml 開頭(命名空間)與最後分節設定, 包住 bodyXml 成為 docx; sect 給字串時取代最後分節設定
function buildDocx(bodyXml, opt = {}) {
    let fpTemplate = opt.fpTemplate || './src/templates/temp_tpc.docx'
    let tpl = unzipSync(new Uint8Array(fs.readFileSync(path.resolve(fpTemplate))))
    let docTpl = strFromU8(tpl['word/document.xml'])
    let head = docTpl.slice(0, docTpl.indexOf('<w:body>') + '<w:body>'.length)
    if (!/xmlns:wp=/.test(head)) {
        head = head.replace('<w:document ', '<w:document xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ')
    }
    let sect = typeof opt.sect === 'string' ? opt.sect : getSectFinal(fpTemplate)
    let doc = `${head}${bodyXml}${sect}</w:body></w:document>`
    return zipSync({ ...tpl, 'word/document.xml': strToU8(doc) })
}


//getSectFinal: 取模板最後之分節設定
function getSectFinal(fpTemplate = './src/templates/temp_tpc.docx') {
    let tpl = unzipSync(new Uint8Array(fs.readFileSync(path.resolve(fpTemplate))))
    return strFromU8(tpl['word/document.xml']).match(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/)[0]
}


//readPart: 取 docx 內之檔案文字
function readPart(u8, name) {
    return strFromU8(unzipSync(u8)[name])
}


//listParas: 依序列出 document.xml 之頂層段落與表格標記, 供斷言段落屬性
function listParas(doc) {
    let out = []
    let depth = 0
    for (let m of doc.matchAll(/<w:tbl\b[^>]*>|<\/w:tbl>|<w:p\b[^>]*\/>|<w:p\b[^>]*>[\s\S]*?<\/w:p>/g)) {
        let x = m[0]
        if (x.startsWith('<w:tbl')) {
            depth++
            continue
        }
        if (x === '</w:tbl>') {
            depth--
            continue
        }
        if (depth > 0) {
            continue
        }
        let pPr = (x.match(/^<w:p\b[^>]*>\s*(<w:pPr>[\s\S]*?<\/w:pPr>)/) || [])[1] || ''
        let runs = x.match(/<w:r\b[^>]*>[\s\S]*?<\/w:r>/g) || []
        out.push({
            xml: x,
            text: textOf(x),
            pStyle: (pPr.match(/<w:pStyle w:val="([^"]+)"\/>/) || [])[1] || '',
            keepNext: /<w:keepNext\/>/.test(pPr),
            keepLines: /<w:keepLines\/>/.test(pPr),
            pageBreakBefore: /<w:pageBreakBefore\/>/.test(pPr),
            sectPr: (pPr.match(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/) || [''])[0],
            pageBreakRun: /<w:br w:type="page"\/>/.test(x),
            nSeq: (x.match(/<w:instrText[^>]*> SEQ /g) || []).length,
            directBold: runs.some((r) => /<w:t[ >]/.test(r) && /<w:b\/>/.test(r)),
            hasPic: /<w:drawing\b/.test(x),
        })
    }
    return out
}


//checkXml: 檢查標籤成對與 OOXML 子元素順序(pPr、rPr、sectPr、style), 回傳錯誤字串陣列
//  僅比對順序表內之元素; 標籤成對以堆疊比對開始與結束標籤, 並檢查標籤外之文字不含 < 與未成實體之 &
let order = {
    'w:pPr': 'pStyle keepNext keepLines pageBreakBefore framePr widowControl numPr suppressLineNumbers pBdr shd tabs suppressAutoHyphens kinsoku wordWrap overflowPunct topLinePunct autoSpaceDE autoSpaceDN bidi adjustRightInd snapToGrid spacing ind contextualSpacing mirrorIndents suppressOverlap jc textDirection textAlignment textboxTightWrap outlineLvl divId cnfStyle rPr sectPr pPrChange',
    'w:rPr': 'rStyle rFonts b bCs i iCs caps smallCaps strike dstrike outline shadow emboss imprint noProof snapToGrid vanish webHidden color spacing w kern position sz szCs highlight u effect bdr shd fitText vertAlign rtl cs em lang eastAsianLayout specVanish oMath',
    'w:sectPr': 'HF footnotePr endnotePr type pgSz pgMar paperSrc pgBorders lnNumType pgNumType cols formProt vAlign noEndnote titlePg textDirection bidi rtlGutter docGrid printerSettings sectPrChange',
    'w:style': 'name aliases basedOn next link autoRedefine hidden uiPriority semiHidden unhideWhenUsed qFormat locked personal personalCompose personalReply rsid pPr rPr tblPr trPr tcPr tblStylePr',
}
for (let k of Object.keys(order)) {
    order[k] = order[k].split(' ').map((v) => `w:${v}`)
}
function checkXml(xml) {
    let errs = []
    let stack = []
    let re = /<\?[\s\S]*?\?>|<!--[\s\S]*?-->|<(\/?)([A-Za-z_][\w.:-]*)((?:\s+[\w.:-]+\s*=\s*(?:"[^"]*"|'[^']*'))*)\s*(\/?)>/g
    let iLast = 0
    let m
    while ((m = re.exec(xml)) !== null) {
        let txt = xml.slice(iLast, m.index)
        if (/</.test(txt) || /&(?!(?:amp|lt|gt|quot|apos|#\d+|#x[0-9a-fA-F]+);)/.test(txt)) {
            errs.push(`invalid text near ${m.index}: ${txt.slice(0, 40)}`)
        }
        iLast = m.index + m[0].length
        if (m[2] === undefined) {
            continue //處理指示或註解
        }
        let name = m[2]
        if (m[1] === '/') {
            let top = stack.pop()
            if (!top || top.name !== name) {
                errs.push(`unbalanced </${name}> near ${m.index}`)
            }
            continue
        }
        let parent = stack[stack.length - 1]
        if (parent && order[parent.name]) {
            let nm = (parent.name === 'w:sectPr' && /^w:(header|footer)Reference$/.test(name)) ? 'w:HF' : name
            let idx = order[parent.name].indexOf(nm)
            if (idx >= 0) {
                if (idx < parent.last) {
                    errs.push(`<${name}> after <${parent.lastName}> in <${parent.name}>`)
                }
                parent.last = idx
                parent.lastName = name
            }
        }
        if (m[4] !== '/') {
            stack.push({ name, last: -1, lastName: '' })
        }
    }
    if (/</.test(xml.slice(iLast))) {
        errs.push('invalid text at the end')
    }
    if (stack.length > 0) {
        errs.push(`unclosed <${stack[stack.length - 1].name}>`)
    }
    return errs
}


//simulateWordToc: 模擬 Word 更新後之目錄(超連結＋定位字元＋PAGEREF 欄位), 並於目標段落插入 _Toc 書籤, 供不需 Word 之核對測試
//  tamper: 'text' 改目錄第 2 項文字、'page' 圖目錄第 2 項頁碼改 0、'anchor' 表目錄第 1 項連結改指第 2 項、'dropFig' 移除圖目錄欄位、'bodyFooter' 正文節加頁尾指定
//  endSeparate: 欄位結束置於獨立段落; emptyToc: 目錄無項目(Word 找不到項目時之情形)
function simulateWordToc(u8, info, opt = {}) {
    let { tamper = '', endSeparate = false, emptyToc = false } = opt
    let z = unzipSync(u8)
    let doc = strFromU8(z['word/document.xml'])
    let k = 0
    let addBm = (text) => {
        let name = `_Toc${100 + k}`
        let id = 900 + k
        k++
        let done = false
        doc = doc.replace(/<w:p\b[^>]*>[\s\S]*?<\/w:p>/g, (x) => {
            if (done || textOf(x).trim() !== text || /<w:bookmarkStart\b[^>]*_Toc/.test(x)) {
                return x
            }
            done = true
            return x.replace(/^(<w:p\b[^>]*>\s*(?:<w:pPr>[\s\S]*?<\/w:pPr>)?)/, `$1<w:bookmarkStart w:id="${id}" w:name="${name}"/>`).replace(/<\/w:p>$/, `<w:bookmarkEnd w:id="${id}"/></w:p>`)
        })
        return name
    }
    let entry = (text, anchor, page) => `<w:hyperlink w:anchor="${anchor}" w:history="1"><w:r><w:rPr><w:noProof/></w:rPr><w:t xml:space="preserve">${esc(text)}</w:t></w:r><w:r><w:rPr><w:noProof/><w:webHidden/></w:rPr><w:tab/></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGEREF ${anchor} \\h </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>${page}</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:hyperlink>`
    for (let kind of ['toc', 'fig', 'tab']) {
        if (!info.lists.includes(kind)) {
            continue
        }
        let items = info.expect[kind]
        let anchors = items.map((t) => addBm(t))
        let reInstr = kind === 'toc' ? 'TOC \\\\o [^<]*' : `TOC \\\\h \\\\z \\\\c &quot;${info.labels[kind]}&quot;[^<]*`
        let re = new RegExp(`<w:p>(<w:r><w:fldChar w:fldCharType="begin" w:dirty="true"/></w:r><w:r><w:instrText xml:space="preserve"> ${reInstr}</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>)<w:r><w:t>Update field to build the table</w:t></w:r>(<w:r><w:fldChar w:fldCharType="end"/></w:r>)</w:p>`)
        doc = doc.replace(re, (all, pre, end) => {
            if (tamper === 'dropFig' && kind === 'fig') {
                return '<w:p/>'
            }
            if (emptyToc && kind === 'toc') {
                return `<w:p>${pre}<w:r><w:t>No table of contents entries found.</w:t></w:r>${end}</w:p>`
            }
            let ps = items.map((t, i) => {
                let text = t
                let anchor = anchors[i]
                let page = String(i + 1)
                if (tamper === 'text' && kind === 'toc' && i === 1) {
                    text = `${t}X`
                }
                if (tamper === 'page' && kind === 'fig' && i === 1) {
                    page = '0'
                }
                if (tamper === 'anchor' && kind === 'tab' && i === 0) {
                    anchor = anchors[1]
                }
                return entry(text, anchor, page)
            })
            let x = ps.map((v, i) => `<w:p>${i === 0 ? pre : ''}${v}${(i === ps.length - 1 && !endSeparate) ? end : ''}</w:p>`).join('')
            return endSeparate ? `${x}<w:p>${end}</w:p>` : x
        })
    }
    if (tamper === 'bodyFooter') {
        doc = doc.replace(/(<w:sectPr\b[^>]*>)((?:(?!<\/w:sectPr>)[\s\S])*<\/w:sectPr>)(?=\s*<\/w:body>)/, '$1<w:footerReference w:type="default" r:id="rId8"/>$2')
    }
    return zipSync({ ...z, 'word/document.xml': strToU8(doc) })
}


let parts = { run, para, heading, body, pic, capFig, capTab, table, cover }


export { parts, buildDocx, getSectFinal, readPart, listParas, checkXml, simulateWordToc, textOf }
