import { zipSync, strToU8 } from 'fflate'
import isstr from 'wsemi/src/isstr.mjs'


//docxXml: docx 之 XML 字串處理共用工具(添加目錄 docxToc.mjs 與圖名表名同頁 docxKeep.mjs 共用)
//  段落以正則切分, 不解析為 DOM; 文字方塊內之段落為巢狀段落, 呼叫端須自行排除或拒絕


/**
 * 還原XML文字之實體
 *
 * @param {String} s 輸入XML文字字串
 * @returns {String} 回傳還原後字串
 */
function unesc(s) {
    return s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, '\'').replace(/&amp;/g, '&')
}


/**
 * 跳脫XML文字與屬性值之特殊字元
 *
 * @param {*} s 輸入資料，轉字串後跳脫
 * @returns {String} 回傳跳脫後字串
 */
function esc(s) {
    return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')
}


/**
 * 跳脫正則表達式之特殊字元
 *
 * @param {*} s 輸入資料，轉字串後跳脫
 * @returns {String} 回傳跳脫後字串
 */
function escRe(s) {
    return String(s).replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
}


/**
 * 取XML片段之文字(依序串接各w:t之內容)
 *
 * @param {String} x 輸入XML片段字串
 * @returns {String} 回傳文字字串
 */
function textOf(x) {
    return unesc((x.match(/<w:t(?: [^>]*)?>[^<]*<\/w:t>/g) || []).map((t) => t.replace(/<[^>]+>/g, '')).join(''))
}


/**
 * 產生w:t元素(前後有空白時加xml:space="preserve")
 *
 * @param {String} s 輸入文字字串
 * @returns {String} 回傳w:t元素字串
 */
function tElem(s) {
    return `<w:t${/^\s|\s$/.test(s) ? ' xml:space="preserve"' : ''}>${esc(s)}</w:t>`
}


/**
 * 取段落之段落屬性w:pPr(須為段落之第1個子元素)
 *
 * @param {String} p 輸入段落XML字串
 * @returns {String} 回傳w:pPr元素字串，無則回傳空字串
 */
function pPrOf(p) {
    return (p.match(/^<w:p\b[^>]*>\s*(<w:pPr>[\s\S]*?<\/w:pPr>)/) || [])[1] || ''
}


/**
 * 依序列出document.xml之段落與其是否位於表格內
 *
 * @param {String} doc 輸入document.xml字串
 * @returns {Array} 回傳段落陣列，元素為{start,end,xml,inTbl}，start與end為段落於doc之起訖位置
 */
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


/**
 * 建立段落改寫之登錄表
 *
 * 以段落起點為鍵登錄改寫函數(依登錄順序套用)與插於段落前之內容，apply時由後往前套用，使各段落之位置不受前段改寫影響。
 *
 * @returns {Object} 回傳{addEdit(p,fn),addPre(p,s),apply(doc)}，apply回傳改寫後之doc
 */
function createEdits() {
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
    let apply = (doc) => {
        for (let { p, fns, pre } of [...edits.values()].sort((a, b) => b.p.start - a.p.start)) {
            let x = p.xml
            for (let fn of fns) {
                x = fn(x)
            }
            doc = doc.slice(0, p.start) + pre.join('') + x + doc.slice(p.end)
        }
        return doc
    }
    return { addEdit, addPre, apply }
}


/**
 * 以unzipSync之結果寫出docx，可取代部分檔案之內容
 *
 * 檔案順序沿用原docx；圖檔已壓縮故不再壓，其餘以壓縮等級6寫出。
 *
 * @param {Object} z 輸入fflate之unzipSync回傳物件
 * @param {Object} [parts={}] 輸入欲取代之檔案物件，鍵為檔名、值為文字內容
 * @returns {Uint8Array} 回傳docx內容
 */
function writeDocx(z, parts = {}) {
    let out = {}
    for (let [k, v] of Object.entries(z)) {
        let d = isstr(parts[k]) ? strToU8(parts[k]) : v
        out[k] = /^word\/media\//.test(k) ? [d, { level: 0 }] : d
    }
    return zipSync(out, { level: 6 })
}


export { unesc, esc, escRe, textOf, tElem, pPrOf, scanParas, createEdits, writeDocx }
