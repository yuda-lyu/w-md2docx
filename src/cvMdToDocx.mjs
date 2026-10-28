import fs from 'fs'
import os from 'os'
import path from 'path'
import get from 'lodash-es/get.js'
import isestr from 'wsemi/src/isestr.mjs'
import isobj from 'wsemi/src/isobj.mjs'
import fsIsFile from 'wsemi/src/fsIsFile.mjs'
import WMd2html from 'w-md2html/src/WMd2html.mjs'
import WHtml2docx from 'w-html2docx/src/WHtml2docx.mjs'
import addDocxToc from './addDocxToc.mjs'
import { prepDocxKeep } from './docxKeep.mjs'
import { toErrText, retryBusy, runExclusive, assertReadable } from './utils.mjs'


/**
 * Markdown檔轉Docx檔
 *
 * 內部流程為md -> html -> docx。docx階段係調用本機Microsoft Word(win32com)，故僅能於已安裝Word之Windows執行，且同時間僅一份轉檔作業(內部佇列自動排隊)。
 *
 * 中介html一律以本機絕對路徑引用圖片(w-md2html之imgConvertToBase64: false，optMd2html不得覆寫)，不將圖片內嵌為base64，故圖片數量與總量不受字串長度及堆積上限限制；Word讀取圖檔後由w-html2docx斷開連結並內嵌於docx，docx不依賴原圖。png、jpg維持原格式與位元組，svg以向量(svg)加點陣預覽(png)內嵌。md引用之圖檔不存在時於md轉html階段reject並列出全部缺圖之絕對路徑；Word解析不到之圖(如網址取不到)則於html轉docx階段reject，訊息每行列出一個無法解析之來源(`unresolved linked picture: <來源>`)。Word處理每張圖約需0.6秒。
 *
 * 轉檔器htmlToDocx.exe之取得與校驗由w-html2docx負責：其不含於npm包，首次轉檔時下載並校驗後存入快取(位置與cwd無關，見w-html2docx README〈外部執行檔〉)；取不到時reject字串帶[E_OFFLINE_MISSING]等錯誤碼，本函數於其後附指引。預設模板則由w-html2docx以cwd查找，故未給fpInTemp時須於含node_modules/w-html2docx之專案根執行。
 *
 * w-html2docx(≥1.0.29)於轉檔器失敗(Word未安裝、模板無法開啟、另存失敗等)時以離開碼1 reject並帶轉檔器回報之原因，resolve前亦自行確認docx已產生；本函數將其訊息加上前綴原樣傳出，並再以產物實體檔之存在與大小把關一次，不會回報假成功。
 *
 * keepCaption與toc須將docx整份讀入記憶體，單檔上限為2 GiB-1位元組(fs.readFileSync)，超過時於讀檔前以指出原因之訊息reject並刪除本次產出之docx。
 *
 * @param {String} fpInMd 輸入來源Markdown檔位置字串
 * @param {String} fpOutDocx 輸入轉出Docx檔位置字串
 * @param {Object} [opt={}] 輸入設定物件，預設{}
 * @param {String} [opt.fpInTemp=''] 輸入Docx模板檔位置字串，未給則由w-html2docx使用其內建模板，預設''
 * @param {String} [opt.fpOutHtml=''] 輸入另存中介Html檔位置字串，此檔為供Word轉檔之中介檔，以本機絕對路徑引用圖片而不自足(圖檔移走即無法顯示)，要自足之Html請改用w-md2html或cvMdTo之out為'html'；未給則中介檔產於系統暫存夾並於轉檔後刪除，預設''
 * @param {Object} [opt.optMd2html={}] 輸入傳予w-md2html之設定物件(如imgWidthMax、fontSizeScale等)，其imgConvertToBase64恆為false，不採用給定值，預設{}
 * @param {Object} [opt.optHtml2docx={}] 輸入傳予w-html2docx之設定物件(如imgRatioWidthMax、fontFamilies等)，其imgHeightReserveLines未給時預設4：行格線模板之圖片下方預留4格，供圖區塊間之空段落與至多2行之圖名使用，使整頁高之圖片與其圖名可同頁，無行格線之模板不生效，給null停用(見w-html2docx)，預設{}
 * @param {Boolean|Object} [opt.keepCaption=false] 輸入是否使圖名、表名與其圖、表同頁(不依附目錄)：轉檔後圖名表名設段落內不分頁，並與其下之表格或圖片、或其上之圖片設為同頁(可越過空段落；上下皆有圖表時依同類圖名表名之多數位置，平手時圖名取上方、表名取下方，見docxKeep.mjs之planKeepNext)，於添加目錄之前執行；true用預設，物件可給{labels}(見docxKeep.mjs之prepDocxKeep，與toc之labels各自獨立)；處理失敗時刪除本次產出之docx並reject，預設false
 * @param {Boolean|Object} [opt.toc=false] 輸入是否添加章節目錄、圖目錄、表目錄並重編頁碼(封面無頁碼、目錄為大寫羅馬數字、正文自1起)，true用預設，物件為addDocxToc之設定(見addDocxToc.mjs)；添加失敗時刪除本次產出之docx並reject，預設false
 * @returns {Promise} 回傳Promise，resolve回傳結果物件{fpOutDocx,sizeDocx,sizeHtml,ms,[fpOutHtml],[keepCaption],[toc]}，keepCaption為{figs,tabs,keepNext,keepLines,changed,warns,ms}(圖名數、表名數、新增與下段同頁及段落內不分頁之段落數、是否改寫、提示、耗時)，toc為addDocxToc之結果(文件無標題時僅含skip與ms)，reject回傳錯誤訊息
 * @example
 *
 * import cvMdToDocx from 'w-md2docx/src/cvMdToDocx.mjs'
 *
 * let r = await cvMdToDocx('./test/report.md', './test/report.docx', {
 *     fpInTemp: './src/templates/temp_tpc.docx',
 *     optMd2html: {
 *         imgWidthMax: '500px',
 *     },
 *     toc: true,
 * })
 * console.log(r)
 * // => { fpOutDocx: '…', sizeDocx: 41210, sizeHtml: 12345, ms: 12342, toc: { skip: '', cover: true, levels: [2, 4], lists: ['toc'], toc: 26, … } }
 *
 */
async function cvMdToDocx(fpInMd, fpOutDocx, opt = {}) {

    let msStart = Date.now()

    //check fpInMd
    if (!isestr(fpInMd)) {
        return Promise.reject('fpInMd must be a non-empty string')
    }
    fpInMd = path.resolve(fpInMd)
    if (!fsIsFile(fpInMd)) {
        return Promise.reject(`fpInMd[${fpInMd}] does not exist`)
    }

    //check fpOutDocx
    if (!isestr(fpOutDocx)) {
        return Promise.reject('fpOutDocx must be a non-empty string')
    }
    fpOutDocx = path.resolve(fpOutDocx)

    //fpOutHtml (未指定則產於系統暫存夾, 轉檔後刪除)
    let fpOutHtml = get(opt, 'fpOutHtml', '')
    let bKeepHtml = isestr(fpOutHtml)
    if (bKeepHtml) {
        fpOutHtml = path.resolve(fpOutHtml)
    }
    else {
        //檔名唯一, 避免與被鎖定之同名殘檔衝突造成 EBUSY
        fpOutHtml = path.join(os.tmpdir(), `wmd2docx_${process.pid}_${Date.now()}_${Math.random().toString(36).slice(2, 8)}.html`)
    }

    //optMd2html
    let optMd2html = get(opt, 'optMd2html', {})
    if (!isobj(optMd2html)) {
        optMd2html = {}
    }

    //optHtml2docx (複製一份, 避免改動呼叫端物件)
    let optHtml2docx = get(opt, 'optHtml2docx', {})
    if (!isobj(optHtml2docx)) {
        optHtml2docx = {}
    }
    optHtml2docx = { ...optHtml2docx }

    //imgHeightReserveLines (未給時預設 4; 給 null 停用, 給其他值原樣交 w-html2docx 判斷, 其視無效值為不啟用)
    //why: 行格線模板(如 temp_tpc.docx)下 Word 將段落高進位為行距之整數倍, 原圖高上限(版心高之 0.937 倍)進位後擠掉其下之圖名,
    //     2026-10-10 Word 16 實測整頁高長圖 4 張與圖名皆不同頁, 預留 4 格後皆同頁; 4 = 圖區塊間之空段落 1 格 + 圖名至多 2 行(12pt 各 1 格) + 餘裕 1 格
    //     (段前段後距離、頁首頁尾高於邊界時每頁實際格數少於推算), 無行格線之模板不生效
    optHtml2docx.imgHeightReserveLines = get(optHtml2docx, 'imgHeightReserveLines', 4)

    //fpInTemp (未給則交由 w-html2docx 使用內建模板)
    let fpInTemp = get(opt, 'fpInTemp', '')
    if (isestr(fpInTemp)) {
        fpInTemp = path.resolve(fpInTemp)
        if (!fsIsFile(fpInTemp)) {
            return Promise.reject(`fpInTemp[${fpInTemp}] does not exist`)
        }
        optHtml2docx.fpInTemp = fpInTemp
    }

    try {

        //md -> html
        //note: w-md2html 係以 md 所在資料夾解析其內引用之相對路徑圖片; imgConvertToBase64 恆為 false(置於展開之後, 呼叫端不得覆寫),
        //      中介 html 以本機絕對路徑引用圖片, 由 Word 讀檔後經 w-html2docx 斷開連結內嵌, 故圖片量不受字串長度與堆積上限限制,
        //      但中介 html 不自足(圖檔移走即無法顯示), 且缺圖於此階段 reject
        fs.mkdirSync(path.dirname(fpOutHtml), { recursive: true })
        let errHtml = null
        await retryBusy(() => WMd2html(fpInMd, fpOutHtml, { ...optMd2html, imgConvertToBase64: false }))
            .catch((err) => {
                errHtml = toErrText(err)
            })
        if (errHtml !== null) {
            return Promise.reject(`Failed to convert md to html: ${errHtml}`)
        }

        //驗證 html(產物實體檔為準, 不憑回傳值)
        if (!fsIsFile(fpOutHtml) || fs.statSync(fpOutHtml).size === 0) {
            return Promise.reject('Failed to convert md to html: html was not generated or is empty')
        }
        let sizeHtml = fs.statSync(fpOutHtml).size

        //html -> docx(排隊逐一執行, 避免同時調用本機 Word)
        //note: 轉檔器之取得與校驗由 WHtml2docx 內部的 ensureBins 負責, 此處不得事前檢查,
        //      否則下層之取得機制永遠不會被執行
        fs.mkdirSync(path.dirname(fpOutDocx), { recursive: true })
        let errDocx = null
        await runExclusive(() => retryBusy(() => WHtml2docx(fpOutHtml, fpOutDocx, optHtml2docx)))
            .catch((err) => {
                errDocx = toErrText(err)
            })
        if (errDocx !== null) {
            //預設模板以 cwd 查找, 查找失敗之訊息為 can not find folder for tmp.docx, 保留原文並補上 cwd 與執行位置提示
            if (errDocx.indexOf('can not find folder for tmp.docx') >= 0) {
                errDocx = `${errDocx} (cwd=${path.resolve()}, w-html2docx resolves its default template from cwd, please run from the project root that contains node_modules/w-html2docx or give fpInTemp)`
            }
            //轉檔器取不到(離線缺檔、校驗不符、下載失敗等)之訊息以 [E_xxx] 錯誤碼開頭, 保留原文並附說明檔指引
            if (/^\[E_[A-Z0-9_]+\] /.test(errDocx)) {
                errDocx = `${errDocx} (the external executable htmlToDocx.exe of w-html2docx is not available, see the section "外部執行檔" of the README of w-html2docx for W_BIN_CACHE, W_BIN_MIRROR, W_BIN_OFFLINE and WHtml2docx.ensureBins())`
            }
            return Promise.reject(`Failed to convert html to docx: ${errDocx}`)
        }

        //驗證 docx(第二道防線)
        //note: w-html2docx ≥1.0.29 於轉檔器失敗時已自行 reject(離開碼1, 帶轉檔器原因), resolve 前亦確認 docx 已產生且非空,
        //      此處再以產物實體檔為準把關一次, 涵蓋其 resolve 後檔案又消失等非預期情形, 不憑回傳值
        if (!fsIsFile(fpOutDocx) || fs.statSync(fpOutDocx).size === 0) {
            return Promise.reject('docx was not generated: please make sure Microsoft Word is installed and no stale WINWORD process is locking the files')
        }
        let sizeDocx = fs.statSync(fpOutDocx).size

        //圖名表名與其圖表同頁(不依附目錄)
        //note: 於添加目錄之前獨立執行, 使目錄略過(文件無標題)或目錄之 keepWithObject 為 false 時仍套用; 兩者同開時目錄之同一處理遇已有之設定即原樣回傳
        let keepCaption = get(opt, 'keepCaption', false)
        let rKeep = null
        if (keepCaption === true || isobj(keepCaption)) {
            let msKeep = Date.now()
            let errKeep = null
            try {
                assertReadable(fpOutDocx, 'docx') //整份讀入記憶體, 超過 fs.readFileSync 上限時於讀檔前以明確訊息拋出
                let r = prepDocxKeep(new Uint8Array(fs.readFileSync(fpOutDocx)), isobj(keepCaption) ? keepCaption : {})
                if (r.info.changed) {
                    //以暫存檔改名取代, 寫入中斷時不留半份 docx
                    let fpTmp = `${fpOutDocx}.${process.pid}_${Date.now()}_${Math.random().toString(36).slice(2, 8)}.tmp`
                    try {
                        fs.writeFileSync(fpTmp, r.u8)
                        await retryBusy(() => fs.renameSync(fpTmp, fpOutDocx))
                    }
                    finally {
                        if (fsIsFile(fpTmp)) {
                            try {
                                fs.unlinkSync(fpTmp)
                            }
                            catch (err) {
                                //殘檔被鎖, 略過
                            }
                        }
                    }
                }
                rKeep = { ...r.info, ms: Date.now() - msKeep }
            }
            catch (err) {
                errKeep = toErrText(err)
            }
            if (errKeep !== null) {
                //why: 已產出之 docx 缺所要求之同頁設定, 留著會被以檔案存在判斷成敗之呼叫端當成成功, 故刪除後才 reject
                try {
                    fs.unlinkSync(fpOutDocx)
                }
                catch (err) {
                    //被鎖則略過, 仍以 reject 回報失敗
                }
                return Promise.reject(`Failed to keep the captions with their figures and tables: ${errKeep}`)
            }
            sizeDocx = fs.statSync(fpOutDocx).size
        }

        //添加目錄並重編頁碼
        //note: 須於轉檔佇列之外呼叫, addDocxToc 內部之 Word 更新另行排隊; 若於佇列內呼叫, 兩者互等而永不完成
        let toc = get(opt, 'toc', false)
        let rToc = null
        if (toc === true || isobj(toc)) {
            let errToc = null
            rToc = await addDocxToc(fpOutDocx, isobj(toc) ? toc : {})
                .catch((err) => {
                    errToc = toErrText(err)
                })
            if (errToc !== null) {
                //why: 已產出之 docx 缺所要求之目錄, 留著會被以檔案存在判斷成敗之呼叫端當成成功, 故刪除後才 reject
                try {
                    fs.unlinkSync(fpOutDocx)
                }
                catch (err) {
                    //被鎖則略過, 仍以 reject 回報失敗
                }
                return Promise.reject(`Failed to add the TOC: ${errToc}`)
            }
            sizeDocx = fs.statSync(fpOutDocx).size
        }

        let rt = {
            fpOutDocx,
            sizeDocx,
            sizeHtml,
            ms: Date.now() - msStart,
        }
        if (bKeepHtml) {
            rt.fpOutHtml = fpOutHtml
        }
        if (rKeep !== null) {
            rt.keepCaption = rKeep
        }
        if (rToc !== null) {
            rt.toc = rToc
        }

        return rt

    }
    finally {

        //刪除中介 html(使用者指定保留者不刪; 被鎖就略過, 留待系統暫存清理)
        if (!bKeepHtml && fsIsFile(fpOutHtml)) {
            try {
                fs.unlinkSync(fpOutHtml)
            }
            catch (err) {
                //殘檔被鎖, 略過
            }
        }

    }

}


export default cvMdToDocx
