import cvMdToDocx from './cvMdToDocx.mjs'
import cvMdTo from './cvMdTo.mjs'
import ApiServer from './ApiServer.mjs'
import ApiClient from './ApiClient.mjs'


/**
 * Markdown轉Docx工具集
 *
 * 匯整四個入口：cvMdToDocx(md檔轉docx檔)、cvMdTo(md內容轉html/docx之base64)、ApiServer(hapi轉檔服務)、ApiClient(呼叫轉檔服務之用戶端)。
 *
 * @example
 *
 * import WMd2docx from 'w-md2docx/src/WMd2docx.mjs'
 *
 * let r = await WMd2docx.cvMdToDocx('./test/report.md', './test/report.docx')
 * console.log(r)
 * // => { fpOutDocx: '…', sizeDocx: 39856, sizeHtml: 12345, ms: 8342 }
 *
 */
let WMd2docx = {
    cvMdToDocx,
    cvMdTo,
    ApiServer,
    ApiClient,
}


export default WMd2docx
