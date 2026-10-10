import path from 'path'
import { execFileSync } from 'child_process'


//runWithFakeHtml2docx: 以子程序執行 script, 子程序內 'w-html2docx/src/WHtml2docx.mjs' 經模組解析掛鉤換成假函數,
//  假函數將環境變數 FAKE_HTML2DOCX_DOCX 所指之 docx 複製至輸出位置, 使 md -> docx 不需 Word, 可測目錄之接線、略過與失敗清理;
//  另將收到之設定物件(第 3 參數)之複本依序推入 globalThis.fakeHtml2docxOpts, 供 script 檢查傳予 w-html2docx 之設定
//why: ESM 之 import 無法於同一程序內替換, 且以子程序執行使 cwd 與環境變數不影響 mocha 程序
//  script 以 '@@RESULT@@' + JSON 輸出結果; 回傳解析後之物件
function runWithFakeHtml2docx(script, opt = {}) {
    let fakeCode = `import fs from 'node:fs'; export default async function(fpIn, fpOut, opt) { (globalThis.fakeHtml2docxOpts ||= []).push(structuredClone(opt)); fs.copyFileSync(process.env.FAKE_HTML2DOCX_DOCX, fpOut); return 'ok' }`
    let fakeUrl = `data:text/javascript,${encodeURIComponent(fakeCode)}`
    let hooksCode = `export async function resolve(s, c, n) { if (s === 'w-html2docx/src/WHtml2docx.mjs') { return { url: ${JSON.stringify(fakeUrl)}, shortCircuit: true } } return n(s, c) }`
    let registerCode = `import { register } from 'node:module'; register(${JSON.stringify(`data:text/javascript,${encodeURIComponent(hooksCode)}`)})`
    let out = execFileSync(process.execPath, ['--import', `data:text/javascript,${encodeURIComponent(registerCode)}`, '--input-type=module', '-e', script], {
        cwd: path.resolve(opt.cwd || '.'),
        env: { ...process.env, ...(opt.env || {}), FAKE_HTML2DOCX_DOCX: path.resolve(opt.fpDocx) },
        encoding: 'utf8',
        windowsHide: true,
        timeout: 170000,
    })
    let i = out.lastIndexOf('@@RESULT@@')
    if (i < 0) {
        throw new Error(`no result from the child process: ${out}`)
    }
    return JSON.parse(out.slice(i + '@@RESULT@@'.length))
}


export default runWithFakeHtml2docx
