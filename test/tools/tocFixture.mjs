import fs from 'fs'
import os from 'os'
import path from 'path'
import crypto from 'crypto'
import { parts } from './docxFixture.mjs'


//addDocxToc 之測試共用(unit-addDocxToc.test.mjs 與 unit-cvMdToDocx.test.mjs 之〈addDocxToc Word更新〉)
//why: 呼叫 Word 之測試併於 unit-cvMdToDocx.test.mjs 依序執行(規則帳本 R-TEST-1), 夾具與檢查函數於此只定義一次


let { heading, body, capTab, table } = parts


//sha: 檔案內容之 sha256, 供檢查原檔是否改變
let sha = (fp) => crypto.createHash('sha256').update(fs.readFileSync(fp)).digest('hex')


//nTmpToc: 系統暫存夾內 addDocxToc 之暫存檔(wmd2docx_toc_ 開頭)數
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


//scanPs1: 掃描(非遞迴)資料夾內之 .ps1, 供檢查腳本未落地; 套件目錄另遞迴(略過 node_modules、.git、dist、docs、tmp 與 test/_tmp)
let scanPs1 = (fd, deep) => {
    let rs = []
    for (let d of fs.readdirSync(fd, { withFileTypes: true })) {
        let fp = path.join(fd, d.name)
        if (d.isDirectory()) {
            if (deep && !['node_modules', '.git', 'dist', 'docs', 'tmp', '_tmp'].includes(d.name)) {
                rs.push(...scanPs1(fp, deep))
            }
        }
        else if (/\.ps1$/i.test(d.name) && !d.name.startsWith('__PSScriptPolicyTest_')) { //__PSScriptPolicyTest_*.ps1 為 PowerShell 啟動時自建之檢查檔, 瞬間即刪(並行之他檔測試亦會啟動 PowerShell), 非本套件所寫
            rs.push(fp)
        }
    }
    return rs.sort()
}


//snapPs1: 系統暫存夾(非遞迴)與套件目錄(遞迴)之 .ps1 清單快照
let snapPs1 = () => JSON.stringify([scanPs1(os.tmpdir(), false), scanPs1(path.resolve('.'), true)])


export { sha, nTmpToc, bodyWord, scanPs1, snapPs1 }
