import { execFileSync } from 'child_process'


//wordProcs: 列出當前 WINWORD.EXE 之 PID(字串陣列); 僅 Windows
//why: addDocxToc 以 PowerShell 經 COM 驅動 Word, 收尾(Quit)失敗會被腳本吞掉而使 WINWORD 殘留, 測試須檢查執行後無本次新啟動之 WINWORD
function wordPids() {
    if (process.platform !== 'win32') {
        return []
    }
    let out = execFileSync('tasklist', ['/FI', 'IMAGENAME eq WINWORD.EXE', '/FO', 'CSV', '/NH'], { encoding: 'latin1', windowsHide: true })
    return out.split(/\r?\n/).filter((l) => /^"WINWORD\.EXE"/i.test(l)).map((l) => l.split('","')[1])
}


//waitWordGone: 等候 before 之外之 WINWORD 結束(Word 於 Quit 後需短暫時間離場), 回傳逾時仍存在之新 PID
async function waitWordGone(before, timeoutMs = 15000) {
    let t0 = Date.now()
    let left = []
    while (true) {
        left = wordPids().filter((p) => !before.includes(p))
        if (left.length === 0 || Date.now() - t0 > timeoutMs) {
            return left
        }
        await new Promise((resolve) => setTimeout(resolve, 300))
    }
}


export { wordPids, waitWordGone }
