import { execFileSync } from 'child_process'


//hasWord: 本機是否可調用 Microsoft Word(Windows + Word.Application COM 已註冊)
//why: docx 階段與目錄更新須真 Word, 無 Word 之機器實轉區塊整段略過;
//     不以 htmlToDocx.exe 是否存在為條件——缺檔時 w-html2docx 於轉檔時自動下載, 以其為條件會使剛安裝之機器漏測實轉
function hasWord() {
    if (process.platform !== 'win32') {
        return false
    }
    try {
        execFileSync('reg', ['query', 'HKCR\\Word.Application'], { stdio: 'ignore', windowsHide: true })
        return true
    }
    catch (err) {
        return false
    }
}


export default hasWord
