# 累積經驗

本檔為執行工作時特有之累積經驗,留給後續之嘗試、執行與判斷,由agent自行編列、擴充與維護.過長時抽出舊者為`experience_<年月日>.md`,並於本檔註明何時抽出.

## Word 與測試(2026-10-29)

- **Word 為單一實例,不可由 mocha `--parallel` 之不同檔同時調用**:兩個檔同時轉檔時,後到者得 `com_error -2147418111 接收者已拒絕這個呼叫`.失敗後遺留 WINWORD(視窗標題為 `wmd2docx_<pid>_<ts>_<亂數>.html - Word`,持有 4 份文件、COM 一律回 -2147418111,graceful 關閉失敗)、`src/templates/~$mp_tpc.docx` 擁有者檔、系統暫存夾之中介 html、`test/_tmp/...` 鎖死(`EPERM`/`EBUSY`),之後所有 Word 測試連鎖失敗.處置:以視窗標題確認為自己之殘留(標題含自己 worker 之 pid 與時間戳)後 `taskkill /F /PID <pid> /T`,回驗 WINWORD/htmlToDocx/cscript 皆消失,再刪 `~$mp_tpc.docx`、暫存 html、`test/_tmp` 殘夾.預防:新增之 Word 長測試併入 `unit-cvMdToDocx.test.mjs` 同檔依序,不另開檔(R-TEST-1).
- `quiet-run --lock` 等鎖時完全沒有 log 檔(要等取得鎖才建檔);他套件長時間持鎖時一次等候可達 6 分鐘以上,故 Word 實驗應合併成單次取鎖,不要拆小.
- **稀疏檔測大小上限**:`fs.openSync` + `fs.ftruncateSync(h, n)` 產 2.2 GB 檔瞬間完成(0 ms)、不佔磁碟(`df` 不變),`fs.statSync().size` 即為 n,`fs.readFileSync` 即擲 `ERR_FS_FILE_TOO_LARGE`;`unlink` 亦瞬間.`test/tools/fakeHtml2docx.mjs` 之 `size` 選項(環境變數 `FAKE_HTML2DOCX_SIZE`)即以此產出假 docx,測 2 GiB 與 base64 字串上限皆不實寫.
- 假轉檔器(`runWithFakeHtml2docx`)另記錄所收之中介 html(`globalThis.fakeHtml2docxHtml`),可 `reject` 指定字串;假的 `WHtml2docx` 無 `ensureBins` 屬性,故用假轉檔器之子程序不可呼叫 `checkDocxReady`(health).
- `process.env` 為 fsEnsureBins 於呼叫時讀取(`W_BIN_CACHE`、`W_BIN_MIRROR`、`W_BIN_OFFLINE`),單元測試可暫改後還原;快取位置 `<W_BIN_CACHE>/w-html2docx/<sha256>/htmlToDocx.exe`,預設為 `node_modules/w-html2docx/bin-cache/<sha256>/htmlToDocx.exe`(無 pkg 層).`ensureBins({ offline: true, mirror: '' })` 不連網、不用 `W_BIN_MIRROR`,快取內校驗不符之檔會被刪除.

## 防毒軟體與 Word 驅動腳本(2026-11-02,Gitea #53 建議 3:vbs 改 PowerShell,腳本不落地)

- 緣由:Microsoft Defender 即時保護以機器學習把獨立之 `updateDocxToc.vbs` 判為 `Trojan:Script/ObfusScript.A!ml`(誤報;推測成因為「腳本驅動 Word 開檔另存＋`On Error Resume Next`」近似巨集型惡意程式),任何程式一讀取即隔離,使用端只見 `updateDocxToc.vbs is not found`.處置:改以 Windows 內建 PowerShell 5.1 經 COM 驅動 Word,腳本以 `-EncodedCommand` 帶於命令列、不寫成檔案,並刪除該 vbs(規則 R-AV-1).實測本機(Defender 即時保護與行為監控開啟、無 ASR 規則)跑 PowerShell 版數十次,`Get-MpThreatDetection` 恆為 0,新增/改動之原始碼檔 `MpCmdRun -Scan -ScanType 3 -File` 皆 `found no threats`.**此為本機實測,不保證企業端點(EDR、ASR 規則「阻擋可能經過模糊處理的指令碼」、AppLocker/WDAC 受限語言模式)亦放行**;受限語言模式下 `New-Object -ComObject` 被禁(模擬實測),無解.
- **PowerShell 5.1 + COM 實測要點**:
  - `-EncodedCommand <b64>` 之後**不可再接任何引數**(`無法處理命令,因為已經以 -Command 或 -EncodedCommand 指定了命令`,離開碼 0xFFFD0000),`$args` 不可用,故路徑只能走環境變數;`execProcess` 無 `env` 選項,於呼叫期間暫設 `process.env`、**落定後**才還原(`codeCmd: 'auto'` 時子程序要等系統字碼頁查詢完才啟動,呼叫一返回即還原會讓子程序讀不到).含空白、中文、`"&^%!;,='`$` 之路徑經環境變數原樣到達.
  - `-ExecutionPolicy` 不影響 `-EncodedCommand`(執行原則只管腳本檔),不需也不要加 `-ExecutionPolicy Bypass`(是攻擊腳本之典型旗標,徒增風險);`-NoProfile -NonInteractive` 足矣;視窗已由 `execProcess` 的 `windowsHide` 隱藏.
  - **stderr 被導向時 `-EncodedCommand` 會先輸出 CLIXML**(`#< CLIXML` 加「正在準備模組以便第一次使用」之進度物件),使離開碼 0 之輸出也被污染;腳本**第一行**設 `$ProgressPreference = 'SilentlyContinue'` 即消失.未捕捉之例外仍以 CLIXML 的 `<S S="Error">` 出現,`fromClixml` 還原為文字.
  - 輸出編碼:腳本設 `[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)`(無 BOM),stdout 為 UTF-8,**stderr 仍為系統字碼頁(cp950)**;`execProcess` 的 `codeCmd: 'auto'` 對兩串流各自判定,中文原因皆不亂碼.錯誤原因取最內層例外之 `Message`(Word 的 COMException 說明,與 vbs 的 `Err.Description` 同文),換行折為空格.
  - `exit N` 於 `-EncodedCommand` 即行程離開碼;`exit` 於 try 內仍會執行 finally.收尾用 `do { … break … } while ($false)` 單趟迴圈,`break` 可自 catch 跳出.
  - **`$word.Quit(0)` 在 Word 16 + PS 5.1 擲「引數應該是 PSReference。請使用 [ref]。」**(`$doc.Close(0)` 卻可),被收尾之 `catch {}` 吞掉 → **每次 WINWORD 殘留一個(約 130 MB)**;改 `Quit([ref]0)`、`Close([ref]0)` 後零殘留.教訓:吞錯之收尾必須有「執行後無殘留行程」之測試(`test/tools/wordProcs.mjs` 的 `waitWordGone`),不能只看輸出與離開碼.
  - 受保護(唯讀)文件使 `Repaginate` 擲「該文件已編輯鎖定狀態中」,PS 與 vbs 皆為 `error: update failed:`,可作階段字頭測試之夾具(settings.xml 加 `<w:documentProtection w:edit="readOnly" w:enforcement="1"/>`).
  - 逾時:`execProcess` 結束 `powershell.exe` 之樹,COM 啟動之 WINWORD 不在其中而殘留(700 ms 逾時實測 1 個殘留),與 cscript 作法相同;須人工結束.
  - 耗時:Word 更新同一份文件各 5 次(交錯)中位數兩輪方向相反——cscript 2523 ms 對 PowerShell 1852 ms、cscript 2335 ms 對 PowerShell 2814 ms,無一致差異;PowerShell 第一次(冷啟動)較慢可達 4~5 s;不連 Word 的用法錯誤路徑 55 ms 對 265 ms.
- 舊 vbs 之作業守則(不讀、不複製、不雜湊該檔,Grep 加 `!**/*.vbs`)隨該檔刪除而不再適用;`test/tools/figurePages.vbs`(量測圖名頁碼)是另一個獨立 .vbs,仍以 cscript 執行,未動,見規則帳本〈已知限制〉.

## 圖片與 Word 行為(2026-10-29,Word 16 實測)

- w-md2html 1.1.12 路徑模式:md 語法圖之路徑含空白須用 `![x](<圖 資料夾/a.png>)`(CommonMark 角括號寫法),否則整段當文字;`%XX` 不解碼(資產路徑 `pics/%E5%9C%96.png` 以原字串在工作夾建檔,Word 以字面 `%` 之反斜線絕對路徑可解析);`&`、`%23` 亦可;缺圖訊息為 `N image file(s) do not exist:` + 逐行絕對路徑(順序不為契約,去重).
- w-html2docx 1.0.40:`<img src="http://127.0.0.1:1/x.png">`(不可達網址)約 19 秒後 reject `unresolved linked picture: http://127.0.0.1:1/x.png`;內容非圖片但副檔名為 `.png` 之檔 Word 不報錯(docx 照產出).
- **顯示尺寸(小於版心者)**:帶 96 dpi `pHYs` 之 PNG 與過去之 base64 方案相同;不帶 `pHYs` 之 PNG(例 `genTallPng` 產者)路徑方案依 96 dpi(40 px → 30 pt),過去之 base64 方案經 sharp 重編碼後為每像素 1 點(40 px → 39.75 pt);大於版心者兩方案皆縮至版心寬.帶 72/150 dpi 之 PNG 兩方案相同(差 ≤ 0.25 pt).
- 舊的 `transErrAsset` 解析 `fp[…] does not exist`,w-md2html 1.1.12 起不再產出該格式(缺圖改為 `N image file(s) do not exist:`),升版後相關測試立即紅(rmApiClient 之 `allowMissingAssets=true` 放行亦不再可能).

## 大量圖片實測(2026-10-29,不入測試套件)

- 720 張 1600×900 雜訊 PNG(上方 112 列隨機位元組,每張 545,189 B,合計 392,535,821 B),`cvMdToDocx(md, docx, { fpInTemp, keepCaption: true })`,**預設堆積**(4288 MB):成功,417.1 s(含 keepCaption 4.2 s;md→html 未單獨量測,審計 E3a 為 1.4 s;扣除後 Word 階段約 412 s,每張約 0.57 s);中介 html 169,283 B;docx 392,757,546 B;720 個 drawing 皆 `r:embed`、`word/media` 720、`r:link` 0、External 0,依序之媒體位元組皆等於來源圖;顯示尺寸皆 5274310×2966799 EMU;node 峰值 RSS 約 1307 MB(取樣)、`resourceUsage().maxRSS` 1681 MB(含 keepCaption 之 `unzipSync`/`zipSync`),WINWORD 峰值 534 MB.修前(審計):409 MB 直接 `Invalid string length`、154 MB 預設堆積 OOM 中止(離開碼 134).
- 腳本在 `EV/scripts/`(`gen-imgs.mjs`、`mk-md.mjs`、`bulk-run.mjs`、`bulk-verify.mjs`),實測用完即刪大圖與 docx;實測時以 `src/` 快照執行,實測期間另跑變異測試於複本.
- 改前改後(同輪,`cmp-report.mjs`:改前=git HEAD 之 src 配升級後依賴,改後=現 src):`test/report.md`(含 `cocktail.svg`)圖數 1/1、顯示 415.30×415.30 pt 相同,docx 100,916 → 200,787 B(svg + Word 產 2500 px 預覽 png,取代 1067 px png,**檔案變大**);混合圖 5/5,除不帶 pHYs 之小 PNG(39.75 → 30.00 pt)外顯示尺寸差 ≤ 0.5 pt.

## 其他

- `package.json` 經 `npm i` 後檔尾多一個換行(原無),版本欄位不變.
- git 工作副本為 `autocrlf`,`git diff` 提示 `LF will be replaced by CRLF` 屬常態;收尾以 `git ls-files --eol` 確認行尾未被改動.
