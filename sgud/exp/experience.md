# 累積經驗

本檔為執行工作時特有之累積經驗,留給後續之嘗試、執行與判斷,由agent自行編列、擴充與維護.過長時抽出舊者為`experience_<年月日>.md`,並於本檔註明何時抽出.

## Word 與測試(2026-10-29)

- **Word 為單一實例,不可由 mocha `--parallel` 之不同檔同時調用**:兩個檔同時轉檔時,後到者得 `com_error -2147418111 接收者已拒絕這個呼叫`.失敗後遺留 WINWORD(視窗標題為 `wmd2docx_<pid>_<ts>_<亂數>.html - Word`,持有 4 份文件、COM 一律回 -2147418111,graceful 關閉失敗)、`src/templates/~$mp_tpc.docx` 擁有者檔、系統暫存夾之中介 html、`test/_tmp/...` 鎖死(`EPERM`/`EBUSY`),之後所有 Word 測試連鎖失敗.處置:以視窗標題確認為自己之殘留(標題含自己 worker 之 pid 與時間戳)後 `taskkill /F /PID <pid> /T`,回驗 WINWORD/htmlToDocx/cscript 皆消失,再刪 `~$mp_tpc.docx`、暫存 html、`test/_tmp` 殘夾.預防:新增之 Word 長測試併入 `unit-cvMdToDocx.test.mjs` 同檔依序,不另開檔(R-TEST-1).
- `quiet-run --lock` 等鎖時完全沒有 log 檔(要等取得鎖才建檔);他套件長時間持鎖時一次等候可達 6 分鐘以上,故 Word 實驗應合併成單次取鎖,不要拆小.
- **稀疏檔測大小上限**:`fs.openSync` + `fs.ftruncateSync(h, n)` 產 2.2 GB 檔瞬間完成(0 ms)、不佔磁碟(`df` 不變),`fs.statSync().size` 即為 n,`fs.readFileSync` 即擲 `ERR_FS_FILE_TOO_LARGE`;`unlink` 亦瞬間.`test/tools/fakeHtml2docx.mjs` 之 `size` 選項(環境變數 `FAKE_HTML2DOCX_SIZE`)即以此產出假 docx,測 2 GiB 與 base64 字串上限皆不實寫.
- 假轉檔器(`runWithFakeHtml2docx`)另記錄所收之中介 html(`globalThis.fakeHtml2docxHtml`),可 `reject` 指定字串;假的 `WHtml2docx` 無 `ensureBins` 屬性,故用假轉檔器之子程序不可呼叫 `checkDocxReady`(health).
- `process.env` 為 fsEnsureBins 於呼叫時讀取(`W_BIN_CACHE`、`W_BIN_MIRROR`、`W_BIN_OFFLINE`),單元測試可暫改後還原;快取位置 `<W_BIN_CACHE>/w-html2docx/<sha256>/htmlToDocx.exe`,預設為 `node_modules/w-html2docx/bin-cache/<sha256>/htmlToDocx.exe`(無 pkg 層).`ensureBins({ offline: true, mirror: '' })` 不連網、不用 `W_BIN_MIRROR`,快取內校驗不符之檔會被刪除.

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
