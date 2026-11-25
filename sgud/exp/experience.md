# 累積經驗

本檔為執行工作時特有之累積經驗,留給後續之嘗試、執行與判斷,由agent自行編列、擴充與維護.過長時抽出舊者為`experience_<年月日>.md`,並於本檔註明何時抽出.

## Word 與測試(2026-10-29)

- **Word 為單一實例,不可由 mocha `--parallel` 之不同檔同時調用**:兩個檔同時轉檔時,後到者得 `com_error -2147418111 接收者已拒絕這個呼叫`.失敗後遺留 WINWORD(視窗標題為 `wmd2docx_<pid>_<ts>_<亂數>.html - Word`,持有 4 份文件、COM 一律回 -2147418111,graceful 關閉失敗)、`src/templates/~$mp.docx` 擁有者檔、系統暫存夾之中介 html、`test/_tmp/...` 鎖死(`EPERM`/`EBUSY`),之後所有 Word 測試連鎖失敗.處置:以視窗標題確認為自己之殘留(標題含自己 worker 之 pid 與時間戳)後 `taskkill /F /PID <pid> /T`,回驗 WINWORD/htmlToDocx/cscript 皆消失,再刪 `~$mp.docx`、暫存 html、`test/_tmp` 殘夾.預防:新增之 Word 長測試併入 `unit-cvMdToDocx.test.mjs` 同檔依序,不另開檔(R-TEST-1).
- `quiet-run --lock` 等鎖時完全沒有 log 檔(要等取得鎖才建檔);他套件長時間持鎖時一次等候可達 6 分鐘以上,故 Word 實驗應合併成單次取鎖,不要拆小.
- **稀疏檔測大小上限**:`fs.openSync` + `fs.ftruncateSync(h, n)` 產 2.2 GB 檔瞬間完成(0 ms)、不佔磁碟(`df` 不變),`fs.statSync().size` 即為 n,`fs.readFileSync` 即擲 `ERR_FS_FILE_TOO_LARGE`;`unlink` 亦瞬間.`test/tools/fakeHtml2docx.mjs` 之 `size` 選項(環境變數 `FAKE_HTML2DOCX_SIZE`)即以此產出假 docx,測 2 GiB 與 base64 字串上限皆不實寫.
- 假轉檔器(`runWithFakeHtml2docx`)另記錄所收之中介 html(`globalThis.fakeHtml2docxHtml`),可 `reject` 指定字串;假的 `WHtml2docx` 無 `ensureBins` 屬性,故用假轉檔器之子程序不可呼叫 `checkDocxReady`(health).
- `process.env` 為 fsEnsureBins 於呼叫時讀取(`W_BIN_CACHE`、`W_BIN_MIRROR`、`W_BIN_OFFLINE`),單元測試可暫改後還原;快取位置 `<W_BIN_CACHE>/w-html2docx/<sha256>/htmlToDocx.exe`,預設為 `node_modules/w-html2docx/bin-cache/<sha256>/htmlToDocx.exe`(無 pkg 層).`ensureBins({ offline: true, mirror: '' })` 不連網、不用 `W_BIN_MIRROR`,快取內校驗不符之檔會被刪除.

## 跨行程同時驅動 Word(2026-11-10,md407)

- 症狀(全庫連動跟跑,3/3 輪紅):`unit-cvMdToDocx`〈實轉〉before hook 得 `Failed to convert html to docx: code[1]:\nerror: AttributeError: Open.Range`(或逾時 180 s);`unit-addDocxToc`〈Word更新〉於 `assert.strict.equal(nTmpToc(), n0)` 失敗(實際比預期多 1,不是 `res.toc`/`res.tab`——斷言訊息 `3 !== 2` 為 actual/expected,易誤讀),系統暫存夾留下 `wmd2docx_toc_<pid>_…_in.docx`(其 pid 段可認出是哪個行程留的).
- **機轉(實測)**:w-html2docx 之 `htmlToDocx.exe` 以 win32com `Dispatch('Word.Application')` 取得 Word——先連執行中之實例;PowerShell 先開一個隱藏 Word(命令列 `/Automation -Embedding`)並開一份文件,再轉檔:轉檔期間全機只有該 WINWORD、其文件數 1→3→0,轉畢該 WINWORD 結束.兩行程同時起跑而 `addDocxToc` 晚約 1.5 s 時兩邊皆失敗(`update failed: 不可在值為 Null 的運算式上呼叫方法` 與 `Open.Range`),晚 0 s(轉檔尚在 md→html)或 3 s 時皆過,依序皆過;時間窗窄,全套中是否撞上取決於各檔起跑時序與負載.
- **分辨依賴升版之最小步驟**:`npm pack <pkg>@<舊>`、`@<新>` 解到暫存夾,`diff -rq` 比 `src/`、`srcPython/`、`bin/manifest.json`(exe 之 sha256)與 dist(只差版本橫幅即等價);再查本套件 import 的是 `src/` 還是 main(dist).本輪四個升版之自有套件 `src/` 皆逐檔相同;以副本(`git ls-files` 複製、`npm ci --min-release-age=0`、`npm i <pkg>@<舊> --no-save`)重跑同一並行實驗,新舊版失敗率相同,即與升版無關.
- 收尾 WINWORD 時以命令列含 `/Automation` 或 `-Embedding` 判定為自動化啟動者才結束(`Get-CimInstance Win32_Process` 之 `CommandLine`),不動使用者手動開啟之 Word.

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

## 模板與樣式名稱(2026-11-07,Gitea #96,規則 R-TPL-1)

- **繁體中文版 Word 會更名與內建樣式本地化名稱同名之自訂樣式**:「目錄標題」是內建樣式 `TOC Heading` 之本地化名稱.模板內自訂樣式取此名,經 w-html2docx(Word 開檔轉檔)後即成「目錄標題1」,故 `idOf('目錄標題')` 找不到、模板自有之樣式不被沿用,而套件另建之同名樣式於 `addDocxToc` 之 Word 另存時再被更名為「目錄標題2」.檔案內(`styles.xml`)的名稱不會預先出現衝突,**只有經 Word 轉檔後才看得到**,故預設樣式名稱須以真 Word 驗證.偵測法:比對模板全部樣式名稱與 Word 產出 docx 之樣式名稱,模板有而產出無者即被更名(`names-survive` 式比對,已併入 `unit-cvMdToDocx` 之真 Word 案);本模板 137 個樣式(86 個自訂)中僅「目錄標題」被更名,其餘 27 個報告樣式名皆保留.
- **更名樣式名稱的做法**:`styles.xml` 的 `<w:name w:val>` 先依對照表換名(字元連結樣式為「<名稱> 字元」),再把 `styles.xml`、`numbering.xml`、`document.xml` 內剩餘之舊前綴(皆為 id 與其參照:`styleId`、`basedOn`、`next`、`link`、`pStyle`)整批換為 `Rpt`;驗證是把對照表與前綴反向套回後與原檔逐字相同,其他部件逐位元組相同.經此作法不動任何排版屬性.
- **舊模板含業主與個人資訊之處**:不只樣式名稱,`word/_rels/settings.xml.rels` 有 `attachedTemplate` 外部連結(含本機路徑與計畫名)並由 `settings.xml` 以 `<w:attachedTemplate r:id>` 參照,`docProps/core.xml` 有作者與最後修改者,`docProps/app.xml` 有套用範本名稱.去識別時須整批清,只改樣式名稱不夠;驗收以「全部部件掃描業主字樣、個人名、本機路徑、外部連結」為準.
- **內建模板 + `toc: true` + 真 Word 會被拒**:w-html2docx 內建模板(1.0.48)含頁首頁尾,轉出之 docx 末節帶 `headerReference`/`footerReference`(even/default/first),`addDocxToc` 於 Word 更新後核對時回 `TOC check failed (1): the cover section should have no footer`.用 git HEAD 之 src 複本(置於 repo 之 `./tmp/` 下,bare import 由 repo 根 `node_modules` 解析,不需 junction)實測相同,與 `titleStyle` 無關,屬修改前即有之限制.真 Word 之 toc 案故用 `temp.docx` 或其衍生模板,內建模板只在 prep 層(無 Word)測.
- **衍生「舊名樣式」模板的測法**:不附任何帶舊名之模板檔,測試內以 fflate 取 `temp.docx`、把 `styles.xml` 的 `<w:name w:val="報告目錄標題"/>` 換為其他名稱再寫到 `test/_tmp/<案名>/`,即得「只有舊名樣式」之模板;`prep` 層與真 Word 層皆可用.
- **以 git HEAD 之 src 複本做「修改前」對照**:`git archive HEAD src | tar -x -C ./tmp/<夾>`,複本內之 `import` 由 `./tmp/<夾>` 往上找到 repo 根之 `node_modules`,探測腳本以環境變數指定 `SRC`;不用 `git stash`/`checkout`.

## 轉檔器 htmlToDocx.exe 之集中取得與跨程序只取得一次(2026-11-25,R-TEST-2)

- **盤查所列之站點須逐檔核實**:外部盤查稱 `unit-cvMdTo`、`unit-cvMdToDocxImg`、`api-rmApiServer` 真轉檔時隱式取得轉檔器,實查三者皆經 `runWithFakeHtml2docx`(子行程換假轉檔器)或 `out=html`,不取得.真隱式取得者只有 `unit-cvMdToDocx`(〈實轉〉、〈圖片內嵌於docx〉、〈轉檔器取得失敗…〉第2案之子行程);`unit-utils`〈checkDocxReady〉於載入時判斷預設快取有無轉檔器,故隨他檔之取得時序,2 個案略過(264 passing、3 pending)或執行(266 passing、1 pending).
- **修前實測**(wsemi 1.9.74,清空 `node_modules/w-html2docx/bin-cache`,平行全套,以 `NODE_OPTIONS=--import` 預載包裝 `globalThis.fetch` 計對 `raw.githubusercontent.com` 之取檔次數):原佈局只下載 1 次(單一 worker),但 `unit-utils` 之 2 案略過.把新佈局(三個檔各於 `before` 取得)配舊 wsemi,同法實測下載 2 次(兩個 worker 相隔約 1.3 秒),即無鎖時多個消費檔同時冷啟動必重複下載;3 行程同取本機鏡像之案,舊 wsemi 鏡像被取 3 次,1.9.75 取 1 次且其餘 2 行程各有等候事件.
- **w-html2docx 之 dist 不帶鎖(1.0.72;1.0.73 起已帶,本套件仍經 src)**:1.0.72 之 `dist/w-html2docx.umd.js` 內嵌建置時之 wsemi,搜尋 `.lock-`、`lock-wait`、`E_LOCK_TIMEOUT` 皆 0 處(1.0.73 之 dist 已含:壓縮檔中 `lock-wait` 出現 3 次、`E_LOCK_TIMEOUT` 出現 1 次,2026-11-25 實查);其 `src/ensureBins.mjs` 引用 `wsemi/src/fsEnsureBins.mjs`,由使用端 `node_modules/wsemi` 解析(w-html2docx 無巢狀 `node_modules`),故經 `src` 取得者隨使用端升級 wsemi 而帶鎖.產品碼 `src/cvMdToDocx.mjs` 與 `binEnsure.mjs` 皆引 `w-html2docx/src/WHtml2docx.mjs`;`unit-binEnsure` 以版本、同一份 wsemi(realpath)、import 陳述(行首錨定)、鎖逾時、多行程計數五處守住,工具改引 dist 之變異使 3 案轉紅.獨立複審發現:若鎖案與多行程案不經工具自身匯出之 `getBin`(測試檔與子行程各自 import `WHtml2docx`),改引 dist 之變異五案全綠——故測試與子行程皆經 `getBin`.
- **取得單位之預算**:wsemi 之 `lockBudget`(`fsEnsureBinsCore.mjs`)= `n × totalMs + min(totalMs, 60000)`,單一檔無鏡像 n=1(預設 totalMs 600000 → 660000 毫秒),有鏡像 n=2;鎖之持有上限與等候上限同值,等候者每換一個持有者重新起算.此算式原未匯出,測試端曾依原始碼複寫;wsemi 1.9.76 起匯出 `fsEnsureBinsBudget`,已改呼叫之、複寫已移除(見下條).等候案以 `totalMs=1000`(預算 2000 毫秒,實測 2.0~3.3 秒:另含輪詢間隔至多約 1.5 秒)與有鏡像之 `totalMs=2000`(預算 6000 毫秒,無鏡像為 4000,故下限 5900 可分辨)對齊.等候者之 `totalMs` 不得小於持鎖者(w-html2docx 隱式取得用預設 600000).
- **改呼叫 `fsEnsureBinsBudget` 與並行取兩倍(2026-11-25,t633e)**:`binBudgetMs(opt)` 以 w-html2docx 之 `bin/manifest.json`(相對 `pkgRoot`)、`names`、`pkgRoot`(`require.resolve('w-html2docx/package.json')` 之夾)呼叫 wsemi,選項與 `getBin` 同一組;`pkgRoot` 為必填(清單給物件時亦然).`BIN_ENSURE_TIMEOUT_MS` = 2 ×預算+60000(並行時等候者最壞先等他方一個預算、他方失敗才自行取得再一個預算,規格 E5-3).實測變異:並行倍數改 1、`totalMs` 被忽略、`mirror` 被忽略、時限不隨 `W_BIN_MIRROR`、複寫算式含 `Math.min(` 皆使新測試轉紅(依序 1、3、2、1、1 案);**忠實複寫算式(無 `Math.min`)而留著不用之匯入**行為層測試抓不到——單一檔清單下複寫與 wsemi 輸出逐位相同,等價於行為層;「移除 `fsEnsureBinsBudget` 之 import」由〈工具不自寫算式〉之 import 行首錨定攔,「import 留著不用」只由 eslint `no-unused-vars` 攔(實測 1 錯誤),「import 留著並以 `void` 等帶過」兩者皆攔不到(殘餘風險,屬審查);`names`、`offline` 對此清單(單一檔、`pinned`)之預算無影響,屬等價變異.以舊測試(改前之斷言 `binBudgetMs(true)`=1260000、時限=預算+60000)對新工具必紅,屬行為有意之改變.
- **不連外網之鎖案作法**:①他方之鎖:在空快取夾內建 `.lock-<鍵>/owner.json`(鍵為 `'f/<sha256>/<name>'`〔Windows 轉小寫〕之 SHA-256 前 32 碼;持有者為本行程、持有上限 1 小時),`offline:true` 且無鏡像時 wsemi 只等候不建鎖,逾預算得 `E_LOCK_TIMEOUT`(帶 `holderPid`);②多行程:本機 `http` 鏡像(`listen(0)`;路徑 `/<pkg>/<sha256>/<name>`,內容為已取得之轉檔器)計次,第 1 個請求之回應扣住,至子行程回報 2 個已進入等候(`onProgress` 之 `wait` 事件,經 stdout 一行 JSON)或出現第 2 個請求才放行,以事件同步不以延時;子行程輸出須等 `stdout.write` 之 callback 再 `exit`.
- **證明新測試有鑑別力之副本根**:`./tmp/<夾>/` 下自建 `node_modules/{w-html2docx,wsemi@舊版}` 與 `test/` 副本,在該夾為工作目錄跑(最近之 `node_modules` 優先,故巢狀舊 wsemi 勝出;`unit-binEnsure` 之版本案可印證解析結果);變異用 Edit 改副本,不動 repo.
- **離線空快取找隱式取得者**:`W_BIN_CACHE=<空夾>`、`W_BIN_OFFLINE=1` 跑全套(R-TEST-2 之盤點;2026-11-25 實測 1 分鐘):不容忍之 `before` 因 `E_OFFLINE_MISSING` 失敗(3 處)、容忍型 `before` 印出 `could not be obtained`(2 行)、起子程序之案以 `[E_OFFLINE_MISSING]` 略過(1 案),三種訊號合計即隱式需要轉檔器者,與站點全集一致.
- **產品面之連帶**:wsemi ≥ 1.9.75 起,離線且無鏡像之取得遇他方持鎖會等候(`checkDocxReady` 之 `exeFound` 診斷在他方下載期間會多等,至多一個預算),等候結束後讀快取.

## 其他

- `package.json` 經 `npm i` 後檔尾多一個換行(原無),版本欄位不變.
- git 工作副本為 `autocrlf`,`git diff` 提示 `LF will be replaced by CRLF` 屬常態;收尾以 `git ls-files --eol` 確認行尾未被改動.
- **`git stash push/pop` 會把還原之檔由 LF 改成 CRLF**(2026-10-10 做「新測試在舊碼會紅」驗證時,`src/rmApiServer.mjs` 還原後 `git ls-files --eol` 由 `w/lf` 變 `w/crlf`,其餘檔不變;`git diff --stat` 因行尾正規化看不出,易夾帶進 commit).故先紅後綠一律用 `git show HEAD:<檔>` 或 `git archive`(見〈模板與樣式名稱〉末條)輸出到 `./tmp/` 另行載入,不動工作檔;已動到者以 node 把 `\r\n` 轉回 `\n`,收尾以 `git ls-files --eol <檔>` 確認 `w/lf`.
- **Word 相依之期望值須以真 Word 實跑核對**(2026-10-10):本機有 Word 16 之前(2026-09-01～10-10)寫下之 Word 相依測試期望值與 README/JSDoc 範例輸出皆未實跑過;補實跑即發現錯誤,例如 toc 測試曾寫死 `\o "1-2"`,實際 Word 匯入 html 時 h1/h2 為段落直設之大綱階層 2/3(非 1、2),改為自 docx 讀大綱階層(`test/unit-cvMdToDocx.test.mjs` 該案註解).Word 相依之主張(版面、目錄、頁碼、欄位 XML)直接以真 Word 實測,不只靠讀碼推論;不需 Word 之接線/失敗路徑用 `test/tools/fakeHtml2docx.mjs` 與 `test/tools/docxFixture.mjs`.
