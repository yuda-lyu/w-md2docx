# 規則帳本

本檔為套件設計與開發階段之規則帳本(藍圖),為要求程式碼須長成什麼樣,或由規格產出之測試(項目/說明/流程)有哪些,包含軸、單一擁有者、已知限制、破壞性變更紀錄,以及對應是哪些測試檔案或指令等,由agent自行編列、擴充與維護.當前已做到哪些記於`sgud/state/state.md`.過長時依類型拆為`rulebook_<類型>.md`,本檔只留索引.

## 一、規則(以規則為單位,非以缺陷站點為單位)

### R-IMG-1 轉 docx 之中介 html 恆以本機絕對路徑引用圖,不內嵌 base64

- 規則:`cvMdToDocx` 傳予 `WMd2html` 之 `imgConvertToBase64` 恆為 `false`,置於展開 `optMd2html` 之後,呼叫端不得覆寫;圖由 Word 讀檔,經 w-html2docx 斷開連結內嵌於 docx.`fpOutHtml` 為供 Word 之中介檔,不自足.
- 盤點:Grep `imgConvertToBase64` `src/`.
- 站點全集:`src/cvMdToDocx.mjs:120`(false,已套)、`src/cvMdTo.mjs:154`(true,回傳用,見 R-IMG-2,刻意相反).
- 判準:`WMd2html(` 之呼叫處恆明寫 `imgConvertToBase64`,且其值依該 html 之用途(供 Word → false;回傳給呼叫端 → true).
- 測試:`test/unit-cvMdToDocxImg.test.mjs`(中介 html 無 `data:`、`optMd2html` 給 true 不得覆寫、`fpOutHtml` 與交予轉檔器者相同)、`test/unit-cvMdToDocx.test.mjs`〈圖片內嵌於docx〉(真 Word:`word/media` 圖數、`r:embed`、無 `TargetMode="External"`、原檔位元組與順序、與 base64 方案同輪雙產之顯示尺寸).

### R-IMG-2 `cvMdTo` 回傳之 html 恆自足

- 規則:`out` 含 html 時以 `WMd2html`(`imgConvertToBase64: true`,不論 `optMd2html`)另行產製於轉 docx 之前;轉 docx 之中介 html 另存 `<name>.word.html`(除錯模式 `keepWork` 時保留).`both` 之 html 與 `html` 逐字相同.
- 站點:`src/cvMdTo.mjs:154`(html)、`src/cvMdTo.mjs` 之 `cvMdToDocx(...fpOutHtml: <name>.word.html)`.
- 測試:`test/unit-cvMdTo.test.mjs`(`out=html` 與 `both` 自足、與 html 同輸入逐字相同、Word 階段之中介 html 為絕對路徑、`both` 圖片過大不呼叫轉檔器).

### R-IMG-3 缺圖訊息單一出口

- 規則:w-md2html 缺圖 reject 之 `N image file(s) do not exist:` + 逐行絕對路徑,於 `cvMdTo`(內容 API,有工作夾)一律經 `transErrAsset` 轉為不含工作夾路徑之可行動訊息(單數 `the referenced file (x) does not exist, please make sure it is included in assets`、複數 `the referenced files (x, y) do not exist, please make sure they are included in assets`);缺圖路徑不全在工作夾內或非缺圖訊息則原樣;`cvMdToDocx`(檔案 API)原樣傳出絕對路徑.
- 站點:`src/cvMdTo.mjs` 之 html 分支與 docx 分支(皆 `transErrAsset`);`src/utils.mjs` `transErrAsset`.
- 測試:`test/unit-utils.test.mjs`〈transErrAsset〉、`test/unit-cvMdTo.test.mjs`(單/複數、docx/both/html 皆不呼叫轉檔器)、`test/unit-cvMdToDocxImg.test.mjs`(原樣絕對路徑)、`test/api-rmApiServer.test.mjs`(500 + 訊息).
- `rmApiClient.cvMdTo` 之 `allowMissingAssets` 已於本輪移除(轉檔端對缺圖一律 reject,放行無意義).

### R-LIM-1 整份讀入記憶體之 docx 於讀檔前先檢查 `assertReadable`

- 規則:`fs.readFileSync` 單檔上限 2 GiB-1 位元組(`ERR_FS_FILE_TOO_LARGE`,訊息不指出是哪個檔);須整份讀入之 docx 於讀檔前呼叫 `assertReadable(fp, what)`,超過時以指出原因之訊息 reject,行程不中止.
- 盤點:Grep `readFileSync(` `src/`.
- 站點全集與判準:
  - `src/cvMdToDocx.mjs:171`(keepCaption,已套)
  - `src/addDocxToc.mjs:148`(toc 輸入 docx,已套)、`:169`(Word 更新後之輸出 docx,已套;**未測**——需 Word 產出 > 2 GiB 之 docx)
  - `src/cvMdTo.mjs:214/224`(回傳用 docx/html):由 R-LIM-2 管(合計上限遠小於 2 GiB),不另套
  - `src/rmApiClient.mjs:171`(資產)、`:300`(模板)、`:251`(md 文字):用戶端請求端,資產合計受服務端 `maxMb`(預設 256 MB)限制;單檔 > 2 GiB 或請求 JSON > 字串上限仍為通用錯誤(**已知不修**,範圍外)
  - `docxKeep.mjs`、`docxToc.mjs`、`docxXml.mjs`:記憶體內處理,不讀檔(不適用)
- 測試:`test/unit-utils.test.mjs`〈assertReadable〉(邊界 2^31-1 通過、2^31 拋錯,稀疏檔)、`test/unit-cvMdToDocx.test.mjs`(keepCaption、toc 經假轉檔器產 2^31 稀疏檔)、`test/unit-addDocxToc.test.mjs`.

### R-LIM-2 回傳 base64 之合計於讀檔與轉 base64 前先檢查 `assertB64Fit`

- 規則:`cvMdTo`(及其上之 `POST /api/convert`)將 docx 與 html 以 base64 置於同一物件(服務端為單一 JSON 字串),受 V8 字串長度上限(`buffer.constants.MAX_STRING_LENGTH` = 536,870,888)限制;各檔 base64 字元數(每 3 位元組 4 字元,含進位)合計 > 上限 - 1 MiB(預留給 JSON 外殼與其他欄位)時,於讀檔前以指出原因之訊息 reject(約 401,866,734 位元組之檔案合計).
- 站點:`src/cvMdTo.mjs:207`;`src/utils.mjs` `assertB64Fit`.服務端不另檢查(由 cvMdTo reject → `500 {success:false,error}`).
- 測試:`test/unit-utils.test.mjs`〈assertB64Fit〉(邊界、合計)、`test/unit-cvMdTo.test.mjs`(docx 410,000,000 B 稀疏檔;docx 401,866,000 B + html 合計)、`test/api-rmApiServer.test.mjs`(500 + 服務存活 + 之後的請求照常).
- 更大之 docx:改串流/檔案回傳為另案(主控 2026-10-29 裁定本輪只給明確錯誤).

### R-DIAG-1 `exeFound`/`exePath` 僅為診斷,不得作守門

- 規則:`checkDocxReady()`(async)以 `WHtml2docx.ensureBins({ offline: true, mirror: '' })` 判斷轉檔器是否已備妥於 w-html2docx 之快取(只看快取、不連網、不用鏡像、命中亦校驗);任何失敗皆視為未備妥;`ready` 僅以作業系統為 Windows 判定.不得以此事前攔截轉檔(否則下層取得機制永遠不被執行,見記憶 no-upper-layer-precheck).
- 站點:`src/utils.mjs` `checkDocxReady`;使用處 `src/rmApiServer.mjs` `getState`(`GET /api/health`).
- 測試:`test/unit-utils.test.mjs`〈checkDocxReady〉(快取空、校驗不符、命中、鏡像不取用)、`test/api-rmApiServer.test.mjs`(health 欄位).

### R-HINT-1 w-html2docx 失敗訊息之提示(原文保留、附指引)

- 規則:`cvMdToDocx` 對 `Failed to convert html to docx: ` 之內容:含 `can not find folder for tmp.docx` → 附 `(cwd=…, w-html2docx resolves its default template from cwd, …)`;以 `[E_XXX] ` 開頭(轉檔器取不到,如 `E_OFFLINE_MISSING`、`E_HASH_MISMATCH`)→ 附指引 w-html2docx README〈外部執行檔〉;其餘(含 `unresolved linked picture: …`、`code[1]:…com_error`)原樣.
- 站點:`src/cvMdToDocx.mjs:144-151`.
- 測試:`test/unit-cvMdToDocx.test.mjs`(假轉檔器逐案表;真 w-html2docx 於離線且快取空夾之子程序;錯 cwd 之預設模板).

### R-AV-1 不以獨立腳本檔驅動 Word(Gitea #53,2026-11-02 由「隔離訊息」改為「不落地」)

- 規則:`addDocxToc` 以 Windows PowerShell 5.1(`powershell.exe`,Windows 內建)經 COM 驅動 Word 更新欄位,腳本以 `-NoProfile -NonInteractive -EncodedCommand <UTF-16LE base64>` 帶於命令列,**不寫成任何腳本檔**(.ps1、.vbs、暫存腳本皆無);原因:獨立之 `updateDocxToc.vbs` 曾被 Microsoft Defender 以機器學習誤判為 `Trojan:Script/ObfusScript.A!ml`(任何程式一讀取即隔離),且 VBScript 已公告淘汰.
  - 路徑不拼進腳本字串:由環境變數 `W_MD2DOCX_SRC`、`W_MD2DOCX_DST` 帶入(`execPs` 於呼叫期間暫設 `process.env`、落定後還原;`-EncodedCommand` 之後不可再接引數,實測會被拒);含空白、中文與 `&^%!;,='` `` ` `` `$"` 之路徑原樣通過.
  - 輸出格式、離開碼、訊息字頭與舊 vbs 相同:成功 `toc=<n> tof=<n> pages=<n>` + `ok`、離開碼 0;失敗 `error: <原因>`、離開碼 1,字頭 `cannot start Microsoft Word: `/`open failed: `/`update failed: `/`save failed: `(另存失敗時統計行在前);缺環境變數為用法錯誤、離開碼 2.原因文字之換行折為空格(舊 vbs 保留換行).
  - 各階段 try/catch 分段(對應 `On Error Resume Next` 與分段檢查);失敗先印錯誤、再關閉文件與 Word;`$word.Quit([ref]0)`、`$doc.Close([ref]0)` **須 `[ref]`**(實測 `Quit(0)` 擲「引數應該是 PSReference」,被收尾 `catch {}` 吞掉,WINWORD 每次殘留一個);`$ProgressPreference` 置首行(否則 stderr 被導向時先輸出 CLIXML 進度);stdout 設 UTF-8,`execProcess` 以 `codeCmd: 'auto'`(stdout UTF-8、stderr 系統字碼頁各自判定);`fromClixml` 還原未捕捉例外之 CLIXML.
  - src 與打包不含 .vbs/.ps1;腳本僅含 ASCII;編碼後命令列 5,384 字元(上限 32,767).
- 盤點:Grep `cscript|wscript|\.vbs|\.ps1|powershell` `src/`(glob 加 `!**/*.vbs`);`ls src` 無 .vbs/.ps1.
- 站點全集:`src/addDocxToc.mjs:26`(`psScript`)、`:97`(`execPs`)、`:126`(`runWord`)——全套件驅動 Word 之腳本僅此一處.測試資產 `test/tools/figurePages.vbs`(量測圖名頁碼,`imgFixture.mjs` 以 cscript 執行)為**另一個獨立 .vbs**,不在本規則範圍,見〈已知限制〉.
- 判準:`src/` 無 .vbs/.ps1、原始碼不呼叫 cscript/wscript、`package.json` 不提 .vbs;執行前後系統暫存夾與套件目錄無新增 .ps1;成功與各階段失敗後無本次新啟動之 WINWORD.
- 測試:`test/unit-addDocxToc.test.mjs`(腳本 ASCII/不含路徑/命令列長度、無腳本檔、`fromClixml`、`execPs` 環境變數原樣帶入與還原、離開碼/未捕捉例外/逾時、用法錯誤與 Word 無法啟動之字頭)、`test/unit-cvMdToDocx.test.mjs`(〈addDocxToc Word更新〉無 .ps1 落地與暫存檔清除〔md407 自 `unit-addDocxToc` 移入,R-TEST-1〕;〈Word更新腳本(PowerShell)〉真 Word:成功/open/update/save 各字頭與無 WINWORD 殘留;〈powershell.exe無法啟動〉;〈README之TOC節〉).
- 不做/限制:受限語言模式(Constrained Language Mode,AppLocker/WDAC)下 `New-Object -ComObject` 被禁,無法以此法驅動 Word(模擬實測 `Cannot create type`);逾時時 `powershell.exe` 樹被結束而 COM 啟動之 WINWORD 不在其中,可能殘留(與舊 cscript 作法相同);`execProcess` 無 `env` 選項,故暫設 `process.env`(呼叫端須序列化,`addDocxToc` 經 `runExclusive`).
- 作業守則:不再有須避讀之 .vbs;`test/tools/figurePages.vbs` 未動.

### R-TPL-1 隨套件附帶之模板與預設樣式名稱須中性,且經真 Word 驗證不被更名(Gitea #96,2026-11-07)

- 規則:
  1. `src/templates/` 只附中性模板 `temp.docx`:樣式名稱不帶業主、計畫或其簡稱(README〈Templates〉之表為其清單),樣式 id 前綴為 `Rpt`,文件屬性之作者/最後修改者/套用範本名稱皆空,無附加模板(`attachedTemplate`)與其外部連結.不附任何帶業主樣式之模板(原模板已於本輪移除,需要舊外觀之使用端自備模板並給 `fpInTemp` 與 `titleStyle`);`rmApiServer` 之 `templateDef` 預設 `temp.docx`.
  2. 預設樣式名稱皆為中性:`toc.titleStyle`=`報告目錄標題`、`toc.headStyles`=`['章標題','節標題','小節標題']`;其名稱出現在產出 docx 之樣式清單.
  3. 樣式名稱不得為 Word 內建樣式之本地化名稱:繁體中文版 Word 以「目錄標題」為內建樣式 TOC Heading 之本地化名稱,自訂樣式同名時於轉檔時被更名為「目錄標題1」(2026-11-07 Word 16 實測),使模板自有之樣式不被沿用(`idOf` 找不到),另建者於另存時亦被更名(「目錄標題2」).故預設名稱不用「目錄標題」;更動預設名稱或模板樣式名稱時,須以真 Word 確認轉檔後仍為原名(模板全部樣式名稱經 Word 轉檔後皆保留,28 個報告樣式中僅「目錄標題」曾被更名).
  4. 無舊名相容:模板只有舊名樣式而未給 `titleStyle` 者,三目錄標題改用新建之置中粗體樣式(README〈Breaking changes〉8、9).
- 排版不變:`temp.docx` 與原模板之差異僅為樣式名稱、樣式 id 前綴、`settings.xml` 之 `attachedTemplate` 與其 rels、`docProps` 之作者/最後修改者/範本名稱;其餘部件逐位元組相同,`styles.xml`/`numbering.xml`/`document.xml` 將對照表反向套回後與原模板逐字相同(2026-11-07 驗證,腳本不入 repo,因原模板已移除而不能重跑);以真 Word 對同一 md(封面、2 層標題、圖、表、`toc`、`keepCaption`)用原模板與 `temp.docx` 各轉一次(`titleStyle` 各用其自有之目錄標題樣式):`toc` 回傳(項目數、頁數 6、`levels`、`lists`)相同、圖尺寸相同、全部 `sectPr`(頁面、邊界、行格線、頁碼)相同,`document.xml` 將 pStyle/rStyle/tblStyle 之 id 換為樣式名稱、新名反向還原後逐字相同(25,103 字元;Word 另存會自行重編部分樣式 id 如 `Rpt11`→`aff2`,故不以 id 比對).
- 盤點:Grep `titleStyle|headStyles` `src/`;`ls src/templates`(只有 `temp.docx`);Grep 模板外之 `\.docx` 範例路徑(`README.md`、`g.mjs`、`src/` JSDoc、`test/`)皆為 `temp.docx`.
- 站點全集:`src/docxToc.mjs:19`(`tocDef.titleStyle`)、`:20`(`headStyles`)、`:363-371`(模板無該樣式時另建並於呼叫端指定時提示)、`src/addDocxToc.mjs:166-167`(JSDoc)、`src/rmApiServer.mjs:76,138`(`templateDef`)、`src/cvMdToDocx.mjs:43,97`(舉例)、`README.md`(範例、〈Templates〉、〈Breaking changes〉8、9)、`g.mjs:12`.
- 測試:`test/unit-docxToc.test.mjs`〈未給titleStyle(temp.docx與內建模板)〉〈模板只有舊名樣式…〉〈temp.docx(本套件所附之中性模板)〉(自訂段落樣式清單、樣式 id 前綴、參照皆有對應、文件屬性與附加模板為空、A4/邊界/行格線、預設名稱不為「目錄標題」);`test/unit-cvMdToDocx.test.mjs`〈toc:true且未給titleStyle(temp.docx)〉〈…(模板只有舊名樣式)〉(真 Word:三目錄標題之樣式名稱、模板全部樣式名稱經轉檔後原名保留、樣式 id 前綴;衍生模板於 `test/_tmp`);`test/api-rmApiServer.test.mjs`、`test/api-rmApiClient.test.mjs`(模板清單只有 `temp.docx`).
- 限制:w-html2docx 內建模板含頁首頁尾,其上 `toc: true` 於真 Word 被拒(`TOC check failed (1): the cover section should have no footer`),修改前即存在、與樣式名稱無關,範圍外未修,見〈已知限制〉;故真 Word 之測試用 `temp.docx` 與其衍生模板.

### R-TEST-1 實際啟動 Word 之測試只在 `unit-cvMdToDocx.test.mjs`(同檔依序),無例外

- 規則:mocha `--parallel` 之不同檔不得同時驅動 Word.凡實際啟動 Word 之測試(經 w-html2docx 轉檔、經 `addDocxToc` 或 `execPs(psScript)` 更新欄位、`getFigurePages`〔`figurePages.vbs`〕量頁碼)一律置於 `test/unit-cvMdToDocx.test.mjs`,同檔依序;其他測試檔以假轉檔器(`runWithFakeHtml2docx`)、`simulateWordToc`、未註冊之 COM 類別或「於啟動 Word 前即 reject 之輸入」代之.新增之 Word 測試不論長短皆併入該檔.
- why:本套件之 `runExclusive` 只在同一行程內排隊.不同行程同時驅動 Word 會互相破壞:w-html2docx 之轉檔器以 win32com `Dispatch` 連上他行程已啟動之 Word,轉畢 `Quit` 而關閉對方之文件(2026-11-10 md407 實測:轉檔期間無新 WINWORD、轉畢對方 WINWORD 消失).實測兩行程(`addDocxToc` 與 `cvMdToDocx`)同時起跑、`addDocxToc` 晚 1.5 s 者主 repo(w-html2docx 1.0.51)與副本(1.0.50)各 2/2 兩邊皆失敗(`addDocxToc`:`update failed: 不可在值為 Null 的運算式上呼叫方法` 且 `wmd2docx_toc_*_in.docx` 被鎖而殘留於系統暫存夾;轉檔:`AttributeError: Open.Range`),依序各 0/2.過去另見 `com_error -2147418111 接收者已拒絕這個呼叫`、`-2147023174 RPC 伺服器無法使用`,失敗後可遺留 WINWORD 與開啟中之文件(鎖住 `test/_tmp` 與 `src/templates/~$mp.docx`).原 `unit-addDocxToc`〈Word更新〉為「短時間、維持現狀」之例外,2026-11-10 全庫連動跟跑時實際發生(3/3 輪紅),md407 移入本檔.
- 盤點:Grep `hasWord\(\)|addDocxToc\(|execPs\(psScript|cvMdToDocx\(|WHtml2docx\(|getFigurePages\(|cvMdTo\(` `test/*.test.mjs`,逐處判是否實際啟動 Word(假轉檔器之子行程內者、於 Word 前 reject 者〔輸入檢查、2 GiB、無標題 skip、前提不符〕、`psScript` 換為未註冊類別者、打假服務者不算).
- 站點全集(2026-11-10):實際啟動 Word 者皆在 `test/unit-cvMdToDocx.test.mjs`——〈實轉〉、〈addDocxToc Word更新〉、〈圖片內嵌於docx〉、〈轉檔器取得失敗與預設模板查找失敗之訊息〉之「給fpInTemp時…有Word則轉檔成功」.不啟動 Word 者:`unit-addDocxToc`(輸入檢查、2 GiB、無標題、前提不符;Word 無法啟動案用 `Md2docxNoSuchApp.Application`)、`unit-cvMdTo`/`unit-cvMdToDocxImg`/`api-rmApiServer`(假轉檔器,toc 案皆無標題 skip)、`api-rmApiClient`(docx 請求皆打假服務)、`unit-docxToc`(`simulateWordToc`)、`unit-docxKeep`、`unit-utils`.
- 判準:上列 Grep 之每一處,實際啟動 Word 者所在檔為 `unit-cvMdToDocx.test.mjs`;否則為違規.
- 共用夾具:`addDocxToc` 之測試夾具與檢查函數(`bodyWord`、`sha`、`nTmpToc`、`snapPs1`)只定義於 `test/tools/tocFixture.mjs`,`unit-addDocxToc` 與 `unit-cvMdToDocx` 共用.
- 測試分層:`unit-` 前綴含真 Word 者沿用既有慣例(`hasWord()` 守門).

## 二、改動範圍 → 測試檔對照

| 改動 | 必跑 |
|---|---|
| `src/utils.mjs` | `unit-utils`、`unit-cvMdTo`、`unit-cvMdToDocx`、`unit-addDocxToc`、`api-rmApiServer`、`api-rmApiClient` |
| `src/cvMdToDocx.mjs` | `unit-cvMdToDocx`、`unit-cvMdToDocxImg`、`unit-cvMdTo`、`api-rmApiServer` |
| `src/cvMdTo.mjs` | `unit-cvMdTo`、`unit-cvMdToDocx`、`api-rmApiServer`、`api-rmApiClient` |
| `src/addDocxToc.mjs`、`docxToc.mjs` | `unit-addDocxToc`、`unit-docxToc`、`unit-cvMdToDocx`(含 R-AV-1 之 PowerShell 腳本案、R-TPL-1 之真 Word 案) |
| `src/templates/*`(模板)、`src/rmApiServer.mjs` 之 `templateDef` | `unit-docxToc`、`unit-cvMdToDocx`(真 Word,R-TPL-1)、`api-rmApiServer`、`api-rmApiClient`,以及凡以模板為 fixture 者(`test/tools/docxFixture.mjs`、圖高/行格線案) |
| `src/docxKeep.mjs`、`docxXml.mjs` | `unit-docxKeep`、`unit-cvMdToDocx` |
| `src/rmApiServer.mjs`、`rmApiClient.mjs` | `api-rmApiServer`、`api-rmApiClient` |
| 升 w-md2html / w-html2docx | 全套 `npm test`(真 Word,經 `quiet-run --lock`) |
| 新增或搬移啟動 Word 之測試、`test/tools/tocFixture.mjs` | `unit-cvMdToDocx`、`unit-addDocxToc`,並依 R-TEST-1 之盤點確認他檔不啟動 Word;收尾跑全套(`--parallel` 之跨檔交錯只在全套出現) |

## 三、破壞性變更紀錄

見 `README.md`〈Breaking changes〉(前輪:中介 html 不自足、內嵌圖之格式與尺寸、缺圖一律 reject 與 `allowMissingAssets` 移除、解析不到之圖 reject、`cvMdTo` 回傳自足 html、`checkDocxReady` 改 async 與 `getFpExe` 移除、硬上限之明確錯誤;md96:`toc.titleStyle` 預設改為中性之 `報告目錄標題`、原模板移除並改附中性 `temp.docx`,見 R-TPL-1).

## 四、已知限制(不修,留待另案)

- `cvMdTo`/服務回傳 docx 以 base64 置於 JSON,合計 ≤ 約 401,866,734 位元組;改串流/檔案回傳為另案.
- `rmApiClient` 夾帶資產/模板之請求端 JSON 受字串上限限制;資產單檔 > 2 GiB 時為通用錯誤.
- 內容非圖片但副檔名為 `.png` 之檔,Word 不報錯(w-html2docx 不偵測).
- w-md2html:同一行多個 `<img src="">` 擲 `size(s)!==2`;單引號/無引號 `src` 不處理(上游另案).
- 不帶 `pHYs` 且小於版心之 PNG 依 96 dpi 顯示(每像素 0.75 點),過去為 1 點(待主控裁示是否需補償,見報告).
- w-html2docx 內建模板(未給 `fpInTemp`)含頁首頁尾,其上 `toc: true` 於真 Word 以 `Failed to add the TOC: TOC check failed (1): the cover section should have no footer` reject(2026-11-07 以 w-html2docx 1.0.48 實測,git HEAD 之 src 亦同,與 `titleStyle` 無關);`docxToc.mjs` 之 prep 層(無 Word)對內建模板可處理,故差在 Word 更新後封面節繼承頁尾.範圍外,未修,待主控裁示是否另案.
- 跨行程同時驅動 Word(兩個 node 行程各自呼叫 `cvMdToDocx`/`addDocxToc`,或其他程式同時以 COM 驅動 Word)不受本套件之佇列保護:`runExclusive` 只排同一行程;w-html2docx 之轉檔器(`srcPython/htmlToDocx.py` 之 `wc.Dispatch`)連上執行中之 Word 並於轉畢 `Quit`,使對方之文件被關、轉換失敗(md407 實測,1.0.50 與 1.0.51 相同).建議 w-html2docx 改 `DispatchEx`(每次獨立實例),待主控裁示是否立案;本套件 README 僅寫「conversions and TOC updates share one queue」,未述明限於同一行程.
- `test/tools/figurePages.vbs`(測試資產:以 Word 量圖與圖名之頁碼,`test/tools/imgFixture.mjs` 以 cscript 執行)仍是獨立 .vbs;因 `package.json` 無 `files` 且無 `.npmignore`,`test/` 隨 npm 包發布,該檔因而仍在發布內容中.R-AV-1 只處理 `src/` 驅動 Word 之腳本,此檔待主控裁示(改為 `execPs` 之 PowerShell 腳本,或自發布內容排除 `test/`).
