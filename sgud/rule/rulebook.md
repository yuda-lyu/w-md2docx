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

### R-TEST-1 呼叫 Word 之測試併於 `unit-cvMdToDocx.test.mjs` 依序執行

- 規則:Word 為單一實例,mocha `--parallel` 之不同檔同時調用會互相拒絕(`com_error -2147418111 接收者已拒絕這個呼叫`),失敗後遺留 WINWORD 與開啟中之文件(鎖住 `test/_tmp` 與 `src/templates/~$mp_tpc.docx`).新增之長時間 Word 測試併入該檔(同檔依序),不另開檔.`unit-addDocxToc`/`unit-docxToc` 之 Word 測試為短時間,維持現狀(**殘留風險**:該二檔與 `unit-cvMdToDocx` 於 `--parallel` 開頭同時啟動,2026-10-29 另有一次在舊碼複本樹之全套出現 `com_error -2147023174 RPC 伺服器無法使用`,見 `sgud/exp/experience.md`;其餘全套 3 次未再現).
- 測試分層:`unit-` 前綴含真 Word 者沿用既有慣例(`hasWord()` 守門).

## 二、改動範圍 → 測試檔對照

| 改動 | 必跑 |
|---|---|
| `src/utils.mjs` | `unit-utils`、`unit-cvMdTo`、`unit-cvMdToDocx`、`unit-addDocxToc`、`api-rmApiServer`、`api-rmApiClient` |
| `src/cvMdToDocx.mjs` | `unit-cvMdToDocx`、`unit-cvMdToDocxImg`、`unit-cvMdTo`、`api-rmApiServer` |
| `src/cvMdTo.mjs` | `unit-cvMdTo`、`unit-cvMdToDocx`、`api-rmApiServer`、`api-rmApiClient` |
| `src/addDocxToc.mjs`、`docxToc.mjs` | `unit-addDocxToc`、`unit-docxToc`、`unit-cvMdToDocx` |
| `src/docxKeep.mjs`、`docxXml.mjs` | `unit-docxKeep`、`unit-cvMdToDocx` |
| `src/rmApiServer.mjs`、`rmApiClient.mjs` | `api-rmApiServer`、`api-rmApiClient` |
| 升 w-md2html / w-html2docx | 全套 `npm test`(真 Word,經 `quiet-run --lock`) |

## 三、破壞性變更紀錄

見 `README.md`〈Breaking changes〉(本輪:中介 html 不自足、內嵌圖之格式與尺寸、缺圖一律 reject 與 `allowMissingAssets` 移除、解析不到之圖 reject、`cvMdTo` 回傳自足 html、`checkDocxReady` 改 async 與 `getFpExe` 移除、硬上限之明確錯誤).

## 四、已知限制(不修,留待另案)

- `cvMdTo`/服務回傳 docx 以 base64 置於 JSON,合計 ≤ 約 401,866,734 位元組;改串流/檔案回傳為另案.
- `rmApiClient` 夾帶資產/模板之請求端 JSON 受字串上限限制;資產單檔 > 2 GiB 時為通用錯誤.
- 內容非圖片但副檔名為 `.png` 之檔,Word 不報錯(w-html2docx 不偵測).
- w-md2html:同一行多個 `<img src="">` 擲 `size(s)!==2`;單引號/無引號 `src` 不處理(上游另案).
- 不帶 `pHYs` 且小於版心之 PNG 依 96 dpi 顯示(每像素 0.75 點),過去為 1 點(待主控裁示是否需補償,見報告).
