# 狀態

本檔記載當前已開發功能之最新狀態(各功能做到哪裡、對應之規則帳本條目與測試檔、已知缺口與待辦、最近一次驗證或發布之版本),只記現況、不寫過程(過程見 `sgud/exp/experience.md`).規則帳本(`sgud/rule/rulebook.md`)為藍圖,與本檔比對即知尚未做到者.

- 套件版本:`package.json` 1.0.18(本輪修正未改版本、未發布,由主控發布);依賴 `w-md2html` `^1.1.12`、`w-html2docx` `^1.0.40`、`wsemi` `^1.9.25`(實裝 1.9.28).
- 最近一次驗證:2026-10-29 全套 `npm test`(真 Word 16、經 `quiet-run --lock`)249 passing、1 pending(既有:有 Word 時略過「無Word時…」)、0 failing;`eslint --ext .mjs src test` 0 error.

## 功能

| 功能 | 狀態 | 規則帳本 | 測試 | 缺口/待辦 |
|---|---|---|---|---|
| `cvMdToDocx` md → html → docx | 完成.中介 html 恆以本機絕對路徑引用圖(不內嵌 base64),Word 讀檔後內嵌;圖量不受字串/堆積上限限制 | R-IMG-1 | `unit-cvMdToDocxImg`、`unit-cvMdToDocx`〈圖片內嵌於docx〉 | 不帶 `pHYs` 且小於版心之 PNG 依 96 dpi 顯示(過去 1 點/像素),待主控裁示是否補償 |
| `cvMdToDocx` 錯誤訊息 | 完成.缺圖(md 階段)、解析不到之圖(Word 階段)、轉檔器取不到 `[E_…]`(附 README 指引)、預設模板查找失敗(附 cwd 提示) | R-IMG-3、R-HINT-1 | `unit-cvMdToDocxImg`、`unit-cvMdToDocx` | — |
| `keepCaption`、`toc` 後處理 | 完成.docx > 2 GiB-1 位元組時讀檔前明確 reject(刪除本次產出之 docx) | R-LIM-1 | `unit-cvMdToDocx`、`unit-addDocxToc` | `addDocxToc` 中 Word 更新後之輸出 docx 檢查未測(需 Word 產 > 2 GiB 之 docx) |
| `cvMdTo` md 內容 → html/docx(base64) | 完成.回傳 html 恆自足(含 `both`),先於 docx 產製;base64 合計 > 約 401,866,734 位元組時明確 reject | R-IMG-2、R-IMG-3、R-LIM-2 | `unit-cvMdTo`、`unit-utils`、`unit-cvMdToDocx` | 更大之 docx 須改串流/檔案回傳(另案) |
| `rmApiServer` | 完成.`GET /api/health` 之 `exeFound`/`exePath` 為 w-html2docx 快取之診斷(只看快取);超限回 500 + 明確訊息、服務不中止 | R-DIAG-1、R-LIM-2 | `api-rmApiServer` | — |
| `rmApiClient` | 完成.缺圖於送出前 reject;`allowMissingAssets` 已移除 | R-IMG-3 | `api-rmApiClient` | 夾帶資產/模板之請求 JSON 受字串上限限制,超限為通用錯誤(已知不修) |
| `addDocxToc`、`docxKeep`、`docxToc` | 未動(僅 `addDocxToc` 加大小檢查) | — | `unit-addDocxToc`、`unit-docxKeep`、`unit-docxToc` | — |

## 大量圖片實測(2026-10-29,不入測試套件)

720 張(393 MB)雜訊 PNG、預設堆積:成功,417 s,docx 392,757,546 B,720 圖皆內嵌且位元組等於來源,node 峰值 RSS 約 1.3 GB、WINWORD 534 MB;詳見 `sgud/exp/experience.md`〈大量圖片實測〉.

## 待辦/待裁示

- 不帶 `pHYs` 之小 PNG 之顯示尺寸(0.75 點/像素 vs 過去 1 點/像素)是否補償(w-md2html 寫入 `width`/`height` 或轉檔後調整).
- 服務/`cvMdTo` 回傳 > 約 401 MB 之 docx:改串流/檔案回傳(破壞 API,另案).
