# 回饋

任務中發現之自有套件、全域技能、全域設定之調整與抽提提案;一條一列,只增不改,處置結果另記一列並寫明承接哪一列.

| 日期 | 類別 | 對象 | 內容 | 依據 |
|---|---|---|---|---|
| 2026-11-10 | 自有套件 | w-html2docx | `srcPython/htmlToDocx.py:406` 以 `wc.Dispatch('Word.Application')` 取得 Word,win32com 會先連上執行中之實例;轉畢於 `:433` `app.Quit(0)`(不儲存).結果:他行程(或使用者)已開啟之 Word 被借用並結束,其文件被關;並行時兩邊皆失敗.建議改 `win32com.client.DispatchEx('Word.Application')`(每次獨立實例,只 Quit 自己啟動者).驗證:PowerShell 先開隱藏 Word 並開一份文件,轉檔期間應出現第 2 個 WINWORD,轉畢原 WINWORD 仍在且文件數不變;兩行程並行轉檔 0 失敗.1.0.50 與 1.0.51(最新)相同 | md407 實測(attach 實驗:轉檔期間全機只有對方之 WINWORD,其文件數 1→3→0,轉畢該行程結束;並行實驗兩版本各 2/2 失敗) |
| 2026-11-10 | 全域技能 | role-coder-for-test §5 | 建議補一條:mocha `--parallel` 下,用到全機單一資源(Word 等 COM 單一實例、固定檔案鎖)之測試須集中於同一測試檔依序執行(或以跨行程鎖保護),不因「時間短」留例外;以系統暫存夾之全域計數檢查殘留時,注意他行程之同名前綴. | 本套件 R-TEST-1 原留「短時間 Word 測試維持現狀」之例外,2026-11-10 全庫連動跟跑 3/3 輪紅;同類套件(w-html2docx、w-docx-split、w-docx2pdf、w-docx2html)皆用 Word |
| 2026-11-25 | 自有套件 | wsemi | `fsEnsureBins` 之取得預算(`fsEnsureBinsCore.mjs` 之 `lockBudget`,鎖之持有與等候上限)未匯出,包裝端與測試端為使時限對齊(mocha `before` 時限、等候上限)只能依原始碼複寫其算法(`n × totalMs + min(totalMs, 60000)`),wsemi 日後調整算法時複寫處靜默失準.建議匯出純函數(依 manifest 之單位與 opt 回傳預算毫秒),或於 `onProgress` 之 `wait` 事件帶預算. | `test/tools/binEnsure.mjs` 之 `binBudgetMs`;`unit-binEnsure.test.mjs`〈時限對齊下載預算〉與〈他方持有鎖時等候〉以實測對齊 |
| 2026-11-25 | 自有套件 | w-html2docx | `dist`(UMD)內嵌建置時之 wsemi,使用端升級 wsemi(帶跨程序鎖)對經 `dist` 之取得無效:`dist/w-html2docx.umd.js` 內 `.lock-`、`lock-wait`、`E_LOCK_TIMEOUT` 皆 0 處,經 `src/WHtml2docx.mjs`(其 `src/ensureBins.mjs` 引用使用端之 `wsemi/src/fsEnsureBins.mjs`)者才帶鎖.建議 README〈外部執行檔〉寫明「要跨程序只下載一次須經 `src` 入口,或待其升版重建 `dist`」,並於 wsemi 升版連動時重建 `dist`. | 1.0.72 之 `dist` 搜尋結果;`unit-binEnsure` 之改引 `dist` 變異使三案轉紅 |
| 2026-11-25 | 全域技能 | role-coder-for-test(§9 或新節:共用取得工具之測試) | 建議補三點:①共用取得工具須由其**自身之綁定**被鎖案與多行程案走過(測試與子行程皆經工具匯出之同一函數取得),不能只測測試檔自己之 import——獨立複審以工具改引 `dist` 之變異使原五案全綠而發現;②「副本根巢狀舊版相依」證明新測試有鑑別力之作法(`./tmp/<夾>/node_modules/<舊版>`,最近之 `node_modules` 優先,不需 junction)與「本機 http 鏡像以事件閘門計次」寫法;③找隱式取得者以 `W_BIN_OFFLINE=1`＋空 `W_BIN_CACHE` 實跑全套,不憑靜態盤查(外部盤查曾把三個用假轉檔器之檔列為真轉檔隱式取得) | 本套件經驗〈轉檔器 htmlToDocx.exe 之集中取得…〉;複審變異表(改引 dist 全綠);本機同步之 role-coder-for-test 尚無共用取得工具之測試一節(疑未同步,內容待主控核對) |
