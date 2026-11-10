# 回饋

任務中發現之自有套件、全域技能、全域設定之調整與抽提提案;一條一列,只增不改,處置結果另記一列並寫明承接哪一列.

| 日期 | 類別 | 對象 | 內容 | 依據 |
|---|---|---|---|---|
| 2026-11-10 | 自有套件 | w-html2docx | `srcPython/htmlToDocx.py:406` 以 `wc.Dispatch('Word.Application')` 取得 Word,win32com 會先連上執行中之實例;轉畢於 `:433` `app.Quit(0)`(不儲存).結果:他行程(或使用者)已開啟之 Word 被借用並結束,其文件被關;並行時兩邊皆失敗.建議改 `win32com.client.DispatchEx('Word.Application')`(每次獨立實例,只 Quit 自己啟動者).驗證:PowerShell 先開隱藏 Word 並開一份文件,轉檔期間應出現第 2 個 WINWORD,轉畢原 WINWORD 仍在且文件數不變;兩行程並行轉檔 0 失敗.1.0.50 與 1.0.51(最新)相同 | md407 實測(attach 實驗:轉檔期間全機只有對方之 WINWORD,其文件數 1→3→0,轉畢該行程結束;並行實驗兩版本各 2/2 失敗) |
| 2026-11-10 | 全域技能 | role-coder-for-test §5 | 建議補一條:mocha `--parallel` 下,用到全機單一資源(Word 等 COM 單一實例、固定檔案鎖)之測試須集中於同一測試檔依序執行(或以跨行程鎖保護),不因「時間短」留例外;以系統暫存夾之全域計數檢查殘留時,注意他行程之同名前綴. | 本套件 R-TEST-1 原留「短時間 Word 測試維持現狀」之例外,2026-11-10 全庫連動跟跑 3/3 輪紅;同類套件(w-html2docx、w-docx-split、w-docx2pdf、w-docx2html)皆用 Word |
