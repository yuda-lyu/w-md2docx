<div pretitle style="font-size:14pt; text-align:center;">
    測試範例
</div>

# 計畫報告

## 一、摘要

本計畫旨在提升 **系統效能** 與 *資料處理品質*。  
專案期間共完成三大目標：

1. 資料流程自動化
2. 系統效能優化
3. 使用者介面改善

> 「流程讓人專注，工具讓人加速。」—— 計畫紀錄語

---

## 二、背景與目的

### 2.1 背景說明

隨著資料量逐年增加，需要一套能：

- 自動化處理
- 智慧化調整
- 高效能儲存

的整合方案。

### 2.2 計畫目的（含核取方塊）

- [x] 建置自動化流程  
- [x] 建立監控警示  
- [ ] 完成第三方 API 串接（下一期目標）

---

## 三、工作項目與進度

### 3.1 工作項目列表

測試無序巢狀清單

- **資料前處理**
  - 格式轉換
  - 欄位標準化
  - 雜訊清理

- **資料分析模組**
  - 統計分析
  - 時間序列模型
    - ARIMA
    - Prophet

- **視覺化平台**
  - 儀表板建立
  - 自訂圖層呈現

- **資料流程自動化**：各自獨立運行，且兼顧完整性與流程自動化之執行方式。

- **使用者介面改善**：通過多人協作與回饋等方式，協助使用者介面改善。

- **統計分析**

    須說明統計分析之數學模型，以及輸出入之資料格式。

- **時間序列模型**

    須說明所採用之時間序列模型，包含數學基礎、演算法與回測之統計分析結果。

測試有序巢狀清單

1. **資料前處理**
    - 格式轉換
    - 欄位標準化
    - 雜訊清理

1. **資料分析模組**
    - 統計分析
    - 時間序列模型
      - ARIMA
      - Prophet

1. **視覺化平台**
    - 儀表板建立
    - 自訂圖層呈現

1. **資料流程自動化**：各自獨立運行，且兼顧完整性與流程自動化之執行方式。

1. **使用者介面改善**：通過多人協作與回饋等方式，協助使用者介面改善。

1. **統計分析**

    須說明統計分析之數學模型，以及輸出入之資料格式。

1. **時間序列模型**

    須說明所採用之時間序列模型，包含數學基礎、演算法與回測之統計分析結果。

---

## 四、成果內容

### 4.1 表格測試

| 項目 | 完成度 | 備註 |
|------|--------|------|
| 自動化流程 | 100% | 已完成 |
| 系統效能優化 | 95% | 尚可持續改善 |
| 介面改善 | 80% | 預計下期延伸 |

---

### 4.1.1 表格字型大小標記（data-for="tab"）

以空 div 標記於表格前，data-style-th／data-style-td 分別設定該表 th／td 字型大小，處理後標記自動移除。

大字級表（th 16pt、td 13pt）：

<div data-for="tab" data-style-th="font-size:16pt;" data-style-td="font-size:13pt;"></div>

| 欄位 | 說明 |
|------|------|
| 甲 | 大字級示範 |
| 乙 | 大字級示範 |

小字級表（僅設 td 8pt，th 維持預設）：

<div data-for="tab" data-style-td="font-size:8pt;"></div>

| 代號 | 數量 |
|------|------|
| A | 1 |
| B | 2 |

---

### 4.1.2 逐欄樣式標記（data-style-th-{n}／data-style-td-{n}）

n 為 0 起算之欄索引（4 欄則 n=0~3），套用至該欄各列之 th／td。可與全域 data-style-th／data-style-td 併用，逐欄設定後套故可覆寫全域。

下表：全欄 td 10pt（全域）；第 0 欄不換行（轉 docx 時文字不斷行）；第 1 欄表頭置中；第 2 欄字級縮小覆寫全域。

<div data-for="tab" data-style-td="font-size:10pt;" data-style-td-0="white-space:nowrap;" data-style-th-1="text-align:center;" data-style-td-2="font-size:8pt;"></div>

| 項目代號 | 狀態 | 備註說明 |
|------|------|------|
| A-001-XYZ | 完成 | 第二欄字級較小 |
| B-002-ABC | 進行中 | 第二欄字級較小 |

---

### 4.2 程式碼區塊（JavaScript）

```js
import dayjs from 'dayjs'

function calcPerformance(a, b) {
  const diff = b - a
  return {
    start: a,
    end: b,
    improved: diff,
    time: dayjs().format('YYYY-MM-DD HH:mm:ss')
  }
}

console.log(calcPerformance(100, 260))
```

### 4.3 程式碼區塊（bash）

```bash
#!/bin/bash
echo "Start processing..."
python analyze.py --input=data.json --output=report.csv
echo "Done."
```

---

## 五、圖片、連結與註腳

### 5.1 相關連結

- 官方網站：https://example.com  
- GitHub Repo: [點我連結](https://github.com/)  
- 參考文獻：見文末註腳[^ref1]  

### 5.2 圖片插入

![測試圖片](./cocktail.svg)

---

## 六、結論與後續建議

### 6.1 成果總結

本計畫成功達成主要目標，並建立：

- 模組化程式架構
- 可擴充資料處理流程
- 具彈性的視覺化平台

### 6.2 後續建議

1. 增加第三方 API 整合
2. 加強 AI 模型訓練與部署
3. 系統導入自動化測試（CI/CD）

---

## 七、附錄

### 7.1 引用區塊層級

> 第一層引用
>> 第二層引用
>>> 第三層引用

---

### 7.2 數學公式

行內公式：$E = mc^2$

區塊公式：

$$
y = ax^2 + bx + c
$$

---

### 7.3 刪除隱藏元素

<div style="margin:1rem 0;">
    <div style="margin:1rem 0;">顯示文字</div>
    <div style="margin:1rem 0; display:none;">不能顯示文字</div>
</div>

---

### 7.4 特殊元素

前段文字
<div data-for="br"></div>
後段文字

---

## 註腳

[^ref1]: 本註腳示範 Markdown footnote 語法。
