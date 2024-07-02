# 使用 Excel Lab 載入和使用 LAMBDA Module的指南

## 介紹

此文件將指導您如何使用一組預定義的 LAMBDA 函數模組。這些設計是為了讓您更好的進行資料表的Json轉換和應用。

## 準備工作

1. **安裝 Excel Lab**： 如果您尚未安裝 Excel Lab，請參考以下步驟操作（可能會因Excel版本而有不同的操作步驟）：
   * 打開 Excel。
   * 前往 `開發人員` 標籤。
   * 選擇 `增益集`。
   * 點選`在Microsoft AppSource中尋找更多增益集`
   * 搜尋並安裝 `Excel Labs, a Microsoft Garage project` 增益集。
2. **載入模組**：
   * 打開 Excel Lab。
   * 選擇 `Modules` 標籤。
   * 點選 `+NEW`按鈕來新增一個模組，將模組命名為MacacaEAFE，然後把MacacaEAFE.txt內的內容貼上，點選儲存圖示以保存模組內容。

## LAMBDA 函數模組

### 1. GetDataRange

**功能**：取得有效資料的範圍，在使用Macaca add-in取得Json時很有幫助。

**參數**：

* `startCell`：第一個資料標頭儲存格
* `endColumn`：最後一個資料標頭儲存格

**範例**：

`=GetDataRange(A1, D1)`

### 2. TOJSON

**功能**：將一組指定的header和資料範圍轉換為 JSON array。

**參數**：

* `headers`：標頭範圍
* `data`：資料範圍

**範例**：

`=TOJSON(A1:D1, A2:D10)`
</code></div></div></pre>

### 3. ToCSharpEnum

**功能**：將指定的兩個column轉換為 C# 的Enum。

**參數**：

* `items`：Enum name的資料範圍
* `values`：Enum value的資料範圍

**範例**：

`=ToCSharpEnum(A1:A10, B1:B10)`

### 7. TOJSONARRAY

**功能**：將一個範圍的值視為array element並轉換為 JSON array。

**參數**：

* `range`：要轉換的值範圍

**範例**：

`=TOJSONARRAY(A1:A10)`

### 8. TOJSONDICTIONARY

**功能**：將key-value pair轉換為 JSON dictionary。

**參數**：

* `key`：字典鍵
* `value`：字典值

**範例**：

`=TOJSONDICTIONARY("name", A1)`

### 9. TOJSONOBJECT

**功能**：指定一個範圍的key和一個範圍的value，轉換為一個 JSON 物件。

**參數**：

* `keysRange`：key的範圍
* `valuesRange`：value的範圍

**範例**：

`=TOJSONOBJECT(A1:A10, B1:B10)`
</code></div></div></pre>

### 10. TOISO8601

**功能**：將日期時間值轉換為 ISO 8601 格式。

**參數**：

* `dateTime`：日期時間值
* `timezone`：時區偏移（小時）

**範例**：

`=TOISO8601(NOW(), 8)`
</code></div></div></pre>

### 11. CHECKINVALIDDATA

**功能**：檢查範圍中的無效資料（空單元格或包含空格的單元格）。

**參數**：

* `range`：要檢查的資料範圍

**範例**：

`=CHECKINVALIDDATA(A1:A10)`
</code></div></div></pre>

### NoData 定義

**功能**：定義不同類型的 NoData 函數，在資料表中標示無數據狀況。

**範例**：

```
=NoData_Number()
=NoData_Text()
=NoData_ENUM()
=NoData_Bool()
=NoData_Json()
=NoData_Array()
```

</code></div></div></pre>
