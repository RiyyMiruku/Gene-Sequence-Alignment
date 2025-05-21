# Gene Sequence Alignment (VBA for Word)

本專案是一個以 VBA 製作的基因序列比對工具，可在 Word 中進行序列比對與相似度計算。

## 功能說明

- **序列比對**：使用 Needleman-Wunsch 演算法對兩條序列進行全域比對。
- **自動去除非字母字元**：輸入的序列會自動去除非英文字母，只保留 A-Z。
- **比對結果顯示**：顯示比對後的兩條序列、相似度百分比、總長度，以及不同位置的差異。
- **序列反轉**：可一鍵反轉任一輸入序列，方便進行反向比對。
- **進度條**：比對時會顯示進度條。

## 使用方式

1. 開啟 Word，載入本 VBA 專案。
2. 執行 `ShowCompareForm` 巨集，開啟比對表單。
3. 在表單的兩個輸入框分別輸入欲比對的序列（僅限英文字母）。
4. 點擊「比對」按鈕（CommandButton1），即可看到比對結果與詳細差異。
5. 可使用「反轉」按鈕（CommandButton2/3）反轉任一序列。

## 主要程式邏輯

- `ShowCompareForm`：顯示比對表單。
- `CommandButton1_Click`：執行序列比對、計算相似度、顯示結果。
- `AlignSequences`：Needleman-Wunsch 全域比對演算法實作。
- `RemoveNonLetters`：去除輸入中的非英文字母。
- `AlignTextForDisplay`：將比對結果格式化為易讀的間隔字串。
- `CommandButton2_Click`/`CommandButton3_Click`：反轉輸入序列。

## 檔案說明

- `Module2.bas`：巨集入口，負責顯示表單。
- `UserForm2.frm`/`UserForm2.frx`：比對表單及其設計與程式碼。
- `README.md`：專案說明文件。

---

如需進一步自訂比對參數（配分、懲罰分數），可直接修改表單程式碼中的相關變數。