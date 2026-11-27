# ProjectTrace

備線核對台，用來查閱 Excel 線表、核對線長與接線內容，並記錄完成狀態和現場修改。

## 啟動

1. 安裝 Node.js。
2. 在 Windows 檔案總管雙擊 `wiring-demo/start-local.cmd`，或在專案根目錄執行 `node wiring-demo/server.cjs --start`。
3. 開啟 [localhost:8764](http://localhost:8764/)，輸入工號後即可操作。

## 操作

- 「匯入 Excel」可一次選取多份 `.xls` 或 `.xlsx` 線表。匯入格式須包含 NO、線名、前接頭與長度欄位。
- 在「線表檔案」側欄切換來源，再從備線清單選取一筆，查看兩端接頭、剝線長與線材示意圖。
- 記錄線長是否需要增減、接線核對結果與現場備註。「匯出核對報告」會下載含操作工號與修改履歷的 Excel。
- 核對資料儲存在目前瀏覽器。透過 localhost 使用時，服務也會將操作事件追加到 `wiring-demo/logs/tracking.log`。

公開版內建虛構線表供試用。原始線表、備份、匯出報告與執行紀錄屬於本機資料。

Excel 讀寫使用隨附的 SheetJS Community Edition 0.20.3，授權見 [SheetJS LICENSE](wiring-demo/vendor/LICENSE.txt)。
