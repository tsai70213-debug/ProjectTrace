# 推送紀錄

## 第 1 次推送（2026-11-27）

分支 `main`

- 建立 ProjectTrace 備線核對台，支援 Excel 匯入、線材示意圖、線長調整與接線狀態核對。
- 提供工號識別、Excel 核對報告、瀏覽器紀錄移轉及 localhost 操作日誌。
- 公開版內建虛構線表，原始工作簿、備份、匯出報告與日誌保留於本機。

新增檔案

- `wiring-demo/app.js`、`index.html`、`styles.css` 與 `demo-data.js` 包含核對操作、介面及虛構範例。
- `wiring-demo/server.cjs`、`start-local.cmd` 與 `vendor/` 提供本機啟動、日誌及 Excel 讀寫。
- `.gitignore`、`README.md` 與 `PUSH_LOG.md` 記錄檔案排除規則、使用方式與推送歷程。

檢核

- [x] 檢查待提交檔案與機密格式，未發現 API 金鑰或權杖。
- [x] 通過 JavaScript 語法與 Git 待提交差異檢查。
- [x] 在 Chrome 驗證工號、線表切換、Excel 匯出及追蹤紀錄。
- [x] 已推送初始內容提交 `c8961aa`，並確認 GitHub 上的公開 `main` 分支與本機提交一致。
