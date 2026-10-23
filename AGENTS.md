# 磁鐵好好玩（專案藍圖）

> 本檔為跨 Agent 通用的專案藍圖（AGENTS.md 開放標準）。任何 Agent 的每個 session 都應先讀本檔＋`handoff.md`。

## 專案簡介
翰林版國小自然科學三年級上學期 **第 2 單元「磁鐵好好玩」** 的完整教學專案，**4 節課**版本：
教案（DOCX）、教學簡報（PPTX 24 頁）、學習單 2 份、測驗卷、互動教學網頁
（課程地圖＋3 個互動實驗＋簡報播放器＋10 題測驗＋5 部影片＋抽籤計時器＋147 張圖庫）、
兩人即時配對搶答遊戲，並已部署到 GitHub Pages。

製作流程走 `teaching-site` 技能（教案→簡報→網頁→測驗→對戰→上線→驗收）。

## 關鍵時程
- 第 1 節：10/26（一）　磁鐵能吸引什麼？
- 第 2 節：10/29（四）　磁力有多大？磁極在哪裡？
- 第 3 節：11/2（一）　同極相斥、異極相吸
- 第 4 節：11/5（四）　生活中的磁鐵＋測驗與對戰

## 目標與路線圖
- [x] 階段 0：專案初始化（藍圖、交接檔、repo、Obsidian 筆記）
- [x] 階段 1：確認教學範圍與節數（**4 節**，6 節內容壓縮；課綱 **INe-Ⅱ-7**）
- [x] 階段 2：內容與素材（147 張圖重新編號 m001–m147、5 部影片驗證、迷思清單）
- [x] 階段 3：教案 docx（8 頁、13 節）＋學習單 2 份＋測驗卷（docx＋pdf）
- [x] 階段 4：教學簡報 24 頁（16:9）＋1600×900 圖檔（PowerPoint COM 匯出）
- [x] 階段 5：教學網頁（11 區）與對戰（battle.html／server.js／30 題題庫）
- [x] 階段 6：驗收（`check_site.js` 15 項全過、瀏覽器 0 console error）
- [x] 階段 7：上線 GitHub Pages → <https://rayilin.github.io/magnets-fun/>
- [ ] 階段 8：旁白動畫 4 段（MP4，各 12～15 秒）— 未做
- [ ] 階段 9：對戰後端上 Render（`render.yaml`＋`battle-server.js` 已備，待老師授權部署）

## 資料夾結構
```
磁鐵好好玩/                        ← 本專案根目錄（獨立 git repo，已公開）
├── AGENTS.md                     專案藍圖（本檔）
├── handoff.md                    交接檔
├── 規劃表.md                     核定版規劃（4 節、素材分配）
├── render.yaml                   Render 藍圖（對戰後端）
├── battle-server.js              Render 啟動入口 → 教學網頁/server.js
├── .github/workflows/pages.yml   GitHub Pages 部署（UNIT_DIR=教學網頁）
├── 教案/
│   ├── 教案_磁鐵好好玩.docx / .pdf
│   ├── 學習單1_磁鐵能吸引什麼.docx / .pdf
│   ├── 學習單2_磁極與相斥相吸.docx / .pdf
│   └── 測驗卷_磁鐵好好玩.docx / .pdf
├── 簡報/
│   ├── 磁鐵好好玩_教學簡報.pptx          ← 本版 24 頁
│   └── 磁鐵的性質與運用_教學簡報.pptx    ← 舊版 15 頁（奇幻森林風，保留未刪）
├── 工具/                         產生程式（python-docx／python-pptx）
│   ├── build_images.py           147 張圖 → media/img＋thumb＋images.js
│   ├── build_deck.py             24 頁簡報
│   ├── build_plan.py             教案
│   ├── build_sheets.py           學習單＋測驗卷
│   └── build_questions.py        對戰 30 題題庫
└── 教學網頁/
    ├── index.html                主網頁（11 區，單檔、file:// 可開）
    ├── battle.html               對戰頁
    ├── server.js                 對戰伺服器（區網／Render 共用）
    ├── questions.json            30 題題庫（Pages 部署時排除）
    ├── battle-api.js             雲端後端位址（空字串＝區網模式）
    ├── start-battle.bat          老師一鍵啟動區網對戰
    ├── slides/slide01–24.jpg     簡報圖（1600×900）
    ├── media/img/                147 張原圖（m001–m147）
    ├── media/thumb/              147 張縮圖（360px）
    ├── media/images.js           圖庫清單（給 index.html 用）
    └── downloads/                教案、簡報、學習單、測驗卷
```

## 重建方式
```powershell
# 簡報（產 pptx 後用 PowerPoint COM 匯出 24 張 JPG）
python 工具\build_deck.py
$pp=New-Object -ComObject PowerPoint.Application
$pres=$pp.Presentations.Open("$pwd\簡報\磁鐵好好玩_教學簡報.pptx",$true,$false,$false)
$pres.Export("$pwd\教學網頁\slides\_raw","JPG",1600,900); $pres.Close(); $pp.Quit()

# 教案／學習單／測驗卷（docx 產完用 Word COM 轉 PDF，FileFormat=17）
python 工具\build_plan.py
python 工具\build_sheets.py

# 題庫 / 圖庫
python 工具\build_questions.py
python 工具\build_images.py
```
- **不要用 LibreOffice 轉中文 PDF**：這台機器上會出現字型子集損壞（整頁變成重複的「磁鐵」二字）。
  一律用 **PowerPoint COM**（簡報→JPG）與 **Word COM**（docx→PDF）。
- 圖庫來源：`D:\##115學年度桃源訓導\新增資料夾\02磁鐵好好玩\02-0X-XXX-名稱.jpg`（147 張）。

## 同步層級（本專案初始化至第 3 層級）

| 層級 | 平台 | 位置 | 讀取時機 |
|------|------|------|---------|
| L1 | 本地 | `AGENTS.md`＋`handoff.md` | 每個 session |
| L2 | GitHub | `rayilin/magnets-fun`（**公開**） | 指定時 |
| L3 | Obsidian | `磁鐵好好玩/專案工作流程.md` | 有需要時 |

## 已上線網址
- 教學網頁：<https://rayilin.github.io/magnets-fun/>
- 對戰頁：<https://rayilin.github.io/magnets-fun/battle.html>（後端未部署時走區網模式）

## 與其他專案的關係
- 位於 `C:\Users\rayi\Documents\自然科學教學專欄\` 底下，但**是獨立的 git repo**；
  上層 `natural-science-teaching-column` 已把本資料夾加入 `.gitignore`。
- 參考範本：`..\115-1_三上自然_植物的構造和生長\`（已完成上線）
- 可重用技能：`teaching-site`、`officecli`、`hyperframes`（動畫）、`draw-free`／`draw-gemini`

## 工作約定
- 任何 Agent、任何電腦：**開工先讀 `handoff.md`，收工必更新 `handoff.md`**
- 修改共用檔案前先讀最新內容，避免覆蓋其他 Agent 的變更
- 所有回應與文件使用繁體中文
- 修改前先確認計畫，優先保留原有資料結構

## 安全與隱私（不可違反）
- **不把 API key、密碼、憑證寫進 repo**，也不要貼進 `AGENTS.md`／`handoff.md`；一律放 `.env` 並列入 `.gitignore`
- **學生資料只用座號**，不出現姓名、學號、班級以外的個資、照片或聯絡方式
- repo 已公開：`questions.json`（對戰答案）仍會出現在 GitHub 檔案清單，
  但 Pages 部署已排除，網站上抓不到（`https://rayilin.github.io/magnets-fun/questions.json` 回 404）
