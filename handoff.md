# 交接檔（handoff.md）

> 任何 Agent、任何電腦接手前**必讀**；收工時**必更新**。本檔只放交接必需的精簡資訊，詳細脈絡放 Obsidian（L3）。

## ⏯️ 目前做到哪
**教學網站已上線（第一版完成）。** 翰林三上 **第 2 單元「磁鐵好好玩」4 節課**版本全部產出，
GitHub Pages 部署成功：<https://rayilin.github.io/magnets-fun/>

## 🚦 目前狀態
| 項目 | 狀態 |
|------|------|
| 教案 docx（8 頁、13 節） | ✅ `教案/教案_磁鐵好好玩.docx`（＋PDF） |
| 學習單 2 份 | ✅ 學習單①（2 頁）②（1 頁），docx＋PDF |
| 測驗卷（10 題＋答案解析） | ✅ docx＋PDF |
| 簡報 24 頁 16:9 | ✅ `簡報/磁鐵好好玩_教學簡報.pptx` |
| 簡報圖檔 1600×900 ×24 | ✅ `教學網頁/slides/slide01–24.jpg`（PowerPoint COM 匯出） |
| 教學網頁（11 區） | ✅ `教學網頁/index.html`，0 console error |
| 測驗語音朗讀 | ✅「🔊 念題目」按鈕（Web Speech API、zh-TW、rate 0.9），換題／作答自動停止 |
| 對戰（battle.html／server.js／30 題） | ✅ 區網模式可用；雲端後端**尚未部署** |
| 素材圖庫 147 張 | ✅ `media/img`＋`media/thumb`＋`images.js` |
| 旁白動畫 4 段 MP4 | ❌ **未做**（規劃排在第二批） |
| GitHub Pages | ✅ 公開 repo、workflow 部署成功 |
| Render 對戰後端 | ⏸ 檔案已備（`render.yaml`＋`battle-server.js`），等老師登入 Render 按 Blueprint |

## 🔗 網址
- 首頁：<https://rayilin.github.io/magnets-fun/>
- 對戰：<https://rayilin.github.io/magnets-fun/battle.html>
- 答案檔：`.../questions.json` → **404（正確，已被 Pages 排除）**
- repo：<https://github.com/rayilin/magnets-fun>（PUBLIC）

## 📌 重要修正紀錄
- **單元編號**：先前文件寫「第 4 單元」是錯的。翰林三上「磁鐵好好玩」是**第 2 單元**
  （翰林官方頻道影片標題 CH2、素材庫前綴 02 皆可佐證）。教案／簡報／學習單／網頁已全部改為 2-1／2-2／2-3。
- **中文 PDF 絕對不要用 LibreOffice 轉**：會有字型子集損壞（整頁變成重複的「磁鐵」二字，甚至圖片上疊字）。
  本機 11:30 左右就踩過一次。正確做法：
  - pptx → JPG：PowerPoint COM `Presentations.Open(...).Export(dir,"JPG",1600,900)`
  - docx → PDF：Word COM `Document.SaveAs([ref]$path,[ref]17)`
- **測驗題數**：網頁內建測驗是 **10 題**（`check_site.js` 的標準；超過會被判失敗），
  對戰題庫另備 **30 題**（含解析）。規劃表原寫「測驗 30 題」→ 已調整為 10＋30。
- **測驗語音**：用瀏覽器內建 `speechSynthesis`（不須 API key、不須網路音檔）。
  Chrome／Edge 在 Windows 有 Microsoft 雅婷／涵涵（zh-TW）；iPad Safari 亦有中文語音。
  按鈕會念「第 N 題＋題目＋A/B/C/D 選項」，再按一次停止；作答或換題會自動 `cancel()`。
  不支援的瀏覽器直接不顯示按鈕。若在 iPad 上沒聲音，先確認系統「設定 → 輔助使用 → 朗讀內容」與音量。

## ➡️ 下一步（依優先順序）
1. **對戰後端**：Render → New + → Blueprint → 選 `rayilin/magnets-fun` → Apply。
   拿到網址後改 `教學網頁/battle-api.js` 的 `window.BATTLE_API = 'https://xxx.onrender.com';` 再 push。
2. **旁白動畫 4 段**（吸引什麼／磁極／相斥相吸／生活應用，各 12–15 秒）：
   用 `hyperframes` 技能，輸出放 `教學網頁/media/video/`，並在 `index.html` 加一段播放區塊
   （目前 `SLIDES/QUIZ/VIDEOS` 都沒有 mp4，`check_site.js` 不會要求）。
3. 上課前實機驗收：單槍投影（簡報 24 頁）、平板開網頁（測驗／圖庫載入）、
   對戰兩人配對（區網版 `start-battle.bat`）。

## 🧰 重建指令（詳見 AGENTS.md）
```powershell
python 工具\build_deck.py      # 24 頁簡報
python 工具\build_plan.py      # 教案
python 工具\build_sheets.py    # 學習單＋測驗卷
python 工具\build_questions.py # 30 題題庫
python 工具\build_images.py    # 147 張圖 → media/
```

## ⚠️ 注意事項
- 本資料夾是**獨立的 git repo**，位在上層 repo `natural-science-teaching-column` 裡面；
  上層已把 `磁鐵好好玩/` 加入 `.gitignore`
- `簡報/` 裡另有一份**舊版 15 頁**（`磁鐵的性質與運用_教學簡報.pptx`，奇幻森林風）＋`簡報/src/`，
  與本版無關，**不要刪**；`簡報/src/bg/`、`簡報/src/ops.json` 已列入 `.gitignore`
- 影片 5 部已用 YouTube oembed 驗證（yTaWEgU_wm0／_vpZXNY20PI／CH_mmg-FN6I／uSjNRToSThA／icPBJ6Nr9Gs）
- 一律繁體中文；學生資料只用座號
- 簡報內容為自行編寫＋課綱查證（INe-Ⅱ-7），**未逐字對照教科書頁面**，定稿前建議對一次

## 🕐 最後更新
- 時間：2026-10-23 14:12
- 更新者：OpenCode（DeepSeek V4.1 Flash Free） @ DESKTOP-4475O68
- Git：✅ 已推（`main`，commit `e24aed0` 加入測驗語音朗讀）
- Pages：✅ 部署成功（run 39898832971／39909937037 success）
