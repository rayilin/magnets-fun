/* 對戰後端位址設定
 *
 * 留空字串 = 使用同一個網站的 /api（老師電腦跑 start-battle.bat 的區網模式）。
 * 現在填入的是 Render 上的雲端對戰伺服器（免費方案，閒置 15 分鐘會休眠，
 * 第一次連線要等 30～60 秒，頁面會自動顯示「正在喚醒」並重試）。
 *
 * 也可以不改這個檔案，直接在網址加上 ?api=https://xxx.onrender.com
 */
window.BATTLE_API = 'https://magnets-battle.onrender.com';
