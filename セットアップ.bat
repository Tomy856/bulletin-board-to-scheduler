@echo off
cd /d %~dp0
echo [1/2] playwright をインストール中...
call npm install
echo.
echo [2/2] Playwright ブラウザをインストール中...
call npx playwright install chromium
echo.
echo ========================================
echo [OK] セットアップ完了！
echo Chrome起動.bat を実行してください。
echo ========================================
pause
