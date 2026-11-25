@echo off
chcp 65001 >nul
echo ========================================================
echo   航運行政 Excel 資料前處理專案 - GitHub 一鍵上傳工具
echo ========================================================
echo.
echo 本機 Git 倉庫已建立完成，作者已設定為：
echo 2s7kk6k2vk@privaterelay.appleid.com
echo.

set /p REPO_URL="請輸入您的 GitHub 遠端倉庫網址 (例如 https://github.com/username/repo.git): "

if "%REPO_URL%"=="" (
    echo [錯誤] 倉庫網址不可為空！
    pause
    exit /b
)

echo.
echo [1/3] 設定遠端倉庫 origin...
"C:\Users\user\AppData\Local\Programs\Git\cmd\git.exe" remote remove origin 2>nul
"C:\Users\user\AppData\Local\Programs\Git\cmd\git.exe" remote add origin %REPO_URL%

echo [2/3] 切換至 main 分支...
"C:\Users\user\AppData\Local\Programs\Git\cmd\git.exe" branch -M main

echo [3/3] 正在推播至 GitHub (若跳出瀏覽器登入視窗，請點選登入授權)...
"C:\Users\user\AppData\Local\Programs\Git\cmd\git.exe" push -u origin main

if %ERRORLEVEL% equ 0 (
    echo.
    echo ========================================================
    echo   [成功] 專案已成功推播上傳至 GitHub！
    echo ========================================================
) else (
    echo.
    echo ========================================================
    echo   [提示] 推播未完成，請檢查倉庫網址或存取權限。
    echo ========================================================
)

echo.
pause
