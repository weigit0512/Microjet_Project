@echo off
REM ============================================================
REM  Microjet_Agent.exe 打包腳本
REM  用法：雙擊本檔，或在終端機執行 build.bat
REM ============================================================
setlocal

echo.
echo [1/3] 檢查 PyInstaller ...
python -m pip show pyinstaller >nul 2>&1
if errorlevel 1 (
    echo     未安裝，開始安裝 pyinstaller ...
    python -m pip install pyinstaller || goto :fail
)

echo [2/3] 清理舊產出 ...
if exist build rmdir /S /Q build
if exist dist rmdir /S /Q dist

echo [3/3] 執行 PyInstaller 打包 ...
python -m PyInstaller Microjet_Agent.spec --clean --noconfirm || goto :fail

echo.
echo ============================================================
echo  打包完成！產出檔：dist\Microjet_Agent.exe
echo  雙擊即可啟動伺服器並自動開啟瀏覽器 http://127.0.0.1:5001
echo ============================================================
pause
exit /b 0

:fail
echo.
echo [ERROR] 打包失敗，請檢視上方訊息。
pause
exit /b 1
