@echo off
chcp 65001 >nul
cd /d "%~dp0"

set "DEPLOY=C:\Users\USER\Desktop\JsonEditor\JsonEditor"

REM close running JsonEditor.exe to release file locks
taskkill /F /IM JsonEditor.exe >nul 2>&1

echo === 打包中 (Nuitka)，請稍候 ===
.venv\Scripts\python.exe -m nuitka --standalone --windows-console-mode=disable --enable-plugin=pyside6 --windows-icon-from-ico=icon.ico --include-data-files=icon.ico=icon.ico --include-data-files=assets/tabler-icons.ttf=assets/tabler-icons.ttf --experimental=force-dependencies-pefile --include-package=openpyxl --output-dir=nuitka_dist --output-filename=JsonEditor --assume-yes-for-downloads main.py
if errorlevel 1 (
    echo.
    echo 打包失敗
    pause
    exit /b 1
)

echo.
echo === 部署中 ===
robocopy "nuitka_dist\main.dist" "%DEPLOY%" /MIR /XF config.json /R:2 /W:3 /NFL /NDL /NJH /NJS /NP
if errorlevel 8 (
    echo.
    echo 部署失敗，先關閉 JsonEditor.exe 再試
    pause
    exit /b 1
)

echo.
echo 完成：%DEPLOY%\JsonEditor.exe
pause
