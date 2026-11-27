@echo off
chcp 65001 >nul
title Build ViSpellCheck Pro
cd /d "%~dp0"
echo ==========================================
echo      BUILD ViSpellCheck Pro - Windows
echo ==========================================
where py >nul 2>nul
if errorlevel 1 (
  echo [LOI] Chua tim thay Python. Cai Python 3.10-3.12 va chon Add Python to PATH.
  pause
  exit /b 1
)
py -m pip install --upgrade pip
if errorlevel 1 goto failed
py -m pip install -r requirements.txt
if errorlevel 1 goto failed
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
py -m PyInstaller --clean --noconfirm ViSpellCheck.spec
if errorlevel 1 goto failed
echo.
echo BUILD THANH CONG: %CD%\dist\ViSpellCheck.exe
echo Copy file nay sang may Windows 10/11 de kiem tra.
pause
exit /b 0
:failed
echo.
echo BUILD THAT BAI. Kiem tra ket noi Internet va thong bao loi phia tren.
pause
exit /b 1
