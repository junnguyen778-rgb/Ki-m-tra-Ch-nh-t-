@echo off
chcp 65001 >nul
title Tao file cai dat ViSpellCheck Pro
cd /d "%~dp0"
echo Thu muc dang chay: %CD%
REM Can Python va Inno Setup 6: https://jrsoftware.org/isdl.php
set "MISSING="
for %%F in (app.py doc_scan.py office_live.py spell_engine.py ViSpellCheck.spec installer.iss) do if not exist "%%F" (echo [LOI] Thieu file: %%F & set MISSING=1)
if defined MISSING goto :missing
if exist requirements.txt (
  py -m pip install -r requirements.txt
) else (
  py -m pip install python-docx openpyxl pyinstaller pywin32
)
if errorlevel 1 goto :err
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
py -m PyInstaller --clean --noconfirm ViSpellCheck.spec
if errorlevel 1 goto :err
set ISCC="%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe"
if not exist %ISCC% set ISCC="%ProgramFiles%\Inno Setup 6\ISCC.exe"
if not exist %ISCC% goto :noinno
%ISCC% installer.iss
if errorlevel 1 goto :err
echo.
echo Xong! File cai dat: %CD%\dist\ViSpellCheck-Setup.exe
pause
exit /b 0
:missing
echo Hay copy DAY DU cac file tren vao CUNG MOT THU MUC roi chay lai.
pause
exit /b 1
:noinno
echo Chua cai Inno Setup 6. Tai tai https://jrsoftware.org/isdl.php
pause
exit /b 1
:err
echo Tao file cai dat that bai.
pause
exit /b 1
