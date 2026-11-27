@echo off
chcp 65001 >nul
title Build ViSpellCheck Pro
cd /d "%~dp0"
echo ==========================================
echo      BUILD ViSpellCheck Pro - Windows
echo ==========================================
echo Thu muc dang chay: %CD%
set "MISSING="
for %%F in (app.py doc_scan.py office_live.py spell_engine.py ViSpellCheck.spec) do if not exist "%%F" (echo [LOI] Thieu file: %%F & set MISSING=1)
if defined MISSING goto :missing
where py >nul 2>nul
if errorlevel 1 (
  echo [LOI] Chua tim thay Python. Cai Python 3.10-3.12 va chon Add Python to PATH.
  pause
  exit /b 1
)
py -m pip install --upgrade pip
if exist requirements.txt (
  py -m pip install -r requirements.txt
) else (
  echo [CANH BAO] Khong thay requirements.txt, cai thu vien truc tiep.
  py -m pip install python-docx openpyxl pyinstaller pywin32
)
if errorlevel 1 goto failed
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
py -m PyInstaller --clean --noconfirm ViSpellCheck.spec
if errorlevel 1 goto failed
echo.
echo BUILD THANH CONG: %CD%\dist\ViSpellCheck.exe
pause
exit /b 0
:missing
echo.
echo Hay copy DAY DU cac file tren vao CUNG MOT THU MUC roi chay lai.
pause
exit /b 1
:failed
echo.
echo BUILD THAT BAI. Kiem tra ket noi Internet va thong bao loi phia tren.
pause
exit /b 1
