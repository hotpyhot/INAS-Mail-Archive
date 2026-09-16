@echo off
setlocal EnableExtensions
cd /d "%~dp0"
set "APP_NAME=INAS_Mail_Archive_Community_Edition_Ver1_0_0"
if not exist ".venv\Scripts\python.exe" py -m venv .venv
call ".venv\Scripts\activate.bat"
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python -m PyInstaller ^
  --noconfirm ^
  --clean ^
  --onefile ^
  --console ^
  --noupx ^
  --name "%APP_NAME%" ^
  --icon "assets\inas_mail_archive.ico" ^
  --version-file "version_info.txt" ^
  --add-data "assets\inas_mail_archive.ico;assets" ^
  --add-data "assets\inas_mail_archive.png;assets" ^
  --add-data "locales\ja.json;locales" ^
  --add-data "locales\en.json;locales" ^
  --add-data "locales\vi.json;locales" ^
  --collect-all tkinterdnd2 ^
  --collect-all pystray ^
  --hidden-import PIL._tkinter_finder ^
  --paths src ^
  src\main.py
if errorlevel 1 goto :error
echo Build complete: dist\%APP_NAME%.exe
exit /b 0
:error
echo Build failed.
exit /b 1
