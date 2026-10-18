@echo off
setlocal
cd /d "%~dp0"
title Building DataTransformer ...

echo [1/3] Installing dependencies: PyMySQL, psycopg2-binary, pymongo, openpyxl, PySide6, paramiko, Nuitka ...
python -m pip install --upgrade pymysql psycopg2-binary pymongo openpyxl PySide6 paramiko cryptography nuitka -i https://pypi.mirrors.ustc.edu.cn/simple/
if errorlevel 1 (
    echo Dependency install failed. Check network / pip mirror and retry.
    pause
    exit /b 1
)

echo [2/3] Building with Nuitka (onefile mode, first build is slow) ...
python -m nuitka ^
  --standalone ^
  --onefile ^
  --enable-plugin=pyside6 ^
  --windows-console-mode=disable ^
  --nofollow-import-to=pandas ^
  --include-package=pymysql ^
  --include-package=psycopg2 ^
  --include-package-data=psycopg2 ^
  --include-package=pymongo ^
  --include-package=bson ^
  --include-package=paramiko ^
  --include-package=cryptography ^
  --include-package=et_xmlfile ^
  --output-dir=build ^
  --output-filename=DataTransformer ^
  data_transformer.py

if errorlevel 1 (
    echo.
    echo Build failed. Troubleshooting:
    echo   1. Remove --onefile above to build folder mode:
    echo      output at build\data_transformer.dist\DataTransformer.exe
    echo   2. Make sure Visual Studio Build Tools ^(Desktop C++^) or MinGW64 is installed.
    pause
    exit /b 1
)

echo [3/3] Verifying onefile artifact ...
set "BUILT_SIZE="
for %%F in ("%~dp0build\DataTransformer.exe") do set "BUILT_SIZE=%%~zF"
if not defined BUILT_SIZE goto :bad_size
if %BUILT_SIZE% LSS 10000000 goto :bad_size
echo [OK] build\DataTransformer.exe = %BUILT_SIZE% bytes (complete onefile build)
certutil -hashfile "%~dp0build\DataTransformer.exe" SHA256
echo [3/3] Done: build\DataTransformer.exe
pause
exit /b 0

:bad_size
echo.
echo [ERROR] build\DataTransformer.exe is missing or only %BUILT_SIZE% bytes.
echo   A complete PySide6 + paramiko onefile build is about 44,000,000+ bytes.
echo   This usually means one of:
echo     1. the build was interrupted, so the payload was NOT attached to the launcher;
echo     2. only the standalone/launcher exe was copied -- in onefile mode copy the
echo        FULL build\DataTransformer.exe, or in folder mode copy the WHOLE
echo        build\data_transformer.dist folder along with the exe;
echo     3. the file transfer was truncated on the target machine.
echo   Rebuild and verify the file size / SHA-256 before distributing.
pause
exit /b 1
