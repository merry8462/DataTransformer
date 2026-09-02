@echo off
setlocal
cd /d "%~dp0"
title Building DataTransformer ...

echo [1/3] Installing dependencies: PyMySQL, psycopg2-binary, openpyxl, PySide6, Nuitka ...
python -m pip install --upgrade pymysql psycopg2-binary openpyxl PySide6 nuitka -i https://pypi.mirrors.ustc.edu.cn/simple/
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
  --include-package=et_xmlfile ^
  --output-dir=build ^
  --output-filename=DataTransformer ^
  data_transformer.py

if errorlevel 1 (
    echo.
    echo Build failed. Troubleshooting:
    echo   1. Remove --onefile above to build folder mode:
    echo      output at build\DataTransformer.dist\DataTransformer.exe
    echo   2. Make sure Visual Studio Build Tools ^(Desktop C++^) or MinGW64 is installed.
    pause
    exit /b 1
)

echo [3/3] Done: build\DataTransformer.exe
pause
