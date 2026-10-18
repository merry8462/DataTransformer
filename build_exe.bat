@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"
title Building DataTransformer ...

rem 若打包后的 exe 双击无法启动(或一闪而过),把下面的 attach 改成 disable 后重新打包试一次,用于二分定位。
set "CONSOLE_MODE=attach"

rem 若打包产物崩溃且怀疑是 Nuitka 版本问题,把下面改成 nuitka==4.2.2(或你验证过的稳定版)再构建
set "NUITKA_SPEC=nuitka==4.2.2"

rem 用法: build_exe.bat        单文件模式,产物 build\DataTransformer.exe
rem      build_exe.bat folder  目录模式(去掉 --onefile),产物 build\data_transformer.dist\DataTransformer.exe
if /i "%~1"=="folder" (
    set "ONEFILE="
    set "BUILD_MODE=folder"
) else (
    set "ONEFILE=--onefile"
    set "BUILD_MODE=onefile"
)

set "LOG=%~dp0build\nuitka_build.log"
if not exist "%~dp0build" md "%~dp0build"
> "%LOG%" echo [DataTransformer build log] %DATE% %TIME%

echo [1/3] Installing dependencies: PyMySQL, psycopg2-binary, pymongo, openpyxl, PySide6, paramiko ...
python -m pip install --upgrade pymysql psycopg2-binary pymongo openpyxl PySide6 paramiko cryptography -i https://pypi.mirrors.ustc.edu.cn/simple/
if errorlevel 1 (
    echo Dependency install failed. Check network / pip mirror and retry.
    pause
    exit /b 1
)

echo [1/3] Installing Nuitka: %NUITKA_SPEC% ...
python -m pip install %NUITKA_SPEC% -i https://pypi.mirrors.ustc.edu.cn/simple/
if errorlevel 1 (
    echo Nuitka install failed. Check network / pip mirror and retry.
    pause
    exit /b 1
)

rem 版本横幅:同时打印到控制台并追加到构建日志,便于事后对照工具链组合
set "PYVER=unknown"
for /f "usebackq delims=" %%V in (`python -c "import sys;print(sys.version.split()[0])" 2^>nul`) do set "PYVER=%%V"
set "PYSIDEVER=unknown"
for /f "usebackq delims=" %%V in (`python -c "import PySide6;print(PySide6.__version__)" 2^>nul`) do set "PYSIDEVER=%%V"
set "NUITKAVER=unknown"
set "NUITKA_FIRST="
for /f "usebackq delims=" %%V in (`python -m nuitka --version 2^>nul`) do (
    if not defined NUITKA_FIRST (
        set "NUITKAVER=%%V"
        set "NUITKA_FIRST=1"
    )
)

call :log "================ build banner ================"
call :log "  python       : %PYVER%"
call :log "  nuitka       : %NUITKAVER%"
call :log "  NUITKA_SPEC  : %NUITKA_SPEC%"
call :log "  PySide6      : %PYSIDEVER%"
call :log "  CONSOLE_MODE : %CONSOLE_MODE%"
call :log "  --onefile    : %ONEFILE%"
call :log "  build mode   : %BUILD_MODE%"
call :log "  build log    : %LOG%"
call :log "=============================================="

echo [2/3] Building with Nuitka (%BUILD_MODE% mode, first build is slow) ...
echo       Nuitka output is appended to build\nuitka_build.log; its last 30 lines are echoed when it ends.
call :log "--- nuitka start: %DATE% %TIME% ---"

python -m nuitka ^
  --standalone ^
  %ONEFILE% ^
  --enable-plugin=pyside6 ^
  --windows-console-mode=%CONSOLE_MODE% ^
  --nofollow-import-to=pandas ^
  --include-package=pymysql ^
  --include-package=psycopg2 ^
  --include-package-data=psycopg2 ^
  --include-package=pymongo ^
  --include-package=bson ^
  --include-package=paramiko ^
  --include-package=cryptography ^
  --include-package=et_xmlfile ^
  --include-package=src ^
  --output-dir=build ^
  --output-filename=DataTransformer ^
  data_transformer.py >> "%LOG%" 2>&1
set "BUILD_RC=%ERRORLEVEL%"
call :log "--- nuitka finished, exit code %BUILD_RC% ---"

echo ---- last 30 lines of build\nuitka_build.log ----
powershell -NoProfile -Command "Get-Content -LiteralPath '%LOG%' -Encoding UTF8 -Tail 30"
if not "%BUILD_RC%"=="0" goto :build_failed

echo.
echo [3/3] Verifying %BUILD_MODE% artifact ...
set "ARTIFACT="
set "BUILT_SIZE="
if defined ONEFILE set "ARTIFACT=%~dp0build\DataTransformer.exe"
if not defined ONEFILE set "ARTIFACT=%~dp0build\data_transformer.dist\DataTransformer.exe"
for %%F in ("%ARTIFACT%") do set "BUILT_SIZE=%%~zF"
if not defined BUILT_SIZE goto :bad_size
if defined ONEFILE if %BUILT_SIZE% LSS 10000000 goto :bad_size
if defined ONEFILE call :log "[OK] build\DataTransformer.exe = %BUILT_SIZE% bytes (complete onefile build)"
if not defined ONEFILE call :log "[OK] build\data_transformer.dist\DataTransformer.exe = %BUILT_SIZE% bytes (folder build)"
certutil -hashfile "%ARTIFACT%" SHA256
if defined ONEFILE call :log "[3/3] Done: build\DataTransformer.exe"
if not defined ONEFILE (
    call :log "[3/3] Done: build\data_transformer.dist\DataTransformer.exe"
    echo   Folder build: distribute the WHOLE build\data_transformer.dist folder, not the exe alone.
    call :log "  分发时必须复制整个 build\data_transformer.dist 文件夹,单独的 exe 无法运行。"
)
pause
exit /b 0

:build_failed
echo.
echo Build failed. Troubleshooting:
echo   1. Try folder mode to tell compilation apart from onefile packaging:
echo      build_exe.bat folder  -^> build\data_transformer.dist\DataTransformer.exe
echo   2. Clean rebuild: delete build\ and %%LOCALAPPDATA%%\Nuitka\Nuitka\Cache
echo      ^(clcache / module-cache / scons-msvc-config^), then rebuild.
echo   3. Make sure Visual Studio Build Tools ^(Desktop C++^) or MinGW64 is installed.
echo   4. Prefer the Nuitka stable release if the installed one is a pre-release:
echo      set NUITKA_SPEC=nuitka==4.2.2 at the top of this script, then rebuild.
call :log "--- build FAILED, see the log above ---"
pause
exit /b 1

:bad_size
echo.
if not exist "%ARTIFACT%" (
    echo [ERROR] %ARTIFACT% not found. Build the exe first:
    echo   未找到构建产物,请先完成上面的 Nuitka 构建步骤。
    goto :bad_size_hint
)
echo [ERROR] %ARTIFACT% is only %BUILT_SIZE% bytes.
echo   A complete PySide6 + paramiko onefile build is about 44,000,000+ bytes.
:bad_size_hint
echo   This usually means one of:
echo     1. the build was interrupted, so the payload was NOT attached to the launcher;
echo     2. only the standalone/launcher exe was copied -- in onefile mode copy the
echo        FULL build\DataTransformer.exe, or in folder mode copy the WHOLE
echo        build\data_transformer.dist folder along with the exe;
echo     3. the file transfer was truncated on the target machine.
echo   Rebuild and verify the file size / SHA-256 before distributing.
pause
exit /b 1

:log
>> "%LOG%" echo %~1
echo %~1
exit /b 0
