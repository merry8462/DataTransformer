@echo off
chcp 65001 >nul
cd /d "%~dp0"
title 正在打包 DataTransformer ...

echo [1/3] 安装运行依赖与 Nuitka ...
python -m pip install --upgrade pymysql psycopg2-binary openpyxl PySide6 nuitka -i https://pypi.mirrors.ustc.edu.cn/simple/
if errorlevel 1 (
    echo 依赖安装失败,请检查网络或 pip 镜像后重试。
    pause
    exit /b 1
)

echo [2/3] 使用 Nuitka 打包(onefile 单文件模式,首次打包较慢)...
rem 说明:Nuitka 会自动分析 _psycopg.pyd 的 DLL 依赖,把 psycopg2.libs 里的
rem libpq/libssl/libcrypto 一并打进 exe;pyside6 插件负责携带 Qt 运行库与插件。

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
    echo 打包失败。常见解决办法:
    echo   1. 卸载 onefile:把上面命令中的 --onefile 删除后重试,会生成 build\DataTransformer.dist\DataTransformer.exe
    echo   2. 确认已安装 Visual Studio Build Tools(C++ 桌面开发)或 MinGW64
    pause
    exit /b 1
)

echo [3/3] 打包完成: build\DataTransformer.exe
pause
