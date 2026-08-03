@echo off
chcp 65001 >nul
setlocal EnableDelayedExpansion

echo ===========================================
echo   本地文档脱敏工具 - 一键打包脚本
echo ===========================================
echo.

:: 检查 Python
python --version >nul 2>&1
if errorlevel 1 (
    echo [错误] 未找到 python 命令。
    echo 请先安装 Python 3.10+，并在安装时勾选 "Add Python to PATH"。
    pause
    exit /b 1
)

python -c "import sys; sys.exit(0 if sys.version_info >= (3,10) else 1)"
if errorlevel 1 (
    echo [错误] Python 版本过低，需要 3.10 或更高版本。
    pause
    exit /b 1
)

echo [1/4] Python 版本检查通过。

:: 创建构建虚拟环境
set VENV_DIR=%~dp0build-venv
if not exist "%VENV_DIR%" (
    echo [2/4] 正在创建打包专用虚拟环境...
    python -m venv "%VENV_DIR%"
) else (
    echo [2/4] 复用已有打包虚拟环境。
)

call "%VENV_DIR%\Scripts\activate.bat"

:: 安装依赖
echo [3/4] 安装打包依赖...
python -m pip install --upgrade pip -q
python -m pip install -r requirements.txt -q
if errorlevel 1 (
    echo [错误] 依赖安装失败，请检查网络连接。
    pause
    exit /b 1
)

:: 生成图标
echo [3/4] 生成应用图标...
python tools\make_icon.py

:: 执行 PyInstaller
echo [4/4] 开始打包，请稍候...
python -m PyInstaller --noconfirm --clean masktool_gui.spec
if errorlevel 1 (
    echo [错误] PyInstaller 打包失败。
    pause
    exit /b 1
)

echo.
echo ===========================================
echo   打包成功！
echo   可执行文件位于：dist\本地文档脱敏工具.exe
echo ===========================================

:: 打开输出目录
if exist "dist" start explorer "dist"

pause
