@echo off
chcp 65001 > nul
setlocal enabledelayedexpansion

echo ==========================================
echo 打包前检查：要求 Python 3.8（用于生成兼容 Win7 的可执行）
echo ==========================================
echo.

REM 优先尝试 py -3.8
set "PYTHON_EXE="
for /f "usebackq tokens=*" %%i in (`where py 2^>nul`) do set HAVE_PY=1
if defined HAVE_PY (
    for /f "usebackq tokens=*" %%v in (`py -3.8 -c "import sys; print(sys.executable)" 2^>nul`) do set "PYTHON_EXE=%%v"
)

REM 若未找到 py -3.8，再检测系统 python 是否为 3.8
if not defined PYTHON_EXE (
    for /f "usebackq tokens=*" %%v in (`python -c "import sys; print('%d.%d' % sys.version_info[:2])" 2^>nul`) do set PYVER=%%v
    if defined PYVER (
        if "%PYVER:~0,3%"=="3.8" (
            set "PYTHON_EXE=python"
        )
    )
)

if not defined PYTHON_EXE (
    echo 错误：未检测到 Python 3.8。请安装 Python 3.8 并确保可通过 "py -3.8" 或 "python" 调用。
    echo 如果系统上只有其它版本 Python，请在目标打包机上安装 Python 3.8 后重试。
    pause
    exit /b 1
)

echo 使用 Python 可执行路径： %PYTHON_EXE%
echo.

REM 创建虚拟环境（使用指定 Python 3.8）
if not exist venv (
    echo 正在用 %PYTHON_EXE% 创建虚拟环境...
    "%PYTHON_EXE%" -m venv venv
    if errorlevel 1 (
        echo 错误：创建虚拟环境失败
        pause
        exit /b 1
    )
)

REM 激活虚拟环境
call venv\Scripts\activate.bat
if errorlevel 1 (
    echo 错误：激活虚拟环境失败
    pause
    exit /b 1
)

REM 升级 pip/setuptools/wheel（尽量用与 Python3.8 兼容的版本）
echo.
echo 升级 pip、setuptools、wheel...
python -m pip install --upgrade pip setuptools wheel
echo.

REM 安装依赖
echo 安装 requirements.txt 中的依赖（请确保 requirements.txt 针对 Python3.8）...
if not exist requirements.txt (
    echo 错误：requirements.txt 不存在，请在脚本目录创建并指定兼容的版本（例如 openpyxl==3.10.10 等）
    pause
    exit /b 1
)
pip install -r requirements.txt
if errorlevel 1 (
    echo 错误：依赖安装失败，请检查 requirements.txt 中的包版本是否支持 Python 3.8
    echo 建议：在命令行中手动运行 "pip install -r requirements.txt" 看具体错误并调整版本
    pause
    exit /b 1
)

REM 安装 PyInstaller（固定兼容版本）
echo.
echo 安装 PyInstaller==5.13.2...
pip install PyInstaller==5.13.2
if errorlevel 1 (
    echo 警告：PyInstaller 安装失败，尝试使用 cx_Freeze 备选方案...
    goto use_cx_freeze
)

REM 检查待打包脚本
set "SRC=下载抓鱼单v3.0_win7兼容下载打印.py"
if not exist "%SRC%" (
    echo 错误：未找到 %SRC%
    pause
    exit /b 1
)

REM PyInstaller 打包
echo.
echo 正在用 PyInstaller 打包...
PyInstaller --onefile ^
    --windowed ^
    --hidden-import=openpyxl ^
    --hidden-import=playwright ^
    --hidden-import=requests ^
    --exclude-module=numpy ^
    --exclude-module=pandas ^
    --name "抓鱼单导出工具" ^
    --distpath "dist" ^
    "%SRC%"

if errorlevel 1 (
    echo 警告：PyInstaller 打包失败，尝试使用 cx_Freeze...
    goto use_cx_freeze
)

goto pack_success

:use_cx_freeze
echo.
echo 使用 cx_Freeze 进行打包（备选）...
pip install cx_Freeze==6.15.10
if errorlevel 1 (
    echo 错误：cx_Freeze 安装失败，打包终止
    pause
    exit /b 1
)

cxfreeze --target-dir dist "%SRC%" --base-name Win32GUI
if errorlevel 1 (
    echo 错误：cx_Freeze 打包失败
    pause
    exit /b 1
)

:pack_success
echo.
echo ==========================================
echo 打包完成！
echo ==========================================
if exist "dist\抓鱼单导出工具.exe" (
    echo ✓ 主程序生成成功：dist\抓鱼单导出工具.exe
) else (
    echo ⚠ 未检测到目标 exe，请检查打包日志
)
echo.
echo 输出目录：%cd%\dist\
echo 分发时请包含：dist\抓鱼单导出工具.exe 及说明文档/requirements.txt
echo.
pause