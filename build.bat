@echo off
chcp 65001 >nul
setlocal

cd /d "%~dp0"

echo ========================================
echo   客户产品 Excel 汇总工具 - 打包脚本
echo ========================================
echo.

where py >nul 2>nul
if %errorlevel%==0 (
    set PYTHON=py -3.10
) else (
    set PYTHON=python
)

echo [1/3] 安装打包依赖...
%PYTHON% -m pip install -r requirements-build.txt -q
if errorlevel 1 (
    echo 依赖安装失败，请确认已安装 Python 3.10+ 且 pip 可用。
    pause
    exit /b 1
)

echo [2/3] 开始打包（约 1-3 分钟）...
%PYTHON% -m PyInstaller build.spec --noconfirm --clean
if errorlevel 1 (
    echo 打包失败，请查看上方错误信息。
    pause
    exit /b 1
)

echo.
echo [3/3] 打包完成！
echo.
echo 程序目录：
echo   dist\客户产品Excel汇总工具\
echo.
echo 发给同事时，请将整个文件夹压缩为 zip 发送。
echo 同事解压后，双击「客户产品Excel汇总工具.exe」即可使用。
echo 无需安装 Python。
echo.
pause
