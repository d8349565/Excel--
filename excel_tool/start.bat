@echo off
chcp 65001 >nul 2>&1

echo ================================================
echo Excel/CSV 汇总工具启动脚本
echo ================================================
echo.

echo 正在检查 Python 环境...
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo 错误: 未找到 Python，请确保已安装 Python 3.9+
    pause
    exit /b 1
)

echo 正在检查依赖包...
python -c "import flask, pandas, openpyxl" >nul 2>&1
if %errorlevel% neq 0 (
    echo 正在安装依赖包...
    pip install -r requirements.txt
    if %errorlevel% neq 0 (
        echo 错误: 依赖包安装失败
        pause
        exit /b 1
    )
)

echo.
echo 启动 Excel/CSV 汇总工具...
echo 浏览器访问地址: http://localhost:5000
echo 按 Ctrl+C 停止服务
echo.

cd excel_tool
python app.py

pause