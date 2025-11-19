@echo off
chcp 65001 >nul
ECHO =========================================
ECHO 健康数据分析工具 v1.0
ECHO =========================================
ECHO.

REM 检查Python是否已安装
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo 错误: 未检测到Python，请先安装Python 3.8+
    pause
    exit /b 1
)

echo 正在启动健康数据分析工具...
echo.

REM 启动主程序 (使用python.exe显示控制台窗口，便于调试)
python.exe main.py

echo.
echo 程序已关闭。
pause
