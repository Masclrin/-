@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo 启动宏执行框架...
python "宏执行框架2.1.py"
if errorlevel 1 (
  echo.
  echo 启动失败，请先运行 1_安装依赖.bat。
  pause
)
