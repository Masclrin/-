@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo [1/2] 安装 Python 依赖...
pip install -r requirements.txt
if errorlevel 1 (
  echo.
  echo 依赖安装失败，请确认已安装 Python 并勾选 Add Python to PATH。
  pause
  exit /b 1
)
echo.
echo 依赖安装完成。
pause
