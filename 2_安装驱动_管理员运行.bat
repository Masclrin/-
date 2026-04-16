@echo off
chcp 65001 >nul
cd /d "%~dp0\Interception\command line installer"
echo [1/2] 安装 Interception 驱动...
install-interception.exe /install
if errorlevel 1 (
  echo.
  echo 驱动安装失败。请右键本文件，选择“以管理员身份运行”。
  pause
  exit /b 1
)
echo.
echo 驱动安装完成，请重启电脑后再启动宏。
pause
