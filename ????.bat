@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo 正在启动博客...
start "" http://127.0.0.1:5178
node server.js
pause
