@echo off
chcp 65001 >nul
title n8n Avtomatizatsiya Tizimini Ishga Tushirish

echo.
echo ╔══════════════════════════════════════════════════════════╗
echo ║         n8n Avtomatizatsiya Tizimi (Local)               ║
echo ╚══════════════════════════════════════════════════════════╝
echo.

set "NODE_PATH=C:\Program Files\nodejs"
if exist "%NODE_PATH%\npx.cmd" (
    set PATH=%NODE_PATH%;%PATH%
)

echo [1/2] n8n ishga tushirilmoqda...
echo Brauzerda http://localhost:5678 manzili ochiladi.
echo.

npx n8n

if errorlevel 1 (
    echo [XATO] n8n ishga tushmadi. Node.js o'rnatilganini tekshiring.
    pause
)
