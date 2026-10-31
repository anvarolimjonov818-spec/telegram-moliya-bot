@echo off
chcp 65001 >nul
title AI Telegram Bot - Setup

echo.
echo ╔══════════════════════════════════════════════╗
echo ║        AI Telegram Bot - Sozlash            ║
echo ╚══════════════════════════════════════════════╝
echo.

REM Python yo'lini yangilash
set "PYTHON_PATH="
for %%i in (python.exe) do set "PYTHON_PATH=%%~$PATH:i"

if "%PYTHON_PATH%"=="" (
    REM Standart o'rnatish joylarini tekshirish
    if exist "%LOCALAPPDATA%\Programs\Python\Python312\python.exe" (
        set "PYTHON_PATH=%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
    ) else if exist "C:\Python312\python.exe" (
        set "PYTHON_PATH=C:\Python312\python.exe"
    ) else (
        echo [XATO] Python topilmadi!
        echo Python yuklab oling: https://python.org
        pause
        exit /b 1
    )
)

echo [OK] Python topildi.
echo.
echo [1/3] Kerakli paketlar o'rnatilmoqda...
"%PYTHON_PATH%" -m pip install -r requirements.txt --quiet --upgrade
if errorlevel 1 (
    echo [XATO] Paketlar o'rnatilmadi!
    pause
    exit /b 1
)
echo [OK] Paketlar o'rnatildi.

echo.
echo [2/3] Gemini API Key tekshirilmoqda...
"%PYTHON_PATH%" -c "from config import GEMINI_API_KEY; exit(0 if GEMINI_API_KEY != 'YOUR_GEMINI_API_KEY' else 1)" 2>nul
if errorlevel 1 (
    echo.
    echo ┌─────────────────────────────────────────────┐
    echo │  GEMINI API KEY KERAK!                      │
    echo │                                             │
    echo │  1. https://aistudio.google.com/apikey ga  │
    echo │     boring (Google akkaunt bilan kirish)    │
    echo │  2. "Create API key" tugmasini bosing       │
    echo │  3. Kalitni nusxalang                       │
    echo │  4. config.py faylida:                      │
    echo │     GEMINI_API_KEY = "AIza..."              │
    echo │     qatorini to'ldiring                     │
    echo └─────────────────────────────────────────────┘
    echo.
    echo config.py faylini ochamiz...
    timeout /t 2 >nul
    start notepad config.py
    echo.
    echo config.py ni saqlagan vaqtda Enter bosing...
    pause
)

echo.
echo [3/3] Bot ishga tushirilmoqda...
echo.
echo ╔══════════════════════════════════════════════╗
echo ║  Bot ishlayapti! Telegramda /start bosing   ║
echo ║  To'xtatish uchun: Ctrl+C                  ║
echo ╚══════════════════════════════════════════════╝
echo.

"%PYTHON_PATH%" sheets_ai_bot.py

if errorlevel 1 (
    echo.
    echo [XATO] Bot to'xtadi. Xato uchun yuqoridagi xabarlarni o'qing.
    pause
)
