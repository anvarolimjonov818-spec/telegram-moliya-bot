# 🌐 Telegram Botni 24/7 Cloud Serverga Tekin Joylashtirish

Sizning AI Telegram botingiz kompyuteringiz o'chiq bo'lganda ham 24 soat uzluksiz ishlashi uchun barcha fayllar va sozlamalar tayyorlandi.

---

## 📦 Loyiha Fayllari:
1. **`sheets_ai_bot.py`** — Ovozli xabarlarni (Whisper API) eshitib, AI orqali mos ustunlar va oylik varaqlarga joylovchi bot kodi.
2. **`config.py`** — Telegram Bot Token, Gemini API va Groq API kalitlari.
3. **`requirements.txt`** — Server uchun zaruriy Python kutubxonalari (`python-telegram-bot`, `openpyxl`, `aiohttp`).
4. **`Procfile`** — 24/7 Background Worker moduli.

---

## 🚀 1-USUL: Render.com da 24/7 Tekin Serverga Ulash (Eng Tavsiya Etilgan)

1. [GitHub.com](https://github.com) ga kiring va **"New repository"** tugmasini bosing (masalan, `my-sheets-bot`).
2. Telegrambot papkasidagi fayllarni yuklang.
3. [Render.com](https://render.com) ga kirib GitHub orqali ro'yxatdan o'ting.
4. **New +** -> **Background Worker** tanlang.
5. GitHub reponizni tanlab quyidagi sozlamalarni kiriting:
   - **Name**: `my-sheets-bot`
   - **Build Command**: `pip install -r requirements.txt`
   - **Start Command**: `python sheets_ai_bot.py`
6. **Create Background Worker** tugmasini bosing.

 Bot 24/7 rejimida serverda ishga tushadi!
