# ============================================================
#  AI Telegram Bot — Konfiguratsiya (.env orqali xavfsiz yuklash)
# ============================================================
import os
from pathlib import Path
from dotenv import load_dotenv

# Loyihaning asosiy papkasi (.env faylini joriy va yuqori papkadan qidiradi)
BASE_DIR = Path(__file__).resolve().parent
ENV_PATH = BASE_DIR / ".env"

# .env faylini yuklaymiz (agar mavjud bo'lsa)
load_dotenv(dotenv_path=ENV_PATH)
# Yuqori papkadagi .env ham bo'lishi mumkin (fallback)
load_dotenv(dotenv_path=BASE_DIR.parent / ".env")

# ─── ASOSIY MAXFIY KALITLAR (Faqat .env yoki Muhit O'zgaruvchilaridan) ───
BOT_TOKEN = os.getenv("BOT_TOKEN", "").strip()
GROQ_API_KEY = os.getenv("GROQ_API_KEY", "").strip()
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "").strip()
GCP_API_KEY = os.getenv("GCP_API_KEY", "").strip()

# ─── QO'SHIMCHA SOZLAMALAR ───
BOT_NAME = os.getenv("BOT_NAME", "Yuksalish Bot")
MAX_HISTORY = int(os.getenv("MAX_HISTORY", "10"))
IMAGE_WIDTH = int(os.getenv("IMAGE_WIDTH", "1024"))
IMAGE_HEIGHT = int(os.getenv("IMAGE_HEIGHT", "1024"))
DEFAULT_USER_ID = int(os.getenv("DEFAULT_USER_ID", "1453337970"))
