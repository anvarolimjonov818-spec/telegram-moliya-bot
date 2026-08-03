#!/usr/bin/env python3
# ============================================================
#  AI Telegram Bot — Multimodal (Matn + Rasm qo'llab-quvvatlash)
#  - Matnli va Rasmli so'rovlar
#  - Gemini 2.0 Flash (Multimodal Vision) + Groq (fallback)
#  - Pollinations.ai rasm generatsiyasi
# ============================================================

import asyncio
import logging
import re
import random
import traceback
import base64
import aiohttp

from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    Application, CommandHandler, MessageHandler,
    CallbackQueryHandler, ContextTypes, filters
)
from telegram.constants import ChatAction

from config import (
    BOT_TOKEN, GEMINI_API_KEY, GROQ_API_KEY,
    BOT_NAME, MAX_HISTORY, IMAGE_WIDTH, IMAGE_HEIGHT
)

# ─── Logging ────────────────────────────────────────────────
logging.basicConfig(
    format="%(asctime)s | %(levelname)s | %(message)s",
    level=logging.INFO,
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler("bot.log", encoding="utf-8"),
    ]
)
logger = logging.getLogger(__name__)

# ─── API URLs ────────────────────────────────────────────────
GEMINI_URL = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent"
GROQ_URL = "https://api.groq.com/openai/v1/chat/completions"

SYSTEM_PROMPT = """Siz O'zbek tilidagi aqlli AI yordamchisiz — "{name}".
Quyidagi qoidalarga rioya qiling:
1. Asosan O'zbek tilida javob bering. Foydalanuvchi boshqa tilda yozsa, shu tilda javob bering.
2. Do'stona, samimiy va foydali bo'ling.
3. Foydalanuvchi rasm yuborib fonini o'zgartirish, chizish yoki rasm yaratishni so'rasa —
   javobingizda tushuntirish bering va OXIRIDA alohida qatorda faqat:
   IMAGE_PROMPT: <inglizcha juda batafsil, chiroyli professional rasm tavsifi>
   deb yozing.
4. Oddiy va tushunarli formatdan foydalaning.""".format(name=BOT_NAME)

# ─── Foydalanuvchi tarixi ────────────────────────────────────
user_histories: dict[int, list] = {}
user_modes: dict[int, str] = {}

IMAGE_KEYWORDS_UZ = [
    'rasm', 'surat', 'chiz', "ko'rsatib ber", "ko'rsat",
    'tasvirla', 'yasab ber', 'yarat', 'rasm yasa', 'surat yasa',
    'rasmini', 'suratini', 'chiqar', 'generate', 'image', 'draw',
    'fonini', 'orqa fon', 'fon', 'background'
]

def is_image_request(text: str) -> bool:
    lower = text.lower()
    return any(kw in lower for kw in IMAGE_KEYWORDS_UZ)

# ─── Gemini Multimodal API (Matn + Rasm) ────────────────────
async def call_gemini(history: list, image_b64: str = None) -> str:
    """Gemini API ga matn yoki rasm bilan so'rov"""
    contents = []
    
    for h in history:
        role = "model" if h["role"] in ("assistant", "model") else "user"
        contents.append({"role": role, "parts": [{"text": h["content"]}]})

    # Agar rasm bo'lsa, so'nggi xabarga rasm ma'lumotini qo'shamiz
    if image_b64 and contents:
        contents[-1]["parts"].append({
            "inline_data": {
                "mime_type": "image/jpeg",
                "data": image_b64
            }
        })

    body = {
        "system_instruction": {"parts": [{"text": SYSTEM_PROMPT}]},
        "contents": contents,
        "generationConfig": {
            "temperature": 0.85,
            "topK": 40,
            "topP": 0.95,
            "maxOutputTokens": 2048,
        }
    }

    headers = {
        "Content-Type": "application/json",
        "x-goog-api-key": GEMINI_API_KEY,
    }

    async with aiohttp.ClientSession() as session:
        async with session.post(
            GEMINI_URL, json=body, headers=headers,
            timeout=aiohttp.ClientTimeout(total=45)
        ) as resp:
            if resp.status == 429:
                raise RateLimitError("Gemini limit tugadi")
            if resp.status != 200:
                text = await resp.text()
                raise Exception(f"Gemini {resp.status}: {text[:150]}")
            data = await resp.json()
            return data["candidates"][0]["content"]["parts"][0]["text"]

# ─── Groq API ────────────────────────────────────────────────
async def call_groq(history: list, image_b64: str = None) -> str:
    """Groq API ga so'rov"""
    messages = [{"role": "system", "content": SYSTEM_PROMPT}]
    for h in history:
        messages.append({"role": h["role"], "content": h["content"]})

    body = {
        "model": "llama-3.3-70b-versatile",
        "messages": messages,
        "temperature": 0.85,
        "max_tokens": 2048,
    }

    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {GROQ_API_KEY}",
    }

    async with aiohttp.ClientSession() as session:
        async with session.post(
            GROQ_URL, json=body, headers=headers,
            timeout=aiohttp.ClientTimeout(total=30)
        ) as resp:
            if resp.status == 429:
                raise RateLimitError("Groq limit tugadi")
            if resp.status != 200:
                text = await resp.text()
                raise Exception(f"Groq {resp.status}: {text[:150]}")
            data = await resp.json()
            return data["choices"][0]["message"]["content"]

class RateLimitError(Exception):
    pass

# ─── Aqlli AI Chaqirish ──────────────────────────────────────
async def ask_ai(user_id: int, text: str, image_b64: str = None) -> str:
    """Gemini ga so'rov (rasm bilan yoki rasmsiz). Limit bo'lsa Groq ga o'tadi."""
    if user_id not in user_histories:
        user_histories[user_id] = []

    history = user_histories[user_id]
    history.append({"role": "user", "content": text})

    if len(history) > MAX_HISTORY:
        history = history[-MAX_HISTORY:]
        user_histories[user_id] = history

    providers = []
    if GEMINI_API_KEY:
        providers.append(("Gemini", call_gemini))
    if GROQ_API_KEY and not image_b64:  # Groq text-only fallback
        providers.append(("Groq", call_groq))

    if not providers:
        history.pop()
        raise Exception("Hech bir API key o'rnatilmagan! config.py ni tekshiring.")

    last_error = ""
    for name, call_fn in providers:
        try:
            logger.info(f"  → {name} ga so'rov yuborilmoqda (Rasm: {'Ha' if image_b64 else 'Yo\'q'})...")
            reply = await call_fn(history, image_b64=image_b64)
            logger.info(f"  ✅ {name} javob berdi")
            history.append({"role": "assistant", "content": reply})
            return reply
        except RateLimitError:
            logger.warning(f"  ⚠️ {name} limit — keyingisiga o'tilmoqda...")
            last_error = f"{name} limiti tugadi"
            continue
        except Exception as e:
            logger.error(f"  ❌ {name} xato: {e}")
            last_error = str(e)
            continue

    history.pop()
    if "limit" in last_error.lower():
        raise Exception(f"⏳ API limiti tugadi. Birozdan so'ng qayta urinib ko'ring.")
    raise Exception(f"❌ Xato: {last_error[:200]}")

# ─── Rasm generatsiya ────────────────────────────────────────
def extract_image_prompt(ai_text: str) -> tuple[str, str]:
    match = re.search(r'IMAGE_PROMPT:\s*(.+)', ai_text, re.IGNORECASE)
    if match:
        img_prompt = match.group(1).strip()
        clean_text = ai_text[:match.start()].strip()
        return clean_text, img_prompt
    return ai_text, ""

async def generate_image_url(prompt: str) -> str:
    seed = random.randint(1, 99999)
    encoded = prompt.replace(' ', '%20').replace(',', '%2C')
    return (
        f"https://image.pollinations.ai/prompt/{encoded}"
        f"?width={IMAGE_WIDTH}&height={IMAGE_HEIGHT}"
        f"&seed={seed}&nologo=true&enhance=true&model=flux"
    )

async def download_image(url: str) -> bytes | None:
    try:
        async with aiohttp.ClientSession() as session:
            async with session.get(url, timeout=aiohttp.ClientTimeout(total=60)) as resp:
                if resp.status == 200:
                    return await resp.read()
    except Exception as e:
        logger.error(f"Rasm yuklab olishda xato: {e}")
    return None

# ─── Telegram Handlerlari ────────────────────────────────────
async def safe_reply(message, text, **kwargs):
    try:
        return await message.reply_text(text, **kwargs)
    except Exception:
        try:
            return await message.reply_text(text)
        except Exception as e:
            logger.error(f"Xabar yuborishda xato: {e}")
            return await message.reply_text("Xabar yuborishda xato yuz berdi.")

async def start_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    user_modes[user.id] = 'auto'

    active = []
    if GEMINI_API_KEY:
        active.append("Gemini 2.0 Flash")
    if GROQ_API_KEY:
        active.append("Groq")
    providers_text = " + ".join(active) if active else "❌ Hech biri"

    keyboard = [
        [
            InlineKeyboardButton("💬 Suhbat", callback_data="mode_text"),
            InlineKeyboardButton("🎨 Rasm", callback_data="mode_image"),
            InlineKeyboardButton("✨ Avto", callback_data="mode_auto"),
        ],
        [InlineKeyboardButton("❓ Yordam", callback_data="help")],
    ]

    welcome = (
        f"👋 Salom, {user.first_name}!\n\n"
        f"Men {BOT_NAME} — sizning aqlli yordamchingizman.\n\n"
        "🔹 Har qanday savolga javob beraman\n"
        "🔹 Rasm va suratlar yarataman\n"
        "🔹 *Yangi:* Rasm yuborsangiz, uni tahlil qilaman va orqa fonini yangilab chizib beraman!\n\n"
        f"⚡ AI: {providers_text}\n\n"
        "Quyidan rejimni tanlang yoki savol / rasm yuboring:"
    )
    await safe_reply(update.message, welcome, reply_markup=InlineKeyboardMarkup(keyboard))

async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = (
        "📖 Buyruqlar va Imkoniyatlar:\n\n"
        "/start — Botni qayta ishga tushirish\n"
        "/help — Ushbu yordam\n"
        "/rasm <tavsif> — Rasm yaratish\n"
        "/tozala — Suhbat tarixini o'chirish\n"
        "/rejim — Joriy rejimni ko'rsatish\n\n"
        "🖼 *Rasmlar bilan ishlash:*\n"
        "Botga mahsulot yoki ob'ekt rasmini yuboring va caption (izoh) sifatida:\n"
        '"Shu rasmni orqa fonini chiroyli qilib moslab ber"\n'
        "deb yozing. Bot rasmni tahlil qilib, unga mos yangi professional rasm yaratadi!\n"
    )
    await safe_reply(update.message, text)

async def clear_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    user_histories.pop(user_id, None)
    await safe_reply(update.message, "✅ Suhbat tarixi tozalandi. Yangi suhbat boshlang!")

async def mode_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    mode = user_modes.get(user_id, 'auto')
    mode_names = {'auto': '✨ Avtomatik', 'text': '💬 Suhbat', 'image': '🎨 Rasm'}
    keyboard = [[
        InlineKeyboardButton("💬 Suhbat", callback_data="mode_text"),
        InlineKeyboardButton("🎨 Rasm", callback_data="mode_image"),
        InlineKeyboardButton("✨ Avto", callback_data="mode_auto"),
    ]]
    await safe_reply(
        update.message,
        f"⚙️ Joriy rejim: {mode_names.get(mode, mode)}\n\nRejimni o'zgartirish:",
        reply_markup=InlineKeyboardMarkup(keyboard)
    )

async def rasm_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    prompt = ' '.join(context.args)
    if not prompt:
        await safe_reply(update.message, "❗ Rasm tavsifini kiriting.\nMasalan: /rasm Toshkent kechasi")
        return
    await handle_image_generation(update, context, prompt, user_prompt=prompt)

async def handle_image_generation(update, context, eng_prompt, user_prompt="", caption_text=""):
    msg = await update.message.reply_text("🎨 Rasm yaratilmoqda, kuting...")
    await context.bot.send_chat_action(update.effective_chat.id, ChatAction.UPLOAD_PHOTO)
    try:
        img_url = await generate_image_url(eng_prompt)
        img_data = await download_image(img_url)
        if img_data:
            await msg.delete()
            display_text = user_prompt or eng_prompt[:80]
            caption = caption_text or f"🎨 {display_text}"
            short_prompt = eng_prompt[:180]
            keyboard = [[InlineKeyboardButton("🔄 Qayta yaratish", callback_data=f"regen:{short_prompt}")]]
            await update.message.reply_photo(photo=img_data, caption=caption, reply_markup=InlineKeyboardMarkup(keyboard))
        else:
            await msg.edit_text("❌ Rasm yaratishda xato. Qayta urinib ko'ring.")
    except Exception as e:
        logger.error(f"Rasm xatosi: {e}")
        try:
            await msg.edit_text(f"❌ Xato: {str(e)[:100]}")
        except:
            pass

# ─── RASM (PHOTO) HANDLERI ────────────────────────────────────
async def handle_photo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Foydalanuvchi rasm yuborganda ishlaydi"""
    if not update.message or not update.message.photo:
        return

    user = update.effective_user
    user_id = user.id
    caption = (update.message.caption or "").strip()

    if not caption:
        caption = "Ushbu rasmni tahlil qil va unga mos professional orqa fon bilan yangi rasm tavsifini yarat."

    logger.info(f"User {user_id} rasm yubordi. Izoh: {caption[:60]}")
    await context.bot.send_chat_action(update.effective_chat.id, ChatAction.TYPING)

    msg = await update.message.reply_text("🔍 Rasm tahlil qilinmoqda...")

    try:
        # Eng yuqori sifatli rasmni yuklab olamiz
        photo_file = await update.message.photo[-1].get_file()
        photo_bytes = await photo_file.download_as_bytearray()
        image_b64 = base64.b64encode(photo_bytes).decode('utf-8')

        prompt = (
            f"Foydalanuvchi rasm va quyidagi so'rovni yubordi: '{caption}'. "
            "Rasmdagi mahsulot/ob'ektni tahlil qil va foydalanuvchining so'roviga mos professional, o'ta jozibali yangi rasm uchun "
            "inglizcha tavsif yoz. Javobing oxirida alohida qatorda: IMAGE_PROMPT: <inglizcha tavsif> deb yoz."
        )

        reply = await ask_ai(user_id, prompt, image_b64=image_b64)
        clean_text, img_prompt = extract_image_prompt(reply)

        await msg.delete()

        if clean_text:
            await send_long_message(update, clean_text)

        if img_prompt:
            await context.bot.send_chat_action(update.effective_chat.id, ChatAction.UPLOAD_PHOTO)
            await handle_image_generation(
                update, context,
                eng_prompt=img_prompt,
                user_prompt=caption,
                caption_text=f"🎨 {caption[:60]}"
            )
        else:
            # Agar AI IMAGE_PROMPT bermasa, avto prompt shakllantirish
            fallback_prompt = f"Professional product photography of the item in image, {caption}, studio lighting, highly detailed 8k"
            await handle_image_generation(update, context, eng_prompt=fallback_prompt, user_prompt=caption)

    except Exception as e:
        logger.error(f"Rasm qayta ishlashda xato: {e}\n{traceback.format_exc()}")
        try:
            await msg.edit_text(f"❌ Rasm tahlilida xato: {str(e)[:200]}")
        except:
            await safe_reply(update.message, f"❌ Xato: {str(e)[:200]}")

# ─── MATN (TEXT) HANDLERI ────────────────────────────────────
async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not update.message or not update.message.text:
        return

    user = update.effective_user
    text = update.message.text.strip()
    user_id = user.id
    mode = user_modes.get(user_id, 'auto')

    logger.info(f"User {user_id} ({user.first_name}): {text[:60]}")

    # Rasm rejimi
    if mode == 'image':
        await context.bot.send_chat_action(update.effective_chat.id, ChatAction.TYPING)
        try:
            ai_reply = await ask_ai(user_id, f"Convert this to an English image generation prompt (only the prompt, nothing else): {text}")
            img_prompt = ai_reply.strip().strip('"').strip("'")
        except Exception:
            img_prompt = text
        await handle_image_generation(update, context, img_prompt, user_prompt=text)
        return

    # Matn rejimi
    if mode == 'text':
        await context.bot.send_chat_action(update.effective_chat.id, ChatAction.TYPING)
        try:
            reply = await ask_ai(user_id, text)
            await send_long_message(update, reply)
        except Exception as e:
            await safe_reply(update.message, f"❌ {str(e)[:300]}")
        return

    # Avtomatik rejim
    await context.bot.send_chat_action(update.effective_chat.id, ChatAction.TYPING)
    try:
        reply = await ask_ai(user_id, text)
        clean_text, img_prompt = extract_image_prompt(reply)

        if clean_text:
            await send_long_message(update, clean_text)

        if img_prompt:
            await context.bot.send_chat_action(update.effective_chat.id, ChatAction.UPLOAD_PHOTO)
            await handle_image_generation(update, context, eng_prompt=img_prompt, user_prompt=text, caption_text=f"🎨 {text[:60]}")
        elif is_image_request(text) and not img_prompt:
            fallback_prompt = f"beautiful illustration of: {text}, high quality, artistic"
            await handle_image_generation(update, context, eng_prompt=fallback_prompt, user_prompt=text)

    except Exception as e:
        logger.error(f"Handler xatosi: {e}")
        await safe_reply(update.message, f"❌ {str(e)[:300]}")

async def send_long_message(update, text):
    max_len = 4000
    if len(text) <= max_len:
        await safe_reply(update.message, text)
        return
    parts = []
    while len(text) > max_len:
        split_at = text.rfind('\n', 0, max_len)
        if split_at == -1:
            split_at = max_len
        parts.append(text[:split_at])
        text = text[split_at:].lstrip()
    if text:
        parts.append(text)
    for part in parts:
        await safe_reply(update.message, part)

async def button_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    user_id = query.from_user.id
    data = query.data

    if data.startswith("mode_"):
        mode = data.replace("mode_", "")
        user_modes[user_id] = mode
        mode_names = {'auto': '✨ Avtomatik', 'text': '💬 Suhbat', 'image': '🎨 Rasm'}
        try:
            await query.edit_message_text(f"✅ Rejim: {mode_names.get(mode, mode)}\n\nEndi xabar yoki rasm yuboring!")
        except:
            pass

    elif data == "help":
        try:
            await query.edit_message_text(
                "📖 Yordam\n\n💬 Suhbat — savolga javob\n🎨 Rasm — rasm yaratadi\n🖼 Rasm yuboring — bot uni tahlil qilib yangi fon chizadi!\n\nSavolingizni yozing!"
            )
        except:
            pass

    elif data.startswith("regen:"):
        prompt = data[6:]
        msg = await query.message.reply_text("🎨 Qayta yaratilmoqda...")
        try:
            img_url = await generate_image_url(prompt)
            img_data = await download_image(img_url)
            if img_data:
                await msg.delete()
                keyboard = [[InlineKeyboardButton("🔄 Qayta yaratish", callback_data=f"regen:{prompt}")]]
                await query.message.reply_photo(photo=img_data, caption="🎨 Yangi variant", reply_markup=InlineKeyboardMarkup(keyboard))
            else:
                await msg.edit_text("❌ Xato yuz berdi.")
        except Exception as e:
            try:
                await msg.edit_text(f"❌ Xato: {str(e)[:100]}")
            except:
                pass

async def error_handler(update, context: ContextTypes.DEFAULT_TYPE):
    logger.error(f"Xato: {context.error}")

# ─── Main ───────────────────────────────────────────────────
def main():
    logger.info("🤖 Bot ishga tushmoqda...")

    active = []
    if GEMINI_API_KEY:
        active.append(f"Gemini ({GEMINI_API_KEY[:8]}...)")
    if GROQ_API_KEY:
        active.append(f"Groq ({GROQ_API_KEY[:8]}...)")

    if active:
        logger.info(f"⚡ AI provayderlar: {', '.join(active)}")
    else:
        logger.error("❌ Hech bir API key o'rnatilmagan! config.py ni tekshiring.")

    app = Application.builder().token(BOT_TOKEN).build()

    app.add_handler(CommandHandler("start", start_command))
    app.add_handler(CommandHandler("help", help_command))
    app.add_handler(CommandHandler("yordam", help_command))
    app.add_handler(CommandHandler("tozala", clear_command))
    app.add_handler(CommandHandler("rasm", rasm_command))
    app.add_handler(CommandHandler("rejim", mode_command))
    app.add_handler(CallbackQueryHandler(button_callback))
    
    # Text va Photo handlerlar
    app.add_handler(MessageHandler(filters.PHOTO, handle_photo))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))
    
    app.add_error_handler(error_handler)

    logger.info("✅ Multimodal Bot tayyor! Telegram da /start bosing.")
    app.run_polling(drop_pending_updates=True)

if __name__ == "__main__":
    main()
