#!/usr/bin/env python3
# ============================================================
#  AI Voice & Text Note to Professional Excel (.xlsx) Bot
#  - Ovozli va Matnli xabarlarni tahlil qilish
#  - Gemini + Groq (Llama 3.3 70B) bilan strukturaga ajratish
#  - Avtomatik OYMA-OY alohida varaqlar (Worksheets) yaratish (2026-08 (Avgust)...)
#  - Kategoriyalar bo'yicha Oylik JAMI Hisobotlar
#  - Istalgan qatorni ko'rish va o'chirish
# ============================================================

import asyncio
import logging
import re
import json
import csv
import os
import datetime
import base64
import aiohttp

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

from telegram import (
    Update, InlineKeyboardButton, InlineKeyboardMarkup,
    ReplyKeyboardMarkup, KeyboardButton
)
from telegram.ext import (
    Application, CommandHandler, MessageHandler,
    CallbackQueryHandler, ContextTypes, filters
)
from telegram.constants import ChatAction

from config import BOT_TOKEN, GEMINI_API_KEY, GROQ_API_KEY

logging.basicConfig(
    format="%(asctime)s | %(levelname)s | %(message)s",
    level=logging.INFO,
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler("sheets_bot.log", encoding="utf-8"),
    ]
)
logger = logging.getLogger(__name__)

EXCEL_FILE_PATH = "Mening_Moliya_Jadvalim.xlsx"
CSV_FILE_PATH = "Mening_Moliya_Jadvalim.csv"

HEADERS = [
    "Sana", "Vaqt", "Turi", "Shaxs / Mijoz",
    "Tushum (so'm)", "Xarajat (so'm)", "Nasiya (so'm)",
    "Kategoriya", "Izoh / Tafsilotlar"
]

COLUMN_WIDTHS = {
    "A": 15,  # Sana
    "B": 10,  # Vaqt
    "C": 14,  # Turi
    "D": 22,  # Shaxs / Mijoz
    "E": 18,  # Tushum
    "F": 18,  # Xarajat
    "G": 18,  # Nasiya
    "H": 18,  # Kategoriya
    "I": 50   # Izoh / Tafsilotlar
}

MONTH_NAMES_UZ = {
    "01": "Yanvar", "02": "Fevral", "03": "Mart", "04": "Aprel",
    "05": "May", "06": "Iyun", "07": "Iyul", "08": "Avgust",
    "09": "Sentabr", "10": "Oktabr", "11": "Noyabr", "12": "Dekabr"
}

def get_month_sheet_title(date_str: str) -> str:
    """Sana berilganda (masalan: 2026-08-01) oylik varaq nomini qaytaradi: '2026-08 (Avgust)'"""
    try:
        parts = date_str.split("-")
        year = parts[0]
        month_num = parts[1]
        month_name = MONTH_NAMES_UZ.get(month_num, month_num)
        return f"{year}-{month_num} ({month_name})"
    except:
        return "Barcha Yozuvlar"

def setup_sheet_header(ws):
    header_font = Font(name="Arial", size=11, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)

    if ws.max_row == 0 or (ws.max_row == 1 and ws.cell(1, 1).value is None):
        ws.append(HEADERS)
        for col_num, header in enumerate(HEADERS, 1):
            cell = ws.cell(row=1, column=col_num)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = center_align

    for col_letter, width in COLUMN_WIDTHS.items():
        ws.column_dimensions[col_letter].width = width

    ws.row_dimensions[1].height = 28

def create_styled_excel():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Barcha Yozuvlar"
    setup_sheet_header(ws)
    wb.save(EXCEL_FILE_PATH)
    return wb

def init_excel_and_csv():
    if not os.path.exists(EXCEL_FILE_PATH):
        create_styled_excel()
    if not os.path.exists(CSV_FILE_PATH):
        with open(CSV_FILE_PATH, mode='w', newline='', encoding='utf-8-sig') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(HEADERS)

init_excel_and_csv()

def get_all_rows():
    rows = []
    if os.path.exists(CSV_FILE_PATH):
        with open(CSV_FILE_PATH, mode='r', encoding='utf-8-sig') as f:
            reader = csv.reader(f, delimiter=';')
            headers = next(reader, None)
            for idx, row in enumerate(reader, 1):
                if len(row) >= 9:
                    rows.append({
                        "id": idx,
                        "sana": row[0],
                        "vaqt": row[1],
                        "turi": row[2],
                        "shaxs": row[3],
                        "tushum": row[4],
                        "xarajat": row[5],
                        "nasiya": row[6],
                        "kategoriya": row[7],
                        "izoh": row[8]
                    })
    return rows

def delete_specific_row(data_row_id: int):
    deleted_info = None

    if os.path.exists(CSV_FILE_PATH):
        with open(CSV_FILE_PATH, mode='r', encoding='utf-8-sig') as f:
            rows = list(csv.reader(f, delimiter=';'))

        header = rows[0] if rows else []
        data_rows = rows[1:] if len(rows) > 1 else []

        if 1 <= data_row_id <= len(data_rows):
            removed = data_rows.pop(data_row_id - 1)
            deleted_info = {
                "id": data_row_id,
                "sana": removed[0],
                "vaqt": removed[1],
                "turi": removed[2],
                "summa": removed[4] if int(removed[4] or 0) > 0 else removed[5],
                "izoh": removed[8]
            }

            with open(CSV_FILE_PATH, mode='w', newline='', encoding='utf-8-sig') as f:
                writer = csv.writer(f, delimiter=';')
                writer.writerow(header)
                writer.writerows(data_rows)

    if os.path.exists(EXCEL_FILE_PATH):
        wb = openpyxl.load_workbook(EXCEL_FILE_PATH)

        # Barcha varaqlardan ushbu qatorni izlab o'chirish
        for sheet_name in wb.sheetnames:
            if sheet_name == "Kategoriya Hisoboti":
                continue
            ws = wb[sheet_name]
            excel_row = data_row_id + 1
            if 2 <= excel_row <= ws.max_row:
                ws.delete_rows(excel_row)

        update_excel_summary_sheet(wb)
        wb.save(EXCEL_FILE_PATH)

    return deleted_info

def delete_last_entry():
    rows = get_all_rows()
    if rows:
        return delete_specific_row(len(rows))
    return None

def get_category_totals(target_month_prefix: str = None):
    """Kategoriyalar bo'yicha hisobot (agar target_month_prefix masalan '2026-08' berilsa, faqat o'sha oy uchun)"""
    tushum_by_cat = {}
    xarajat_by_cat = {}
    nasiya_by_cat = {}

    total_tushum = 0
    total_xarajat = 0
    total_nasiya = 0

    if os.path.exists(CSV_FILE_PATH):
        with open(CSV_FILE_PATH, mode='r', encoding='utf-8-sig') as f:
            reader = csv.reader(f, delimiter=';')
            headers = next(reader, None)
            for row in reader:
                if len(row) < 9:
                    continue
                date_val = row[0]
                if target_month_prefix and not date_val.startswith(target_month_prefix):
                    continue

                tushum = int(row[4]) if str(row[4]).isdigit() else 0
                xarajat = int(row[5]) if str(row[5]).isdigit() else 0
                nasiya = int(row[6]) if str(row[6]).isdigit() else 0
                cat = row[7].strip() or "Boshqa"

                if tushum > 0:
                    tushum_by_cat[cat] = tushum_by_cat.get(cat, 0) + tushum
                    total_tushum += tushum
                if xarajat > 0:
                    xarajat_by_cat[cat] = xarajat_by_cat.get(cat, 0) + xarajat
                    total_xarajat += xarajat
                if nasiya > 0:
                    nasiya_by_cat[cat] = nasiya_by_cat.get(cat, 0) + nasiya
                    total_nasiya += nasiya

    return {
        "tushum_by_cat": tushum_by_cat,
        "xarajat_by_cat": xarajat_by_cat,
        "nasiya_by_cat": nasiya_by_cat,
        "total_tushum": total_tushum,
        "total_xarajat": total_xarajat,
        "total_nasiya": total_nasiya,
        "sof_foyda": total_tushum - total_xarajat
    }

def update_excel_summary_sheet(wb):
    if "Kategoriya Hisoboti" in wb.sheetnames:
        del wb["Kategoriya Hisoboti"]

    ws_sum = wb.create_sheet("Kategoriya Hisoboti")
    stats = get_category_totals()

    font_title = Font(name="Arial", size=12, bold=True, color="FFFFFF")
    fill_header = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    font_bold = Font(name="Arial", size=10, bold=True)
    font_normal = Font(name="Arial", size=10)
    thin_border = Border(left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'), top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9'))

    ws_sum.append(["Kategoriya", "Tushum (so'm)", "Xarajat (so'm)", "Nasiya (so'm)"])
    for col_i in range(1, 5):
        cell = ws_sum.cell(row=1, column=col_i)
        cell.font = font_title
        cell.fill = fill_header
        cell.alignment = Alignment(horizontal="center", vertical="center")

    ws_sum.column_dimensions["A"].width = 25
    ws_sum.column_dimensions["B"].width = 20
    ws_sum.column_dimensions["C"].width = 20
    ws_sum.column_dimensions["D"].width = 20
    ws_sum.row_dimensions[1].height = 26

    all_cats = sorted(list(set(list(stats["tushum_by_cat"].keys()) + list(stats["xarajat_by_cat"].keys()) + list(stats["nasiya_by_cat"].keys()))))

    for cat in all_cats:
        t_val = stats["tushum_by_cat"].get(cat, 0)
        x_val = stats["xarajat_by_cat"].get(cat, 0)
        n_val = stats["nasiya_by_cat"].get(cat, 0)
        ws_sum.append([cat, t_val, x_val, n_val])
        row_i = ws_sum.max_row
        ws_sum.row_dimensions[row_i].height = 20
        ws_sum.cell(row=row_i, column=1).font = font_normal
        ws_sum.cell(row=row_i, column=1).border = thin_border
        for c_idx, val in enumerate([t_val, x_val, n_val], 2):
            cell = ws_sum.cell(row=row_i, column=c_idx)
            cell.font = font_normal
            cell.number_format = '#,##0'
            cell.border = thin_border

    ws_sum.append(["JAMI TOTAL", stats["total_tushum"], stats["total_xarajat"], stats["total_nasiya"]])
    tot_row = ws_sum.max_row
    ws_sum.row_dimensions[tot_row].height = 24
    fill_tot = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    for col_i in range(1, 5):
        cell = ws_sum.cell(row=tot_row, column=col_i)
        cell.font = font_bold
        cell.fill = fill_tot
        cell.border = thin_border
        if col_i > 1:
            cell.number_format = '#,##0'

    ws_sum.append([])
    ws_sum.append(["SOF FOYDA (Tushum - Xarajat)", stats["sof_foyda"], "", ""])
    profit_row = ws_sum.max_row
    fill_profit = PatternFill(start_color="E2EFDA" if stats["sof_foyda"] >= 0 else "FCE4D6", fill_type="solid")
    ws_sum.cell(row=profit_row, column=1).font = font_bold
    ws_sum.cell(row=profit_row, column=2).font = font_bold
    ws_sum.cell(row=profit_row, column=2).number_format = '#,##0'
    ws_sum.cell(row=profit_row, column=1).fill = fill_profit
    ws_sum.cell(row=profit_row, column=2).fill = fill_profit

GEMINI_URL = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent"
GROQ_URL = "https://api.groq.com/openai/v1/chat/completions"
GROQ_WHISPER_URL = "https://api.groq.com/openai/v1/audio/transcriptions"

SYSTEM_PROMPT = """Siz moliya va eslatmalarni strukturaga ajratuvchi mutaxassis AI yordamchisiz.
Foydalanuvchi ovozli yoki matnli xabarda nima degan bo'lsa (lotin yoki kirill alifbosida), uni quyidagi JSON formatiga o'giring.

MUHIM MOLLYA VA SUMMA QOIDALARI:
1. O'zbekcha aytilgan summalarni to'g'ri raqamga o'giring:
   - "yigirma ming" / "игирма мин" / "20 ming" / "20k" = 20000
   - "o'n ming" / "10 ming" = 10000
   - "ellik ming" / "50 ming" = 50000
   - "yuz ming" / "100 ming" = 100000
   - "bir million" / "1 mln" = 1000000
   - "yigirma ming so'm" -> 20000 so'm (20 so'm emas!).
2. Kirillcha va xato yozilgan so'zlarni to'g'irlang (masalan: "бензенгай" -> Benzin / Transport, "игирма мин" -> 20000 so'm).

Javobingiz FAQAT va FAQAT toza JSON bo'lishi shart (hech qanday markdown prefiksiz, ```json siz):

{
  "turi": "Tushum" yoki "Xarajat" yoki "Nasiya",
  "shaxs": "Mijoz yoki kompaniya/shaxs ismi (bo'lmasa 'Noma'lum')",
  "tushum_som": 0,
  "xarajat_som": 0,
  "nasiya_som": 0,
  "kategoriya": "Savdo/Transport/Oziq-ovqat/Kommunal/Boshqa",
  "izoh": "Tafsilotli qisqa izoh"
}"""

async def transcribe_audio_with_groq(audio_bytes: bytes) -> str:
    headers = {"Authorization": f"Bearer {GROQ_API_KEY}"}
    form = aiohttp.FormData()
    form.add_field('file', audio_bytes, filename='voice.ogg', content_type='audio/ogg')
    form.add_field('model', 'whisper-large-v3-turbo')
    form.add_field('prompt', "O'zbek tilida moliya va xarajatlar: benzinga 20000 so'm, yigirma ming so'm, ellik ming so'm, yuz ming so'm, tushum, xarajat, nasiya.")

    async with aiohttp.ClientSession() as session:
        async with session.post(GROQ_WHISPER_URL, data=form, headers=headers, timeout=aiohttp.ClientTimeout(total=45)) as resp:
            if resp.status != 200:
                err_text = await resp.text()
                raise Exception(f"Whisper error ({resp.status}): {err_text[:150]}")
            result = await resp.json()
            return result.get("text", "").strip()

async def call_gemini_parser(prompt_text: str, audio_b64: str = None) -> str:
    parts = [{"text": prompt_text}]
    if audio_b64:
        parts.append({"inline_data": {"mime_type": "audio/ogg", "data": audio_b64}})

    body = {
        "system_instruction": {"parts": [{"text": SYSTEM_PROMPT}]},
        "contents": [{"role": "user", "parts": parts}],
        "generationConfig": {"temperature": 0.1, "responseMimeType": "application/json"}
    }
    headers = {"Content-Type": "application/json", "x-goog-api-key": GEMINI_API_KEY}

    async with aiohttp.ClientSession() as session:
        async with session.post(GEMINI_URL, json=body, headers=headers, timeout=aiohttp.ClientTimeout(total=45)) as resp:
            if resp.status == 429:
                raise Exception("Gemini limit 429")
            if resp.status != 200:
                err = await resp.text()
                raise Exception(f"Gemini error ({resp.status}): {err[:150]}")
            data = await resp.json()
            return data["candidates"][0]["content"]["parts"][0]["text"].strip()

async def call_groq_parser(prompt_text: str) -> str:
    body = {
        "model": "llama-3.3-70b-versatile",
        "messages": [
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": prompt_text}
        ],
        "temperature": 0.1,
        "response_format": {"type": "json_object"}
    }
    headers = {"Content-Type": "application/json", "Authorization": f"Bearer {GROQ_API_KEY}"}

    async with aiohttp.ClientSession() as session:
        async with session.post(GROQ_URL, json=body, headers=headers, timeout=aiohttp.ClientTimeout(total=30)) as resp:
            if resp.status != 200:
                err = await resp.text()
                raise Exception(f"Groq error ({resp.status}): {err[:150]}")
            data = await resp.json()
            return data["choices"][0]["message"]["content"].strip()

async def parse_note_with_ai(text: str = "", audio_b64: str = None) -> dict:
    now = datetime.datetime.now()
    date_str = now.strftime("%Y-%m-%d")
    time_str = now.strftime("%H:%M")
    prompt_text = f"Bugungi sana: {date_str} {time_str}.\nFoydalanuvchi xabari: '{text}'"

    raw_json = None
    last_error = ""

    if GEMINI_API_KEY:
        try:
            logger.info("→ Gemini API chaqirilmoqda...")
            raw_json = await call_gemini_parser(prompt_text, audio_b64)
            logger.info("✅ Gemini muvaffaqiyatli javob berdi")
        except Exception as e:
            logger.warning(f"⚠️ Gemini ishlamadi: {e}")
            last_error = str(e)

    if not raw_json and GROQ_API_KEY:
        try:
            logger.info("→ Groq Llama AI ga o'tildi...")
            raw_json = await call_groq_parser(prompt_text)
            logger.info("✅ Groq Llama muvaffaqiyatli javob berdi")
        except Exception as e:
            logger.error(f"❌ Groq xato: {e}")
            last_error = str(e)

    if not raw_json:
        raise Exception(f"AI javob bera olmadi. ({last_error[:100]})")

    raw_json = re.sub(r'^```json\s*', '', raw_json)
    raw_json = re.sub(r'\s*```$', '', raw_json)
    parsed = json.loads(raw_json)

    parsed["sana"] = date_str
    parsed["vaqt"] = time_str
    return parsed

def append_row_to_worksheet(ws, row_data):
    """Katakchaga chiroyli dizayn bilan qator qo'shish"""
    ws.append(row_data)
    last_row = ws.max_row

    thin_border = Border(left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'), top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9'))
    fill_color = "F9FBFD" if last_row % 2 == 0 else "FFFFFF"
    row_fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type="solid")
    font_data = Font(name="Arial", size=10)

    ws.row_dimensions[last_row].height = 22

    for col_idx, val in enumerate(row_data, 1):
        cell = ws.cell(row=last_row, column=col_idx)
        cell.font = font_data
        cell.fill = row_fill
        cell.border = thin_border

        if col_idx in (1, 2, 3):
            cell.alignment = Alignment(horizontal="center", vertical="center")
        elif col_idx in (5, 6, 7):
            cell.alignment = Alignment(horizontal="right", vertical="center")
            cell.number_format = '#,##0'
        else:
            cell.alignment = Alignment(horizontal="left", vertical="center")

def append_to_excel_and_csv(data: dict):
    # 1. CSV ga yozish
    with open(CSV_FILE_PATH, mode='a', newline='', encoding='utf-8-sig') as f:
        writer = csv.writer(f, delimiter=';')
        writer.writerow([
            data.get("sana", ""),
            data.get("vaqt", ""),
            data.get("turi", "Tushum"),
            data.get("shaxs", "Noma'lum"),
            data.get("tushum_som", 0),
            data.get("xarajat_som", 0),
            data.get("nasiya_som", 0),
            data.get("kategoriya", "Boshqa"),
            data.get("izoh", "")
        ])

    # 2. Excel (.xlsx) ga Avtomatik Oylik Varaqlarga yozish
    if not os.path.exists(EXCEL_FILE_PATH):
        wb = create_styled_excel()
    else:
        wb = openpyxl.load_workbook(EXCEL_FILE_PATH)

    row_data = [
        data.get("sana", ""),
        data.get("vaqt", ""),
        data.get("turi", "Tushum"),
        data.get("shaxs", "Noma'lum"),
        int(data.get("tushum_som", 0)),
        int(data.get("xarajat_som", 0)),
        int(data.get("nasiya_som", 0)),
        data.get("kategoriya", "Boshqa"),
        data.get("izoh", "")
    ]

    # A) "Barcha Yozuvlar" varag'iga qo'shish
    ws_all = wb["Barcha Yozuvlar"] if "Barcha Yozuvlar" in wb.sheetnames else wb.active
    setup_sheet_header(ws_all)
    append_row_to_worksheet(ws_all, row_data)

    # B) Avtomatik tegishli Oy varag'iga (masalan: "2026-08 (Avgust)") qo'shish!
    month_sheet_name = get_month_sheet_title(data.get("sana", ""))
    if month_sheet_name not in wb.sheetnames:
        ws_month = wb.create_sheet(month_sheet_name)
        setup_sheet_header(ws_month)
    else:
        ws_month = wb[month_sheet_name]

    append_row_to_worksheet(ws_month, row_data)

    # C) Kategoriya Hisoboti varag'ini yangilash
    update_excel_summary_sheet(wb)
    wb.save(EXCEL_FILE_PATH)

def get_action_keyboard():
    keyboard = [
        [
            InlineKeyboardButton("📋 Yozuvlar Ro'yxati", callback_data="show_rows_list"),
            InlineKeyboardButton("📈 Kategoriyalar Hisoboti", callback_data="category_report"),
        ],
        [
            InlineKeyboardButton("📊 Excel Faylini Yuklash (.xlsx)", callback_data="download_excel"),
            InlineKeyboardButton("🌐 Google Sheets Onlayn", url="https://sheets.google.com")
        ]
    ]
    return InlineKeyboardMarkup(keyboard)

def get_reply_menu():
    keyboard = [
        [KeyboardButton("📋 Yozuvlarni ko'rish va o'chirish"), KeyboardButton("📈 Kategoriyalar Hisoboti")],
        [KeyboardButton("📊 Excel Jadvalini Yuklash (.xlsx)"), KeyboardButton("ℹ️ Yordam")]
    ]
    return ReplyKeyboardMarkup(keyboard, resize_keyboard=True)

async def send_rows_list(update: Update, context: ContextTypes.DEFAULT_TYPE):
    rows = get_all_rows()
    chat_id = update.effective_chat.id

    if not rows:
        text = "⚠️ Jadvalda hali hech qanday yozuv yo'q."
        if update.callback_query:
            await update.callback_query.edit_message_text(text, reply_markup=get_action_keyboard())
        else:
            await update.message.reply_text(text, reply_markup=get_action_keyboard())
        return

    recent_rows = rows[-10:]

    msg = "📋 **JADVALDAGI YOZUVLAR RO'YXATI (O'chirish uchun tugmani bosing):**\n"
    msg += "━━━━━━━━━━━━━━━━━━━━━━\n\n"

    keyboard = []

    for r in recent_rows:
        summa = int(r['tushum'] or 0) or int(r['xarajat'] or 0) or int(r['nasiya'] or 0)
        msg += f"**[{r['id']}]** 📅 `{r['sana']}` | 📌 **{r['turi']}**: `{summa:,}` so'm\n"
        msg += f"   📝 Izoh: *{r['izoh']}*\n\n"

        keyboard.append([InlineKeyboardButton(f"🗑 {r['id']}-qatorni o'chirish [{r['izoh'][:15]}]", callback_data=f"del_row_{r['id']}")])

    keyboard.append([InlineKeyboardButton("📊 Excel Faylini Yuklash", callback_data="download_excel")])
    reply_markup = InlineKeyboardMarkup(keyboard)

    if update.callback_query:
        try:
            await update.callback_query.edit_message_text(msg, parse_mode="Markdown", reply_markup=reply_markup)
        except:
            await context.bot.send_message(chat_id=chat_id, text=msg, parse_mode="Markdown", reply_markup=reply_markup)
    else:
        await update.message.reply_text(msg, parse_mode="Markdown", reply_markup=reply_markup)

async def send_category_report(update: Update, context: ContextTypes.DEFAULT_TYPE):
    now = datetime.datetime.now()
    current_month_prefix = now.strftime("%Y-%m")
    month_name_uz = MONTH_NAMES_UZ.get(now.strftime("%m"), "")

    stats = get_category_totals(target_month_prefix=current_month_prefix)

    t_cats = stats["tushum_by_cat"]
    x_cats = stats["xarajat_by_cat"]
    n_cats = stats["nasiya_by_cat"]

    report = f"📊 **{now.year}-{now.strftime('%m')} ({month_name_uz.upper()}) OYLIK HISOBOT** 📈\n"
    report += "━━━━━━━━━━━━━━━━━━━━━━\n\n"

    report += "🟢 **TUSHUMLAR BO'YICHA:**\n"
    if t_cats:
        for cat, val in t_cats.items():
            report += f"  • {cat}: `{val:,}` so'm\n"
    else:
        report += "  • *(Bu oy tushum bo'lmagan)*\n"
    report += f"➡️ **Jami Tushum:** `{stats['total_tushum']:,}` so'm\n\n"

    report += "🔴 **XARAJATLAR BO'YICHA:**\n"
    if x_cats:
        for cat, val in x_cats.items():
            report += f"  • {cat}: `{val:,}` so'm\n"
    else:
        report += "  • *(Bu oy xarajat bo'lmagan)*\n"
    report += f"➡️ **Jami Xarajat:** `{stats['total_xarajat']:,}` so'm\n\n"

    if n_cats:
        report += "⏳ **NASIYALAR BO'YICHA:**\n"
        for cat, val in n_cats.items():
            report += f"  • {cat}: `{val:,}` so'm\n"
        report += f"➡️ **Jami Nasiya:** `{stats['total_nasiya']:,}` so'm\n\n"

    report += "━━━━━━━━━━━━━━━━━━━━━━\n"
    profit = stats['sof_foyda']
    sign = "+" if profit >= 0 else ""
    report += f"💵 **OYLIK SOF FOYDA:** `{sign}{profit:,}` so'm"

    chat_id = update.effective_chat.id
    if update.callback_query:
        try:
            await update.callback_query.edit_message_text(report, parse_mode="Markdown", reply_markup=get_action_keyboard())
        except:
            await context.bot.send_message(chat_id=chat_id, text=report, parse_mode="Markdown", reply_markup=get_action_keyboard())
    else:
        await update.message.reply_text(report, parse_mode="Markdown", reply_markup=get_action_keyboard())

async def start_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    welcome = (
        "👋 Assalomu alaykum!\n\n"
        "Men sizning **Professional Excel / Google Sheets AI Yordamchingizman**.\n\n"
        "🎙 Menga **ovozli xabar (golosovoy)** yuboring yoki matn ko'rinishida yozing:\n"
        "👉 *\"Benzinga yigirma ming to'lov qildim\"*\n"
        "👉 *\"Bugun Shoxruxga 150000 so'mga atir sotdim\"*\n\n"
        "📅 **OYMA-OY AVTOMATIK SAQLASH:**\n"
        "Siz yangi jadval ochishingiz shart emas! Bot har bir oyni avtomatik ravishda **Excel ichida alohida oylik varaqlar (2026-08 (Avgust), 2026-09 (Sentabr)...)** ko'rinishida guruhlab boradi! 📁✨\n\n"
        "📈 **Oylik Hisobot:** Pastdagi '📈 Kategoriyalar Hisoboti' tugmasini bossangiz, shu oyning tushum va xarajatlarini ko'rasiz."
    )
    await update.message.reply_text(welcome, parse_mode="Markdown", reply_markup=get_reply_menu())

async def send_excel_file(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat_id = update.effective_chat.id
    await context.bot.send_chat_action(chat_id, ChatAction.UPLOAD_DOCUMENT)

    file_to_send = EXCEL_FILE_PATH if os.path.exists(EXCEL_FILE_PATH) else CSV_FILE_PATH

    if os.path.exists(file_to_send):
        with open(file_to_send, 'rb') as doc:
            await context.bot.send_document(
                chat_id=chat_id,
                document=doc,
                filename="Mening_Moliya_Jadvalim.xlsx",
                caption="📊 **Sizning Moliya Excel Jadvalingiz**\n\n*(Excel ichida har bir oy uchun alohida varaqlar: '2026-08 (Avgust)', 'Barcha Yozuvlar' va 'Kategoriya Hisoboti' mavjud)*",
                parse_mode="Markdown"
            )
    else:
        await context.bot.send_message(chat_id=chat_id, text="⚠️ Hali jadvalga ma'lumot kiritilmagan.")

async def handle_delete_specific(update: Update, context: ContextTypes.DEFAULT_TYPE, row_id: int):
    chat_id = update.effective_chat.id
    deleted = delete_specific_row(row_id)

    if deleted:
        msg_text = (
            f"🗑 **{row_id}-qator muvaffaqiyatli o'chirildi!** ❌\n\n"
            f"📅 **Sana:** `{deleted['sana']}` ({deleted['vaqt']})\n"
            f"📌 **Turi:** {deleted['turi']}\n"
            f"💸 **Summa:** {deleted['summa']} so'm\n"
            f"📝 **Izoh:** {deleted['izoh']}"
        )
    else:
        msg_text = f"⚠️ {row_id}-sonli qator topilmadi yoki allaqachon o'chirilgan."

    await context.bot.send_message(chat_id=chat_id, text=msg_text, parse_mode="Markdown")
    await send_rows_list(update, context)

async def handle_delete_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text.strip()
    parts = text.split()

    if len(parts) >= 2 and parts[1].isdigit():
        row_id = int(parts[1])
        await handle_delete_specific(update, context, row_id)
    else:
        deleted = delete_last_entry()
        if deleted:
            msg_text = f"🗑 **Oxirgi yozuv ({deleted['id']}-qator) o'chirildi!** ❌\n📝 *{deleted['izoh']}*"
        else:
            msg_text = "⚠️ O'chirish uchun jadvalda yozuv topilmadi."
        await update.message.reply_text(msg_text, parse_mode="Markdown")

async def handle_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    data = query.data

    if data in ("download_excel", "download_csv"):
        await send_excel_file(update, context)
    elif data == "delete_last":
        rows = get_all_rows()
        if rows:
            await handle_delete_specific(update, context, rows[-1]["id"])
    elif data == "show_rows_list":
        await send_rows_list(update, context)
    elif data == "category_report":
        await send_category_report(update, context)
    elif data.startswith("del_row_"):
        row_id = int(data.replace("del_row_", ""))
        await handle_delete_specific(update, context, row_id)

async def handle_text_or_voice(update: Update, context: ContextTypes.DEFAULT_TYPE):
    msg = update.message
    if not msg:
        return

    text_input = msg.text or ""

    if text_input in ("📊 Excel Jadvalini Yuklash (.xlsx)", "📊 Jadvalni yuklab olish (Excel)", "/jadval"):
        await send_excel_file(update, context)
        return
    elif text_input in ("📈 Kategoriyalar Hisoboti", "/hisobot", "/statistika"):
        await send_category_report(update, context)
        return
    elif text_input in ("📋 Yozuvlarni ko'rish va o'chirish", "/royxat", "/yozuvlar"):
        await send_rows_list(update, context)
        return
    elif text_input.startswith(("/ochirish", "/bekor")):
        await handle_delete_command(update, context)
        return
    elif text_input == "ℹ️ Yordam":
        await start_command(update, context)
        return

    await context.bot.send_chat_action(msg.chat.id, ChatAction.TYPING)
    wait_msg = await msg.reply_text("🎙 AI xabarni eshitib, tegishli oy jadvaliga joylamoqda...")

    audio_b64 = None

    try:
        if msg.voice or msg.audio:
            file_obj = await (msg.voice or msg.audio).get_file()
            audio_bytes = await file_obj.download_as_bytearray()
            audio_b64 = base64.b64encode(audio_bytes).decode('utf-8')

            try:
                logger.info("🎙 Groq Whisper API ga ovozli fayl yuborilmoqda...")
                transcribed_text = await transcribe_audio_with_groq(bytes(audio_bytes))
                logger.info(f"✅ Ovoz matnga o'girildi: '{transcribed_text}'")
                text_input = transcribed_text
            except Exception as w_err:
                logger.warning(f"Whisper xatosi: {w_err}")
                text_input = msg.caption or "Ovozli xabarni tahlil qil"

        res = await parse_note_with_ai(text=text_input, audio_b64=audio_b64)

        append_to_excel_and_csv(res)

        await wait_msg.delete()

        month_label = get_month_sheet_title(res.get('sana', ''))

        reply_text = (
            f"✅ **Excel jadvalingizning '{month_label}' varag'iga saqlandi!** 📊\n\n"
            f"🗣 **Eshitilgan xabar:** *\"{text_input}\"*\n\n"
            f"📅 **Sana:** `{res.get('sana')}` ({res.get('vaqt')})\n"
            f"📌 **Turi:** {res.get('turi')}\n"
            f"👤 **Shaxs/Mijoz:** {res.get('shaxs')}\n"
            f"💰 **Tushum:** {int(res.get('tushum_som', 0)):,} so'm\n"
            f"💸 **Xarajat:** {int(res.get('xarajat_som', 0)):,} so'm\n"
            f"⏳ **Nasiya:** {int(res.get('nasiya_som', 0)):,} so'm\n"
            f"🏷 **Kategoriya:** {res.get('kategoriya')}\n"
            f"📝 **Izoh:** {res.get('izoh')}"
        )
        await msg.reply_text(
            reply_text,
            parse_mode="Markdown",
            reply_markup=get_action_keyboard()
        )

    except Exception as e:
        logger.error(f"Xato: {e}")
        try:
            await wait_msg.edit_text(f"❌ Xato: {str(e)[:200]}")
        except:
            await msg.reply_text(f"❌ Xato: {str(e)[:200]}")

def main():
    logger.info("🤖 Voice & Text to Styled Excel (.xlsx) Bot with Monthly Sheets support...")
    app = Application.builder().token(BOT_TOKEN).build()

    app.add_handler(CommandHandler("start", start_command))
    app.add_handler(CommandHandler("jadval", send_excel_file))
    app.add_handler(CommandHandler("hisobot", send_category_report))
    app.add_handler(CommandHandler("statistika", send_category_report))
    app.add_handler(CommandHandler("royxat", send_rows_list))
    app.add_handler(CommandHandler("yozuvlar", send_rows_list))
    app.add_handler(CommandHandler("ochirish", handle_delete_command))
    app.add_handler(CommandHandler("bekor", handle_delete_command))
    app.add_handler(CallbackQueryHandler(handle_callback))
    app.add_handler(MessageHandler(filters.VOICE | filters.AUDIO | (filters.TEXT & ~filters.COMMAND), handle_text_or_voice))

    logger.info("✅ Bot tayyor!")
    app.run_polling(drop_pending_updates=True)

if __name__ == "__main__":
    main()
