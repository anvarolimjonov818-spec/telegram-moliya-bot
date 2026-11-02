#!/usr/bin/env python3
# ============================================================
#  AI Voice & Text to Professional Excel (.xlsx) Bot
#  - Ultra-Aniqlikdagi O'zbek Nutqi va Murakkab Sonlar (480.000 / To'rt yuz sakson ming)
#  - STT Shovqinlarini (Noise Words) Avtomatik Tozalash
#  - Avtomatik OYMA-OY alohida varaqlar (Worksheets)
#  - HAR BIR OY SAHIFASIDA FAQAT SHU OYNING KATEGORIYA HISOBOTI (K,L,M,N ustunlarda)
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

try:
    from config import BOT_TOKEN as CFG_BOT_TOKEN, GEMINI_API_KEY as CFG_GEMINI_KEY, GROQ_API_KEY as CFG_GROQ_KEY
except ImportError:
    CFG_BOT_TOKEN, CFG_GEMINI_KEY, CFG_GROQ_KEY = "", "", ""

BOT_TOKEN = os.environ.get("BOT_TOKEN", CFG_BOT_TOKEN).strip()
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", CFG_GEMINI_KEY).strip()
GROQ_API_KEY = os.environ.get("GROQ_API_KEY", CFG_GROQ_KEY).strip()

logging.basicConfig(
    format="%(asctime)s | %(levelname)s | %(message)s",
    level=logging.INFO,
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler("sheets_bot.log", encoding="utf-8"),
    ]
)
logger = logging.getLogger(__name__)

DEFAULT_USER_ID = 1453337970
USER_DATA_DIR = os.path.join(os.path.dirname(__file__), "user_data")
os.makedirs(USER_DATA_DIR, exist_ok=True)

HEADERS = [
    "Sana", "Vaqt", "Turi", "Shaxs / Mijoz", "Telefon",
    "Tushum (so'm)", "Xarajat (so'm)", "Nasiya (so'm)",
    "Qaytarish Sanasi", "Kategoriya", "Izoh / Tafsilotlar"
]

COLUMN_WIDTHS = {
    "A": 13, "B": 8, "C": 15, "D": 18, "E": 16,
    "F": 17, "G": 17, "H": 17, "I": 16, "J": 18,
    "K": 38, "L": 5, "M": 26, "N": 17, "O": 17, "P": 17
}

MONTH_NAMES_UZ = {
    "01": "Yanvar", "02": "Fevral", "03": "Mart", "04": "Aprel",
    "05": "May", "06": "Iyun", "07": "Iyul", "08": "Avgust",
    "09": "Sentabr", "10": "Oktabr", "11": "Noyabr", "12": "Dekabr"
}

def get_user_paths(user_id=None):
    uid = str(user_id or DEFAULT_USER_ID).strip()
    u_dir = os.path.join(USER_DATA_DIR, uid)
    os.makedirs(u_dir, exist_ok=True)

    csv_path = os.path.join(u_dir, "Mening_Moliya_Jadvalim.csv")
    excel_path = os.path.join(u_dir, "Mening_Moliya_Jadvalim.xlsx")
    debts_path = os.path.join(u_dir, "qarz_eslatmalari.json")
    fastfood_path = os.path.join(u_dir, "fastfood_smena.json")

    # Boshlang'ich foydalanuvchi ma'lumotlarini saqlash
    if uid == str(DEFAULT_USER_ID):
        root_csv = os.path.join(os.path.dirname(__file__), "Mening_Moliya_Jadvalim.csv")
        root_excel = os.path.join(os.path.dirname(__file__), "Mening_Moliya_Jadvalim.xlsx")
        root_debts = os.path.join(os.path.dirname(__file__), "qarz_eslatmalari.json")

        if os.path.exists(root_csv) and not os.path.exists(csv_path):
            import shutil
            shutil.copy2(root_csv, csv_path)
        if os.path.exists(root_excel) and not os.path.exists(excel_path):
            import shutil
            shutil.copy2(root_excel, excel_path)
        if os.path.exists(root_debts) and not os.path.exists(debts_path):
            import shutil
            shutil.copy2(root_debts, debts_path)

    # Yangi foydalanuvchi uchun toza alohida shablonlar
    if not os.path.exists(csv_path):
        with open(csv_path, mode='w', newline='', encoding='utf-8-sig') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(HEADERS)
    if not os.path.exists(excel_path):
        create_styled_excel(excel_path)
    if not os.path.exists(debts_path):
        with open(debts_path, "w", encoding="utf-8") as f:
            json.dump([], f)
    if not os.path.exists(fastfood_path):
        with open(fastfood_path, "w", encoding="utf-8") as f:
            json.dump([], f)

    return {
        "user_id": uid,
        "csv": csv_path,
        "excel": excel_path,
        "debts": debts_path,
        "fastfood": fastfood_path
    }

DEBT_REMINDERS_FILE = "qarz_eslatmalari.json"
SUBSCRIBERS_FILE = "bot_subscribers.json"
DIGEST_STATE_FILE = "digest_state.json"

def add_subscriber(chat_id: int):
    if not chat_id:
        return
    subs = set()
    if os.path.exists(SUBSCRIBERS_FILE):
        try:
            with open(SUBSCRIBERS_FILE, "r", encoding="utf-8") as f:
                for s in json.load(f):
                    subs.add(int(s))
        except:
            pass
    for r in load_debt_reminders(chat_id):
        c = r.get("chat_id")
        if c:
            subs.add(int(c))
    if int(chat_id) not in subs:
        subs.add(int(chat_id))
        try:
            with open(SUBSCRIBERS_FILE, "w", encoding="utf-8") as f:
                json.dump(list(subs), f)
        except Exception as e:
            logger.error(f"add_subscriber error: {e}")

def get_all_subscribers() -> list:
    subs = set()
    if os.path.exists(SUBSCRIBERS_FILE):
        try:
            with open(SUBSCRIBERS_FILE, "r", encoding="utf-8") as f:
                for s in json.load(f):
                    subs.add(int(s))
        except:
            pass
    if os.path.exists(USER_DATA_DIR):
        for uid in os.listdir(USER_DATA_DIR):
            if uid.isdigit():
                subs.add(int(uid))
    return list(subs)

def load_debt_reminders(user_id=None) -> list:
    paths = get_user_paths(user_id)
    if os.path.exists(paths["debts"]):
        try:
            with open(paths["debts"], "r", encoding="utf-8") as f:
                return json.load(f)
        except:
            return []
    return []

def save_debt_reminders(reminders: list, user_id=None):
    paths = get_user_paths(user_id)
    with open(paths["debts"], "w", encoding="utf-8") as f:
        json.dump(reminders, f, ensure_ascii=False, indent=2)

def calculate_person_debt_balance(person_name: str, user_id=None) -> dict:
    total_berildi = 0
    total_qaytdi = 0
    p_clean = (person_name or "").strip().lower()
    paths = get_user_paths(user_id)

    if os.path.exists(paths["csv"]):
        with open(paths["csv"], mode='r', encoding='utf-8-sig') as f:
            reader = csv.reader(f, delimiter=';')
            headers = next(reader, None)
            for row in reader:
                if len(row) < 9:
                    continue
                row_person = str(row[3] or "").strip().lower()
                row_turi = str(row[2] or "").strip().lower()
                row_izoh = str(row[8] or "").strip().lower()
                nasiya_s = int(row[6]) if str(row[6]).isdigit() else 0
                tushum_s = int(row[4]) if str(row[4]).isdigit() else 0

                if p_clean in row_person or p_clean in row_izoh:
                    if "qaytdi" in row_turi or "qaytar" in row_izoh:
                        summa = nasiya_s if nasiya_s > 0 else tushum_s
                        total_qaytdi += summa
                    elif "nasiya" in row_turi or "berildi" in row_turi or nasiya_s > 0:
                        total_berildi += nasiya_s

    qoldiq = total_berildi - total_qaytdi
    return {
        "person": person_name,
        "total_berildi": total_berildi,
        "total_qaytdi": total_qaytdi,
        "qoldiq": max(0, qoldiq)
    }

# ==================== FAST FOOD SMENA & TABEL MODULI ====================

def load_fastfood_smenas(user_id=None) -> list:
    paths = get_user_paths(user_id)
    if os.path.exists(paths["fastfood"]):
        try:
            with open(paths["fastfood"], "r", encoding="utf-8") as f:
                return json.load(f)
        except:
            return []
    return []

def save_fastfood_smenas(smenas: list, user_id=None):
    paths = get_user_paths(user_id)
    with open(paths["fastfood"], "w", encoding="utf-8") as f:
        json.dump(smenas, f, ensure_ascii=False, indent=2)

def append_fastfood_smena_to_excel(smena_data: dict, user_id=None):
    paths = get_user_paths(user_id)
    excel_path = paths["excel"]
    if not os.path.exists(excel_path):
        wb = create_styled_excel(excel_path)
    else:
        wb = openpyxl.load_workbook(excel_path)

    sheet_title = "Fast Food Tabel"
    if sheet_title not in wb.sheetnames:
        ws = wb.create_sheet(sheet_title)
        headers = [
            "Sana", "Vaqt", "Xodimlar (Tafsilot)",
            "Jami Ish Haqi (З/П)", "Topshirildi (Сдано)", "Kartadan Yechildi",
            "Karta / Avanslar & Xarajatlar", "Izoh"
        ]
        ws.append(headers)
        header_fill = PatternFill(start_color="1F2937", end_color="1F2937", fill_type="solid")
        header_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
        for col_idx in range(1, len(headers) + 1):
            cell = ws.cell(row=1, column=col_idx)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
        col_widths = {"A": 13, "B": 9, "C": 45, "D": 20, "E": 20, "F": 20, "G": 40, "H": 30}
        for col_l, width in col_widths.items():
            ws.column_dimensions[col_l].width = width
    else:
        ws = wb[sheet_title]

    workers = smena_data.get("hodimlar", [])
    w_strs = []
    for w in workers:
        ism = w.get("ism", "")
        lavozim = f" ({w['lavozim']})" if w.get("lavozim") else ""
        soat = f" [{w['soat']} soat]" if w.get("soat") else ""
        haq = f": {int(w.get('ish_haqi', 0)):,}"
        w_strs.append(f"{ism}{lavozim}{soat}{haq}")
    w_details = "; ".join(w_strs)

    card_items = smena_data.get("karta_avans_va_xarajatlar", [])
    c_strs = []
    for c in card_items:
        c_strs.append(f"{c.get('shaxs', '')} {c.get('turi', '')} {int(c.get('summa', 0)):,} ({c.get('izoh', '')})")
    c_details = "; ".join(c_strs)

    row_data = [
        smena_data.get("sana", ""),
        smena_data.get("vaqt", ""),
        w_details,
        int(smena_data.get("jami_ish_haqi", 0)),
        int(smena_data.get("kassa_topshirildi_sdano", 0)),
        int(smena_data.get("kartadan_yechildi", 0)),
        c_details,
        smena_data.get("izoh", "")
    ]
    ws.append(row_data)

    row_idx = ws.max_row
    thin_border = Border(
        left=Side(style='thin', color='CBD5E1'),
        right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='CBD5E1'),
        bottom=Side(style='thin', color='CBD5E1')
    )
    for col_idx in range(1, len(row_data) + 1):
        cell = ws.cell(row=row_idx, column=col_idx)
        cell.border = thin_border
        cell.font = Font(name="Segoe UI", size=10)
        if col_idx in (4, 5, 6):
            cell.number_format = "#,##0"
            cell.alignment = Alignment(horizontal="right")
        else:
            cell.alignment = Alignment(horizontal="left", vertical="center")

    wb.save(excel_path)
    try:
        from tabel_sync import sync_fastfood_tabel_file
        sync_fastfood_tabel_file(user_id)
    except Exception as t_err:
        logger.error(f"Tabel sync error: {t_err}")

def calculate_fastfood_summary(user_id=None, month_prefix=None, start_day=1, end_day=31):
    smenas = load_fastfood_smenas(user_id)
    if not month_prefix:
        month_prefix = datetime.datetime.now().strftime("%Y-%m")

    filtered = []
    for s in smenas:
        s_date = s.get("sana", "")
        if s_date.startswith(month_prefix):
            try:
                day = int(s_date.split("-")[2])
                if start_day <= day <= end_day:
                    filtered.append(s)
            except:
                filtered.append(s)

    workers_map = {}
    total_zp = 0
    total_sdano = 0
    total_karta_yechildi = 0
    total_avans = 0
    total_xarajat = 0

    for s in filtered:
        total_zp += int(s.get("jami_ish_haqi", 0))
        total_sdano += int(s.get("kassa_topshirildi_sdano", 0))
        total_karta_yechildi += int(s.get("kartadan_yechildi", 0))

        for w in s.get("hodimlar", []):
            name = (w.get("ism") or "Xodim").strip()
            if name not in workers_map:
                workers_map[name] = {
                    "name": name,
                    "total_ish_haqi": 0,
                    "shifts_count": 0,
                    "avanslar": 0,
                    "soatlar": 0
                }
            workers_map[name]["total_ish_haqi"] += int(w.get("ish_haqi", 0))
            workers_map[name]["shifts_count"] += 1
            if w.get("soat"):
                try:
                    workers_map[name]["soatlar"] += float(w["soat"])
                except:
                    pass

        for c in s.get("karta_avans_va_xarajatlar", []):
            summa = int(c.get("summa", 0))
            c_type = (c.get("turi") or "").lower()
            target_shaxs = (c.get("shaxs") or "").strip()

            if "avans" in c_type:
                total_avans += summa
                matched = False
                for w_name in workers_map:
                    if target_shaxs.lower() in w_name.lower() or w_name.lower() in target_shaxs.lower():
                        workers_map[w_name]["avanslar"] += summa
                        matched = True
                        break
                if not matched and target_shaxs:
                    if target_shaxs not in workers_map:
                        workers_map[target_shaxs] = {
                            "name": target_shaxs,
                            "total_ish_haqi": 0,
                            "shifts_count": 0,
                            "avanslar": summa,
                            "soatlar": 0
                        }
                    else:
                        workers_map[target_shaxs]["avanslar"] += summa
            else:
                total_xarajat += summa

    for w in workers_map.values():
        w["qoldiq_to_lanadi"] = w["total_ish_haqi"] - w["avanslar"]

    return {
        "filtered_smenas": filtered,
        "workers": workers_map,
        "total_zp": total_zp,
        "total_sdano": total_sdano,
        "total_karta_yechildi": total_karta_yechildi,
        "total_avans": total_avans,
        "total_xarajat": total_xarajat,
        "month_prefix": month_prefix,
        "start_day": start_day,
        "end_day": end_day
    }

def format_fastfood_shift_message(data: dict, transcribed_text: str = None, newly_added: dict = None) -> str:
    sana = data.get("sana", datetime.datetime.now().strftime("%Y-%m-%d"))
    text = f"🍔 **FAST FOOD KUNLIK SMENA HISOBOTI** 📋\n"
    text += f"📅 **Sana:** `{sana}`\n"
    text += f"━━━━━━━━━━━━━━━━━━━━━━\n\n"

    if transcribed_text:
        text += f"🗣 **Eshitilgan xabar:** *\"{transcribed_text}\"*\n\n"

    if newly_added and newly_added.get("hodimlar"):
        text += "✨ **Yangi qo'shilgan / yangilangan:**\n"
        for w in newly_added["hodimlar"]:
            text += f"  ➕ **{w['ism']}:** `{int(w['ish_haqi']):,}` so'm\n"
        if newly_added.get("sdano"):
            text += f"  💵 **Kassa (Сдано):** `{int(newly_added['sdano']):,}` so'm\n"
        if newly_added.get("kartadan"):
            text += f"  💳 **Kartadan yechildi:** `{int(newly_added['kartadan']):,}` so'm\n"
        text += "\n"

    workers = data.get("hodimlar", [])
    if workers:
        text += "👥 **KUNNING JAMI XODIMLAR RO'YXATI (З/П):**\n"
        for w in workers:
            ism = w.get("ism", "Xodim")
            lavozim = f" ({w['lavozim']})" if w.get("lavozim") and w['lavozim'] != "bo'sh" else ""
            soat = f" [{w['soat']} soat]" if w.get("soat") else ""
            haq = int(w.get("ish_haqi", 0))
            text += f"  • {ism}{lavozim}{soat}: `{haq:,}` so'm\n"
        text += f"💰 **JAMI KUNLIK З/П:** `{int(data.get('jami_ish_haqi', 0)):,}` so'm\n\n"

    sdano = int(data.get("kassa_topshirildi_sdano", 0))
    kartadan = int(data.get("kartadan_yechildi", 0))
    card_items = data.get("karta_avans_va_xarajatlar", [])

    text += "🏦 **KASSA VA TO'LOVLAR:**\n"
    text += f"💵 Kassa topshirildi (Сдано): `{sdano:,}` so'm\n"
    text += f"💳 Kartadan yechildi: `{kartadan:,}` so'm\n"
    if card_items:
        text += f"\n💳 **KARTA / AVANSLAR VA XARAJATLAR:**\n"
        for c in card_items:
            text += f"  • {c.get('shaxs', '')}: `{int(c.get('summa', 0)):,}` so'm ({c.get('izoh', '')})\n"

    text += f"\n✅ *Fast Food tabeliga saqlandi!*"
    return text

def format_fastfood_tabel_report(summary: dict, title: str = "10 KUNLIK SMENA SARHISOBI") -> str:
    workers = list(summary["workers"].values())
    workers.sort(key=lambda x: x["total_ish_haqi"], reverse=True)

    month_parts = summary["month_prefix"].split("-")
    month_name = MONTH_NAMES_UZ.get(month_parts[1], "") if len(month_parts) > 1 else ""

    text = f"🍔 **FAST FOOD TABEL: {title.upper()}** 📊\n"
    text += f"📅 **Davr:** `{summary['start_day']}-{summary['end_day']} {month_name} {month_parts[0]}`\n"
    text += f"━━━━━━━━━━━━━━━━━━━━━━\n\n"

    if not workers:
        text += "⚠️ Ushbu davr uchun smena ma'lumotlari hali kiritilmagan.\n"
        text += "👉 *Smenani kiritish uchun botga ovozli xabar yuboring!*\n\n"
    else:
        text += "👥 **XODIMLAR ISH HAQI VA AVANSLAR:**\n"
        for w in workers:
            name = w['name']
            haq = w['total_ish_haqi']
            avans = w['avanslar']
            qoldiq = w['qoldiq_to_lanadi']
            shifts = w['shifts_count']
            soat_info = f" ({w['soatlar']:.1f} soat)" if w.get('soatlar') else ""

            if avans > 0:
                text += f"• **{name}:** `{haq:,}` so'm{soat_info} ({shifts} smena)\n"
                text += f"  ↳ Avans: `-{avans:,}` | **To'lanadi: {qoldiq:,} so'm**\n"
            else:
                text += f"• **{name}:** `{haq:,}` so'm{soat_info} ({shifts} smena) — To'lanadi: `{qoldiq:,}`\n"
        text += "\n"

    text += f"━━━━━━━━━━━━━━━━━━━━━━\n"
    text += f"💰 **JAMI ISH HAQI FONDI (З/П):** `{summary['total_zp']:,}` so'm\n"
    text += f"💵 **JAMI KASSA TOPSHIRILDI (Сдано):** `{summary['total_sdano']:,}` so'm\n"
    text += f"💳 **JAMI KARTADAN YECHILDI:** `{summary['total_karta_yechildi']:,}` so'm\n"
    if summary['total_avans'] > 0:
        text += f"⚠️ **Jami berilgan avanslar:** `{summary['total_avans']:,}` so'm\n"
    if summary['total_xarajat'] > 0:
        text += f"🧾 **Karta orqali boshqa xarajatlar:** `{summary['total_xarajat']:,}` so'm\n"

    return text

async def check_debt_reminders_task(app):
    """Har 30 daqiqada qarz qaytarish muddati kelganini tekshirib, xabar yuborish"""
    await asyncio.sleep(10)
    while True:
        try:
            reminders = load_debt_reminders()
            today_str = datetime.datetime.now().strftime("%Y-%m-%d")
            updated = False

            for item in reminders:
                due_date = item.get("qaytarish_sanasi")
                is_notified = item.get("ogohlantirildi", False)
                chat_id = item.get("chat_id")
                shaxs = item.get("shaxs")

                if due_date and due_date <= today_str and not is_notified and chat_id:
                    # Qarz holatini hisoblaymiz
                    debt_info = calculate_person_debt_balance(shaxs)
                    qoldiq = debt_info["qoldiq"]
                    total_qaytdi = debt_info["total_qaytdi"]

                    if qoldiq <= 0:
                        # Qarz to'liq to'langan bo'lsa ogohlantirmaymiz
                        item["ogohlantirildi"] = True
                        item["holat"] = "to'liq_to'langan"
                        updated = True
                        continue

                    # Agar qoldiq qolgan bo'lsa ogohlantiramiz
                    tel = item.get('telefon', '')
                    tel_text = f"\n📞 **Telefon:** `{tel}`" if tel else ""

                    msg_text = (
                        f"🔔 **QARZNI QAYTARISH MUDDATI KELDI!** ⏰\n\n"
                        f"👤 **Mijoz / Shaxs:** `{shaxs}`{tel_text}\n"
                        f"📅 **Qarz olingan sana:** `{item.get('berilgan_sana')}`\n"
                        f"🎯 **Qaytarish kerak bo'lgan sana:** `{due_date}`\n"
                        f"💰 **Dastlabki olingan qarz:** `{int(item.get('boshlangich_qarz', 0)):,} so'm`\n"
                    )

                    if total_qaytdi > 0:
                        msg_text += (
                            f"✅ **Qaytarilgan qismi:** `{total_qaytdi:,} so'm`\n"
                            f"⚠️ **QOLGAN QARZ QOLDIG'I:** `{qoldiq:,} so'm`\n\n"
                            f"📌 *Iltimos, {shaxs} bilan bog'lanib, qoldiq qarzni yopishni so'rang.*"
                        )
                    else:
                        msg_text += (
                            f"⚠️ **Hozirgi qarz miqdori:** `{qoldiq:,} so'm`\n"
                            f"❌ Hali biror so'm qaytarilmadi.\n\n"
                            f"📌 *Bugun qarzni qaytarish kuni!*"
                        )

                    try:
                        await app.bot.send_message(chat_id=chat_id, text=msg_text, parse_mode="Markdown")
                        item["ogohlantirildi"] = True
                        item["oxirgi_ogohlantirish"] = today_str
                        updated = True
                    except Exception as send_err:
                        logger.error(f"Eslatma yuborishda xatolik: {send_err}")

            if updated:
                save_debt_reminders(reminders)

        except Exception as e:
            logger.error(f"check_debt_reminders_task xatosi: {e}")

        await asyncio.sleep(1800)  # Har 30 daqiqada bir tekshiradi

EXCEL_FILE_PATH = "Mening_Moliya_Jadvalim.xlsx"
CSV_FILE_PATH = "Mening_Moliya_Jadvalim.csv"

HEADERS = [
    "Sana", "Vaqt", "Turi", "Shaxs / Mijoz", "Telefon",
    "Tushum (so'm)", "Xarajat (so'm)", "Nasiya (so'm)",
    "Qaytarish Sanasi", "Kategoriya", "Izoh / Tafsilotlar"
]

COLUMN_WIDTHS = {
    "A": 13,  # Sana
    "B": 8,   # Vaqt
    "C": 15,  # Turi
    "D": 18,  # Shaxs / Mijoz
    "E": 16,  # Telefon
    "F": 17,  # Tushum
    "G": 17,  # Xarajat
    "H": 17,  # Nasiya
    "I": 16,  # Qaytarish Sanasi
    "J": 18,  # Kategoriya
    "K": 38,  # Izoh / Tafsilotlar
    "L": 5,   # Ajratuvchi bo'sh ustun
    "M": 26,  # Kategoriya Hisoboti
    "N": 17,  # Tushum
    "O": 17,  # Xarajat
    "P": 17   # Nasiya
}

MONTH_NAMES_UZ = {
    "01": "Yanvar", "02": "Fevral", "03": "Mart", "04": "Aprel",
    "05": "May", "06": "Iyun", "07": "Iyul", "08": "Avgust",
    "09": "Sentabr", "10": "Oktabr", "11": "Noyabr", "12": "Dekabr"
}

def clean_stt_noise(text: str) -> str:
    """Whisper tushunarsiz shovqin so'zlarini matndan tozalash (o'zbek va rus tillari uchun)"""
    if not text:
        return ""
    noise_patterns = [
        r'\bget\b', r'\bover\b', r'\bzikim\b', r'\buh\b', r'\bumm\b',
        r'\bsubtitles\b', r'\bby amara\.org\b', r'\bthank you\b', r'\bthanks\b',
        r'\[.*?\]', r'\(.*?\)',
        r'субтитры.*', r'редактор.*', r'перевод.*', r'продолжение следует.*',
        r'спасибо за просмотр.*', r'подписывайтесь на канал.*'
    ]
    cleaned = text
    for np in noise_patterns:
        cleaned = re.sub(np, '', cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r'\s+', ' ', cleaned).strip()
    return cleaned

def fix_uzbek_speech_phonetics(text: str) -> str:
    """Whisper STT tomonidan xato o'girilgan o'zbekcha talaffuz variantlarini to'g'rilash"""
    if not text:
        return text

    replacements = [
        (r'\bsakismi\b|\bsaksen\b', 'sakson'),
        (r'\byettism\b', 'yetmish'),
        (r'\boltmis\b', 'oltmish'),
        (r'\bo\'ttis\b', 'o\'ttiz'),
        (r'\bigirma\b', 'yigirma'),
        (r'\bonbesh\b', 'o\'n besh'),
        (r'\bso\'na\b|\bso\'n\b|\bsom\b', 'so\'m'),
        (r'\bturao\b|\bto\'rao\b|\bo\'kildim\b', 'to\'lov qildim'),
        (r'\bprogramma\b|\bgetu\b|\brizugim\b', 'dastur uchun')
    ]

    fixed = text
    for pattern, repl in replacements:
        fixed = re.sub(pattern, repl, fixed, flags=re.IGNORECASE)
    return clean_stt_noise(fixed)

def get_month_sheet_title(date_str: str) -> str:
    try:
        parts = date_str.split("-")
        year = parts[0]
        month_num = parts[1]
        month_name = MONTH_NAMES_UZ.get(month_num, month_num)
        return f"{year}-{month_num} ({month_name})"
    except:
        return "Barcha Yozuvlar"

def setup_sheet_header(ws):
    ws.sheet_view.zoomScale = 100
    header_font = Font(name="Arial", size=11, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # 1-qator har doim sarlavha (Header) bo'lishi shart
    if ws.cell(row=1, column=1).value != HEADERS[0]:
        ws.insert_rows(1)
        for col_num, header in enumerate(HEADERS, 1):
            cell = ws.cell(row=1, column=col_num, value=header)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = center_align
    else:
        for col_num, header in enumerate(HEADERS, 1):
            cell = ws.cell(row=1, column=col_num, value=header)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = center_align

    for col_letter, width in COLUMN_WIDTHS.items():
        ws.column_dimensions[col_letter].width = width

    ws.row_dimensions[1].height = 28

def create_styled_excel(file_path=None):
    if not file_path:
        file_path = EXCEL_FILE_PATH
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Barcha Yozuvlar"
    setup_sheet_header(ws)
    wb.save(file_path)
    return wb

def init_excel_and_csv(user_id=None):
    paths = get_user_paths(user_id)
    if not os.path.exists(paths["excel"]):
        create_styled_excel(paths["excel"])
    if not os.path.exists(paths["csv"]):
        with open(paths["csv"], mode='w', newline='', encoding='utf-8-sig') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(HEADERS)

init_excel_and_csv()

def get_all_rows(user_id=None):
    paths = get_user_paths(user_id)
    rows = []
    if os.path.exists(paths["csv"]):
        with open(paths["csv"], mode='r', encoding='utf-8-sig') as f:
            reader = csv.reader(f, delimiter=';')
            headers = next(reader, None)
            for idx, row in enumerate(reader, 1):
                if len(row) >= 9:
                    # Yangi format: 11 ustun
                    telefon = row[4] if len(row) >= 11 else ""
                    tushum = row[5] if len(row) >= 11 else row[4]
                    xarajat = row[6] if len(row) >= 11 else row[5]
                    nasiya = row[7] if len(row) >= 11 else row[6]
                    due_date = row[8] if len(row) >= 11 else ""
                    kategoriya = row[9] if len(row) >= 11 else row[7]
                    izoh = row[10] if len(row) >= 11 else row[8]

                    rows.append({
                        "id": idx,
                        "sana": row[0],
                        "vaqt": row[1],
                        "turi": row[2],
                        "shaxs": row[3],
                        "telefon": telefon,
                        "tushum": int(tushum) if str(tushum).isdigit() else 0,
                        "xarajat": int(xarajat) if str(xarajat).isdigit() else 0,
                        "nasiya": int(nasiya) if str(nasiya).isdigit() else 0,
                        "qaytarish_sanasi": due_date,
                        "kategoriya": kategoriya,
                        "izoh": izoh
                    })
    return rows

def render_in_sheet_summary(ws):
    font_header = Font(name="Arial", size=11, bold=True, color="FFFFFF")
    fill_header = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    font_bold = Font(name="Arial", size=10, bold=True)
    font_normal = Font(name="Arial", size=10)
    thin_border = Border(left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'), top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9'))

    # 1. O'ng tarafdagi (M dan P gacha: 13-16 ustunlar) barcha eski xulosalarni tozalaymiz
    max_clean_row = max(ws.max_row + 15, 30)
    for r in range(1, max_clean_row + 1):
        for c in range(13, 17):
            cell = ws.cell(row=r, column=c)
            cell.value = None
            cell.fill = PatternFill(fill_type=None)
            cell.border = Border()

    # 2. Xulosa sarlavhalari (1-qatorda, 13-ustundan boshlab)
    sum_headers = ["Kategoriya (Hisobot)", "Tushum (so'm)", "Xarajat (so'm)", "Nasiya (so'm)"]
    for idx, h_text in enumerate(sum_headers, 13):
        cell = ws.cell(row=1, column=idx, value=h_text)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border

    tushum_by_cat = {}
    xarajat_by_cat = {}
    nasiya_by_cat = {}
    
    total_tushum = 0       # Haqiqiy sof savdo/daromad
    total_xarajat = 0      # Haqiqiy xarajatlar
    total_nasiya_berildi = 0 # Berilgan qarzlar
    total_nasiya_qaytdi = 0  # Qaytgan qarzlar

    def parse_num(v):
        if v is None:
            return 0
        s = str(v).replace(" ", "").replace(",", "").replace("\xa0", "").strip()
        if s.isdigit():
            return int(s)
        try:
            return int(float(s))
        except:
            return 0

    # 3. 2-qatordan boshlab barcha ma'lumotlarni o'qiymiz
    # Yangi ustunlar:
    # 1: Sana, 2: Vaqt, 3: Turi, 4: Shaxs, 5: Telefon
    # 6: Tushum, 7: Xarajat, 8: Nasiya, 9: Qaytarish Sanasi, 10: Kategoriya, 11: Izoh
    for r_i in range(2, ws.max_row + 1):
        sana_val = ws.cell(row=r_i, column=1).value
        if not sana_val or str(sana_val).strip() == "Sana":
            continue

        turi_val = str(ws.cell(row=r_i, column=3).value or "").strip().lower()
        t_c = ws.cell(row=r_i, column=6).value
        x_c = ws.cell(row=r_i, column=7).value
        n_c = ws.cell(row=r_i, column=8).value
        cat_c = ws.cell(row=r_i, column=10).value
        izoh_c = str(ws.cell(row=r_i, column=11).value or "").lower()

        t_val = parse_num(t_c)
        x_val = parse_num(x_c)
        n_val = parse_num(n_c)

        cat = str(cat_c).strip() if cat_c else "Boshqa"
        if not cat:
            cat = "Boshqa"

        is_qarz_qaytdi = ("qaytdi" in turi_val) or ("qarzini qaytardi" in izoh_c) or ("qarz qaytdi" in izoh_c) or ("nasiyani yopdi" in izoh_c)

        if is_qarz_qaytdi:
            qaytgan_summa = n_val if n_val > 0 else t_val
            total_nasiya_qaytdi += qaytgan_summa
        else:
            if t_val > 0:
                tushum_by_cat[cat] = tushum_by_cat.get(cat, 0) + t_val
                total_tushum += t_val
            if x_val > 0:
                xarajat_by_cat[cat] = xarajat_by_cat.get(cat, 0) + x_val
                total_xarajat += x_val
            if n_val > 0:
                nasiya_by_cat[cat] = nasiya_by_cat.get(cat, 0) + n_val
                total_nasiya_berildi += n_val

    all_cats = sorted(list(set(list(tushum_by_cat.keys()) + list(xarajat_by_cat.keys()) + list(nasiya_by_cat.keys()))))

    current_row = 2
    for cat in all_cats:
        t_val = tushum_by_cat.get(cat, 0)
        x_val = xarajat_by_cat.get(cat, 0)
        n_val = nasiya_by_cat.get(cat, 0)

        fill_r = PatternFill(start_color="F9FBFD" if current_row % 2 == 0 else "FFFFFF", fill_type="solid")

        c1 = ws.cell(row=current_row, column=13, value=cat)
        c1.font = font_normal
        c1.fill = fill_r
        c1.border = thin_border
        c1.alignment = Alignment(horizontal="left", vertical="center")

        for c_idx, val in enumerate([t_val, x_val, n_val], 14):
            cell = ws.cell(row=current_row, column=c_idx, value=val)
            cell.font = font_normal
            cell.fill = fill_r
            cell.border = thin_border
            cell.alignment = Alignment(horizontal="right", vertical="center")
            cell.number_format = '#,##0'

        ws.row_dimensions[current_row].height = 20
        current_row += 1

    # JAMI TOTAL
    fill_tot = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    tot_vals = ["JAMI TOTAL", total_tushum, total_xarajat, total_nasiya_berildi]
    for idx, val in enumerate(tot_vals, 13):
        cell = ws.cell(row=current_row, column=idx, value=val)
        cell.font = font_bold
        cell.fill = fill_tot
        cell.border = thin_border
        if idx > 13:
            cell.alignment = Alignment(horizontal="right", vertical="center")
            cell.number_format = '#,##0'
        else:
            cell.alignment = Alignment(horizontal="left", vertical="center")

    ws.row_dimensions[current_row].height = 24
    current_row += 1

    # 1. SOF FOYDA = Haqiqiy Tushum - Xarajat
    sof_foyda = total_tushum - total_xarajat
    fill_profit = PatternFill(start_color="E2EFDA" if sof_foyda >= 0 else "FCE4D6", fill_type="solid")
    c_p1 = ws.cell(row=current_row, column=13, value="SOF FOYDA")
    c_p1.font = font_bold
    c_p1.fill = fill_profit
    c_p1.border = thin_border

    c_p2 = ws.cell(row=current_row, column=14, value=sof_foyda)
    c_p2.font = font_bold
    c_p2.fill = fill_profit
    c_p2.border = thin_border
    c_p2.alignment = Alignment(horizontal="right", vertical="center")
    c_p2.number_format = '#,##0'

    for c_i in range(15, 17):
        c_empty = ws.cell(row=current_row, column=c_i, value="")
        c_empty.fill = fill_profit
        c_empty.border = thin_border

    ws.row_dimensions[current_row].height = 24
    current_row += 1

    # 2. QARZLAR QOLDIG'I (Nasiya) = Berilgan qarz - Qaytarilgan qarz
    nasiya_qoldiq = total_nasiya_berildi - total_nasiya_qaytdi
    fill_debt = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    c_d1 = ws.cell(row=current_row, column=13, value="QARZ QOLDIG'I")
    c_d1.font = font_bold
    c_d1.fill = fill_debt
    c_d1.border = thin_border

    c_d4 = ws.cell(row=current_row, column=16, value=nasiya_qoldiq)
    c_d4.font = font_bold
    c_d4.fill = fill_debt
    c_d4.border = thin_border
    c_d4.alignment = Alignment(horizontal="right", vertical="center")
    c_d4.number_format = '#,##0'

    for c_i in (14, 15):
        c_empty = ws.cell(row=current_row, column=c_i, value="")
        c_empty.fill = fill_debt
        c_empty.border = thin_border

    ws.row_dimensions[current_row].height = 24
    current_row += 1

    # 3. KASSA / CHO'NTAK QOLDIG'I (Real bor pul):
    kassa_balansi = (total_tushum + total_nasiya_qaytdi) - (total_xarajat + total_nasiya_berildi)
    fill_kassa = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")
    c_k1 = ws.cell(row=current_row, column=13, value="KASSA BALANSI")
    c_k1.font = font_bold
    c_k1.fill = fill_kassa
    c_k1.border = thin_border

    c_k2 = ws.cell(row=current_row, column=14, value=kassa_balansi)
    c_k2.font = font_bold
    c_k2.fill = fill_kassa
    c_k2.border = thin_border
    c_k2.alignment = Alignment(horizontal="right", vertical="center")
    c_k2.number_format = '#,##0'

    for c_i in range(15, 17):
        c_empty = ws.cell(row=current_row, column=c_i, value="")
        c_empty.fill = fill_kassa
        c_empty.border = thin_border

    ws.row_dimensions[current_row].height = 24

def delete_specific_row(data_row_id: int, user_id=None):
    paths = get_user_paths(user_id)
    csv_file = paths["csv"]
    excel_file = paths["excel"]
    deleted_info = None

    if os.path.exists(csv_file):
        with open(csv_file, mode='r', encoding='utf-8-sig') as f:
            rows = list(csv.reader(f, delimiter=';'))

        header = rows[0] if rows else []
        data_rows = rows[1:] if len(rows) > 1 else []

        if 1 <= data_row_id <= len(data_rows):
            removed = data_rows.pop(data_row_id - 1)
            tushum_val = int(removed[5]) if len(removed) >= 11 and str(removed[5]).isdigit() else (int(removed[4]) if len(removed) >= 9 and str(removed[4]).isdigit() else 0)
            xarajat_val = int(removed[6]) if len(removed) >= 11 and str(removed[6]).isdigit() else (int(removed[5]) if len(removed) >= 9 and str(removed[5]).isdigit() else 0)
            nasiya_val = int(removed[7]) if len(removed) >= 11 and str(removed[7]).isdigit() else (int(removed[6]) if len(removed) >= 9 and str(removed[6]).isdigit() else 0)
            summa = tushum_val or xarajat_val or nasiya_val

            deleted_info = {
                "id": data_row_id,
                "sana": removed[0],
                "vaqt": removed[1],
                "turi": removed[2],
                "shaxs": removed[3] if len(removed) > 3 else "",
                "summa": summa,
                "izoh": removed[10] if len(removed) >= 11 else (removed[8] if len(removed) > 8 else "")
            }

            with open(csv_file, mode='w', newline='', encoding='utf-8-sig') as f:
                writer = csv.writer(f, delimiter=';')
                writer.writerow(header)
                writer.writerows(data_rows)

    if os.path.exists(excel_file):
        wb = openpyxl.load_workbook(excel_file)
        for sheet_name in wb.sheetnames:
            if sheet_name == "FastFood_Tabel":
                continue
            ws = wb[sheet_name]
            excel_row = data_row_id + 1
            if 2 <= excel_row <= ws.max_row:
                ws.delete_rows(excel_row)
            render_in_sheet_summary(ws)

        wb.save(excel_file)

    return deleted_info

def update_or_add_record(record: dict, user_id=None):
    """Web App orqali yozuvni qo'shish yoki tahrirlash"""
    paths = get_user_paths(user_id)
    csv_file = paths["csv"]
    excel_file = paths["excel"]
    rec_id = record.get("id")

    rows = []
    if os.path.exists(csv_file):
        with open(csv_file, mode='r', encoding='utf-8-sig') as f:
            rows = list(csv.reader(f, delimiter=';'))

    header = rows[0] if rows else HEADERS
    data_rows = rows[1:] if len(rows) > 1 else []

    now = datetime.datetime.now()
    sana = record.get("sana") or now.strftime("%Y-%m-%d")
    vaqt = record.get("vaqt") or now.strftime("%H:%M")

    new_row = [
        sana,
        vaqt,
        record.get("turi", "Tushum"),
        record.get("shaxs", "Noma'lum"),
        record.get("telefon", "") or "",
        str(record.get("tushum_som", 0)),
        str(record.get("xarajat_som", 0)),
        str(record.get("nasiya_som", 0)),
        record.get("qaytarish_sanasi", "") or "",
        record.get("kategoriya", "Boshqa"),
        record.get("izoh", "")
    ]

    if rec_id and 1 <= int(rec_id) <= len(data_rows):
        # Mavjud qatorni yangilash
        data_rows[int(rec_id) - 1] = new_row
    else:
        # Yangi qator qo'shish
        data_rows.append(new_row)

    # CSV ga qayta yozish
    with open(csv_file, mode='w', newline='', encoding='utf-8-sig') as f:
        writer = csv.writer(f, delimiter=';')
        writer.writerow(header)
        writer.writerows(data_rows)

    # FastFood_Tabel varag'i bo'lsa saqlab qolish
    ff_ws_rows = None
    if os.path.exists(excel_file):
        try:
            old_wb = openpyxl.load_workbook(excel_file)
            if "FastFood_Tabel" in old_wb.sheetnames:
                ff_sheet = old_wb["FastFood_Tabel"]
                ff_ws_rows = list(ff_sheet.iter_rows(values_only=True))
        except:
            pass

    # Excelni qayta toza yig'ish
    wb = openpyxl.Workbook()
    ws_all = wb.active
    ws_all.title = "Barcha Yozuvlar"
    setup_sheet_header(ws_all)

    for r in data_rows:
        row_formatted = [
            r[0], r[1], r[2], r[3], r[4],
            int(r[5]) if str(r[5]).isdigit() else 0,
            int(r[6]) if str(r[6]).isdigit() else 0,
            int(r[7]) if str(r[7]).isdigit() else 0,
            r[8], r[9], r[10] if len(r) > 10 else ""
        ]
        append_row_to_worksheet(ws_all, row_formatted)
        m_name = get_month_sheet_title(r[0])
        if m_name not in wb.sheetnames:
            ws_m = wb.create_sheet(m_name)
            setup_sheet_header(ws_m)
        else:
            ws_m = wb[m_name]
        append_row_to_worksheet(ws_m, row_formatted)

    for s in wb.sheetnames:
        render_in_sheet_summary(wb[s])

    if ff_ws_rows:
        ff_ws = wb.create_sheet("FastFood_Tabel")
        for r_vals in ff_ws_rows:
            ff_ws.append(list(r_vals))

    wb.save(excel_file)
    return True

def delete_last_entry(user_id=None):
    rows = get_all_rows(user_id=user_id)
    if rows:
        return delete_specific_row(len(rows), user_id=user_id)
    return None

def get_category_totals(target_month_prefix: str = None, user_id=None):
    paths = get_user_paths(user_id)
    csv_file = paths["csv"]
    tushum_by_cat = {}
    xarajat_by_cat = {}
    nasiya_by_cat = {}

    total_tushum = 0
    total_xarajat = 0
    total_nasiya = 0

    if os.path.exists(csv_file):
        with open(csv_file, mode='r', encoding='utf-8-sig') as f:
            reader = csv.reader(f, delimiter=';')
            headers = next(reader, None)
            for row in reader:
                if len(row) < 9:
                    continue
                date_val = row[0]
                if target_month_prefix and not date_val.startswith(target_month_prefix):
                    continue

                if len(row) >= 11:
                    tushum = int(row[5]) if str(row[5]).isdigit() else 0
                    xarajat = int(row[6]) if str(row[6]).isdigit() else 0
                    nasiya = int(row[7]) if str(row[7]).isdigit() else 0
                    cat = row[9].strip() or "Boshqa"
                else:
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

GEMINI_URL = "https://generativelanguage.googleapis.com/v1beta/models/gemini-3.5-flash:generateContent"
GROQ_URL = "https://api.groq.com/openai/v1/chat/completions"
GROQ_WHISPER_URL = "https://api.groq.com/openai/v1/audio/transcriptions"

SYSTEM_PROMPT = """Siz O'zbek tili nutqi, moliya, fast food oshxonasi va buxgalteriya bo'yicha mutaxassis AI yordamchisiz.
Foydalanuvchi ovozli yoki matnli xabarda nima degan bo'lsa (lotin yoki kirill alifbosida), uni tushunib, quyidagi 2 TA REJIMDAN BIRIDA toza JSON qaytaring.

REJIM 1: FAST FOOD VA ISHCHI HODIMLAR SMENASI (JUDA MUHIM!)
Agar xabarda fast food oshxonasi, smena, ishchi hodimlar ismi va ularning ishlagan puli/soati, kassa topshirish (sdano/сдано), kartadan yechish (karta/снял с карты) yoki ishchilarga berilgan avanslar aytilgan bo'lsa:
(Masalan:
 - "Bugun smenada Nurbek 7700, Jurabek 4500, Mansur 3900, Gobiya 5600, Islom 4500. Kassa topshirildi 13400, kartadan yechildi 13000, Nurbekka 10000 avans berildi"
 - "Smena: Nurbek 10 soat 7700, Jurabek 8 soat 4500, kassa 25000 topshirildi"
 - "Kartadan Polatga 5000 avans berildi, dostavkaga 10000 to'landi"
 - "Bugun smenani yopdik, master Nurbek 7700, Gobiya 5600, sdano 15000")

FAQAT quyidagi JSON formatini qaytaring:
{
  "is_fastfood_smena": true,
  "sana": "YYYY-MM-DD",
  "hodimlar": [
    {
      "ism": "Xodim ismi (masalan: Nurbek, Jurabek, Polat, Mansur, Gobiya, Islom)",
      "lavozim": "master yoki oshpaz yoki bo'sh",
      "soat": null,
      "ish_haqi": 7700
    }
  ],
  "jami_ish_haqi": 26200,
  "kassa_topshirildi_sdano": 13400,
  "kartadan_yechildi": 13000,
  "karta_avans_va_xarajatlar": [
    {
      "shaxs": "Xodim ismi",
      "turi": "avans",
      "summa": 10000,
      "izoh": "Nurbekka avans berildi"
    }
  ],
  "izoh": "Smena haqida to'liq xulosa",
  "toza_matn": "Shovqinsiz chiroyli o'zbekcha matn"
}

REJIM 2: STANDART MOLIYA AMALLARI (ODDIY DAROMAD VA XARAJATLAR)
MOLIYAVIY OPERATSIYA TURLARI:
1. "Tushum" — Haqiqiy sof savdo, daromad, xizmat haqi yoki maosh tushumi bo'lganda (Masalan: "500 ming savdo bo'ldi", "Mijozdan 300 ming pul oldim").
2. "Xarajat" — Oziq-ovqat, benzin, ijara, maosh berish, do'kon xarajati kabi pul sarflanganda (Masalan: "Benzinga 50 ming to'ladim", "Tushlikka 40 ming sarflandi").
3. "Nasiya berildi" — Biror shaxsga qarz yoki nasiyaga tovar berilganda (Masalan: "Doniyorga 200 ming qarz berdim", "Ali 100 ming nasiyaga oldi").
4. "Nasiya qaytdi" — Avval berilgan qarz yoki nasiya qaytarib berilganda (Masalan: "Doniyor qarzini qaytardi 200 ming", "Alidan qarzimni oldim 100 ming", "Nasiyani yopdi").
   * DIQQAT: Qarz qaytishi SOF DAROMAD EMAS! Uning turi qat'iy "Nasiya qaytdi" bo'lishi shart!

QARZNI QAYTARISH MUDDATI:
Agar xabarda qarz qachon qaytarilishi aytilgan bo'lsa (masalan: "29 oktyabr qaytaradi", "juma kuni beradi", "10 kundan keyin", "oy oxirida"):
- "qaytarish_sanasi" maydoniga YYYY-MM-DD ko'rinishida sana yozing (Masalan: "2026-10-29").
- Agar aytilmagan bo'lsa -> null bo'lsin.

TELEFON RAQAM VA QO'SHIMCHA MA'LUMOTLARNI 100% TO'G'RI AJRATISH:
- O'zbekistondagi telefon raqamlari 9 ta raqamdan iborat (masalan: 90 046 16 08 yoki 99 123 45 67 yoki 900461608).
- Foydalanuvchi "telefon nomeri 90 046 16 08" yoki "raqami 900461608" deb aytsa:
  - "telefon" maydoniga FAQAT va FAQAT shu 9 ta raqamni "+998" bilan boshlab yozing: "+998900461608".
  - Agar telefon aytilmagan bo'lsa -> "" (bo'sh qolsin).
- Shaxs ismini aniq yozing.
- "izoh" maydoniga foydalanuvchi aytgan har qanday qo'shimcha tafsilotlarni to'liq kiriting.

O'ZBEKCHA MURAKKAB SONLAR (To'liq hisoblang!):
- "to'rt yuz sakson ming" / "480k" = 480000 so'm
- "uch yuz ellik ming" / "350 ming" = 350000 so'm
- "bir yarim million" = 1500000 so'm
- "bir million" = 1000000 so'm
- "sakson ming" = 80000 so'm

Javobingiz FAQAT toza JSON bo'lishi shart:
{
  "is_fastfood_smena": false,
  "turi": "Tushum" yoki "Xarajat" yoki "Nasiya berildi" yoki "Nasiya qaytdi",
  "shaxs": "Mijoz yoki shaxs ismi (bo'lmasa 'Noma'lum')",
  "telefon": "+998XXXXXXXXX yoki bo'sh",
  "tushum_som": 0,
  "xarajat_som": 0,
  "nasiya_som": 0,
  "kategoriya": "Savdo/Qarz munosabatlari/Oziq-ovqat/Transport/Kommunal/Boshqa",
  "izoh": "Tafsilotli to'liq izoh",
  "toza_matn": "Shovqinsiz chiroyli o'zbekcha jumla",
  "qaytarish_sanasi": "YYYY-MM-DD yoki null"
}
* Eslatma:
- "Nasiya berildi" bo'lsa -> nasiya_som ga summani yozing, tushum va xarajat 0 bo'lsin.
- "Nasiya qaytdi" bo'lsa -> nasiya_som ga summani yozing, tushum va xarajat 0 bo'lsin."""

def save_key_to_env(key_name: str, key_val: str):
    """Yangi API kalitni .env fayliga xavfsiz saqlash"""
    env_path = os.path.join(os.path.dirname(__file__), ".env")
    lines = []
    found = False
    if os.path.exists(env_path):
        with open(env_path, "r", encoding="utf-8") as f:
            lines = f.readlines()
        for idx, line in enumerate(lines):
            if line.strip().startswith(f"{key_name}=") or line.strip().startswith(f"{key_name} ="):
                lines[idx] = f"{key_name}={key_val}\n"
                found = True
                break
    if not found:
        lines.append(f"{key_name}={key_val}\n")
    with open(env_path, "w", encoding="utf-8") as f:
        f.writelines(lines)

def normalize_spoken_numbers(text: str) -> str:
    if not text:
        return text
    t = text.lower()
    replacements = [
        (r"(\d+)\s*(?:yarim|ярим)\s*(?:ming|минг)?", lambda m: str(int(m.group(1)) * 1000 + 500)),
        (r"(\d+)\s*(?:с половиной|и половина)\s*(?:тысяч|тысячи|тыс)?", lambda m: str(int(m.group(1)) * 1000 + 500)),
        (r"\bsakkiz\s*(?:ming\s*)?besh\s*yuz\b|\bсаккиз\s*(?:минг\s*)?беш\s*юз\b|\bвосемь\s*(?:тысяч\s*)?пятьсот\b", "8500"),
        (r"\bsakkiz\s*yarim\s*ming\b|\bсаккиз\s*ярим\s*минг\b|\bвосемь\s*с\s*половиной\s*(?:тысяч)?\b", "8500"),
        (r"\bsakkiz\s*yarim\b|\bсаккиз\s*ярим\b|\bвосемь\s*с\s*половиной\b", "8500"),
        (r"\bsakson\s*besh\s*yuz\b|\bсаксон\s*беш\s*юз\b", "8500"),
        (r"\bto['g]?rt\s*(?:ming\s*)?besh\s*yuz\b|\bтўрт\s*(?:минг\s*)?беш\s*юз\b|\bчетыре\s*(?:тысяч\s*)?пятьсот\b", "4500"),
        (r"\bto['g]?rt\s*yarim\s*ming\b|\bтўрт\s*ярим\s*минг\b|\bчетыре\s*с\s*половиной\s*(?:тысяч)?\b", "4500"),
        (r"\bto['g]?rt\s*yarim\b|\bтўрт\s*ярим\b|\bчетыре\s*с\s*половиной\b", "4500"),
        (r"\bqirq\s*besh\s*yuz\b|\bқирқ\s*беш\s*юз\b", "4500"),
        (r"\bbesh\s*ming\b|\bбеш\s*минг\b|\bпять\s*тысяч\b", "5000"),
        (r"\betti\s*ming\b|\bетти\s*минг\b|\bсемь\s*тысяч\b", "7000"),
        (r"\bsakkiz\s*ming\b|\bсаккиз\s*минг\b|\bвосемь\s*тысяч\b", "8000"),
        (r"\bto['g]?qqiz\s*ming\b|\bтўққиз\s*минг\b|\bдевять\s*тысяч\b", "9000"),
        (r"\bo['g]?n\s*ming\b|\bўн\s*минг\b|\bдесять\s*тысяч\b", "10000"),
        (r"\bo['g]?n\s*besh\s*ming\b|\bўн\s*беш\s*минг\b|\bпятнадцать\s*тысяч\b", "15000"),
        (r"\bo['g]?n\s*yetti\s*ming\b|\bўн\s*етти\s*минг\b|\bсемнадцать\s*тысяч\b", "17000"),
        (r"\byigirma\s*ming\b|\bйигирма\s*минг\b|\bдвадцать\s*тысяч\b", "20000"),
        (r"\byigirma\s*besh\s*ming\b|\bйигирма\s*беш\s*минг\b|\bдвадцать\s*пять\s*тысяч\b", "25000"),
        (r"\bo['g]?ttiz\s*ming\b|\bўттиз\s*минг\b|\bтридцать\s*тысяч\b", "30000"),
    ]
    for pattern, repl in replacements:
        if callable(repl):
            t = re.sub(pattern, repl, t)
        else:
            t = re.sub(pattern, repl, t)
    return t

async def transcribe_audio_with_groq(audio_bytes: bytes, language: str = None) -> str:
    if not GROQ_API_KEY:
        raise ValueError("GROQ_API_KEY topilmadi.")

    headers = {
        "Authorization": f"Bearer {GROQ_API_KEY}",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
    }

    form = aiohttp.FormData()
    form.add_field('file', audio_bytes, filename='voice.ogg', content_type='audio/ogg')
    form.add_field('model', 'whisper-large-v3')
    if language:
        form.add_field('language', language)
    form.add_field('temperature', '0.0')
    form.add_field('response_format', 'json')
    form.add_field('prompt', (
        "Нурбек 8500, Журабек 8500, Махмуд 5000, Нодира 4500, Робия 4500, Рохат 4500, "
        "Сдано 10000, Снял с карты 17000, Сантехник 10000, Малиш 5000, Шина 983 25000. "
        "Nurbek 8500, Jorabek 8500, Mahmud 5000, Nodira 4500, Robiya 4500, Rohat 4500, "
        "Sdano 10000, Kassa 10000, Kartadan 17000."
    ))

    async with aiohttp.ClientSession() as session:
        try:
            async with session.post(
                GROQ_WHISPER_URL,
                data=form,
                headers=headers,
                timeout=aiohttp.ClientTimeout(total=45)
            ) as resp:
                if resp.status == 200:
                    result = await resp.json()
                    raw_text = result.get("text", "").strip()
                    cleaned_text = clean_stt_noise(raw_text)
                    return cleaned_text
                else:
                    err_body = await resp.text()
                    raise Exception(f"Groq Whisper xatosi ({resp.status}): {err_body[:150]}")
        except Exception as e:
            logger.error(f"Whisper request error: {e}")
            raise e

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
        "model": "openai/gpt-oss-120b",
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


def parse_with_smart_rules(text: str) -> dict:
    if not text:
        return None
    lower = text.lower()

    # 1. Sana aniqlash (Toshkent vaqti UTC+5)
    tz_tashkent = datetime.timezone(datetime.timedelta(hours=5))
    now_tashkent = datetime.datetime.now(tz_tashkent)
    sana = now_tashkent.strftime("%Y-%m-%d")

    # Kecha so'zi bo'lsa
    if "kecha" in lower or "вчера" in lower:
        sana = (now_tashkent - datetime.timedelta(days=1)).strftime("%Y-%m-%d")
    else:
        d_m = re.search(r'\b(\d{1,2})[\s.-]*(oktyabr|noyabr|dekabr|yanvar|fevral|mart|aprel|may|iyun|iyul|avgust|sentabr|октябр|ноябр|декабр|январ|феврал|март|апрел|май|июн|июл|август|сентабр|\.01|\.02|\.03|\.04|\.05|\.06|\.07|\.08|\.09|\.10|\.11|\.12)\b', lower)
        if d_m:
            day_num = int(d_m.group(1))
            m_word = d_m.group(2)
            month_map = {
                "01": 1, "yan": 1, "янв": 1,
                "02": 2, "fev": 2, "фев": 2,
                "03": 3, "mar": 3, "мар": 3,
                "04": 4, "apr": 4, "апр": 4,
                "05": 5, "may": 5, "май": 5,
                "06": 6, "iyun": 6, "июн": 6,
                "07": 7, "iyul": 7, "июл": 7,
                "08": 8, "avg": 8, "авг": 8,
                "09": 9, "sen": 9, "сен": 9,
                "10": 10, "okt": 10, "окт": 10,
                "11": 11, "noy": 11, "ной": 11,
                "12": 12, "dek": 12, "дек": 12
            }
            m_num = now_tashkent.month
            for k, v in month_map.items():
                if k in m_word:
                    m_num = v
                    break
            sana = f"{now_tashkent.year}-{m_num:02d}-{day_num:02d}"
        else:
            day_only_m = re.search(r'\b(?:число|sana|kuni)\s*(\d{1,2})\b|\b(\d{1,2})[\s-]*(?:число|kuni)\b', lower)
            if day_only_m:
                d_val = int(day_only_m.group(1) or day_only_m.group(2))
                if 1 <= d_val <= 31:
                    cur_ym = sana[:7]
                    sana = f"{cur_ym}-{d_val:02d}"

    # 2. Xodimlar ro'yxati (o'zbekcha va ruscha qo'shimchalar bilan: -ga, -ka, -ning, -ни...)
    suffixes = r'(?:ga|ka|ge|ke|ning|ni|га|ка|нинг|ни|а|у)?'
    workers_list = [
        ("Нурбек", ["нурбек", "нурбик", "nurbek", "nurbik", "nurevika", "nurev", "nurbeg"]),
        ("Журабек", ["журабек", "jorabek", "jo'rabek", "jurabek", "corabek", "corab", "jora", "jo'ra", "жура"]),
        ("Мухмуд", ["мухмуд", "махмуд", "mahmud", "muhmad", "muhammadjon", "maxmud", "маxмуд"]),
        ("Нодира", ["нодира", "nodira", "nodera", "нодера"]),
        ("Робия", ["робия", "robiya", "rabia", "рабия", "gobiya"]),
        ("Рохат", ["рохат", "rohat", "rahat", "рахат"])
    ]

    found_workers = []
    for disp_name, aliases in workers_list:
        for al in aliases:
            pattern = rf'{al}{suffixes}[^\d\n]*?(\d+[\d\s.,]*)'
            m = re.search(pattern, lower)
            if m:
                raw_str = re.sub(r'[^\d]', '', m.group(1))
                if raw_str:
                    num_val = int(raw_str)
                    if num_val < 100:
                        num_val *= 100
                    found_workers.append({'ism': disp_name, 'ish_haqi': num_val, 'lavozim': "bo'sh", 'soat': None})
                    break

    # Сдано (касса topshirildi)
    sdano_m = re.search(r'(?:сдано|сдал|sdano|kassa|касса|topshirildi)[^\d\n]*?(\d+[\d\s.,]*)', lower)
    sdano = int(re.sub(r'[^\d]', '', sdano_m.group(1))) if sdano_m and re.sub(r'[^\d]', '', sdano_m.group(1)) else 0

    # Снял с карты
    snyal_m = re.search(r'(?:снял\s*с\s*карты|снял|kartadan\s*yechildi|kartadan|yechildi|snyal)[^\d\n]*?(\d+[\d\s.,]*)', lower)
    snyal = int(re.sub(r'[^\d]', '', snyal_m.group(1))) if snyal_m and re.sub(r'[^\d]', '', snyal_m.group(1)) else 0

    # Savdo (выручка)
    savdo_m = re.search(r'(?:выручка|viruchka|savdo|савдо)[^\d\n]*?(\d+[\d\s.,]*)', lower)
    savdo = int(re.sub(r'[^\d]', '', savdo_m.group(1))) if savdo_m and re.sub(r'[^\d]', '', savdo_m.group(1)) else 0

    # Karta+ xarajatlari
    karta_expenses = []
    karta_special = [
        ("Сантехник", ["сантехник", "santexnik", "santex"]),
        ("Малиш", ["малиш", "malish"]),
        ("Шина 983", ["шина 983", "шина", "shina 983", "shina", "балон", "balon"]),
        ("Блогер", ["блогер", "bloger"]),
        ("Страховка", ["страховка", "страхов", "straxov"]),
        ("Холод", ["холод", "xolod", "холодильник", "xolodilnik"])
    ]
    for exp_name, exp_aliases in karta_special:
        for al in exp_aliases:
            pattern = rf'{al}{suffixes}[^\d\n]*?(\d+[\d\s.,]*)'
            m = re.search(pattern, lower)
            if m:
                raw_exp = re.sub(r'[^\d]', '', m.group(1))
                if raw_exp:
                    sum_val = int(raw_exp)
                    if sum_val < 100:
                        sum_val *= 1000
                    karta_expenses.append({'shaxs': exp_name, 'summa': sum_val, 'turi': 'xarajat', 'izoh': exp_name})
                    break

    if found_workers or sdano > 0 or snyal > 0 or savdo > 0 or karta_expenses:
        zp = sum(w['ish_haqi'] for w in found_workers)
        return {
            'is_fastfood_smena': True,
            'sana': sana,
            'vaqt': now_tashkent.strftime("%H:%M"),
            'hodimlar': found_workers,
            'jami_ish_haqi': zp,
            'kassa_topshirildi_sdano': sdano,
            'kartadan_yechildi': snyal,
            'savdo_viruchka': savdo,
            'karta_avans_va_xarajatlar': karta_expenses,
            'izoh': "Kunlik smena",
            'toza_matn': text
        }

    # Raqamlar bo'lsa umumiy xarajat yoki kassa deb qabul qilamiz
    nums = re.findall(r'\b\d+[\d\s.,]*\b', text)
    if nums:
        cleaned_nums = [int(re.sub(r'[^\d]', '', n)) for n in nums if re.sub(r'[^\d]', '', n)]
        if cleaned_nums:
            max_sum = max(cleaned_nums)
            if 'ming' in lower and max_sum < 10000:
                max_sum *= 1000
            if any(w in lower for w in ['kassa', 'касса', 'tushum', 'daromad']):
                return {
                    'is_fastfood_smena': True,
                    'sana': sana,
                    'vaqt': now_tashkent.strftime("%H:%M"),
                    'hodimlar': [],
                    'jami_ish_haqi': 0,
                    'kassa_topshirildi_sdano': max_sum,
                    'kartadan_yechildi': 0,
                    'savdo_viruchka': 0,
                    'karta_avans_va_xarajatlar': [],
                    'izoh': text,
                    'toza_matn': text
                }
            else:
                return {
                    'is_fastfood_smena': True,
                    'sana': sana,
                    'vaqt': now_tashkent.strftime("%H:%M"),
                    'hodimlar': [],
                    'jami_ish_haqi': 0,
                    'kassa_topshirildi_sdano': 0,
                    'kartadan_yechildi': 0,
                    'savdo_viruchka': 0,
                    'karta_avans_va_xarajatlar': [{'shaxs': text[:30], 'summa': max_sum, 'turi': 'xarajat', 'izoh': text}],
                    'izoh': text,
                    'toza_matn': text
                }

    return None

async def parse_note_with_ai(text: str = "", audio_b64: str = None) -> dict:
    now = datetime.datetime.now()
    date_str = now.strftime("%Y-%m-%d")
    time_str = now.strftime("%H:%M")
    prompt_text = f"Bugungi sana: {date_str} {time_str}.\nFoydalanuvchi ovozli/matnli xabari: '{text}'"

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
        rule_res = parse_with_smart_rules(text)
        if rule_res:
            logger.info("✅ Qoidalar orqali muvaffaqiyatli tahlil qilindi")
            if not rule_res.get("sana"):
                rule_res["sana"] = date_str
            rule_res["vaqt"] = time_str
            return rule_res
        raise Exception(f"AI javob bera olmadi. ({last_error[:100]})")

    raw_json = re.sub(r'^```json\s*', '', raw_json)
    raw_json = re.sub(r'\s*```$', '', raw_json)
    parsed = json.loads(raw_json)

    if not parsed.get("sana"):
        parsed["sana"] = date_str
    parsed["vaqt"] = time_str
    return parsed

def get_next_data_row(ws):
    r = 1
    while ws.cell(row=r, column=1).value is not None:
        r += 1
    return r

def append_row_to_worksheet(ws, row_data):
    target_row = get_next_data_row(ws)

    thin_border = Border(left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'), top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9'))
    
    turi = str(row_data[2] or "").strip().lower()
    tushum_val = int(row_data[5] or 0) if len(row_data) > 5 else int(row_data[4] or 0)
    xarajat_val = int(row_data[6] or 0) if len(row_data) > 6 else int(row_data[5] or 0)
    nasiya_val = int(row_data[7] or 0) if len(row_data) > 7 else int(row_data[6] or 0)

    if "qaytdi" in turi:
        fill_color = "E2EFDA"  # Yashil (Chunki pul kassaga qaytdi)
    elif "berildi" in turi or ("nasiya" in turi and xarajat_val == 0 and tushum_val == 0):
        fill_color = "FFF2CC"  # Sariq (Nasiya berildi)
    elif "xarajat" in turi or xarajat_val > 0:
        fill_color = "FCE4D6"  # Qizil (Xarajat)
    elif "tushum" in turi or tushum_val > 0:
        fill_color = "E2EFDA"  # Yashil (Daromad)
    else:
        fill_color = "F9FBFD" if target_row % 2 == 0 else "FFFFFF"

    row_fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type="solid")
    font_data = Font(name="Arial", size=10)

    ws.row_dimensions[target_row].height = 22

    for col_idx, val in enumerate(row_data, 1):
        cell = ws.cell(row=target_row, column=col_idx, value=val)
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

def append_to_excel_and_csv(data: dict, user_id=None):
    paths = get_user_paths(user_id)
    csv_file = paths["csv"]
    excel_file = paths["excel"]

    with open(csv_file, mode='a', newline='', encoding='utf-8-sig') as f:
        writer = csv.writer(f, delimiter=';')
        writer.writerow([
            data.get("sana", ""),
            data.get("vaqt", ""),
            data.get("turi", "Tushum"),
            data.get("shaxs", "Noma'lum"),
            data.get("telefon", ""),
            data.get("tushum_som", 0),
            data.get("xarajat_som", 0),
            data.get("nasiya_som", 0),
            data.get("qaytarish_sanasi", "") or "",
            data.get("kategoriya", "Boshqa"),
            data.get("izoh", "")
        ])

    if not os.path.exists(excel_file):
        wb = create_styled_excel(excel_file)
    else:
        wb = openpyxl.load_workbook(excel_file)

    row_data = [
        data.get("sana", ""),
        data.get("vaqt", ""),
        data.get("turi", "Tushum"),
        data.get("shaxs", "Noma'lum"),
        data.get("telefon", "") or "",
        int(data.get("tushum_som", 0)),
        int(data.get("xarajat_som", 0)),
        int(data.get("nasiya_som", 0)),
        data.get("qaytarish_sanasi", "") or "",
        data.get("kategoriya", "Boshqa"),
        data.get("izoh", "")
    ]

    ws_all = wb["Barcha Yozuvlar"] if "Barcha Yozuvlar" in wb.sheetnames else wb.active
    setup_sheet_header(ws_all)
    append_row_to_worksheet(ws_all, row_data)
    render_in_sheet_summary(ws_all)

    date_val = data.get("sana", "")
    month_sheet_name = get_month_sheet_title(date_val)
    if month_sheet_name not in wb.sheetnames:
        ws_month = wb.create_sheet(month_sheet_name)
        setup_sheet_header(ws_month)
    else:
        ws_month = wb[month_sheet_name]

    append_row_to_worksheet(ws_month, row_data)
    render_in_sheet_summary(ws_month)

    wb.save(excel_file)

WEB_APP_URL = os.environ.get("WEB_APP_URL", "https://moliya-webapp.vercel.app").strip()
if not WEB_APP_URL.startswith("http"):
    WEB_APP_URL = "https://" + WEB_APP_URL

from telegram import WebAppInfo

def get_user_webapp_url(user_id=None) -> str:
    uid = str(user_id or DEFAULT_USER_ID)
    base = WEB_APP_URL.split("?")[0]
    return f"{base}?user_id={uid}"

def get_action_keyboard(user_id=None):
    keyboard = [
        [
            InlineKeyboardButton("📅 1-Dekada (1-10)", callback_data="ff_dek_1"),
            InlineKeyboardButton("📅 2-Dekada (11-20)", callback_data="ff_dek_2"),
        ],
        [
            InlineKeyboardButton("📅 3-Dekada (21-31)", callback_data="ff_dek_3"),
            InlineKeyboardButton("📊 To'liq Oy", callback_data="ff_month_full"),
        ],
        [
            InlineKeyboardButton("📥 Excel Yuklash (.xlsx)", callback_data="download_excel"),
            InlineKeyboardButton("🗑 Smenani o'chirish", callback_data="ff_delete_menu")
        ]
    ]
    return InlineKeyboardMarkup(keyboard)

def get_reply_menu(user_id=None):
    keyboard = [
        [KeyboardButton("🍔 Fast Food Tabel"), KeyboardButton("📊 Excel Jadvalini Yuklash (.xlsx)")],
        [KeyboardButton("🗑 Smenani o'chirish"), KeyboardButton("ℹ️ Yordam")]
    ]
    return ReplyKeyboardMarkup(keyboard, resize_keyboard=True)

async def send_rows_list(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat_id = update.effective_chat.id
    user_id = update.effective_user.id if update.effective_user else chat_id
    rows = get_all_rows(user_id=user_id)

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
    chat_id = update.effective_chat.id
    user_id = update.effective_user.id if update.effective_user else chat_id
    now = datetime.datetime.now()
    current_month_prefix = now.strftime("%Y-%m")
    month_name_uz = MONTH_NAMES_UZ.get(now.strftime("%m"), "")

    stats = get_category_totals(target_month_prefix=current_month_prefix, user_id=user_id)

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

def generate_daily_digest(target_date: str = None, user_id=None) -> str:
    """Bugungi kungi Fast Food smenasi va oylik kassa sarhisobi"""
    now = datetime.datetime.now()
    if not target_date:
        target_date = now.strftime("%Y-%m-%d")

    smenas = load_fastfood_smenas(user_id=user_id)
    today_smena = next((s for s in smenas if s.get("sana") == target_date), None)

    parts = target_date.split("-")
    year_num = parts[0]
    month_key = parts[1] if len(parts) > 1 else "10"
    day_num = parts[2] if len(parts) > 2 else "01"
    month_name = MONTH_NAMES_UZ.get(month_key, "")

    text = f"⏰ **KUNLIK FAST FOOD HISOBOTI (DAILY DIGEST)**\n"
    text += f"📅 **Sana:** `{day_num}-{month_name} {year_num}-yil`\n"
    text += f"━━━━━━━━━━━━━━━━━━━━━━\n\n"

    if today_smena:
        text += format_fastfood_shift_message(today_smena)
    else:
        text += f"⚠️ **Bugun (`{target_date}`) hali smena yozuvlari kiritilmadi.**\n\n"
        text += "👉 *Smenani kiritish uchun botga ovozli xabar (golosovoy) yuboring!*"

    return text

async def send_daily_digest(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat_id = update.effective_chat.id
    user_id = update.effective_user.id if update.effective_user else chat_id
    add_subscriber(chat_id)
    today_str = datetime.datetime.now().strftime("%Y-%m-%d")
    report = generate_daily_digest(today_str, user_id=user_id)

    if update.callback_query:
        try:
            await update.callback_query.edit_message_text(report, parse_mode="Markdown", reply_markup=get_action_keyboard())
        except:
            await context.bot.send_message(chat_id=chat_id, text=report, parse_mode="Markdown", reply_markup=get_action_keyboard())
    else:
        await update.message.reply_text(report, parse_mode="Markdown", reply_markup=get_action_keyboard())

async def daily_digest_scheduler_task(app):
    """Har kuni soat 20:00 da (Toshkent vaqti bilan) barcha obunachilarga kunlik avtomatik hisobotni yuboradi"""
    await asyncio.sleep(25)
    while True:
        try:
            # Toshkent vaqti: UTC+5
            utc_now = datetime.datetime.now(datetime.timezone.utc)
            tashkent_now = utc_now + datetime.timedelta(hours=5)
            today_str = tashkent_now.strftime("%Y-%m-%d")
            hour = tashkent_now.hour

            last_sent_date = ""
            if os.path.exists(DIGEST_STATE_FILE):
                try:
                    with open(DIGEST_STATE_FILE, "r", encoding="utf-8") as f:
                        last_sent_date = json.load(f).get("last_digest_date", "")
                except:
                    pass

            if hour >= 20 and last_sent_date != today_str:
                subs = get_all_subscribers()
                if subs:
                    digest_text = generate_daily_digest(today_str)
                    keyboard = get_action_keyboard()
                    for chat_id in subs:
                        try:
                            await app.bot.send_message(
                                chat_id=chat_id,
                                text=digest_text,
                                parse_mode="Markdown",
                                reply_markup=keyboard
                            )
                            logger.info(f"✅ Daily digest chat_id {chat_id} ga yuborildi!")
                        except Exception as send_err:
                            logger.error(f"Daily digest send error ({chat_id}): {send_err}")

                    with open(DIGEST_STATE_FILE, "w", encoding="utf-8") as f:
                        json.dump({"last_digest_date": today_str}, f)

        except Exception as e:
            logger.error(f"daily_digest_scheduler_task error: {e}")

        await asyncio.sleep(600)  # Har 10 daqiqada tekshiradi

async def start_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat_id = update.effective_chat.id
    user_id = update.effective_user.id if update.effective_user else chat_id
    add_subscriber(chat_id)
    try:
        from telegram import MenuButtonCommands
        await context.bot.set_chat_menu_button(chat_id=chat_id, menu_button=MenuButtonCommands())
    except Exception as me_err:
        pass

    welcome = (
        "👋 Assalomu alaykum!\n\n"
        "Men sizning **Fast Food Xodimlar Tabeli & Kassa Hisoboti** yordamchingizman 🍔\n\n"
        "🎙 Menga xodimlar oyligi va kassa hisobotini **ovozli xabar (golosovoy)** yoki matn ko'rinishida yuboring:\n"
        "👉 Masalan:\n"
        "• *\"Нурбек 8500, Журабек 8500, Сдано 10000, Снял с карты 17000\"*\n"
        "• *\"Нурбек 8500\"*\n"
        "• *\"Сантехник 10000\"*\n\n"
        "✨ Barcha yozuvlar avtomatik tarzda **Fast Food Excel Tabeli**ga tushadi!"
    )
    await update.message.reply_text(welcome, parse_mode="Markdown", reply_markup=get_reply_menu())

async def send_excel_file(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat_id = update.effective_chat.id
    user_id = update.effective_user.id if update.effective_user else chat_id
    try:
        from tabel_sync import sync_fastfood_tabel_file
        sync_fastfood_tabel_file(user_id)
    except Exception as sync_e:
        logger.error(f"Sinxronizatsiya xatosi: {sync_e}")

    paths = get_user_paths(user_id)
    await context.bot.send_chat_action(chat_id, ChatAction.UPLOAD_DOCUMENT)

    base_dir = os.path.dirname(os.path.abspath(__file__))
    user_dir = os.path.dirname(paths["excel"])

    tz_tashkent = datetime.timezone(datetime.timedelta(hours=5))
    now = datetime.datetime.now(tz_tashkent)
    cur_month_str = f"{now.month:02d}"
    cur_month_uz = MONTH_NAMES_UZ.get(cur_month_str, "NOYABR").capitalize()

    # Barcha oylar (Oktyabr, Noyabr...) to'liq kiritilgan asosiy Excel fayli
    candidates = [
        os.path.join(user_dir, "Fast_Food_Tabel_2026.xlsx"),
        os.path.join(base_dir, "Fast_Food_Tabel_2026.xlsx"),
        os.path.join(user_dir, f"Fast_Food_Tabel_{cur_month_uz}_{now.year}.xlsx"),
        os.path.join(base_dir, f"Fast_Food_Tabel_{cur_month_uz}_{now.year}.xlsx"),
        os.path.join(user_dir, "Fast_Food_Tabel_Oktyabr_2026.xlsx"),
        os.path.join(base_dir, "Fast_Food_Tabel_Oktyabr_2026.xlsx"),
    ]
    if os.path.exists(r"d:\Antigravity\123"):
        candidates.extend([
            r"d:\Antigravity\123\Fast_Food_Tabel_2026.xlsx",
            r"d:\Antigravity\123\telegrambot\Fast_Food_Tabel_2026.xlsx",
        ])

    file_to_send = None
    for cand in candidates:
        if os.path.exists(cand):
            file_to_send = cand
            break

    if file_to_send:
        with open(file_to_send, 'rb') as doc:
            await context.bot.send_document(
                chat_id=chat_id,
                document=doc,
                filename=f"Fast_Food_Tabel_{cur_month_uz}_{now.year}.xlsx",
                caption=(
                    f"🍔 **Fast Food Xodimlar Tabeli & Kassa Hisoboti ({cur_month_uz} {now.year})**\n\n"
                    f"📅 **Excel fayli ichida har bir oy alohida sahifada:**\n"
                    f"  • `Tabel 10.2026` — Oktyabr oyi to'liq sarhisobi\n"
                    f"  • `Tabel 11.2026` — Noyabr oyi (yangi kiritilgan smenalar)\n\n"
                    f"📌 Yangi oy boshlanganda barcha ma'lumotlar avtomatik tarzda yangi oy varag'iga yozib boriladi."
                ),
                parse_mode="Markdown"
            )
        return

    await context.bot.send_message(chat_id=chat_id, text="⚠️ Hali jadvalga ma'lumot kiritilmagan.")

async def handle_delete_specific(update: Update, context: ContextTypes.DEFAULT_TYPE, row_id: int):
    chat_id = update.effective_chat.id
    user_id = update.effective_user.id if update.effective_user else chat_id
    deleted = delete_specific_row(row_id, user_id=user_id)

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
    chat_id = update.effective_chat.id
    user_id = update.effective_user.id if update.effective_user else chat_id
    text = update.message.text.strip()
    parts = text.split()

    if len(parts) >= 2 and parts[1].isdigit():
        row_id = int(parts[1])
        await handle_delete_specific(update, context, row_id)
    else:
        deleted = delete_last_entry(user_id=user_id)
        if deleted:
            msg_text = f"🗑 **Oxirgi yozuv ({deleted['id']}-qator) o'chirildi!** ❌\n📝 *{deleted['izoh']}*"
        else:
            msg_text = "⚠️ O'chirish uchun jadvalda yozuv topilmadi."
        await update.message.reply_text(msg_text, parse_mode="Markdown")

async def send_fastfood_tabel(update: Update, context: ContextTypes.DEFAULT_TYPE, start_day: int = 1, end_day: int = 10, title: str = "1-DEKADA (1-10 KUN)"):
    chat_id = update.effective_chat.id
    user_id = update.effective_user.id if update.effective_user else chat_id
    now = datetime.datetime.now()
    month_prefix = now.strftime("%Y-%m")
    summary = calculate_fastfood_summary(user_id=user_id, month_prefix=month_prefix, start_day=start_day, end_day=end_day)
    text = format_fastfood_tabel_report(summary, title=title)

    keyboard = [
        [
            InlineKeyboardButton("📅 1-Dekada (1-10)", callback_data="ff_dek_1"),
            InlineKeyboardButton("📅 2-Dekada (11-20)", callback_data="ff_dek_2"),
        ],
        [
            InlineKeyboardButton("📅 3-Dekada (21-31)", callback_data="ff_dek_3"),
            InlineKeyboardButton("📊 To'liq Oy", callback_data="ff_month_full"),
        ],
        [
            InlineKeyboardButton("📊 Excel Jadvalini Yuklash", callback_data="download_excel"),
            InlineKeyboardButton("🗑 Smenani o'chirish", callback_data="ff_delete_menu")
        ]
    ]
    reply_markup = InlineKeyboardMarkup(keyboard)

    if update.callback_query:
        try:
            await update.callback_query.edit_message_text(text, parse_mode="Markdown", reply_markup=reply_markup)
        except:
            await context.bot.send_message(chat_id=chat_id, text=text, parse_mode="Markdown", reply_markup=reply_markup)
    else:
        await update.message.reply_text(text, parse_mode="Markdown", reply_markup=reply_markup)


async def send_fastfood_delete_menu(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat_id = update.effective_chat.id
    user_id = update.effective_user.id if update.effective_user else chat_id
    smenas = load_fastfood_smenas(user_id)
    if not smenas:
        text = "⚠️ Hali smena yozuvlari yo'q."
        if update.callback_query:
            await update.callback_query.edit_message_text(text)
        else:
            await context.bot.send_message(chat_id=chat_id, text=text)
        return

    text = "🗑 **O'CHIRMOQCHI BO'LGAN SMENA SANASINI TANLANG:**\n*(Tugmani bossangiz, o'sha kundagi smena o'chiriladi)*\n"
    keyboard = []
    for s in smenas[-8:]:
        sana = s.get("sana", "")
        zp = s.get("jami_ish_haqi", 0)
        keyboard.append([InlineKeyboardButton(f"❌ {sana} smenasi ({zp:,} so'm)", callback_data=f"del_smena_{sana}")])
    keyboard.append([InlineKeyboardButton("🔙 Tabelga qaytish", callback_data="ff_month_full")])

    if update.callback_query:
        await update.callback_query.edit_message_text(text, parse_mode="Markdown", reply_markup=InlineKeyboardMarkup(keyboard))
    else:
        await context.bot.send_message(chat_id=chat_id, text=text, parse_mode="Markdown", reply_markup=InlineKeyboardMarkup(keyboard))

async def handle_delete_smena(update: Update, context: ContextTypes.DEFAULT_TYPE, smena_sana: str):
    chat_id = update.effective_chat.id
    user_id = update.effective_user.id if update.effective_user else chat_id
    smenas = load_fastfood_smenas(user_id)
    new_smenas = [s for s in smenas if s.get("sana") != smena_sana]
    if len(new_smenas) < len(smenas):
        save_fastfood_smenas(new_smenas, user_id=user_id)
        try:
            from tabel_sync import sync_fastfood_tabel_file
            sync_fastfood_tabel_file(user_id)
        except:
            pass
        msg = f"✅ `{smena_sana}` sanasidagi smena muvaffaqiyatli o'chirildi!"
    else:
        msg = f"⚠️ `{smena_sana}` sanasidagi smena topilmadi."

    if update.callback_query:
        await update.callback_query.edit_message_text(msg, parse_mode="Markdown")
    else:
        await context.bot.send_message(chat_id=chat_id, text=msg, parse_mode="Markdown")
    await send_fastfood_tabel(update, context, start_day=1, end_day=31, title="To'liq Oylik Tabel")

async def handle_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    data = query.data

    user_id = update.effective_user.id if update.effective_user else update.effective_chat.id

    if data in ("download_excel", "download_csv"):
        await send_excel_file(update, context)
    elif data == "delete_last":
        rows = get_all_rows(user_id=user_id)
        if rows:
            await handle_delete_specific(update, context, rows[-1]["id"])
    elif data == "show_rows_list":
        await send_rows_list(update, context)
    elif data == "category_report":
        await send_category_report(update, context)
    elif data == "daily_digest":
        await send_daily_digest(update, context)
    elif data == "ff_dek_1":
        await send_fastfood_tabel(update, context, start_day=1, end_day=10, title="1-Dekada (1-10 kun)")
    elif data == "ff_dek_2":
        await send_fastfood_tabel(update, context, start_day=11, end_day=20, title="2-Dekada (11-20 kun)")
    elif data == "ff_dek_3":
        await send_fastfood_tabel(update, context, start_day=21, end_day=31, title="3-Dekada (21-31 kun)")
    elif data == "ff_delete_menu":
        await send_fastfood_delete_menu(update, context)
    elif data.startswith("del_smena_"):
        smena_sana = data.replace("del_smena_", "")
        await handle_delete_smena(update, context, smena_sana)
    elif data == "ff_month_full":
        await send_fastfood_tabel(update, context, start_day=1, end_day=31, title="To'liq Oylik Tabel")
    elif data.startswith("del_row_"):
        row_id = int(data.replace("del_row_", ""))
        await handle_delete_specific(update, context, row_id)

async def handle_text_or_voice(update: Update, context: ContextTypes.DEFAULT_TYPE):
    msg = update.message
    if not msg:
        return

    user_id = msg.chat.id
    add_subscriber(user_id)
    text_input = msg.text or ""
    original_voice_text = text_input

    # Groq API kalitini chatdan to'g'ridan-to'g'ri o'rnatish
    clean_text = text_input.strip()
    if clean_text.startswith("gsk_") or "gsk_" in clean_text:
        m = re.search(r'(gsk_[a-zA-Z0-9_-]+)', clean_text)
        if m:
            new_groq_key = m.group(1)
            try:
                headers = {"Authorization": f"Bearer {new_groq_key}"}
                async with aiohttp.ClientSession() as session:
                    async with session.get("https://api.groq.com/openai/v1/models", headers=headers, timeout=aiohttp.ClientTimeout(total=10)) as resp:
                        if resp.status == 200:
                            save_key_to_env("GROQ_API_KEY", new_groq_key)
                            global GROQ_API_KEY
                            GROQ_API_KEY = new_groq_key
                            os.environ["GROQ_API_KEY"] = new_groq_key
                            success_msg = (
                                "🎉 **Groq AI Kaliti Muvaffaqiyatli O'rnatildi va Tekshirildi!** ✅\n\n"
                                "🎙 Endi bemalol menga **ovozli xabar (golosovoy)** yuborishingiz mumkin!\n"
                                "Bot barcha xodimlaringiz, smenalar va kassa to'lovlarini eshitib, jadvalga yozib boradi."
                            )
                            await msg.reply_text(success_msg, parse_mode="Markdown")
                            return
                        else:
                            await msg.reply_text("❌ Ushbu Groq kaliti ishlamadi (xato kalit). Iltimos, console.groq.com/keys dan yangi kalit olib yuboring.")
                            return
            except Exception as test_err:
                await msg.reply_text(f"❌ Kalitni tekshirishda xatolik: {test_err}")
                return

    if text_input in ("📊 Excel Jadvalini Yuklash (.xlsx)", "📊 Jadvalni yuklab olish (Excel)", "/jadval", "/excel"):
        await send_excel_file(update, context)
        return
    elif text_input in ("🍔 Fast Food Tabel", "Fast Food Tabel", "/tabel", "/smena"):
        await send_fastfood_tabel(update, context, start_day=1, end_day=10, title="1-Dekada (1-10 kun)")
        return
    elif text_input in ("🗑 Smenani o'chirish", "Smenani o'chirish", "/ochirish", "/bekor"):
        await send_fastfood_delete_menu(update, context)
        return
    elif text_input in ("ℹ️ Yordam", "Yordam", "/help", "/start"):
        await start_command(update, context)
        return
    elif text_input in ("⏰ Kunlik Hisobot (Daily Digest)", "⏰ Kunlik Hisobot", "/kunlik", "/digest"):
        await send_daily_digest(update, context)
        return

    await context.bot.send_chat_action(msg.chat.id, ChatAction.TYPING)
    wait_msg = await msg.reply_text("🎙 AI xabarni eshitib, tegishli oy jadvaliga joylamoqda...")

    audio_b64 = None

    try:
        if msg.voice or msg.audio:
            media = msg.voice or msg.audio
            if media.file_size and media.file_size > 25 * 1024 * 1024:
                await wait_msg.delete()
                await msg.reply_text("❌ Ovozli xabar hajmi juda katta (maksimal 25 MB).")
                return

            audio_bytes = None
            for attempt in range(3):
                try:
                    file_obj = await media.get_file(read_timeout=60, write_timeout=60, connect_timeout=60)
                    byte_arr = await file_obj.download_as_bytearray(read_timeout=60, write_timeout=60, connect_timeout=60)
                    audio_bytes = bytes(byte_arr)
                    audio_b64 = base64.b64encode(audio_bytes).decode("utf-8")
                    break
                except Exception as dl_err:
                    logger.warning(f"Audio yuklab olish xatosi ({attempt+1}-urinish): {dl_err}")
                    if attempt == 2:
                        await wait_msg.delete()
                        await msg.reply_text("❌ Ovozli xabarni Telegram serveridan yuklab bo'lmadi. Qaytadan yuboring.")
                        return
                    await asyncio.sleep(1)

            if not audio_bytes:
                await wait_msg.delete()
                await msg.reply_text("❌ Ovozli xabar bo'sh yoki yuklanmadi.")
                return

            try:
                logger.info("🎤 Groq Whisper API (whisper-large-v3, language='ru') ga ovoz yuborilmoqda...")
                transcribed_text = await transcribe_audio_with_groq(audio_bytes, language="ru")
                logger.info(f"✅ Ovoz matnga o'girildi: '{transcribed_text}'")

                if not transcribed_text:
                    await wait_msg.delete()
                    await msg.reply_text("⚠️ Ovozli xabardan so'zlarni ajratib bo'lmadi. Iltimos, aniqroq gapirib qayta yuboring.")
                    return

                text_input = transcribed_text
            except PermissionError as p_err:
                logger.error(f"Whisper kalit xatosi: {p_err}")
                await wait_msg.delete()
                guide_msg = (
                    "🔑 **Ovozli xabarni eshitish uchun Groq AI kaliti noto'g'ri yoki kiritilmagan!**\n\n"
                    "1. [console.groq.com/keys](https://console.groq.com/keys) saytiga kiring\n"
                    "2. **\"Create API Key\"** tugmasini bosing\n"
                    "3. Chiqqan kalitni (`gsk_...`) nusxalab, **to'g'ridan-to'g'ri shu chatga yuboring**!\n\n"
                    "*(Siz chatga yuborishingiz bilan bot uni .env fayliga avtomatik saqlab oladi)*"
                )
                await msg.reply_text(guide_msg, parse_mode="Markdown", disable_web_page_preview=True)
                return
            except Exception as w_err:
                logger.error(f"Whisper xatosi: {w_err}")
                await wait_msg.delete()
                await msg.reply_text(f"❌ Ovozni tanishda xatolik yuz berdi: {w_err}\nIltimos, qayta yuboring yoki matn sifatida yozing.")
                return

        original_voice_text = text_input
        norm_text = normalize_spoken_numbers(text_input)

        res = await parse_note_with_ai(text=norm_text, audio_b64=audio_b64)
        if not res:
            res = parse_with_smart_rules(norm_text)

        if not res:
            res = {
                "is_fastfood_smena": True,
                "sana": datetime.datetime.now().strftime("%Y-%m-%d"),
                "vaqt": datetime.datetime.now().strftime("%H:%M"),
                "hodimlar": [],
                "jami_ish_haqi": 0,
                "kassa_topshirildi_sdano": 0,
                "kartadan_yechildi": 0,
                "savdo_viruchka": 0,
                "karta_avans_va_xarajatlar": [],
                "izoh": text_input,
                "toza_matn": text_input
            }
        else:
            res["is_fastfood_smena"] = True

        newly_added_info = {
            "hodimlar": res.get("hodimlar", []),
            "sdano": res.get("kassa_topshirildi_sdano", 0),
            "kartadan": res.get("kartadan_yechildi", 0),
            "karta_items": res.get("karta_avans_va_xarajatlar", [])
        }
        has_new_content = bool(
            newly_added_info["hodimlar"] or
            newly_added_info["sdano"] or
            newly_added_info["kartadan"] or
            newly_added_info["karta_items"]
        )

        if not has_new_content:
            await wait_msg.delete()
            no_match_text = (
                f"🗣 **Eshitilgan xabar:** *\"{original_voice_text}\"*\n\n"
                "⚠️ Ushbu xabardan xodim ismi yoki to'lov summasi aniqlanmadi.\n\n"
                "👉 *Iltimos, xodim ismi va summasini aniq ayting:*\n"
                "Masalan: *\"Журабек 8500\"*, *\"Махмуд 5000\"*, *\"Сдано 10000\"*, *\"Снял с карты 17000\"*"
            )
            await msg.reply_text(no_match_text, parse_mode="Markdown")
            return

        # Smenani mavjud smenalar bilan aqlli birlashtiramiz (Merge)
        smenas = load_fastfood_smenas(user_id=user_id)
        s_date = res.get("sana") or datetime.datetime.now().strftime("%Y-%m-%d")
        res["sana"] = s_date
        
        existing_idx = None
        for idx, s in enumerate(smenas):
            if s.get("sana") == s_date:
                existing_idx = idx
                break

        if existing_idx is not None:
            existing = smenas[existing_idx]
            # Xodimlarni yangilash / qo'shish
            w_dict = {w["ism"]: w for w in existing.get("hodimlar", [])}
            for w in res.get("hodimlar", []):
                w_dict[w["ism"]] = w
            existing["hodimlar"] = list(w_dict.values())
            existing["jami_ish_haqi"] = sum(w.get("ish_haqi", 0) for w in existing["hodimlar"])

            if res.get("kassa_topshirildi_sdano", 0) > 0:
                existing["kassa_topshirildi_sdano"] = res["kassa_topshirildi_sdano"]
            if res.get("kartadan_yechildi", 0) > 0:
                existing["kartadan_yechildi"] = res["kartadan_yechildi"]
            if res.get("savdo_viruchka", 0) > 0:
                existing["savdo_viruchka"] = res["savdo_viruchka"]

            k_list = existing.get("karta_avans_va_xarajatlar", [])
            for k in res.get("karta_avans_va_xarajatlar", []):
                k_list.append(k)
            existing["karta_avans_va_xarajatlar"] = k_list
            res = existing
            smenas[existing_idx] = existing
        else:
            smenas.append(res)

        save_fastfood_smenas(smenas, user_id=user_id)

        try:
            from tabel_sync import sync_fastfood_tabel_file
            sync_fastfood_tabel_file(user_id=user_id)
        except Exception as ex_err:
            logger.error(f"Excel sinxronizatsiya xatosi: {ex_err}")

        await wait_msg.delete()

        reply_text = format_fastfood_shift_message(
            res,
            transcribed_text=original_voice_text,
            newly_added=newly_added_info
        )
        await msg.reply_text(
            reply_text,
            parse_mode="Markdown",
            reply_markup=get_action_keyboard(user_id)
        )
        return

    except Exception as e:
        logger.error(f"Xato: {e}")
        try:
            await wait_msg.edit_text(f"❌ Xato: {str(e)[:200]}")
        except:
            await msg.reply_text(f"❌ Xato: {str(e)[:200]}")

from aiohttp import web
import threading

async def run_server_webhook(app, port: int, external_url: str):
    webhook_url = f"{external_url}/telegram"
    logger.info(f"🌐 Setting Telegram Webhook to: {webhook_url}")
    try:
        await app.bot.set_webhook(url=webhook_url, drop_pending_updates=False)
        logger.info("✅ Telegram Webhook successfully set with Telegram API!")
    except Exception as w_err:
        logger.error(f"Failed to set webhook: {w_err}")

    async def telegram_webhook_handler(request):
        try:
            data = await request.json()
            update = Update.de_json(data, app.bot)
            asyncio.create_task(app.process_update(update))
            return web.Response(text="OK", status=200)
        except Exception as e:
            logger.error(f"Webhook update processing error: {e}")
            return web.Response(text="OK", status=200)

    async def health_check(request):
        return web.Response(text="FAST_FOOD_BOT_V2_DEPLOYED_2026_11_02")

    async def self_ping_task():
        await asyncio.sleep(15)
        ping_url = f"{external_url}/health"
        logger.info(f"🔄 Self Keep-Alive pinger target: {ping_url}")
        async with aiohttp.ClientSession() as session:
            while True:
                try:
                    async with session.get(ping_url, timeout=aiohttp.ClientTimeout(total=10)) as resp:
                        logger.info(f"🔄 Self-ping keep-alive status: {resp.status}")
                except Exception as p_err:
                    logger.debug(f"Self-ping notice: {p_err}")
                await asyncio.sleep(240)

    async def serve_index(request):
        index_path = os.path.join(os.path.dirname(__file__), "static", "index.html")
        if os.path.exists(index_path):
            with open(index_path, "r", encoding="utf-8") as f:
                return web.Response(text=f.read(), content_type="text/html")
        return web.Response(text="Web App UI yuklanmoqda...")

    async def api_get_data(request):
        user_id = request.query.get("user_id") or DEFAULT_USER_ID
        rows = get_all_rows(user_id=user_id)
        now = datetime.datetime.now()
        current_month_prefix = now.strftime("%Y-%m")
        stats = get_category_totals(target_month_prefix=current_month_prefix, user_id=user_id)
        reminders = load_debt_reminders(user_id=user_id)

        # Faol qarzdorlar ro'yxatini hisoblaymiz
        debt_persons = {}
        for r in rows:
            shaxs = r.get("shaxs") or "Noma'lum"
            if shaxs not in debt_persons:
                debt_persons[shaxs] = {
                    "shaxs": shaxs,
                    "telefon": r.get("telefon", ""),
                    "berilgan_sana": r.get("sana", ""),
                    "qaytarish_sanasi": r.get("qaytarish_sanasi", ""),
                    "total_berildi": 0,
                    "total_qaytdi": 0
                }
            turi = (r.get("turi") or "").lower()
            izoh = (r.get("izoh") or "").lower()
            nasiya_s = r.get("nasiya") or 0
            tushum_s = r.get("tushum") or 0

            if r.get("telefon") and not debt_persons[shaxs]["telefon"]:
                debt_persons[shaxs]["telefon"] = r.get("telefon")
            if r.get("qaytarish_sanasi") and not debt_persons[shaxs]["qaytarish_sanasi"]:
                debt_persons[shaxs]["qaytarish_sanasi"] = r.get("qaytarish_sanasi")

            if "qaytdi" in turi or "qarzini qaytardi" in izoh or "qarz qaytdi" in izoh or "nasiyani yopdi" in izoh:
                debt_persons[shaxs]["total_qaytdi"] += (nasiya_s if nasiya_s > 0 else tushum_s)
            elif "berildi" in turi or "qarz berdim" in izoh or "nasiyaga berdim" in izoh or nasiya_s > 0:
                debt_persons[shaxs]["total_berildi"] += nasiya_s

        active_debts = []
        today_str = now.strftime("%Y-%m-%d")
        for p, info in debt_persons.items():
            qoldiq = info["total_berildi"] - info["total_qaytdi"]
            if qoldiq > 0:
                is_overdue = bool(info["qaytarish_sanasi"] and info["qaytarish_sanasi"] < today_str)
                active_debts.append({
                    "shaxs": p,
                    "telefon": info["telefon"],
                    "qoldiq": qoldiq,
                    "berilgan_sana": info["berilgan_sana"],
                    "qaytarish_sanasi": info["qaytarish_sanasi"],
                    "muddati_otgan": is_overdue
                })

        total_tushum = stats.get("total_tushum", 0)
        total_xarajat = stats.get("total_xarajat", 0)
        total_nasiya_b = stats.get("total_nasiya", 0)
        sof_foyda = total_tushum - total_xarajat

        sum_qoldiq_debt = sum(d["qoldiq"] for d in active_debts)
        kassa_balansi = total_tushum - total_xarajat - sum_qoldiq_debt

        # Fast food xulosalari
        ff_smenas = load_fastfood_smenas(user_id=user_id)
        ff_dek_1 = calculate_fastfood_summary(user_id=user_id, month_prefix=current_month_prefix, start_day=1, end_day=10)
        ff_dek_2 = calculate_fastfood_summary(user_id=user_id, month_prefix=current_month_prefix, start_day=11, end_day=20)
        ff_dek_3 = calculate_fastfood_summary(user_id=user_id, month_prefix=current_month_prefix, start_day=21, end_day=31)
        ff_month = calculate_fastfood_summary(user_id=user_id, month_prefix=current_month_prefix, start_day=1, end_day=31)

        return web.json_response({
            "transactions": rows,
            "debts": active_debts,
            "summary": {
                "sof_foyda": sof_foyda,
                "nasiya_qoldiq": sum_qoldiq_debt,
                "kassa_balansi": kassa_balansi,
                "kategoriya_xarajatlar": stats.get("xarajat_by_cat", {})
            },
            "fastfood": {
                "smenas": ff_smenas,
                "dekada_1": ff_dek_1,
                "dekada_2": ff_dek_2,
                "dekada_3": ff_dek_3,
                "month": ff_month
            }
        })

    async def api_save_record(request):
        try:
            data = await request.json()
            user_id = request.query.get("user_id") or data.get("user_id") or DEFAULT_USER_ID
            update_or_add_record(data, user_id=user_id)
            return web.json_response({"ok": True})
        except Exception as e:
            return web.json_response({"ok": False, "error": str(e)}, status=500)

    async def api_delete_record(request):
        try:
            rec_id = int(request.match_info['id'])
            user_id = request.query.get("user_id") or DEFAULT_USER_ID
            delete_specific_row(rec_id, user_id=user_id)
            return web.json_response({"ok": True})
        except Exception as e:
            return web.json_response({"ok": False, "error": str(e)}, status=500)

    async def serve_download_excel(request):
        user_id = request.query.get("user_id") or DEFAULT_USER_ID
        paths = get_user_paths(user_id)
        if os.path.exists(paths["excel"]):
            return web.FileResponse(paths["excel"], headers={
                "Content-Disposition": 'attachment; filename="Mening_Moliya_Jadvalim.xlsx"'
            })
        return web.Response(text="Fayl topilmadi", status=404)

    web_app = web.Application()
    web_app.router.add_get('/', serve_index)
    web_app.router.add_get('/health', health_check)
    web_app.router.add_get('/api/data', api_get_data)
    web_app.router.add_post('/api/record', api_save_record)
    web_app.router.add_delete('/api/record/{id}', api_delete_record)
    web_app.router.add_get('/download-excel', serve_download_excel)
    web_app.router.add_post('/telegram', telegram_webhook_handler)

    runner = web.AppRunner(web_app)
    await runner.setup()
    site = web.TCPSite(runner, '0.0.0.0', port)
    await site.start()
    logger.info(f"🚀 Webhook & Health Check Server active on 0.0.0.0:{port}")
    asyncio.create_task(self_ping_task())

    await asyncio.Event().wait()

def main():
    if not BOT_TOKEN:
        print('CRITICAL: BOT_TOKEN is missing! Set BOT_TOKEN in Render Environment variables.', flush=True)
        import sys
        sys.exit(1)
    logger.info("🤖 Voice & Text to Styled Excel (.xlsx) Bot with Robust Webhook Support (480.000 so'm)...")
    app = (
        Application.builder()
        .token(BOT_TOKEN)
        .read_timeout(60)
        .write_timeout(60)
        .connect_timeout(60)
        .pool_timeout(60)
        .get_updates_read_timeout(60)
        .build()
    )

    app.add_handler(CommandHandler("start", start_command))
    app.add_handler(CommandHandler("tabel", send_fastfood_tabel))
    app.add_handler(CommandHandler("smena", send_fastfood_tabel))
    app.add_handler(CommandHandler("kunlik", send_daily_digest))
    app.add_handler(CommandHandler("digest", send_daily_digest))
    app.add_handler(CommandHandler("kunlik_hisobot", send_daily_digest))
    app.add_handler(CommandHandler("jadval", send_excel_file))
    app.add_handler(CommandHandler("hisobot", send_category_report))
    app.add_handler(CommandHandler("statistika", send_category_report))
    app.add_handler(CommandHandler("royxat", send_rows_list))
    app.add_handler(CommandHandler("yozuvlar", send_rows_list))
    app.add_handler(CommandHandler("ochirish", handle_delete_command))
    app.add_handler(CommandHandler("bekor", handle_delete_command))
    app.add_handler(CallbackQueryHandler(handle_callback))
    app.add_handler(MessageHandler(filters.VOICE | filters.AUDIO | (filters.TEXT & ~filters.COMMAND), handle_text_or_voice))

    is_render = os.environ.get("RENDER") or os.environ.get("PORT")
    port = int(os.environ.get("PORT", 8080))
    external_url = os.environ.get("RENDER_EXTERNAL_URL", "https://telegram-moliya-bot.onrender.com").strip()
    if not external_url.startswith("http"):
        external_url = "https://" + external_url

    if is_render:
        async def start_webhook():
            await app.initialize()
            await app.start()
            asyncio.create_task(check_debt_reminders_task(app))
            asyncio.create_task(daily_digest_scheduler_task(app))
            await run_server_webhook(app, port, external_url)

        asyncio.run(start_webhook())
    else:
        import urllib.request, json
        try:
            wh_res = json.loads(urllib.request.urlopen(f"https://api.telegram.org/bot{BOT_TOKEN}/getWebhookInfo").read())
            active_wh = wh_res.get("result", {}).get("url", "")
            if active_wh and "onrender.com" in active_wh:
                logger.info(f"🌐 Bot hozirda 24/7 Cloud Serverda ({active_wh}) faol ishlamoqda. Local polling shart emas!")
                return
        except Exception as wh_chk_err:
            logger.debug(f"Webhook check: {wh_chk_err}")

        logger.info("✅ Bot tayyor (Local Polling)!")
        async def on_startup(application):
            asyncio.create_task(check_debt_reminders_task(application))
            asyncio.create_task(daily_digest_scheduler_task(application))

        app.post_init = on_startup
        app.run_polling(drop_pending_updates=False)

if __name__ == "__main__":
    main()
