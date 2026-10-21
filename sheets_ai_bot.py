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

DEBT_REMINDERS_FILE = "qarz_eslatmalari.json"

def load_debt_reminders() -> list:
    if os.path.exists(DEBT_REMINDERS_FILE):
        try:
            with open(DEBT_REMINDERS_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except:
            return []
    return []

def save_debt_reminders(reminders: list):
    with open(DEBT_REMINDERS_FILE, "w", encoding="utf-8") as f:
        json.dump(reminders, f, ensure_ascii=False, indent=2)

def calculate_person_debt_balance(person_name: str) -> dict:
    total_berildi = 0
    total_qaytdi = 0
    p_clean = (person_name or "").strip().lower()

    if os.path.exists(CSV_FILE_PATH):
        with open(CSV_FILE_PATH, mode='r', encoding='utf-8-sig') as f:
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

                # Agar shaxs ismi mos kelsa yoki izohda ismi qatnashgan bo'lsa
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
    """Whisper tushunarsiz shovqin so'zlarini matndan tozalash"""
    if not text:
        return text
    noise_patterns = [
        r'\bget\b', r'\bover\b', r'\bzikim\b', r'\buh\b', r'\bumm\b',
        r'\bsubtitles\b', r'\bby amara\.org\b', r'\bthank you\b', r'\bthanks\b'
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
                "shaxs": removed[3] if len(removed) > 3 else "",
                "izoh": removed[10] if len(removed) >= 11 else (removed[8] if len(removed) > 8 else "")
            }

            with open(CSV_FILE_PATH, mode='w', newline='', encoding='utf-8-sig') as f:
                writer = csv.writer(f, delimiter=';')
                writer.writerow(header)
                writer.writerows(data_rows)

    if os.path.exists(EXCEL_FILE_PATH):
        wb = openpyxl.load_workbook(EXCEL_FILE_PATH)
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            excel_row = data_row_id + 1
            if 2 <= excel_row <= ws.max_row:
                ws.delete_rows(excel_row)
            render_in_sheet_summary(ws)

        wb.save(EXCEL_FILE_PATH)

    return deleted_info

def update_or_add_record(record: dict):
    """Web App orqali yozuvni qo'shish yoki tahrirlash"""
    rec_id = record.get("id")

    with open(CSV_FILE_PATH, mode='r', encoding='utf-8-sig') as f:
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
    with open(CSV_FILE_PATH, mode='w', newline='', encoding='utf-8-sig') as f:
        writer = csv.writer(f, delimiter=';')
        writer.writerow(header)
        writer.writerows(data_rows)

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

    wb.save(EXCEL_FILE_PATH)
    return True

def delete_last_entry():
    rows = get_all_rows()
    if rows:
        return delete_specific_row(len(rows))
    return None

def get_category_totals(target_month_prefix: str = None):
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

GEMINI_URL = "https://generativelanguage.googleapis.com/v1beta/models/gemini-3.5-flash:generateContent"
GROQ_URL = "https://api.groq.com/openai/v1/chat/completions"
GROQ_WHISPER_URL = "https://api.groq.com/openai/v1/audio/transcriptions"

SYSTEM_PROMPT = """Siz O'zbek tili nutqi, moliya va buxgalteriya iboralari bo'yicha mutaxassis AI yordamchisiz.
Foydalanuvchi ovozli yoki matnli xabarda nima degan bo'lsa (lotin yoki kirill alifbosida), uni tushunib, quyidagi JSON formatiga o'giring.

MOLIYAVIY OPERATSIYA TURLARI (MUHIM QOIDALAR):
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
  - Boshqa hech qanday so'zni, summani yoki sanani telefon raqamiga ARALASHTIRMANG! 9 ta raqamdan ortiq bo'lmasin!
  - Agar telefon aytilmagan bo'lsa -> "" (bo'sh qolsin).
- Shaxs ismini aniq yozing (Masalan: "Doniyorga" desa, ism "Doniyor" bo'lsin, "Don Yorga" emas!).
- "izoh" maydoniga foydalanuvchi aytgan har qanday qo'shimcha tafsilotlar, eslatmalar yoki shartlarni to'liq kiriting.

O'ZBEKCHA MURAKKAB SONLAR (To'liq hisoblang!):
- "to'rt yuz sakson ming" / "480k" = 480000 so'm
- "uch yuz ellik ming" / "350 ming" = 350000 so'm
- "ikki yuz ming" / "200 ming" = 200000 so'm
- "bir yarim million" = 1500000 so'm
- "bir million" = 1000000 so'm
- "sakson ming" = 80000 so'm
- "yigirma besh ming" = 25000 so'm

Javobingiz FAQAT toza JSON bo'lishi shart:
{
  "turi": "Tushum" yoki "Xarajat" yoki "Nasiya berildi" yoki "Nasiya qaytdi",
  "shaxs": "Mijoz yoki shaxs ismi (bo'lmasa 'Noma'lum')",
  "telefon": "+998XXXXXXXXX yoki bo'sh",
  "tushum_som": 0,
  "xarajat_som": 0,
  "nasiya_som": 0,
  "kategoriya": "Savdo/Qarz munosabatlari/Oziq-ovqat/Transport/Kommunal/Boshqa",
  "izoh": "Tafsilotli to'liq izoh",
  "toza_matn": "Shovqinsiz chiroyli toza o'zbekcha jumla",
  "qaytarish_sanasi": "YYYY-MM-DD yoki null"
}
* Eslatma:
- "Nasiya berildi" bo'lsa -> nasiya_som ga summani yozing, tushum va xarajat 0 bo'lsin.
- "Nasiya qaytdi" bo'lsa -> nasiya_som ga summani yozing, tushum va xarajat 0 bo'lsin."""

async def transcribe_audio_with_groq(audio_bytes: bytes) -> str:
    headers = {"Authorization": f"Bearer {GROQ_API_KEY}"}
    form = aiohttp.FormData()
    form.add_field('file', audio_bytes, filename='voice.ogg', content_type='audio/ogg')
    form.add_field('model', 'whisper-large-v3-turbo')
    form.add_field('language', 'uz')
    form.add_field('temperature', '0.0')
    form.add_field('prompt', "O'zbek tilidagi moliya, tushum va xarajatlar: to'rt yuz sakson ming so'm to'lov qildim, 480 ming so'm, uch yuz ellik ming, sakson ming so'm, benzinga yigirma besh ming, tushlikka 45 ming, Alisherdan 500 ming tushum oldim, nasiya qarz 1 million so'm.")

    async with aiohttp.ClientSession() as session:
        async with session.post(GROQ_WHISPER_URL, data=form, headers=headers, timeout=aiohttp.ClientTimeout(total=45)) as resp:
            if resp.status != 200:
                err_text = await resp.text()
                raise Exception(f"Whisper error ({resp.status}): {err_text[:150]}")
            result = await resp.json()
            text = result.get("text", "").strip()
            return fix_uzbek_speech_phonetics(text)

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
        raise Exception(f"AI javob bera olmadi. ({last_error[:100]})")

    raw_json = re.sub(r'^```json\s*', '', raw_json)
    raw_json = re.sub(r'\s*```$', '', raw_json)
    parsed = json.loads(raw_json)

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
    tushum_val = int(row_data[4] or 0)
    xarajat_val = int(row_data[5] or 0)
    nasiya_val = int(row_data[6] or 0)

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

def append_to_excel_and_csv(data: dict):
    with open(CSV_FILE_PATH, mode='a', newline='', encoding='utf-8-sig') as f:
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

    if not os.path.exists(EXCEL_FILE_PATH):
        wb = create_styled_excel()
    else:
        wb = openpyxl.load_workbook(EXCEL_FILE_PATH)

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

    wb.save(EXCEL_FILE_PATH)

WEB_APP_URL = os.environ.get("WEB_APP_URL", "https://moliya-webapp.vercel.app").strip()
if not WEB_APP_URL.startswith("http"):
    WEB_APP_URL = "https://" + WEB_APP_URL

from telegram import WebAppInfo

def get_action_keyboard():
    keyboard = [
        [
            InlineKeyboardButton("📱 Moliya & Qarz Web Ilovasi", web_app=WebAppInfo(url=WEB_APP_URL)),
        ],
        [
            InlineKeyboardButton("📋 Yozuvlar Ro'yxati", callback_data="show_rows_list"),
            InlineKeyboardButton("📈 Kategoriyalar Hisoboti", callback_data="category_report"),
        ],
        [
            InlineKeyboardButton("📊 Excel Faylini Yuklash (.xlsx)", callback_data="download_excel"),
            InlineKeyboardButton("🌐 Web Interfeys (Browser)", url=WEB_APP_URL)
        ]
    ]
    return InlineKeyboardMarkup(keyboard)

def get_reply_menu():
    keyboard = [
        [KeyboardButton("📱 Moliya & Qarz Ilovasi", web_app=WebAppInfo(url=WEB_APP_URL))],
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
        "👉 *\"Dastur uchun to'rt yuz sakson ming so'm to'lov qildim\"*\n"
        "👉 *\"Bugun Shoxruxga 150000 so'mga atir sotdim\"*\n\n"
        "🎙 **ULTRA-ANIQ OVOZ VA MURAKKAB SONLAR TAHLILI:**\n"
        "Bot O'zbek tili va murakkab sonlarni (masalan: *480 000 so'm / to'rt yuz sakson ming*) ultra-yuqori aniqlikda tushunadi va to'g'ri summalarga o'girib saqlaydi! 📁✨\n\n"
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
                caption="📊 **Sizning Moliya Excel Jadvalingiz**\n\n*(Har bir oylik sahifada FAQAT shu oyning Kategoriya Hisoboti alohida ko'rsatilgan)*",
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
            audio_bytes = None
            for attempt in range(3):
                try:
                    file_obj = await (msg.voice or msg.audio).get_file(read_timeout=60, write_timeout=60, connect_timeout=60)
                    audio_bytes = await file_obj.download_as_bytearray(read_timeout=60, write_timeout=60, connect_timeout=60)
                    audio_b64 = base64.b64encode(audio_bytes).decode('utf-8')
                    break
                except Exception as dl_err:
                    logger.warning(f"Audio download attempt {attempt+1} failed: {dl_err}")
                    if attempt == 2:
                        raise dl_err
                    await asyncio.sleep(1)

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

        # Agar qarz berilgan bo'lsa va qaytarish sanasi ko'rsatilgan bo'lsa, eslatmalar ro'yxatiga saqlaymiz
        qaytarish_sanasi = res.get("qaytarish_sanasi")
        if qaytarish_sanasi and ("nasiya" in str(res.get("turi", "")).lower() or "berildi" in str(res.get("turi", "")).lower()):
            reminders = load_debt_reminders()
            reminders.append({
                "chat_id": msg.chat.id,
                "shaxs": res.get("shaxs", "Noma'lum"),
                "telefon": res.get("telefon", ""),
                "boshlangich_qarz": int(res.get("nasiya_som", 0)),
                "berilgan_sana": res.get("sana"),
                "qaytarish_sanasi": qaytarish_sanasi,
                "izoh": res.get("izoh"),
                "ogohlantirildi": False
            })
            save_debt_reminders(reminders)
            logger.info(f"⏰ Qarz eslatmasi saqlandi: {res.get('shaxs')} -> {qaytarish_sanasi}")

        await wait_msg.delete()

        month_label = get_month_sheet_title(res.get('sana', ''))

        display_text = res.get("toza_matn") or clean_stt_noise(text_input)

        reply_text = (
            f"✅ **Excel jadvalingizning '{month_label}' varag'iga saqlandi!** 📊\n"
            f"*(Ushbu oyning Kategoriya Hisoboti o'ng tomonga alohida joylandi)*\n\n"
            f"🗣 **Eshitilgan xabar:** *\"{display_text}\"*\n\n"
            f"📅 **Sana:** `{res.get('sana')}` ({res.get('vaqt')})\n"
            f"📌 **Turi:** {res.get('turi')}\n"
            f"👤 **Shaxs/Mijoz:** {res.get('shaxs')}\n"
        )

        if res.get("telefon"):
            reply_text += f"📞 **Telefon:** `{res.get('telefon')}`\n"

        reply_text += (
            f"💰 **Tushum:** {int(res.get('tushum_som', 0)):,} so'm\n"
            f"💸 **Xarajat:** {int(res.get('xarajat_som', 0)):,} so'm\n"
            f"⏳ **Nasiya:** {int(res.get('nasiya_som', 0)):,} so'm\n"
        )

        if qaytarish_sanasi:
            reply_text += f"⏰ **Qaytarish muddati:** `{qaytarish_sanasi}` (Eslatib turiladi!)\n"

        reply_text += (
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

from aiohttp import web
import threading

async def run_server_webhook(app, port: int, external_url: str):
    webhook_url = f"{external_url}/telegram"
    logger.info(f"🌐 Setting Telegram Webhook to: {webhook_url}")
    try:
        await app.bot.set_webhook(url=webhook_url, drop_pending_updates=True)
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
        return web.Response(text="✅ Telegram Bot is running 24/7!")

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
        rows = get_all_rows()
        now = datetime.datetime.now()
        current_month_prefix = now.strftime("%Y-%m")
        stats = get_category_totals(target_month_prefix=current_month_prefix)
        reminders = load_debt_reminders()

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

        return web.json_response({
            "transactions": rows,
            "debts": active_debts,
            "summary": {
                "sof_foyda": sof_foyda,
                "nasiya_qoldiq": sum_qoldiq_debt,
                "kassa_balansi": kassa_balansi,
                "kategoriya_xarajatlar": stats.get("xarajat_by_cat", {})
            }
        })

    async def api_save_record(request):
        try:
            data = await request.json()
            update_or_add_record(data)
            return web.json_response({"ok": True})
        except Exception as e:
            return web.json_response({"ok": False, "error": str(e)}, status=500)

    async def api_delete_record(request):
        try:
            rec_id = int(request.match_info['id'])
            delete_specific_row(rec_id)
            return web.json_response({"ok": True})
        except Exception as e:
            return web.json_response({"ok": False, "error": str(e)}, status=500)

    async def serve_download_excel(request):
        if os.path.exists(EXCEL_FILE_PATH):
            return web.FileResponse(EXCEL_FILE_PATH, headers={
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
            await run_server_webhook(app, port, external_url)

        asyncio.run(start_webhook())
    else:
        logger.info("✅ Bot tayyor (Local Polling)!")
        async def on_startup(application):
            asyncio.create_task(check_debt_reminders_task(application))

        app.post_init = on_startup
        app.run_polling(drop_pending_updates=True)

if __name__ == "__main__":
    main()
