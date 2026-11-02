# ============================================================
#  Fast Food Xodimlar Tabeli — Avtomatlashtirilgan Excel Sinxronizatori
# ============================================================
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import json
import os
import calendar
import datetime

MONTH_NAMES_UZ = {
    "01": "YANVAR", "02": "FEVRAL", "03": "MART", "04": "APREL",
    "05": "MAY", "06": "IYUN", "07": "IYUL", "08": "AVGUST",
    "09": "SENTABR", "10": "OKTYABR", "11": "NOYABR", "12": "DEKABR"
}

WORKER_COLUMNS = [
    ("Nurbek", "Нурбек", 2, "B"),
    ("Jorabek", "Журабек", 3, "C"),
    ("Mahmud", "Мухмуд", 4, "D"),
    ("Nodira", "Нодира", 5, "E"),
    ("Robiya", "Робия", 6, "F"),
    ("Rohat", "Рохат", 7, "G"),
]

def sync_fastfood_tabel_file(user_id=5472519783):
    """
    fastfood_smena.json dagi barcha ma'lumotlarni oylar bo'yicha ajratib,
    har bir oy uchun alohida chiroyli varaq (Sheet) yaratadi.
    Windows va Linux (Render) serverida cross-platform ishlaydi.
    """
    base_dir = os.path.dirname(os.path.abspath(__file__))
    user_id_str = str(user_id or 5472519783)
    user_dir = os.path.join(base_dir, "user_data", user_id_str)
    os.makedirs(user_dir, exist_ok=True)

    json_path = os.path.join(user_dir, "fastfood_smena.json")
    if not os.path.exists(json_path):
        for alt_uid in ["5472519783", "1453337970", "616405914"]:
            cand = os.path.join(base_dir, "user_data", alt_uid, "fastfood_smena.json")
            if os.path.exists(cand):
                import shutil
                shutil.copy2(cand, json_path)
                break

    smenas = []
    if os.path.exists(json_path):
        try:
            with open(json_path, "r", encoding="utf-8") as f:
                smenas = json.load(f)
        except Exception as e:
            print(f"Error reading {json_path}: {e}")

    # Oylar bo'yicha guruhlaymiz
    months_set = set()
    for s in smenas:
        sana = s.get("sana", "")
        if len(sana) >= 7 and "-" in sana:
            months_set.add(sana[:7])

    # Hozirgi oy va 2026-10 (Oktyabr) doim bo'lishini ta'minlaymiz
    tz_tashkent = datetime.timezone(datetime.timedelta(hours=5))
    now = datetime.datetime.now(tz_tashkent)
    cur_ym = now.strftime("%Y-%m")
    months_set.add("2026-10")
    months_set.add(cur_ym)

    sorted_months = sorted(list(months_set))

    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # Standart bo'sh varaqni o'chiramiz

    header_fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
    header_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
    header_border = Border(
        left=Side(style="thin", color="334155"), right=Side(style="thin", color="334155"),
        top=Side(style="medium", color="0F172A"), bottom=Side(style="medium", color="0F172A")
    )
    thin_border = Border(
        left=Side(style="thin", color="E2E8F0"), right=Side(style="thin", color="E2E8F0"),
        top=Side(style="thin", color="E2E8F0"), bottom=Side(style="thin", color="E2E8F0")
    )
    tot_fill = PatternFill(start_color="FEF08A", end_color="FEF08A", fill_type="solid")
    tot_border = Border(
        left=Side(style="thin", color="CA8A04"), right=Side(style="thin", color="CA8A04"),
        top=Side(style="medium", color="CA8A04"), bottom=Side(style="double", color="CA8A04")
    )
    tot_font = Font(name="Segoe UI", size=11, bold=True, color="713F12")

    headers = [
        "Sana", "Нурбек", "Журабек", "Мухмуд", "Нодира", "Робия", "Рохат",
        "Сдано (Naqd)", "Снял с карты", "З/П (Jami)", "Savdo (Viruchka)",
        "КАРТА + Izohi", "КАРТА + Summa"
    ]

    col_widths = {
        "A": 14, "B": 12, "C": 12, "D": 12, "E": 12, "F": 12, "G": 12,
        "H": 15, "I": 15, "J": 16, "K": 16, "L": 22, "M": 16
    }

    for month_prefix in sorted_months:
        parts = month_prefix.split("-")
        year_val = int(parts[0])
        month_val = int(parts[1])
        month_str = f"{month_val:02d}"
        month_name_uz = MONTH_NAMES_UZ.get(month_str, month_str)

        _, num_days = calendar.monthrange(year_val, month_val)

        ws = wb.create_sheet(title=f"Tabel {month_str}.{year_val}")

        # 1. Sarlavha Banner
        ws.merge_cells("A1:M1")
        title_cell = ws["A1"]
        title_cell.value = f"FAST FOOD XODIMLAR TABELI VA KASSA HISOBOTI — {month_name_uz} {year_val}"
        title_cell.font = Font(name="Segoe UI", size=14, bold=True, color="FFFFFF")
        title_cell.fill = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
        title_cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 40

        # 2. Quyi sarlavha
        ws.merge_cells("A2:M2")
        sub_cell = ws["A2"]
        sub_cell.value = f"Davr: 01.{month_str}.{year_val} — {num_days:02d}.{month_str}.{year_val} | Ishchilar oyligi, Kassa (Сдано), Kartadan yechilgan va Карта+ xarajatlari"
        sub_cell.font = Font(name="Segoe UI", size=10, italic=True, color="64748B")
        sub_cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[2].height = 22

        # 3. Ustun nomlari
        ws.row_dimensions[3].height = 28
        for col_idx, h in enumerate(headers, 1):
            c = ws.cell(row=3, column=col_idx, value=h)
            c.fill = header_fill
            c.font = header_font
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c.border = header_border

        # Oy kunlari ma'lumotlarini tayyorlash
        days_data = {}
        for d in range(1, num_days + 1):
            days_data[d] = {
                "Nurbek": None, "Jorabek": None, "Mahmud": None,
                "Nodira": None, "Robiya": None, "Rohat": None,
                "sdano": None, "snyal": None, "savdo": None,
                "karta_desc": [], "karta_sum": 0
            }

        # Shu oyga tegishli smenalarni to'ldirish
        for s in smenas:
            sana = s.get("sana", "")
            if not sana.startswith(month_prefix):
                continue
            try:
                day = int(sana.split("-")[2])
            except:
                continue

            if day not in days_data:
                continue

            target = days_data[day]
            for w in s.get("hodimlar", []):
                name = (w.get("ism") or "").strip().lower()
                haq = int(w.get("ish_haqi", 0))
                if "нурб" in name or "nurb" in name or "nurev" in name:
                    target["Nurbek"] = (target["Nurbek"] or 0) + haq
                elif "жураб" in name or "jurab" in name or "jorab" in name or "corab" in name:
                    target["Jorabek"] = (target["Jorabek"] or 0) + haq
                elif "махмуд" in name or "мухмуд" in name or "mahmud" in name or "muhmad" in name:
                    target["Mahmud"] = (target["Mahmud"] or 0) + haq
                elif "нодир" in name or "nodir" in name or "noder" in name:
                    target["Nodira"] = (target["Nodira"] or 0) + haq
                elif "роби" in name or "robi" in name or "rabi" in name or "gobi" in name:
                    target["Robiya"] = (target["Robiya"] or 0) + haq
                elif "роха" in name or "roha" in name or "raha" in name:
                    target["Rohat"] = (target["Rohat"] or 0) + haq

            if s.get("kassa_topshirildi_sdano"):
                target["sdano"] = (target["sdano"] or 0) + int(s["kassa_topshirildi_sdano"])
            if s.get("kartadan_yechildi"):
                target["snyal"] = (target["snyal"] or 0) + int(s["kartadan_yechildi"])
            if s.get("savdo_viruchka"):
                target["savdo"] = (target["savdo"] or 0) + int(s["savdo_viruchka"])

            for c_item in s.get("karta_avans_va_xarajatlar", []):
                c_shaxs = c_item.get("shaxs", "")
                c_sum = int(c_item.get("summa", 0))
                c_izoh = c_item.get("izoh", "")
                if c_sum > 0:
                    desc_text = c_shaxs or c_izoh or "Xarajat"
                    target["karta_desc"].append(desc_text)
                    target["karta_sum"] += c_sum

        # Kunlarni jadvalga yozish
        start_row = 4
        for day in range(1, num_days + 1):
            row_num = start_row + day - 1
            d_val = days_data[day]
            ws.row_dimensions[row_num].height = 20
            bg_color = "F8FAFC" if day % 2 == 1 else "FFFFFF"
            row_fill = PatternFill(start_color=bg_color, end_color=bg_color, fill_type="solid")

            # Col 1: Sana
            c = ws.cell(row=row_num, column=1, value=f"{day:02d}.{month_str}.{year_val}")
            c.fill = row_fill
            c.border = thin_border
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.font = Font(name="Segoe UI", size=10, bold=True, color="334155")

            # Cols 2-7: Workers
            w_keys = ["Nurbek", "Jorabek", "Mahmud", "Nodira", "Robiya", "Rohat"]
            has_any_worker = False
            for c_offset, w_k in enumerate(w_keys, 2):
                val = d_val[w_k]
                cell = ws.cell(row=row_num, column=c_offset)
                cell.fill = row_fill
                cell.border = thin_border
                cell.font = Font(name="Segoe UI", size=10)
                cell.alignment = Alignment(horizontal="right", vertical="center")
                if val is not None and val > 0:
                    has_any_worker = True
                    cell.value = val
                    cell.number_format = "#,##0"
                else:
                    cell.value = "-"

            # Col 8: Sdano
            c_sdano = ws.cell(row=row_num, column=8)
            c_sdano.fill = row_fill
            c_sdano.border = thin_border
            c_sdano.alignment = Alignment(horizontal="right", vertical="center")
            c_sdano.font = Font(name="Segoe UI", size=10)
            if d_val["sdano"] is not None and d_val["sdano"] > 0:
                c_sdano.value = d_val["sdano"]
                c_sdano.number_format = "#,##0"
            else:
                c_sdano.value = "-"

            # Col 9: Snyal
            c_snyal = ws.cell(row=row_num, column=9)
            c_snyal.fill = row_fill
            c_snyal.border = thin_border
            c_snyal.alignment = Alignment(horizontal="right", vertical="center")
            c_snyal.font = Font(name="Segoe UI", size=10)
            if d_val["snyal"] is not None and d_val["snyal"] > 0:
                c_snyal.value = d_val["snyal"]
                c_snyal.number_format = "#,##0"
            else:
                c_snyal.value = "-"

            # Col 10: Z/P Formula
            c_zp = ws.cell(row=row_num, column=10)
            c_zp.fill = row_fill
            c_zp.border = thin_border
            c_zp.alignment = Alignment(horizontal="right", vertical="center")
            c_zp.font = Font(name="Segoe UI", size=10, bold=True, color="0F172A")
            if has_any_worker:
                c_zp.value = f"=SUM(B{row_num}:G{row_num})"
                c_zp.number_format = "#,##0"
            else:
                c_zp.value = "-"

            # Col 11: Savdo (Viruchka)
            c_savdo = ws.cell(row=row_num, column=11)
            c_savdo.fill = row_fill
            c_savdo.border = thin_border
            c_savdo.alignment = Alignment(horizontal="right", vertical="center")
            c_savdo.font = Font(name="Segoe UI", size=10)
            if d_val["savdo"] is not None and d_val["savdo"] > 0:
                c_savdo.value = d_val["savdo"]
                c_savdo.number_format = "#,##0"
            else:
                c_savdo.value = "-"

            # Col 12: Karta+ desc
            c_kdesc = ws.cell(row=row_num, column=12)
            c_kdesc.fill = row_fill
            c_kdesc.border = thin_border
            c_kdesc.alignment = Alignment(horizontal="left", vertical="center")
            if d_val["karta_desc"]:
                c_kdesc.value = ", ".join(d_val["karta_desc"])
                c_kdesc.font = Font(name="Segoe UI", size=10, bold=True, color="B91C1C")
            else:
                c_kdesc.value = ""

            # Col 13: Karta+ sum
            c_ksum = ws.cell(row=row_num, column=13)
            c_ksum.fill = row_fill
            c_ksum.border = thin_border
            c_ksum.alignment = Alignment(horizontal="right", vertical="center")
            if d_val["karta_sum"] > 0:
                c_ksum.value = d_val["karta_sum"]
                c_ksum.font = Font(name="Segoe UI", size=10, bold=True, color="B91C1C")
                c_ksum.number_format = "#,##0"
            else:
                c_ksum.value = "-"

        # Bottom Total Row (JAMI)
        total_row = start_row + num_days
        ws.row_dimensions[total_row].height = 26

        lbl = ws.cell(row=total_row, column=1, value="JAMI (ИТОГО):")
        lbl.font = tot_font
        lbl.fill = tot_fill
        lbl.alignment = Alignment(horizontal="center", vertical="center")
        lbl.border = tot_border

        last_day_row = start_row + num_days - 1
        col_letters = ["B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M"]
        for idx, cl in enumerate(col_letters, 2):
            cell = ws.cell(row=total_row, column=idx)
            cell.fill = tot_fill
            cell.font = tot_font
            cell.border = tot_border
            cell.alignment = Alignment(horizontal="right", vertical="center")
            if cl in ("B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "M"):
                cell.value = f"=SUM({cl}{start_row}:{cl}{last_day_row})"
                cell.number_format = "#,##0"
            else:
                cell.value = ""

        # Ustun kengliklari
        for col_l, width in col_widths.items():
            ws.column_dimensions[col_l].width = width

    # Hozirgi oy varag'ini faol (Active) qilib belgilaymiz, shunda fayl ochilganda aynan hozirgi oy avtomatik ochiladi
    cur_sheet_title = f"Tabel {now.month:02d}.{now.year}"
    target_ws = wb[cur_sheet_title] if cur_sheet_title in wb.sheetnames else wb.worksheets[-1]
    wb.active = target_ws
    for s in wb.worksheets:
        s.sheet_view.tabSelected = (s == target_ws)

    # Saqlash joylari (Cross-platform)
    cur_month_str = f"{now.month:02d}"
    cur_month_uz = MONTH_NAMES_UZ.get(cur_month_str, "NOYABR").capitalize()

    targets = [
        os.path.join(user_dir, "Fast_Food_Tabel_2026.xlsx"),
        os.path.join(base_dir, "Fast_Food_Tabel_2026.xlsx"),
        os.path.join(user_dir, f"Fast_Food_Tabel_{cur_month_uz}_{now.year}.xlsx"),
        os.path.join(base_dir, f"Fast_Food_Tabel_{cur_month_uz}_{now.year}.xlsx"),
        os.path.join(user_dir, "Fast_Food_Tabel_Oktyabr_2026.xlsx"),
        os.path.join(base_dir, "Fast_Food_Tabel_Oktyabr_2026.xlsx"),
    ]

    # Agar lokal Windows kompyuterda bo'lsa
    if os.path.exists(r"d:\Antigravity\123"):
        targets.extend([
            r"d:\Antigravity\123\Fast_Food_Tabel_2026.xlsx",
            r"d:\Antigravity\123\Fast_Food_Tabel_Oktyabr_2026.xlsx",
            rf"d:\Antigravity\123\Fast_Food_Tabel_{cur_month_uz}_{now.year}.xlsx",
        ])

    for t in targets:
        try:
            os.makedirs(os.path.dirname(t), exist_ok=True)
            wb.save(t)
        except Exception as save_err:
            pass

    print("Fast Food Excel Tabel successfully synchronized!")

if __name__ == "__main__":
    sync_fastfood_tabel_file(5472519783)
