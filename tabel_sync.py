import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import json
import os

def sync_fastfood_tabel_file(user_id=1453337970):
    json_path = rf"d:\Antigravity\123\telegrambot\user_data\{user_id}\fastfood_smena.json"
    if not os.path.exists(json_path):
        json_path = r"d:\Antigravity\123\telegrambot\user_data\999888777\fastfood_smena.json"
    
    if not os.path.exists(json_path):
        return

    with open(json_path, "r", encoding="utf-8") as f:
        smenas = json.load(f)

    days_data = {}
    for d in range(1, 32):
        days_data[d] = {
            "Nurbek": None,
            "Jorabek": None,
            "Mahmud": None,
            "Nodira": None,
            "Robiya": None,
            "Rohat": None,
            "sdano": None,
            "snyal": None,
            "savdo": None,
            "karta_desc": [],
            "karta_sum": 0
        }

    for s in smenas:
        sana = s.get("sana", "")
        if not sana.startswith("2026-10"):
            continue
        try:
            day = int(sana.split("-")[2])
        except:
            continue
        
        target = days_data[day]
        for w in s.get("hodimlar", []):
            name = (w.get("ism") or "").strip().lower()
            haq = int(w.get("ish_haqi", 0))
            if "нурбек" in name or "nurbek" in name:
                target["Nurbek"] = (target["Nurbek"] or 0) + haq
            elif "журабек" in name or "jorabek" in name or "jo'rabek" in name:
                target["Jorabek"] = (target["Jorabek"] or 0) + haq
            elif "мухмуд" in name or "мухаммаджон" in name or "mahmud" in name:
                target["Mahmud"] = (target["Mahmud"] or 0) + haq
            elif "нодира" in name or "nodira" in name:
                target["Nodira"] = (target["Nodira"] or 0) + haq
            elif "робия" in name or "robiya" in name:
                target["Robiya"] = (target["Robiya"] or 0) + haq
            elif "рохат" in name or "rohat" in name:
                target["Rohat"] = (target["Rohat"] or 0) + haq

        if s.get("kassa_topshirildi_sdano"):
            target["sdano"] = (target["sdano"] or 0) + int(s["kassa_topshirildi_sdano"])
        if s.get("kartadan_yechildi"):
            target["snyal"] = (target["snyal"] or 0) + int(s["kartadan_yechildi"])
        if s.get("savdo_viruchka"):
            target["savdo"] = (target["savdo"] or 0) + int(s["savdo_viruchka"])

        for c in s.get("karta_avans_va_xarajatlar", []):
            c_shaxs = c.get("shaxs", "")
            c_sum = int(c.get("summa", 0))
            c_izoh = c.get("izoh", "")
            if c_sum > 0:
                target["karta_desc"].append(c_shaxs or c_izoh)
                target["karta_sum"] += c_sum

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Tabel 10.2026"

    # Title Banner
    ws.merge_cells("A1:M1")
    title_cell = ws["A1"]
    title_cell.value = "FAST FOOD XODIMLAR TABELI VA KASSA HISOBOTI — OKTYABR 2026"
    title_cell.font = Font(name="Segoe UI", size=14, bold=True, color="FFFFFF")
    title_cell.fill = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 40

    ws.merge_cells("A2:M2")
    sub_cell = ws["A2"]
    sub_cell.value = "📅 Davr: 01.10.2026 — 31.10.2026 | Ishchilar oyligi, Kassa (Sdano), Kartadan yechilgan va Karta+ xarajatlari"
    sub_cell.font = Font(name="Segoe UI", size=10, italic=True, color="64748B")
    sub_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 22

    headers = [
        "Sana", "Нурбек", "Журабек", "Мухмуд", "Нодира", "Робия", "Рохат",
        "Сдано (Naqd)", "Снял с карты", "З/П (Jami Oylik)", "Savdo (Viruchka)",
        "Karta + Izohi", "Karta + Summa"
    ]
    header_fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
    header_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
    header_border = Border(
        left=Side(style="thin", color="334155"), right=Side(style="thin", color="334155"),
        top=Side(style="medium", color="0F172A"), bottom=Side(style="medium", color="0F172A")
    )
    ws.row_dimensions[3].height = 28
    for col_idx, h in enumerate(headers, 1):
        c = ws.cell(row=3, column=col_idx, value=h)
        c.fill = header_fill
        c.font = header_font
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = header_border

    thin_border = Border(
        left=Side(style="thin", color="E2E8F0"), right=Side(style="thin", color="E2E8F0"),
        top=Side(style="thin", color="E2E8F0"), bottom=Side(style="thin", color="E2E8F0")
    )

    start_row = 4
    for day in range(1, 32):
        row_num = start_row + day - 1
        d_val = days_data[day]
        ws.row_dimensions[row_num].height = 20
        bg_color = "F8FAFC" if day % 2 == 1 else "FFFFFF"
        row_fill = PatternFill(start_color=bg_color, end_color=bg_color, fill_type="solid")

        # Col 1: Sana
        c = ws.cell(row=row_num, column=1, value=f"{day:02d}.10.2026")
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

        # Col 11: Savdo
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

    # Bottom Total Row (Row 35)
    total_row = start_row + 31
    ws.row_dimensions[total_row].height = 26
    tot_fill = PatternFill(start_color="FEF08A", end_color="FEF08A", fill_type="solid")
    tot_border = Border(
        left=Side(style="thin", color="CA8A04"), right=Side(style="thin", color="CA8A04"),
        top=Side(style="medium", color="CA8A04"), bottom=Side(style="double", color="CA8A04")
    )
    tot_font = Font(name="Segoe UI", size=11, bold=True, color="713F12")

    lbl = ws.cell(row=total_row, column=1, value="JAMI (ИТОГО):")
    lbl.font = tot_font
    lbl.fill = tot_fill
    lbl.alignment = Alignment(horizontal="center", vertical="center")
    lbl.border = tot_border

    col_letters = ["B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M"]
    for idx, cl in enumerate(col_letters, 2):
        cell = ws.cell(row=total_row, column=idx)
        cell.fill = tot_fill
        cell.font = tot_font
        cell.border = tot_border
        cell.alignment = Alignment(horizontal="right", vertical="center")
        if cl in ("B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "M"):
            cell.value = f"=SUM({cl}4:{cl}34)"
            cell.number_format = "#,##0"
        else:
            cell.value = ""

    col_widths = {
        "A": 14, "B": 12, "C": 12, "D": 12, "E": 12, "F": 12, "G": 12,
        "H": 14, "I": 14, "J": 16, "K": 16, "L": 20, "M": 15
    }
    for col_l, width in col_widths.items():
        ws.column_dimensions[col_l].width = width

    targets = [
        rf"d:\Antigravity\123\telegrambot\user_data\{user_id}\Fast_Food_Tabel_Oktyabr_2026.xlsx",
        r"d:\Antigravity\123\Fast_Food_Tabel_Oktyabr_2026.xlsx",
        r"d:\Antigravity\123\telegrambot\Fast_Food_Tabel_Oktyabr_2026.xlsx"
    ]
    for t in targets:
        os.makedirs(os.path.dirname(t), exist_ok=True)
        wb.save(t)
        print("Tabel Excel synchronized to:", t)

if __name__ == "__main__":
    sync_fastfood_tabel_file(1453337970)
