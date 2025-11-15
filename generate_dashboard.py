import os
import glob
from datetime import datetime
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.chart import BarChart, Reference

FOLDER_INPUT = r"Rezultate Teste"
OUTPUT_EXCEL = "dashboard_final.xlsx"



def parse_quiz_file(file_path: str):
    """Citește fișierul CSV și extrage meta + întrebări + total."""
    with open(file_path, "r", encoding="utf-8-sig") as f:
        lines = [line.strip() for line in f.readlines()]

    # META
    meta_line = lines[0]
    parts = meta_line.split(",")

    name = email = "UNKNOWN"
    timestamp = datetime(2100, 1, 1)

    for p in parts:
        p = p.strip()
        if p.lower().startswith("name:"):
            name = p.split(":", 1)[1].strip()
        elif p.lower().startswith("email:"):
            email = p.split(":", 1)[1].strip()
        elif p.lower().startswith("startdate:"):
            dt = p.split(":", 1)[1].strip()
            timestamp = datetime.strptime(dt, "%Y-%m-%d %H:%M:%S")

    # TABEL + TOTAL
    df = pd.read_csv(file_path, skiprows=1, encoding="utf-8-sig")

    last = df.iloc[-1]
    if str(last["NR_INTREBARE"]).upper().startswith("TOTAL"):
        total = int(last["PUNCTAJ"])
        df_questions = df.iloc[:-1].copy()
    else:
        total = int(df["PUNCTAJ"].sum())
        df_questions = df.copy()

    return {
        "name": name,
        "email": email,
        "timestamp": timestamp,
        "total": total,
        "questions_df": df_questions,
        "source_file": os.path.basename(file_path)
    }



def make_safe_sheet_name(name: str, existing: set) -> str:
    safe = "".join(c for c in name if c.isalnum() or c in (" ", "_")).strip()
    if not safe:
        safe = "Sheet"

    base = safe[:31]
    out = base
    i = 2
    while out in existing:
        suffix = f"_{i}"
        out = base[:31 - len(suffix)] + suffix
        i += 1

    existing.add(out)
    return out



def main():
    csv_files = glob.glob(os.path.join(FOLDER_INPUT, "*.csv"))
    if not csv_files:
        print("❌ Nu există fișiere!")
        return

    records = []
    persons = []
    existing_names = set()

    # Citim toate fișierele
    for file in csv_files:
        data = parse_quiz_file(file)
        sheet_name = make_safe_sheet_name(data["name"], existing_names)

        records.append({
            "Name": data["name"],
            "Email": data["email"],
            "Total": data["total"],
            "Timestamp": data["timestamp"],
            "SheetName": sheet_name,
            "SourceFile": data["source_file"]
        })

        persons.append({
            "sheet_name": sheet_name,
            "name": data["name"],
            "email": data["email"],
            "timestamp": data["timestamp"],
            "total": data["total"],
            "df": data["questions_df"]
        })

    df_dash = pd.DataFrame(records).sort_values("Total", ascending=False)

    # Detectăm duplicatele reale
    df_dash["IS_DUPLICATE"] = df_dash.duplicated(subset=["Name", "Email"], keep=False)

    # ===== CREARE WORKBOOK =====
    wb = Workbook()
    ws_dash = wb.active
    ws_dash.title = "Dashboard"

    # Header
    headers = ["Name", "Email", "Total", "Timestamp"]
    ws_dash.append(headers)
    for c in ws_dash[1]:
        c.font = Font(bold=True)

    red_fill = PatternFill(start_color="FF9999", end_color="FF9999", fill_type="solid")

    # Scriere rânduri + colorarea duplicatelor
    excel_row = 2
    index_to_sheet = {}

    for idx, r in df_dash.iterrows():
        ws_dash.append([r["Name"], r["Email"], r["Total"], r["Timestamp"]])

        # mapare corectă index → sheet
        index_to_sheet[idx] = r["SheetName"]

        # colorare duplicat
        if r["IS_DUPLICATE"]:
            for col in range(1, 5):
                ws_dash.cell(row=excel_row, column=col).fill = red_fill

        excel_row += 1

    ws_dash.freeze_panes = "A2"

    # Total participanți
    ws_dash["F1"] = f"Total participanți: {len(df_dash)}"
    ws_dash["F1"].font = Font(bold=True)

    # Link către sheet-ul de fraudă
    ws_dash["F2"] = "⚠️ Vezi Chestionare Fraudate"
    ws_dash["F2"].hyperlink = "#'Chestionare Fraudate'!A1"
    ws_dash["F2"].style = "Hyperlink"
    ws_dash["F2"].font = Font(color="0000FF", bold=True)


    # ===== HISTOGRAMĂ =====
    intervals = ["0–5", "6–10", "11–15", "16–20", "21–25", "26–30", "31–35", "36–40"]
    bins = [(0,5),(6,10),(11,15),(16,20),(21,25),(26,30),(31,35),(36,40)]

    ws_dash.append([])
    start_row = ws_dash.max_row + 1

    ws_dash[f"G{start_row}"] = "Interval"
    ws_dash[f"H{start_row}"] = "Număr"
    ws_dash[f"G{start_row}"].font = ws_dash[f"H{start_row}"].font = Font(bold=True)

    for i, (label, (lo, hi)) in enumerate(zip(intervals, bins), start=start_row+1):
        count = ((df_dash["Total"] >= lo) & (df_dash["Total"] <= hi)).sum()
        ws_dash[f"G{i}"] = label
        ws_dash[f"H{i}"] = count

    chart = BarChart()
    chart.title = "Distribuția scorurilor"
    chart.x_axis.title = "Interval scor"
    chart.y_axis.title = "Număr participanți"

    data_ref = Reference(ws_dash, min_col=8, min_row=start_row, max_row=start_row+len(bins))
    cats_ref = Reference(ws_dash, min_col=7, min_row=start_row+1, max_row=start_row+len(bins))
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats_ref)

    ws_dash.add_chart(chart, f"J{start_row}")

    # ===== CREARE SHEET-URI INDIVIDUALE =====
    for p in persons:
        ws = wb.create_sheet(title=p["sheet_name"])

        ws.append(["Name", p["name"]])
        ws.append(["Email", p["email"]])
        ws.append(["StartDate", p["timestamp"].strftime("%Y-%m-%d %H:%M:%S")])
        ws.append(["TOTAL", p["total"]])
        ws.append([""])

        for cell in ["A1", "A2", "A3", "A4"]:
            ws[cell].font = Font(bold=True)

        ws["A6"] = "⬅ Back to Dashboard"
        ws["A6"].hyperlink = "#'Dashboard'!A1"
        ws["A6"].style = "Hyperlink"
        ws["A6"].font = Font(color="0000FF", bold=True)

        ws.append([""])

        start = ws.max_row + 1
        ws.append(list(p["df"].columns))
        for c in ws[start]:
            c.font = Font(bold=True)

        for _, row in p["df"].iterrows():
            ws.append(row.tolist())

        ws.freeze_panes = f"A{start+1}"

    # ===== HIPERLINK CORECT (după index, nu nume) =====
    excel_row = 2
    for idx, r in df_dash.iterrows():
        sheet_name = index_to_sheet[idx]
        cell = ws_dash.cell(row=excel_row, column=1)
        cell.hyperlink = f"#'{sheet_name}'!A1"
        cell.style = "Hyperlink"
        excel_row += 1

    # =====================================================================
    #           SHEET PENTRU CHESTIONARE FRAUDATE
    # =====================================================================

    fraud = []

    # detectam aceeasi persoana → scor diferit
    duplicates = df_dash.groupby(["Name", "Email"])
    for (name, email), group in duplicates:
        if len(group) > 1:
            if group["Total"].nunique() > 1:
                for _, r in group.iterrows():
                    fraud.append(["Name+Email duplicat", name, email, r["Total"], r["SourceFile"]])

    # detectam email comun → nume diferite
    by_email = df_dash.groupby("Email")
    for email, group in by_email:
        if len(group) > 1 and group["Name"].nunique() > 1:
            for _, r in group.iterrows():
                fraud.append(["Email folosit de mai multe nume", r["Name"], email, r["Total"], r["SourceFile"]])

    # detectam nume comun → emailuri diferite
    by_name = df_dash.groupby("Name")
    for name, group in by_name:
        if len(group) > 1 and group["Email"].nunique() > 1:
            for _, r in group.iterrows():
                fraud.append(["Nume folosit cu emailuri diferite", name, r["Email"], r["Total"], r["SourceFile"]])

    ws_fraud = wb.create_sheet("Chestionare Fraudate")
    ws_fraud.append(["Tip suspiciune", "Name", "Email", "Total", "Fișier sursă"])

    # Adăugăm buton de back la dashboard
    ws_fraud["A2"] = "⬅ Înapoi la Dashboard"
    ws_fraud["A2"].hyperlink = "#'Dashboard'!A1"
    ws_fraud["A2"].style = "Hyperlink"
    ws_fraud["A2"].font = Font(color="0000FF", bold=True)

    ws_fraud.append([""])  # rând gol după buton


    for c in ws_fraud[1]:
        c.font = Font(bold=True)

    red = PatternFill(start_color="FF6666", end_color="FF6666", fill_type="solid")

    for row in fraud:
        ws_fraud.append(row)
        r = ws_fraud.max_row
        for col in range(1, 6):
            ws_fraud.cell(row=r, column=col).fill = red

    wb.save(OUTPUT_EXCEL)
    print("✔️ Dashboard complet generat cu succes!")


if __name__ == "__main__":
    main()
