import sys
import xlwings as xw
from pathlib import Path
import pandas as pd
from datetime import datetime

# =========================
# フォルダ設定
# Python実行時: run_drm.py の場所
# exe実行時   : run_drm.exe の場所
# =========================

if getattr(sys, "frozen", False):
    BASE_DIR = Path(sys.executable).parent
else:
    BASE_DIR = Path(__file__).resolve().parent

DAT_FOLDER = BASE_DIR / "input_dat"
EXCEL_FILE = BASE_DIR / "sample_workbook.xlsm"

OUTPUT_DIR = BASE_DIR / "output"
OUTPUT_DIR.mkdir(exist_ok=True)

OUTPUT_FILE = OUTPUT_DIR / "drm_summary.xlsx"

# =========================
# 日付変換関数
# =========================

def parse_japanese_datetime(value):
    if value is None:
        return None

    if isinstance(value, datetime):
        return value

    text = str(value).strip()

    try:
        return datetime.strptime(text, "%Y年%m月%d日%H時%M分%S秒")
    except ValueError:
        return None

# =========================
# 事前チェック
# =========================

if not DAT_FOLDER.exists():
    raise FileNotFoundError(f"input_dat フォルダが見つかりません: {DAT_FOLDER}")

if not EXCEL_FILE.exists():
    raise FileNotFoundError(f"xlsmファイルが見つかりません: {EXCEL_FILE}")

dat_files = list(DAT_FOLDER.glob("*.dat"))

print("検索フォルダ: input_dat")
print("datファイル数:", len(dat_files))
print()
print("処理中...")


    

if len(dat_files) == 0:
    raise FileNotFoundError("datファイルが見つかりません。input_dat フォルダを確認してください。")

results = []

# =========================
# Excel起動
# =========================

app = xw.App(visible=False)
app.display_alerts = False
app.screen_updating = False

wb = app.books.open(str(EXCEL_FILE))

try:
    for dat_file in dat_files:
        
        wb.macro("Python_dat_2_csv")(str(dat_file))
        app.calculate()

        sheet = wb.sheets["各種ﾃﾞｰﾀ"]

        date_value = parse_japanese_datetime(sheet.range("E8").value)

        vin_model = str(sheet.range("E9").value).replace(" ", "").replace("　", "").strip()
        vin_number = str(sheet.range("E10").value).replace(" ", "").replace("　", "").strip()

        results.append({
            "datファイル名": dat_file.name,
            "日付": date_value,
            "VIN型式部": vin_model,
            "VIN車番部": vin_number,
            "VIN": vin_model + vin_number,
            "総走行距離（Km）": sheet.range("G24").value,
            "総走行時間（Hour）": sheet.range("I26").value,
            "平均燃費（アイドル含）": sheet.range("E34").value,
            "平均車速（アイドル含）": sheet.range("E37").value,
        })

finally:
    wb.close()
    app.quit()

# =========================
# Excel出力
# =========================

df = pd.DataFrame(results)

with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
    df.to_excel(writer, index=False, sheet_name="Summary")

    ws = writer.sheets["Summary"]

    for cell in ws["B"][1:]:
        cell.number_format = "yyyy/mm/dd hh:mm:ss"

print()
print("=================================")
print("処理完了")
print("=================================")
print("出力ファイル:output/drm_summary.xlsx")
print("出力件数:", len(df))

input("Enterキーで終了...")