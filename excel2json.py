# excel2json.py
import os
import re
import json
from datetime import datetime, timedelta, timezone
import pandas as pd
import warnings
import yaml

warnings.filterwarnings("ignore", category=UserWarning)
pd.set_option("future.no_silent_downcasting", True)

# =========================================
# 設定読み込み
# =========================================
with open(os.path.join(os.path.dirname(__file__), "config.yaml"), encoding="utf-8") as f:
    CONFIG = yaml.safe_load(f)

# -------------------------
# save_name から term / year を推定
# -------------------------
TERM_MAP = {"spring": "前期", "fall": "後期"}

def parse_save_name(save_name):
    """
    'schedule_spring_2026.xlsx' -> ('前期', 2026)
    'schedule_fall_2026.xlsx'   -> ('後期', 2026)
    """
    m = re.match(r"schedule_(spring|fall)_(\d{4})\.xlsx$", save_name)
    if not m:
        raise ValueError(f"save_name がパターンに一致しません: {save_name}")
    season, year_str = m.groups()
    return TERM_MAP[season], int(year_str)

# -------------------------
# 個別JSON出力パス
# -------------------------
def individual_json_path(save_name):
    base = os.path.splitext(save_name)[0]
    return f"docs/{base}.json"

# -------------------------
# display_period 読み込み
# -------------------------
def get_display_period():
    dp = CONFIG["display_period"]
    start = datetime.strptime(str(dp["start_date"]), "%Y-%m-%d")
    end = datetime.strptime(str(dp["end_date"]), "%Y-%m-%d")
    return start, end

# -------------------------
# Excel シート読み込み（1シート）
# -------------------------
def load_sheet(file_path, sheet_name, grade):
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name, engine="openpyxl")
    except Exception as e:
        print(f"⚠ シート読み込みエラー：{file_path} / {sheet_name} → {e}")
        return []

    df = df[pd.to_datetime(df.iloc[:, 1], format="%Y-%m-%d", errors="coerce").notnull()]
    df = df.fillna("").infer_objects(copy=False)

    timetable = []
    for _, row in df.iterrows():
        date = pd.to_datetime(row.iloc[1]).strftime("%Y-%m-%d")
        comment = row.iloc[13] if len(row) > 13 else ""

        if comment:
            timetable.append({
                "grade": grade,
                "date": date,
                "period": 0,
                "courses": "",
                "room": "",
                "comment": comment
            })

        for p in range(1, 6):
            col_c = p * 2 + 1
            col_r = p * 2 + 2
            if col_r >= len(row):
                continue
            cname = row.iloc[col_c]
            room = row.iloc[col_r]
            if not cname:
                continue
            timetable.append({
                "grade": grade,
                "date": date,
                "period": p,
                "courses": cname,
                "room": room,
                "comment": ""
            })
    return timetable

# -------------------------
# 指定ファイル内のシート（全部）ロード
# -------------------------
def load_year_term(file_path, sheet_map):
    """
    sheet_map : { Excelシート名 : grade名 }
    """
    result = []
    for sheet, grade in sheet_map.items():
        part = load_sheet(file_path, sheet, grade)
        result.extend(part)
    return result

# -------------------------
# 助産統合（4年→4年助産補完）
# -------------------------
def add_schedule_to_josan(timetable):
    josan_cfg = CONFIG["josan"]
    source_grade = josan_cfg["source_grade"]
    target_grade = josan_cfg["target_grade"]

    source = {}
    target = {}
    for c in timetable:
        key = c["date"] + str(c["period"])
        if c["grade"] == source_grade:
            source[key] = c
        elif c["grade"] == target_grade:
            target[key] = c
    for key, c in source.items():
        if key in target:
            continue
        new_c = c.copy()
        new_c["grade"] = target_grade
        timetable.append(new_c)
    return timetable

# -------------------------
# JSON / info 保存
# -------------------------
def save_json(data, path):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print(f"✓ {path} を生成（{len(data)} 件）")

def save_info_json(file_path, out_path):
    if not os.path.exists(file_path):
        print(f"⚠ {file_path} がありません（空の info を作成）")
        save_json({"file_path": file_path, "last_modified": None}, out_path)
        return
    jst = timezone(timedelta(hours=9))
    ts = datetime.fromtimestamp(os.stat(file_path).st_mtime, tz=jst)
    save_json({"file_path": file_path, "last_modified": ts.strftime("%Y-%m-%d %H:%M:%S")}, out_path)

# -------------------------
# シート名マップ（config.yaml から構築）
# -------------------------
def sheet_names_for_year(year, term):
    yyyy = f"{year}年度"
    sheets_cfg = CONFIG["sheets"][term]
    return {
        entry["sheet"].format(yyyy=yyyy): entry["grade"]
        for entry in sheets_cfg
    }

# -------------------------
# 日付範囲フィルタ
# -------------------------
def filter_by_date_range(timetable, start_date, end_date):
    result = []
    for c in timetable:
        try:
            d = datetime.strptime(c["date"], "%Y-%m-%d")
        except Exception:
            continue
        if start_date <= d <= end_date:
            result.append(c)
    return result

# -------------------------
# メイン
# -------------------------
if __name__ == "__main__":
    start_date, end_date = get_display_period()
    print(f"◆ 表示期間: {start_date.strftime('%Y-%m-%d')} ～ {end_date.strftime('%Y-%m-%d')}")

    all_timetable = []

    for entry in CONFIG["files"]:
        save_name = entry["save_name"]
        term, year = parse_save_name(save_name)
        sheet_map = sheet_names_for_year(year, term)

        print(f"→ {save_name} : term={term}, year={year}年度")
        file_data = load_year_term(save_name, sheet_map)

        # 個別JSON（フィルタなし）を保存
        ind_path = individual_json_path(save_name)
        save_json(file_data, ind_path)

        all_timetable.extend(file_data)

    # 助産統合
    all_timetable = add_schedule_to_josan(all_timetable)

    # display_period でフィルタ
    filtered = filter_by_date_range(all_timetable, start_date, end_date)
    print(f"◆ フィルタ前: {len(all_timetable)} 件 → フィルタ後: {len(filtered)} 件")

    # 結合フィルタ済み JSON を出力
    save_json(filtered, CONFIG["output"]["schedule_json"])

    # info JSON を出力
    for entry in CONFIG["files"]:
        save_info_json(entry["save_name"], entry["info_json"])
