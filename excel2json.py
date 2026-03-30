# excel2json.py
import os
import json
import re
from datetime import datetime, timedelta, timezone
import pandas as pd
from collections import defaultdict
import warnings
import yaml

warnings.filterwarnings("ignore", category=UserWarning)
pd.set_option("future.no_silent_downcasting", True)

# =========================================
# 設定読み込み
# =========================================
with open(
    os.path.join(os.path.dirname(__file__), "config.yaml"), encoding="utf-8"
) as f:
    CONFIG = yaml.safe_load(f)


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
    min_date = None
    max_date = None
    for _, row in df.iterrows():
        date = pd.to_datetime(row.iloc[1]).strftime("%Y-%m-%d")
        if min_date is None or date < min_date:
            min_date = date
        if max_date is None or date > max_date:
            max_date = date

        comment = row.iloc[13] if len(row) > 13 else ""

        if comment:
            timetable.append(
                {
                    "grade": grade,
                    "date": date,
                    "period": 0,
                    "courses": "",
                    "room": "",
                    "comment": comment,
                }
            )

        for p in range(1, 6):
            col_c = p * 2 + 1  # 講義名の列
            col_r = p * 2 + 2  # 教室の列
            if col_r >= len(row):
                continue
            cname = row.iloc[col_c]
            room = row.iloc[col_r]
            if not cname:
                continue
            timetable.append(
                {
                    "grade": grade,
                    "date": date,
                    "period": p,
                    "courses": cname,
                    "room": room,
                    "comment": "",
                }
            )
    print(
        f"✓ {file_path} / {sheet_name} : {len(timetable)} 件 ({min_date} ～ {max_date})"
    )
    return timetable


# -------------------------
# 指定ファイル内のシート（全部）ロード
# -------------------------
def load_year_term(file_path):
    """
    Excelファイルのシート名から yyyy年度(学年+前期|後期) にマッチするシートを自動検出して読み込む。
    学年の変換は config.yaml の grades を使用。
    """
    try:
        sheet_names = pd.ExcelFile(file_path).sheet_names
    except Exception as e:
        print(f"⚠ ファイルを開けません：{file_path} → {e}")
        return []

    SHEET_PATTERN = re.compile(CONFIG["sheet_pattern"])
    result = []
    for sheet in sheet_names:
        m = SHEET_PATTERN.match(sheet)
        if not m:
            # print(f"✗ シート名不一致: {sheet} （パターン: {CONFIG['sheet_pattern']}）")
            continue
        else:
            # print(f"✓ シート名マッチ: {sheet} → {m.groups()}")
            pass
        _year = m.group(1)
        grade = m.group(2)
        _term = m.group(3)
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
    jst = timezone(timedelta(hours=9))
    ts = datetime.fromtimestamp(os.stat(file_path).st_mtime, tz=jst)
    save_json(
        {"file_path": file_path, "last_modified": ts.strftime("%Y-%m-%d %H:%M:%S")},
        out_path,
    )


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
    # 全ファイル読み込み・連結
    all_timetable = []
    for entry in CONFIG["files"]:
        all_timetable.extend(load_year_term(entry["save_name"]))

    # 助産統合
    all_timetable = add_schedule_to_josan(all_timetable)

    # 日付範囲でフィルタ
    dp = CONFIG["display_period"]
    start_date = datetime.strptime(dp["start_date"], "%Y-%m-%d")
    end_date = datetime.strptime(dp["end_date"], "%Y-%m-%d")
    print(
        f"◆ 表示期間: {start_date.strftime('%Y-%m-%d')} ～ {end_date.strftime('%Y-%m-%d')}"
    )
    filtered = filter_by_date_range(all_timetable, start_date, end_date)
    print(f"◆ フィルタ前: {len(all_timetable)} 件 → フィルタ後: {len(filtered)} 件")

    save_json(filtered, CONFIG["output"]["schedule_json"])
    save_info_json(CONFIG["output"]["schedule_json"], CONFIG["output"]["info_json"])
