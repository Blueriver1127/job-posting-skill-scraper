"""
修復現有 9 個月輸出資料：
  1. 從 long parquet 移除短詞（SKILL_NAME_ZH ≤ 2 字中文）和重複技能
  2. 從固定後的 long parquet + 原始來源檔重建 wide xlsx
     （不依賴可能已損壞的舊 wide 檔案）
"""
import pandas as pd, os, glob, re, warnings, time
warnings.filterwarnings("ignore")

BASE_DIR    = "/Users/annie1127/Library/CloudStorage/OneDrive-個人/RA2024/104_JobData_new"
OUTPUT_BASE = "/Users/annie1127/Downloads/skills_output_all"

MONTH_CONFIGS = [
    ("202502", "02_2025/整理後檔案_縣市", "cleaned_*_combined.csv.xlsx"),
    ("202503", "03_2025/整理後檔案_縣市", "cleaned_*_combined.csv.xlsx"),
    ("202504", "04_2025/整理後檔案_縣市", "cleaned_*_combined.csv.xlsx"),
    ("202505", "05_2025/整理後檔案_縣市", "cleaned_*_combined.csv.xlsx"),
    ("202506", "06_2025/整理後檔案_縣市", "cleaned_*_combined_202506.xlsx"),
    ("202508", "08_2025/整理後檔案_縣市", "cleaned_*_202508.xlsx"),
    ("202509", "09_2025/整理後檔案",      "cleaned_*_202509.xlsx"),
    ("202510", "10_2025/整理後檔案",      "cleaned_*_202510.xlsx"),
    ("202601", "01_2026/清洗後檔案",      "cleaned_*_202601.xlsx"),
]


def is_short_chinese(name):
    if not isinstance(name, str):
        return False
    zh_count = len(re.findall(r'[一-鿿]', name))
    return zh_count > 0 and len(name) <= 2


def fix_long(df):
    n0 = len(df)
    df = df[~df["SKILL_NAME_ZH"].apply(is_short_chinese)].copy()
    n1 = len(df)
    df = df.drop_duplicates(subset=["ID", "SKILL_NAME_ZH"], keep="first")
    n2 = len(df)
    return df, n0 - n1, n1 - n2


_CAT9_ZH = {
    "Cognitive Skills":         "認知技能",
    "Social Skills":            "社交技能",
    "Character Skills":         "特質技能",
    "Financial Skills":         "財務技能",
    "Management Skills":        "管理技能",
    "General Digital Skills":   "數位技能",
    "Technical Support Skills": "技術技能",
    "Advanced Computer Skills": "電腦技能",
    "AI & Big Data Skills":     "AI技能",
    "Unclassified":             "其他技能",
}


def build_wide(long_df, raw_df):
    """從固定後的 long + 原始來源檔建立 wide 格式"""
    raw_df = raw_df.copy()
    raw_df["ID"] = raw_df["工作編號"].astype(str)
    tool_col = "擅長工具" if "擅長工具" in raw_df.columns else "電腦工具"

    long_df = long_df.copy()
    long_df["ID"] = long_df["ID"].astype(str)
    # 舊版 parquet 可能沒有 SKILL_CAT9 欄，補上預設值
    if "SKILL_CAT9" not in long_df.columns:
        long_df["SKILL_CAT9"] = long_df.get("SKILL_TYPE", "Unclassified").map(
            {"Specialized Skill": "Unclassified", "Common Skill": "Unclassified"}
        ).fillna("Unclassified")

    counts     = long_df.groupby("ID").size().reset_index(name="技能數")
    all_skills = long_df.groupby("ID")["SKILL_NAME_ZH"].apply("｜".join).reset_index(name="技能_中文")

    agg_df = counts.merge(all_skills, on="ID", how="left")
    for en_cat, zh_col in _CAT9_ZH.items():
        cat_skills = (
            long_df[long_df["SKILL_CAT9"] == en_cat]
            .groupby("ID")["SKILL_NAME_ZH"].apply("｜".join)
            .reset_index(name=zh_col)
        )
        agg_df = agg_df.merge(cat_skills, on="ID", how="left")

    keep_cols = [c for c in ["ID", "職位名稱", "職位描述", "工作技能", tool_col]
                 if c in raw_df.columns]
    return raw_df[keep_cols].merge(agg_df, on="ID", how="left")


def extract_county(filename):
    name = os.path.basename(filename)
    name = name.replace("cleaned_", "").replace(".xlsx", "")
    name = re.sub(r"_combined(\.csv)?(_\d{6})?$", "", name)
    name = re.sub(r"_\d{6}$", "", name)
    return name


grand_removed_short = 0
grand_removed_dup   = 0
session_start = time.time()

for month, rel_dir, pattern in MONTH_CONFIGS:
    t_month = time.time()
    output_dir = os.path.join(OUTPUT_BASE, month)
    raw_dir    = os.path.join(BASE_DIR, rel_dir)
    print(f"\n【{month}】")

    src_files = sorted(glob.glob(os.path.join(raw_dir, pattern)))
    if not src_files:
        print(f"  ⚠️  找不到來源檔：{os.path.join(raw_dir, pattern)}")
        continue

    month_removed_short = 0
    month_removed_dup   = 0
    fixed_longs = []

    for src_path in src_files:
        county = extract_county(src_path)

        long_path = os.path.join(output_dir, f"skills_{county}_{month}_long.parquet")
        wide_path = os.path.join(output_dir, f"skills_{county}_{month}_wide.xlsx")

        if not os.path.exists(long_path):
            print(f"  ⚠️  找不到 {os.path.basename(long_path)}，跳過")
            continue

        t0 = time.time()
        long_df, rs, rd = fix_long(pd.read_parquet(long_path))
        month_removed_short += rs
        month_removed_dup   += rd

        long_df.to_parquet(long_path, index=False)
        fixed_longs.append(long_df)

        raw_df   = pd.read_excel(src_path)
        wide_df  = build_wide(long_df, raw_df)
        wide_df.to_excel(wide_path, index=False)

        print(f"  {county}: 移除短詞 {rs:,}，去重 {rd:,}，剩 {len(long_df):,} 筆 ({time.time()-t0:.1f}s)")

    if not fixed_longs:
        continue

    # 重建 ALL 整合檔
    combined_long = pd.concat(fixed_longs, ignore_index=True)
    combined_long.to_parquet(os.path.join(output_dir, f"skills_{month}_ALL_long.parquet"), index=False)

    all_wides = []
    for ldf in fixed_longs:
        all_wides.append(pd.read_excel(
            os.path.join(output_dir,
                         f"skills_{ldf['縣市'].iloc[0]}_{month}_wide.xlsx")))
    combined_wide = pd.concat(all_wides, ignore_index=True)
    combined_wide.to_excel(
        os.path.join(output_dir, f"skills_{month}_ALL_wide.xlsx"), index=False)

    sz = os.path.getsize(os.path.join(output_dir, f"skills_{month}_ALL_wide.xlsx")) / 1024 / 1024
    grand_removed_short += month_removed_short
    grand_removed_dup   += month_removed_dup
    print(f"  → ALL: {len(combined_long):,} 筆，wide {sz:.1f} MB，月份耗時 {(time.time()-t_month)/60:.1f} 分鐘")

total_min = (time.time() - session_start) / 60
print(f"\n{'='*55}")
print(f"全部完成！耗時 {total_min:.1f} 分鐘")
print(f"  刪除短/通用詞：{grand_removed_short:,} 筆")
print(f"  刪除重複技能： {grand_removed_dup:,} 筆")
print(f"{'='*55}")
