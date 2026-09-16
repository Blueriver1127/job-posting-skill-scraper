"""
修復 202601 損壞的 wide.xlsx 縣市檔案，並重建整合 ALL_wide.xlsx
"""
import pandas as pd, os, glob, warnings, time
warnings.filterwarnings("ignore")

BASE_DIR   = "/Users/annie1127/Library/CloudStorage/OneDrive-個人/RA2024/104_JobData_new"
OUTPUT_DIR = "/Users/annie1127/Downloads/skills_output_all/202601"
SOURCE_DIR = os.path.join(BASE_DIR, "01_2026/清洗後檔案")

BROKEN = ["嘉義市", "新北市", "臺北市", "臺南市"]


def skills_to_wide_fast(long_df, raw_df):
    raw_df = raw_df.copy()
    raw_df["ID"] = raw_df["工作編號"].astype(str)
    tool_col = "擅長工具" if "擅長工具" in raw_df.columns else "電腦工具"

    counts     = long_df.groupby("ID").size().reset_index(name="技能數")
    all_skills = long_df.groupby("ID")["SKILL_NAME_ZH"].apply("｜".join).reset_index(name="技能_中文")
    sp = (long_df[long_df["SKILL_TYPE"] == "Specialized Skill"]
          .groupby("ID")["SKILL_NAME_ZH"].apply("｜".join).reset_index(name="專業技能"))
    co = (long_df[long_df["SKILL_TYPE"] == "Common Skill"]
          .groupby("ID")["SKILL_NAME_ZH"].apply("｜".join).reset_index(name="通用技能"))

    agg_df = (counts
              .merge(all_skills, on="ID", how="left")
              .merge(sp,         on="ID", how="left")
              .merge(co,         on="ID", how="left"))

    cols = ["ID", "職位名稱", "職位描述", "工作技能", tool_col]
    cols = [c for c in cols if c in raw_df.columns]
    return raw_df[cols].merge(agg_df, on="ID", how="left")


# 修復損壞縣市
for county in BROKEN:
    t0 = time.time()
    src = os.path.join(SOURCE_DIR, f"cleaned_{county}_202601.xlsx")
    if not os.path.exists(src):
        print(f"  找不到來源：{src}"); continue

    print(f"處理 {county}...")
    long_df = pd.read_parquet(os.path.join(OUTPUT_DIR, f"skills_{county}_202601_long.parquet"))
    raw_df  = pd.read_excel(src)
    wide_df = skills_to_wide_fast(long_df, raw_df)
    out = os.path.join(OUTPUT_DIR, f"skills_{county}_202601_wide.xlsx")
    wide_df.to_excel(out, index=False)
    sz = os.path.getsize(out)/1024/1024
    print(f"  {county}: {len(wide_df):,} 筆，{sz:.1f} MB，耗時 {time.time()-t0:.1f}s")

# 重建整合 ALL_wide.xlsx
print("\n重建 ALL_wide.xlsx...")
all_wide = []
wide_files = sorted(glob.glob(os.path.join(OUTPUT_DIR, "skills_*_202601_wide.xlsx")))
for f in wide_files:
    df = pd.read_excel(f)
    all_wide.append(df)
    print(f"  {os.path.basename(f)}: {len(df):,} 筆")

combined = pd.concat(all_wide, ignore_index=True)
out_combined = os.path.join(OUTPUT_DIR, "skills_202601_ALL_wide.xlsx")
combined.to_excel(out_combined, index=False)
sz = os.path.getsize(out_combined)/1024/1024
print(f"\n完成！合計 {len(combined):,} 筆，{sz:.1f} MB")
