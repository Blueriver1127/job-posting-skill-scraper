"""
============================================================
104 職缺技能比對 — 批次月度版（一鍵跑全部縣市）
============================================================

【功能】
    一次讀入多個縣市的 Excel 職缺檔案，
    合併後統一比對（自動機只建一次），
    再依縣市分別輸出 parquet 和 Excel。

【使用前準備】
    pip3 install pandas openpyxl pyarrow ahocorasick-python

【使用方式】
    python3 104_batch_monthly.py

    或指定月份：
    python3 104_batch_monthly.py --month 202604

【設定區】
    修改下方「使用者設定區」：
    - RAW_DIR     : 存放清洗後 Excel 的資料夾（每個縣市一個檔案）
    - LEXICON_PATH: 技能詞庫路徑
    - OUTPUT_DIR  : 輸出資料夾
    - FILE_PATTERN: 檔案名稱規則（用 {month} 代替月份）

【檔案命名規則】
    輸入範例：cleaned_臺北市_202604.xlsx
    輸出範例：skills_臺北市_202604_long.parquet
              skills_臺北市_202604_wide.xlsx
============================================================
"""

import re
import time
import glob
import argparse
import os
import warnings
import ahocorasick
import pandas as pd
import nltk
import jieba
from nltk.stem import PorterStemmer

warnings.filterwarnings("ignore")
nltk.download("wordnet", quiet=True)
nltk.download("omw-1.4", quiet=True)
jieba.setLogLevel(60)

_stemmer = PorterStemmer()


def _stem_english_text(text: str) -> str:
    return re.sub(r"[a-z]+", lambda m: _stemmer.stem(m.group()), text)


def _get_zh_boundaries(text: str) -> set:
    positions = {0}
    pos = 0
    for token in jieba.cut(text):
        pos += len(token)
        positions.add(pos)
    return positions


def setup_jieba_dict(lex: pd.DataFrame):
    """將詞庫中 3 字以上的中文技能詞加入 jieba 自訂詞典"""
    count = 0
    for _, row in lex.iterrows():
        for field in ["Skill_Name_ZH", "Keywords"]:
            val = str(row.get(field, "")).strip()
            if not val or val.lower() == "nan":
                continue
            terms = val.split("｜") if field == "Keywords" else [val]
            for t in terms:
                t = t.strip()
                if has_chinese(t) and len(t) >= 3:
                    jieba.add_word(t, freq=1000)
                    count += 1
    print(f"  jieba 自訂詞典：加入 {count} 個詞條")

# ============================================================
# 【使用者設定區】
# ============================================================
RAW_DIR      = "/Users/annie1127/Library/CloudStorage/OneDrive-個人/RA2024/104_JobData_new/01_2026/清洗後檔案"
LEXICON_PATH = "/Users/annie1127/Downloads/skill_lexicon_v10.xlsx"
OUTPUT_DIR   = "/Users/annie1127/Downloads/skills_output"
FILE_PATTERN = "cleaned_*_{month}.xlsx"   # {month} 會被替換成實際月份，例如 202604
# ============================================================

SOFTWARE_CATEGORIES = {17, 370, 369, 371, 372, 373, 374, 375, 376, 377, 378, 379, 380}


def has_chinese(text: str) -> bool:
    if not isinstance(text, str):
        return False
    return bool(re.search(r"[一-鿿]", text))


def load_lexicon(path: str):
    """
    讀取詞庫並展開 Keywords，回傳：
    - lex_records: list of dict（比 DataFrame.loc 快 10 倍以上）
    - automaton: Aho-Corasick 自動機
    """
    print(f"  讀取詞庫：{path}")
    lex = pd.read_excel(path)

    # 展開關鍵字，建成 list of dict（避免後續用 lex.loc 查詢）
    term_to_skills = {}  # term → list of skill dict

    for _, row in lex.iterrows():
        try:
            cat_code = int(row["Category_Code"])
        except (ValueError, TypeError):
            cat_code = 0

        skill = {
            "SKILL_ID":               str(row["Skill_ID"]),
            "SKILL_NAME":             str(row["Skill_Name"]),
            "SKILL_NAME_ZH":          str(row["Skill_Name_ZH"]),
            "SKILL_TYPE":             str(row["Skill_Type"]),
            "SKILL_CATEGORY":         str(row["Category_Code"]),
            "SKILL_CATEGORY_NAME":    str(row["Category_Name"]),
            "SKILL_SUBCATEGORY":      str(row["Subcategory_Code"]),
            "SKILL_SUBCATEGORY_NAME": str(row["Subcategory_Name"]),
            "IS_SOFTWARE":            cat_code in SOFTWARE_CATEGORIES,
        }

        # 收集所有比對詞
        terms = []
        zh = str(row.get("Skill_Name_ZH", "")).strip()
        if has_chinese(zh):
            terms.append((zh, True))
        else:
            en = str(row.get("Skill_Name", "")).strip().lower()
            if len(en) >= 2:
                terms.append((en, False))

        kw_str = str(row.get("Keywords", "")).strip()
        if kw_str and kw_str.lower() != "nan":
            for kw in kw_str.split("｜"):
                kw = kw.strip()
                if len(kw) < 2:
                    continue
                if has_chinese(kw):
                    terms.append((kw, True))
                else:
                    terms.append((kw.lower(), False))

        # 英文關鍵字加入詞幹形式（讓 managing 能比對到 management）
        extra = []
        for term, is_zh in terms:
            if not is_zh:
                stemmed = _stemmer.stem(term)
                if stemmed != term and len(stemmed) >= 3:
                    extra.append((stemmed, False))
        terms.extend(extra)

        seen = set()
        for term, is_zh in terms:
            if term in seen:
                continue
            seen.add(term)
            entry = (skill, is_zh)
            if term not in term_to_skills:
                term_to_skills[term] = []
            # 避免同技能重複加入
            if not any(e[0]["SKILL_ID"] == skill["SKILL_ID"] for e in term_to_skills[term]):
                term_to_skills[term].append(entry)

    # 建 Aho-Corasick 自動機
    A = ahocorasick.Automaton()
    for term, entries in term_to_skills.items():
        if len(term) >= 2:
            A.add_word(term, entries)
    A.make_automaton()

    print(f"  詞庫比對詞條數：{len(term_to_skills)}")
    return A


def _is_ascii_alnum(c: str) -> bool:
    return c.isascii() and c.isalnum()


def _is_word_boundary(text: str, start: int, end: int) -> bool:
    before_ok = (start == 0) or (not _is_ascii_alnum(text[start - 1]) and text[start - 1] != "_")
    after_ok = (end == len(text)) or (not _is_ascii_alnum(text[end]) and text[end] != "_")
    return before_ok and after_ok


def match_skills(texts: list[str], automaton) -> list[dict]:
    """直接從自動機取 skill dict，不需要查 DataFrame"""
    field_labels = ["職位描述", "工作技能", "電腦工具"]
    seen_ids = set()
    results = []

    for label, raw_text in zip(field_labels, texts):
        if not isinstance(raw_text, str) or not raw_text.strip():
            continue
        text_lower = raw_text.lower()

        for end_idx, entries in automaton.iter(text_lower):
            skill0, is_zh0 = entries[0]
            term_len = len(list(automaton.iter(text_lower))[0][1]) if False else 0

            # 取得 term 長度（從 skill 名稱推算）
            # 實際上 entries 存的是 (skill_dict, is_zh)，term 長度需另外處理
            # 改用記錄 term 長度的方式
            for skill, is_zh in entries:
                skill_id = skill["SKILL_ID"]
                if skill_id in seen_ids:
                    continue
                seen_ids.add(skill_id)
                results.append({**skill, "MATCHED_FROM": label})

    return results


def match_skills_v2(texts: list[str], automaton) -> list[dict]:
    """
    改良版：automaton value 存 (term_len, skill, is_zh) 以支援邊界檢查
    """
    field_labels = ["職位描述", "工作技能", "電腦工具"]
    seen_ids = set()
    results = []

    for label, raw_text in zip(field_labels, texts):
        if not isinstance(raw_text, str) or not raw_text.strip():
            continue
        text_lower = raw_text.lower()
        text_stemmed = _stem_english_text(text_lower)
        scan_texts = [text_lower] if text_lower == text_stemmed else [text_lower, text_stemmed]
        zh_boundaries = _get_zh_boundaries(text_lower)

        for scan_text in scan_texts:
            for end_idx, entries in automaton.iter(scan_text):
                for term_len, skill, is_zh in entries:
                    start_idx = end_idx - term_len + 1

                    if is_zh:
                        if scan_text is text_lower:
                            if start_idx not in zh_boundaries or (end_idx + 1) not in zh_boundaries:
                                continue
                    else:
                        if not _is_word_boundary(scan_text, start_idx, end_idx + 1):
                            continue

                    skill_id = skill["SKILL_ID"]
                    if skill_id in seen_ids:
                        continue
                    seen_ids.add(skill_id)
                    results.append({**skill, "MATCHED_FROM": label})

    return results


def load_lexicon_v2(path: str):
    """改良版：automaton value 存 (term_len, skill, is_zh)"""
    print(f"  讀取詞庫：{path}")
    lex = pd.read_excel(path)

    term_to_entries = {}

    for _, row in lex.iterrows():
        try:
            cat_code = int(row["Category_Code"])
        except (ValueError, TypeError):
            cat_code = 0

        skill = {
            "SKILL_ID":               str(row["Skill_ID"]),
            "SKILL_NAME":             str(row["Skill_Name"]),
            "SKILL_NAME_ZH":          str(row["Skill_Name_ZH"]),
            "SKILL_TYPE":             str(row["Skill_Type"]),
            "SKILL_CATEGORY":         str(row["Category_Code"]),
            "SKILL_CATEGORY_NAME":    str(row["Category_Name"]),
            "SKILL_SUBCATEGORY":      str(row["Subcategory_Code"]),
            "SKILL_SUBCATEGORY_NAME": str(row["Subcategory_Name"]),
            "IS_SOFTWARE":            cat_code in SOFTWARE_CATEGORIES,
        }

        terms = []
        zh = str(row.get("Skill_Name_ZH", "")).strip()
        if has_chinese(zh):
            terms.append((zh, True))
        else:
            en = str(row.get("Skill_Name", "")).strip().lower()
            if len(en) >= 2:
                terms.append((en, False))

        kw_str = str(row.get("Keywords", "")).strip()
        if kw_str and kw_str.lower() != "nan":
            for kw in kw_str.split("｜"):
                kw = kw.strip()
                if len(kw) < 2:
                    continue
                if has_chinese(kw):
                    terms.append((kw, True))
                else:
                    terms.append((kw.lower(), False))

        seen = set()
        for term, is_zh in terms:
            if term in seen:
                continue
            seen.add(term)
            entry = (len(term), skill, is_zh)
            if term not in term_to_entries:
                term_to_entries[term] = []
            if not any(e[1]["SKILL_ID"] == skill["SKILL_ID"] for e in term_to_entries[term]):
                term_to_entries[term].append(entry)

    A = ahocorasick.Automaton()
    for term, entries in term_to_entries.items():
        A.add_word(term, entries)
    A.make_automaton()

    print(f"  詞庫比對詞條數：{len(term_to_entries)}")
    return A


def process_county(df: pd.DataFrame, automaton, county: str) -> pd.DataFrame:
    """處理單一縣市的所有職缺"""
    # 相容兩種欄位名稱（舊：電腦工具，新：擅長工具）
    tool_col = "擅長工具" if "擅長工具" in df.columns else "電腦工具"
    records = []
    for i, row in df.iterrows():
        job_id = str(row["工作編號"])
        texts = [
            str(row.get("職位描述", "")),
            str(row.get("工作技能", "")),
            str(row.get(tool_col, "")),
        ]
        skills = match_skills_v2(texts, automaton)
        for s in skills:
            records.append({"ID": job_id, "縣市": county, **s})

        if (i + 1) % 5000 == 0:
            print(f"    {i + 1:,} / {len(df):,} 筆...")

    return pd.DataFrame(records)


def skills_to_wide(long_df: pd.DataFrame, raw_df: pd.DataFrame) -> pd.DataFrame:
    """長表格轉寬表格"""
    raw_df = raw_df.copy()
    raw_df["ID"] = raw_df["工作編號"].astype(str)

    def agg(group):
        specialized = "｜".join(r["SKILL_NAME_ZH"] for _, r in group.iterrows() if r["SKILL_TYPE"] == "Specialized Skill")
        common      = "｜".join(r["SKILL_NAME_ZH"] for _, r in group.iterrows() if r["SKILL_TYPE"] == "Common Skill")
        return pd.Series({
            "技能數":    len(group),
            "技能_中文": "｜".join(group["SKILL_NAME_ZH"]),
            "專業技能":  specialized,
            "通用技能":  common,
        })

    import warnings
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        agg_df = long_df.groupby("ID").apply(agg, include_groups=False).reset_index()

    tool_col = "擅長工具" if "擅長工具" in raw_df.columns else "電腦工具"
    desc = raw_df[["ID", "職位名稱", "職位描述", "工作技能", tool_col]]
    return desc.merge(agg_df, on="ID", how="left")


def run(month: str, raw_dir: str, lexicon_path: str, output_dir: str):
    os.makedirs(output_dir, exist_ok=True)

    # 找所有符合月份的 Excel 檔案
    pattern = os.path.join(raw_dir, FILE_PATTERN.replace("{month}", month))
    files = sorted(glob.glob(pattern))
    if not files:
        print(f"找不到符合的檔案：{pattern}")
        return
    print(f"找到 {len(files)} 個縣市檔案")

    # 建自動機（只建一次）
    print("\n建立比對自動機...")
    t0 = time.time()
    automaton = load_lexicon_v2(lexicon_path)
    # 設定 jieba 自訂詞典（用原始 DataFrame 版本）
    lex_raw = pd.read_excel(lexicon_path)
    setup_jieba_dict(lex_raw)
    print(f"  完成，耗時 {time.time()-t0:.1f} 秒\n")

    total_jobs = 0
    total_skills = 0
    t_start = time.time()
    all_long = []
    all_wide = []

    for fpath in files:
        fname = os.path.basename(fpath)
        # 從檔名取縣市名稱（例：cleaned_臺北市_202604.xlsx → 臺北市）
        county = fname.replace("cleaned_", "").replace(f"_{month}.xlsx", "")
        print(f"處理：{county}（{fname}）")

        df = pd.read_excel(fpath)
        n_jobs = len(df)
        total_jobs += n_jobs

        t1 = time.time()
        long_df = process_county(df, automaton, county)
        elapsed = time.time() - t1

        n_skills = len(long_df)
        total_skills += n_skills
        avg = n_skills / n_jobs if n_jobs > 0 else 0
        print(f"  {n_jobs} 筆職缺，{n_skills} 筆技能，平均 {avg:.1f} 個，耗時 {elapsed:.1f} 秒")

        all_long.append(long_df)
        all_wide.append(skills_to_wide(long_df, df))

    # 整合所有縣市輸出兩個檔案
    print("\n整合所有縣市...")
    combined_long = pd.concat(all_long, ignore_index=True)
    combined_wide = pd.concat(all_wide, ignore_index=True)

    long_path = os.path.join(output_dir, f"skills_{month}_ALL_long.parquet")
    wide_path = os.path.join(output_dir, f"skills_{month}_ALL_wide.xlsx")

    combined_long.to_parquet(long_path, index=False)
    print(f"  長表格：{long_path}")

    combined_wide.to_excel(wide_path, index=False)
    print(f"  寬表格：{wide_path}")

    total_elapsed = time.time() - t_start
    print(f"\n{'='*50}")
    print(f"全部完成！")
    print(f"  總職缺數：{total_jobs:,} 筆")
    print(f"  總技能記錄：{total_skills:,} 筆")
    print(f"  總耗時：{total_elapsed/60:.1f} 分鐘")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="104 月度批次技能比對")
    parser.add_argument("--month",   default="202604", help="月份，例如 202604")
    parser.add_argument("--raw-dir", default=RAW_DIR)
    parser.add_argument("--lexicon", default=LEXICON_PATH)
    parser.add_argument("--output",  default=OUTPUT_DIR)
    args = parser.parse_args()

    print(f"月份：{args.month}")
    print(f"輸入資料夾：{args.raw_dir}")
    print(f"輸出資料夾：{args.output}\n")
    run(args.month, args.raw_dir, args.lexicon, args.output)
