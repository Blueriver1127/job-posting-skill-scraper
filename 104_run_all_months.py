"""
============================================================
一鍵跑完全部月份的技能比對腳本
每個月份輸出一個資料夾，內含各縣市個別檔案 + 整合檔
============================================================

【輸出結構】
  OUTPUT_BASE/
    202502/
      skills_Changhua_202502_long.parquet
      skills_Changhua_202502_wide.xlsx
      ...（各縣市）
      skills_202502_ALL_long.parquet   ← 整合檔
      skills_202502_ALL_wide.xlsx      ← 整合檔
    202503/
      ...
    （共 9 個月份資料夾）

【執行方式】
  python3 104_run_all_months.py

【注意】約需 6–8 小時，建議睡前開始跑
============================================================
"""

import re
import os
import glob
import time
import warnings
import jieba
import nltk
import ahocorasick
import pandas as pd
from nltk.stem import PorterStemmer

warnings.filterwarnings("ignore")
nltk.download("wordnet", quiet=True)
nltk.download("omw-1.4", quiet=True)
jieba.setLogLevel(60)

# 台灣製造/品質領域重要 2 字詞，jieba 預設不認識會拆開，需強制加入
_JIEBA_2CHAR_WHITELIST = {
    "製程", "量試", "試作", "送樣", "首件", "稼働", "治具",
    "備料", "備件", "標案", "採購", "詢價", "驗收",
    "輪班", "物性", "化性", "採樣",
}

# 有意義的 2 字中文技能關鍵字白名單
# （一般規則：中文 < 3 字會被過濾；此白名單的 2 字詞例外放行）
_ZH_2CHAR_SKILL_ALLOWLIST = {
    # 業務/銷售
    "銷售", "業務", "行銷", "推廣", "開發",
    # 工程/研發
    "研發", "設計", "開發", "繪圖", "建模",
    # 生產/製造/安裝
    "生產", "製造", "組裝", "加工", "備料", "安裝",
    # 品質/管理
    "品管", "稽核", "檢驗", "管控",
    # 採購/物流
    "採購", "倉儲", "物流", "配送",
    # 財務/人資
    "財務", "會計", "人資", "薪資",
    # 維修/技術
    "維修", "保養", "操作", "校正",
    # 翻譯/溝通
    "翻譯", "口譯",
}

# ============================================================
# 【使用者設定區】
# ============================================================
BASE_DIR     = "/Users/annie1127/Library/CloudStorage/OneDrive-個人/RA2024/104_JobData_new"
LEXICON_PATH = "/Users/annie1127/Downloads/skill_lexicon_v13.xlsx"
OUTPUT_BASE  = "/Users/annie1127/Downloads/skills_output_all"
# ============================================================

# 各月份的設定：(月份標籤, 資料夾, glob 樣式)
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

SOFTWARE_CATEGORIES = {17, 370, 369, 371, 372, 373, 374, 375, 376, 377, 378, 379, 380}
_stemmer = PorterStemmer()


def has_chinese(text: str) -> bool:
    return bool(re.search(r"[一-鿿]", str(text)))


def extract_county(filename: str) -> str:
    """從各種格式的檔名中擷取縣市名稱"""
    name = os.path.basename(filename)
    name = name.replace("cleaned_", "").replace(".xlsx", "")
    name = re.sub(r"_combined(\.csv)?(_\d{6})?$", "", name)
    name = re.sub(r"_\d{6}$", "", name)
    return name


def _stem_english_text(text: str) -> str:
    return re.sub(r"[a-z]+", lambda m: _stemmer.stem(m.group()), text)


def _get_zh_boundaries(text: str) -> set:
    positions = {0}
    pos = 0
    for token in jieba.cut(text):
        pos += len(token)
        positions.add(pos)
    return positions


def _is_ascii_alnum(c: str) -> bool:
    return c.isascii() and c.isalnum()


def _is_word_boundary(text: str, start: int, end: int) -> bool:
    before_ok = (start == 0) or (not _is_ascii_alnum(text[start - 1]) and text[start - 1] != "_")
    after_ok = (end == len(text)) or (not _is_ascii_alnum(text[end]) and text[end] != "_")
    return before_ok and after_ok


def load_automaton(lexicon_path: str):
    """載入詞庫、設定 jieba 自訂詞典、建立 Aho-Corasick 自動機"""
    print(f"  讀取詞庫：{lexicon_path}")
    lex = pd.read_excel(lexicon_path)

    # 先把 2 字白名單強制加入 jieba（這些詞 jieba 預設會拆開）
    for w in _JIEBA_2CHAR_WHITELIST:
        jieba.add_word(w, freq=1000)

    # jieba 自訂詞典（3字以上避免誤切）
    jcount = 0
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
            "SKILL_CAT9":             str(row["Skill_Category"]) if pd.notna(row.get("Skill_Category")) else "Unclassified",
            "SKILL_CATEGORY":         str(row["Category_Code"]),
            "SKILL_CATEGORY_NAME":    str(row["Category_Name"]),
            "SKILL_SUBCATEGORY":      str(row["Subcategory_Code"]),
            "SKILL_SUBCATEGORY_NAME": str(row["Subcategory_Name"]),
            "IS_SOFTWARE":            cat_code in SOFTWARE_CATEGORIES,
        }

        terms = []
        zh = str(row.get("Skill_Name_ZH", "")).strip()
        # 中文技能名：≥3 字，或剛好 2 字且在允許白名單內（避免「標」「電」等泛用詞）
        if has_chinese(zh) and (len(zh) >= 3 or zh in _ZH_2CHAR_SKILL_ALLOWLIST):
            terms.append((zh, True))
            jieba.add_word(zh, freq=1000); jcount += 1
        elif not has_chinese(zh):
            en = str(row.get("Skill_Name", "")).strip().lower()
            if len(en) >= 2:
                terms.append((en, False))

        kw_str = str(row.get("Keywords", "")).strip()
        if kw_str and kw_str.lower() != "nan":
            for kw in kw_str.split("｜"):
                kw = kw.strip()
                if has_chinese(kw):
                    # 中文關鍵字：≥3 字，或剛好 2 字且在允許白名單內
                    if len(kw) < 2 or (len(kw) == 2 and kw not in _ZH_2CHAR_SKILL_ALLOWLIST):
                        continue
                    terms.append((kw, True))
                    jieba.add_word(kw, freq=1000); jcount += 1
                else:
                    if len(kw) < 2:
                        continue
                    terms.append((kw.lower(), False))

        # 英文詞加詞幹（門檻拉到 >=7：短詞幹如 avail/optim/activ/intern/execut
        # 太容易撞到職缺文字裡完全不相關的常見英文字，門檻以下寧可漏判不誤判）
        extra = []
        for term, is_zh in terms:
            if not is_zh:
                stemmed = _stemmer.stem(term)
                if stemmed != term and len(stemmed) >= 7:
                    extra.append((stemmed, False))
        terms.extend(extra)

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

    print(f"  詞條數：{len(term_to_entries):,}，jieba 自訂詞：{jcount:,}")
    return A


def _suppress_shorter_matches(candidates: list, scan_text: str):
    """最長匹配優先：若某段文字同時被短詞與完全涵蓋它的長詞命中，
    且兩者對應不同技能，捨棄短詞、只保留長詞（讓比對結果更 specific，
    例如同時命中「開發」與「軟體開發」時只留「軟體開發」）。
    回傳 (保留的候選, 被抑制掉的字面文字集合)——後者用於跨 text_lower / text_stemmed
    兩次掃描共用抑制結果，避免長詞因詞幹化而在另一次掃描裡對不上，短詞因此漏網。"""
    kept = []
    suppressed_texts = set()
    for i, a in enumerate(candidates):
        covered = False
        for j, b in enumerate(candidates):
            if i == j or a["skill"]["SKILL_ID"] == b["skill"]["SKILL_ID"]:
                continue
            if b["term_len"] <= a["term_len"]:
                continue
            if b["start"] <= a["start"] and b["end"] >= a["end"]:
                covered = True
                break
        if covered:
            suppressed_texts.add(scan_text[a["start"]:a["end"] + 1])
        else:
            kept.append(a)
    return kept, suppressed_texts


def match_skills(texts: list, automaton) -> list:
    field_labels = ["職位描述", "工作技能", "工具欄"]
    seen_ids      = set()
    seen_zh_terms = set()  # 同一中文詞只觸發一個技能，避免「追蹤」同時匹配 Ftrace 和 Ltrace
    results = []

    for label, raw_text in zip(field_labels, texts):
        if not isinstance(raw_text, str) or not raw_text.strip():
            continue
        text_lower = raw_text.lower()
        text_stemmed = _stem_english_text(text_lower)
        scan_texts = [text_lower] if text_lower == text_stemmed else [text_lower, text_stemmed]
        zh_boundaries = _get_zh_boundaries(text_lower)
        suppressed_texts = set()  # 跨這個欄位的兩次掃描共用

        for scan_text in scan_texts:
            candidates = []
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
                    candidates.append({
                        "start": start_idx, "end": end_idx,
                        "term_len": term_len, "skill": skill, "is_zh": is_zh,
                    })

            kept, newly_suppressed = _suppress_shorter_matches(candidates, scan_text)
            suppressed_texts |= newly_suppressed

            for cand in kept:
                matched_term = scan_text[cand["start"]:cand["end"] + 1]
                if matched_term in suppressed_texts:
                    continue
                if cand["is_zh"]:
                    if matched_term in seen_zh_terms:
                        continue
                    seen_zh_terms.add(matched_term)
                skill_id = cand["skill"]["SKILL_ID"]
                if skill_id in seen_ids:
                    continue
                seen_ids.add(skill_id)
                results.append({**cand["skill"], "MATCHED_FROM": label})
    return results


def process_county(df: pd.DataFrame, automaton, county: str, month: str) -> pd.DataFrame:
    tool_col = "擅長工具" if "擅長工具" in df.columns else "電腦工具"
    records = []
    for i, row in df.iterrows():
        job_id = str(row["工作編號"])
        texts = [
            str(row.get("職位描述", "")),
            str(row.get("工作技能", "")),
            str(row.get(tool_col, "")),
        ]
        for s in match_skills(texts, automaton):
            records.append({"ID": job_id, "縣市": county, "月份": month, **s})
        if (i + 1) % 5000 == 0:
            print(f"    {i+1:,} / {len(df):,} 筆...")
    return pd.DataFrame(records)


# 9 大技能類別：英文名稱 → 中文欄位名稱
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


def skills_to_wide(long_df: pd.DataFrame, raw_df: pd.DataFrame) -> pd.DataFrame:
    raw_df = raw_df.copy()
    raw_df["ID"] = raw_df["工作編號"].astype(str)
    tool_col = "擅長工具" if "擅長工具" in raw_df.columns else "電腦工具"
    title_col = "職位名稱" if "職位名稱" in raw_df.columns else "104職位名稱"
    title_code_col = "職位名稱碼" if "職位名稱碼" in raw_df.columns else "104職位名稱碼"

    def agg(group):
        result = {
            "技能數":    len(group),
            "技能_中文": "｜".join(group["SKILL_NAME_ZH"]),
        }
        for en_cat, zh_col in _CAT9_ZH.items():
            skills_in_cat = "｜".join(
                r["SKILL_NAME_ZH"] for _, r in group.iterrows()
                if r.get("SKILL_CAT9", "Unclassified") == en_cat
            )
            result[zh_col] = skills_in_cat
        return pd.Series(result)

    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        agg_df = long_df.groupby("ID").apply(agg, include_groups=False).reset_index()

    desc_cols = ["ID", "資料月份", "刊登日期", "工作角色", title_col, title_code_col,
                 "職位描述", "工作技能", tool_col]
    desc_cols = [c for c in desc_cols if c in raw_df.columns]  # 欄位不存在就跳過，不同月份 schema 可能不一致
    desc = raw_df[desc_cols]
    return desc.merge(agg_df, on="ID", how="left")


def run_month(month: str, raw_dir: str, pattern: str, automaton, output_dir: str):
    os.makedirs(output_dir, exist_ok=True)
    files = sorted(glob.glob(os.path.join(raw_dir, pattern)))
    if not files:
        print(f"  ⚠️ 找不到檔案：{os.path.join(raw_dir, pattern)}")
        return

    print(f"\n【{month}】{len(files)} 個縣市，輸出至 {output_dir}")
    all_long, all_wide = [], []
    t_start = time.time()

    for fpath in files:
        county = extract_county(fpath)
        df = pd.read_excel(fpath)
        t1 = time.time()

        long_df = process_county(df, automaton, county, month)
        wide_df = skills_to_wide(long_df, df)
        elapsed = time.time() - t1

        avg = len(long_df) / len(df) if len(df) > 0 else 0
        print(f"  {county}: {len(df):,} 筆 → {len(long_df):,} 技能記錄 (平均 {avg:.1f} 個, {elapsed:.1f}s)")

        # 各縣市個別輸出
        long_df.to_parquet(os.path.join(output_dir, f"skills_{county}_{month}_long.parquet"), index=False)
        wide_df.to_excel(os.path.join(output_dir, f"skills_{county}_{month}_wide.xlsx"), index=False)

        all_long.append(long_df)
        all_wide.append(wide_df)

    # 整合輸出
    combined_long = pd.concat(all_long, ignore_index=True)
    combined_wide = pd.concat(all_wide, ignore_index=True)
    combined_long.to_parquet(os.path.join(output_dir, f"skills_{month}_ALL_long.parquet"), index=False)
    combined_wide.to_excel(os.path.join(output_dir, f"skills_{month}_ALL_wide.xlsx"), index=False)

    total_elapsed = time.time() - t_start
    print(f"  ✓ {month} 完成：{len(combined_long):,} 筆技能記錄，耗時 {total_elapsed/60:.1f} 分鐘")
    return len(combined_long)


if __name__ == "__main__":
    print("=" * 60)
    print("104 全月份技能比對 — 開始")
    print("=" * 60)

    # 建立自動機（只建一次，所有月份共用）
    print("\n建立比對自動機...")
    t0 = time.time()
    automaton = load_automaton(LEXICON_PATH)
    print(f"完成，耗時 {time.time()-t0:.1f} 秒\n")

    grand_total = 0
    session_start = time.time()

    for month, rel_dir, pattern in MONTH_CONFIGS:
        raw_dir = os.path.join(BASE_DIR, rel_dir)
        output_dir = os.path.join(OUTPUT_BASE, month)
        result = run_month(month, raw_dir, pattern, automaton, output_dir)
        if result:
            grand_total += result

    total_min = (time.time() - session_start) / 60
    print("\n" + "=" * 60)
    print(f"全部完成！")
    print(f"  總技能記錄：{grand_total:,} 筆")
    print(f"  總耗時：{total_min:.1f} 分鐘（{total_min/60:.1f} 小時）")
    print(f"  輸出資料夾：{OUTPUT_BASE}")
    print("=" * 60)
