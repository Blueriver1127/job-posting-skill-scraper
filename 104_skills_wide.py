"""
============================================================
104 職缺技能比對腳本（寬表格版）
每一列 = 一個職缺，所有技能整合在同一列
方便人工審閱、直接比對職缺描述與技能分類是否正確
============================================================

【使用前準備】
1. 安裝所需套件：
   pip3 install pandas openpyxl pyarrow ahocorasick-python nltk jieba

2. 修改下方「使用者設定區」的三個路徑：
   - RAW_PATH    : 你的 104 職缺 Excel 路徑
   - LEXICON_PATH: 技能詞庫 Excel 路徑
   - OUTPUT_PATH : 輸出 Excel 檔案的儲存位置（.xlsx）

【執行方式】
   python3 104_skills_wide.py

   或指定路徑（覆蓋預設值）：
   python3 104_skills_wide.py \
       --raw    你的資料.xlsx \
       --lexicon 詞庫.xlsx \
       --output  輸出.xlsx \
       --limit  100          # 只處理前 N 筆（測試用，不加則處理全部）

【輸出欄位說明】
   ID          : 工作編號
   職位名稱      : 職缺標題
   技能數        : 比對到的技能總數
   技能_中文     : 所有技能中文名，用 ｜ 分隔
   專業技能      : Specialized Skill 類（具體技術/領域知識）
   通用技能      : Common Skill 類（溝通、問題解決等軟實力）
   技能_英文     : 所有技能英文名，用 ｜ 分隔
   職位描述      : 原始職缺描述（對照用）
   工作技能      : 原始工作技能欄位（對照用）
   電腦工具      : 原始電腦工具欄位（對照用）
============================================================
"""

import re
import argparse
import warnings
import ahocorasick
import pandas as pd
import nltk
import jieba
from nltk.stem import PorterStemmer

warnings.filterwarnings("ignore", category=DeprecationWarning)
nltk.download("wordnet", quiet=True)
nltk.download("omw-1.4", quiet=True)
jieba.setLogLevel(60)

_stemmer = PorterStemmer()


def _stem_english_text(text: str) -> str:
    """把英文單字替換成詞幹，讓 'managing' 能比對到 'management'"""
    return re.sub(r"[a-z]+", lambda m: _stemmer.stem(m.group()), text)


def _get_zh_boundaries(text: str) -> set:
    """jieba 斷詞邊界，用於防止中文詞跨詞誤抓（如「軟體操作」中的「體操」）"""
    positions = {0}
    pos = 0
    for token in jieba.cut(text):
        pos += len(token)
        positions.add(pos)
    return positions


def setup_jieba_dict(lex: pd.DataFrame):
    """將詞庫中 3 字以上的中文技能詞加入 jieba 自訂詞典（只加 3 字以上避免 2 字詞誤切）"""
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
# 【使用者設定區】修改這裡的路徑
# ============================================================
RAW_PATH     = "/Users/annie1127/Downloads/cleaned_南投縣_202604.xlsx"  # 104 職缺 Excel
LEXICON_PATH = "/Users/annie1127/Downloads/skill_lexicon_v10.xlsx"      # 技能詞庫
OUTPUT_PATH  = "/Users/annie1127/Downloads/skills_output_wide.xlsx"     # 輸出位置（Excel）
# ============================================================

SOFTWARE_CATEGORIES = {17, 370, 369, 371, 372, 373, 374, 375, 376, 377, 378, 379, 380}


def has_chinese(text: str) -> bool:
    if not isinstance(text, str):
        return False
    return bool(re.search(r"[一-鿿]", text))


def load_lexicon(path: str) -> pd.DataFrame:
    """讀取技能詞庫，展開 Keywords 欄位（每個關鍵字各建一列）"""
    lex = pd.read_excel(path)
    rows = []
    for _, row in lex.iterrows():
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

        # 英文關鍵字加入詞幹形式
        extra = []
        for term, is_zh in terms:
            if not is_zh:
                stemmed = _stemmer.stem(term)
                if stemmed != term and len(stemmed) >= 3:
                    extra.append((stemmed, False))
        terms.extend(extra)

        seen_terms = set()
        for term, is_zh in terms:
            if term in seen_terms:
                continue
            seen_terms.add(term)
            new_row = row.copy()
            new_row["_match_term"] = term
            new_row["_has_chinese"] = is_zh
            rows.append(new_row)

    result = pd.DataFrame(rows)
    result = result[result["_match_term"].str.len().fillna(0) >= 2].copy()
    result = result.reset_index(drop=True)
    return result


def build_automaton(lex: pd.DataFrame):
    """建立 Aho-Corasick 自動機"""
    A = ahocorasick.Automaton()
    for idx, row in lex.iterrows():
        term = row["_match_term"]
        if term in A:
            A.get(term).append(idx)
        else:
            A.add_word(term, [idx])
    A.make_automaton()
    return A


def _is_ascii_alnum(c: str) -> bool:
    return c.isascii() and c.isalnum()


def _is_word_boundary(text: str, start: int, end: int) -> bool:
    before_ok = (start == 0) or (not _is_ascii_alnum(text[start - 1]) and text[start - 1] != "_")
    after_ok = (end == len(text)) or (not _is_ascii_alnum(text[end]) and text[end] != "_")
    return before_ok and after_ok


def match_skills(texts: list[str], lex: pd.DataFrame, automaton) -> list[dict]:
    field_labels = ["職位描述", "工作技能", "電腦工具"]
    seen_skill_ids = set()
    results = []

    for label, raw_text in zip(field_labels, texts):
        if not isinstance(raw_text, str) or not raw_text.strip():
            continue
        text_lower = raw_text.lower()
        text_stemmed = _stem_english_text(text_lower)
        scan_texts = [text_lower] if text_lower == text_stemmed else [text_lower, text_stemmed]

        zh_boundaries = _get_zh_boundaries(text_lower)

        for scan_text in scan_texts:
            for end_idx, indices in automaton.iter(scan_text):
                term = lex.loc[indices[0], "_match_term"]
                start_idx = end_idx - len(term) + 1
                has_zh = lex.loc[indices[0], "_has_chinese"]

                if has_zh:
                    if scan_text is text_lower:
                        if start_idx not in zh_boundaries or (end_idx + 1) not in zh_boundaries:
                            continue
                else:
                    if not _is_word_boundary(scan_text, start_idx, end_idx + 1):
                        continue

                for idx in indices:
                    row = lex.loc[idx]
                    skill_id = row["Skill_ID"]
                    if skill_id in seen_skill_ids:
                        continue
                    seen_skill_ids.add(skill_id)
                    results.append({
                        "SKILL_ID": skill_id,
                        "SKILL_NAME": row["Skill_Name"],
                        "SKILL_NAME_ZH": str(row["Skill_Name_ZH"]),
                        "SKILL_TYPE": row["Skill_Type"],
                    })
    return results


def process(raw_path: str, lexicon_path: str, output_path: str, limit: int = None):
    print(f"讀取職缺資料：{raw_path}")
    df = pd.read_excel(raw_path)
    if limit:
        df = df.head(limit)
    print(f"職缺筆數：{len(df)}")

    print(f"讀取技能詞庫：{lexicon_path}")
    lex = load_lexicon(lexicon_path)
    print(f"展開後比對詞條數：{len(lex)}")
    setup_jieba_dict(lex)

    print("建立比對自動機...")
    automaton = build_automaton(lex)
    print("開始比對...")

    # 每筆職缺比對技能，結果存成 dict
    job_skills = {}
    for i, row in df.iterrows():
        job_id = str(row["工作編號"])
        texts = [
            str(row.get("職位描述", "")),
            str(row.get("工作技能", "")),
            str(row.get("電腦工具", "")),
        ]
        job_skills[job_id] = match_skills(texts, lex, automaton)

        if (i + 1) % 20 == 0:
            print(f"  已處理 {i + 1} / {len(df)} 筆")

    # 整合成寬表格：每個職缺一列
    print("\n整合成寬表格...")
    records = []
    for i, row in df.iterrows():
        job_id = str(row["工作編號"])
        skills = job_skills.get(job_id, [])

        specialized = [s["SKILL_NAME_ZH"] for s in skills if s["SKILL_TYPE"] == "Specialized Skill"]
        common      = [s["SKILL_NAME_ZH"] for s in skills if s["SKILL_TYPE"] == "Common Skill"]

        records.append({
            "ID":        job_id,
            "職位名稱":  row.get("職位名稱", ""),
            "技能數":    len(skills),
            "技能_中文": "｜".join(s["SKILL_NAME_ZH"] for s in skills),
            "專業技能":  "｜".join(specialized),
            "通用技能":  "｜".join(common),
            "技能_英文": "｜".join(s["SKILL_NAME"]    for s in skills),
            "職位描述":  row.get("職位描述", ""),
            "工作技能":  row.get("工作技能", ""),
            "電腦工具":  row.get("電腦工具", ""),
        })

    result = pd.DataFrame(records)

    # 統計
    print(f"平均每職缺技能數：{result['技能數'].mean():.1f}")
    print(f"≥5 個技能的職缺比例：{(result['技能數'] >= 5).mean() * 100:.1f}%")

    result.to_excel(output_path, index=False)
    print(f"\n【輸出完成】{output_path}")
    print("\n前 5 筆預覽：")
    print(result[["ID", "職位名稱", "技能數", "技能_中文"]].head(5).to_string())


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="104 職缺技能比對（寬表格版）")
    parser.add_argument("--raw",     default=RAW_PATH,     help="104 職缺 Excel 路徑")
    parser.add_argument("--lexicon", default=LEXICON_PATH, help="技能詞庫 xlsx 路徑")
    parser.add_argument("--output",  default=OUTPUT_PATH,  help="輸出 Excel 路徑（.xlsx）")
    parser.add_argument("--limit",   type=int, default=None, help="只處理前 N 筆（測試用）")
    args = parser.parse_args()

    process(args.raw, args.lexicon, args.output, args.limit)
