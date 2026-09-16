"""
============================================================
104 職缺技能比對腳本（長表格版）
每一列 = 一個職缺 × 一個技能
輸出格式與 Lightcast skills 檔案相同
============================================================

【使用前準備】
1. 安裝所需套件：
   pip3 install pandas openpyxl pyarrow ahocorasick-python nltk jieba

2. 修改下方「使用者設定區」的三個路徑：
   - RAW_PATH    : 你的 104 職缺 Excel 路徑
   - LEXICON_PATH: 技能詞庫 Excel 路徑
   - OUTPUT_PATH : 輸出 parquet 檔案的儲存位置

【執行方式】
   python3 104_skills_long.py

   或指定路徑（覆蓋預設值）：
   python3 104_skills_long.py \
       --raw    你的資料.xlsx \
       --lexicon 詞庫.xlsx \
       --output  輸出.parquet \
       --limit  100          # 只處理前 N 筆（測試用，不加則處理全部）

【輸出欄位說明】
   ID                    : 工作編號
   SKILL_ID              : 技能代碼（Lightcast 格式）
   SKILL_NAME            : 技能英文名
   SKILL_NAME_ZH         : 技能中文名
   SKILL_TYPE            : Specialized Skill / Common Skill
   SKILL_CATEGORY        : 大類代碼
   SKILL_CATEGORY_NAME   : 大類名稱
   SKILL_SUBCATEGORY     : 小類代碼
   SKILL_SUBCATEGORY_NAME: 小類名稱
   IS_SOFTWARE           : 是否為軟體/IT 技能
   MATCHED_FROM          : 從哪個欄位比對到（職位描述/工作技能/電腦工具）
============================================================
"""

import re
import argparse
import ahocorasick
import pandas as pd
import nltk
import jieba
from nltk.stem import PorterStemmer

# 第一次執行時自動下載詞幹資料
nltk.download("wordnet", quiet=True)
nltk.download("omw-1.4", quiet=True)
jieba.setLogLevel(60)  # 關閉 jieba 載入訊息

_stemmer = PorterStemmer()


def _stem_english_text(text: str) -> str:
    """把文字中的英文單字全部替換成詞幹（中文不受影響）
    例：'Managing RF Module Development' → 'manag rf modul develop'
    用途：讓 'managing' 能比對到詞庫的 'management'（兩者詞幹都是 'manag'）
    """
    return re.sub(r"[a-z]+", lambda m: _stemmer.stem(m.group()), text)


def _get_zh_boundaries(text: str) -> set:
    """
    用 jieba 斷詞，取得所有合法的詞段邊界位置（字元索引）。
    用途：避免中文技能詞跨越斷詞邊界被誤抓。
    例：「軟體操作」→ jieba 切成 [軟體, 操作] → 邊界為 {0, 2, 4}
        「體操」從位置 1 到 3，1 不是邊界 → 拒絕比對 ✓
        「體操運動」→ jieba 切成 [體操, 運動] → 邊界為 {0, 2, 4}
        「體操」從位置 0 到 2，兩端都是邊界 → 接受 ✓
    """
    positions = {0}
    pos = 0
    for token in jieba.cut(text):
        pos += len(token)
        positions.add(pos)
    return positions


def setup_jieba_dict(lex: pd.DataFrame):
    """
    將詞庫中 3 字以上的中文技能詞加入 jieba 自訂詞典。
    只加 3 字以上是因為 2 字詞（如「體操」）若加入會導致
    在「軟體操作」等複合詞中被錯誤切出。
    """
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
OUTPUT_PATH  = "/Users/annie1127/Downloads/skills_output_long.parquet"  # 輸出位置
# ============================================================

# IT 相關 Category_Code，這些技能的 IS_SOFTWARE = True
SOFTWARE_CATEGORIES = {17, 370, 369, 371, 372, 373, 374, 375, 376, 377, 378, 379, 380}


def has_chinese(text: str) -> bool:
    if not isinstance(text, str):
        return False
    return bool(re.search(r"[一-鿿]", text))


def load_lexicon(path: str) -> pd.DataFrame:
    """
    讀取技能詞庫，並展開 Keywords 欄位。
    每個技能可能有多個比對關鍵字（用 ｜ 分隔），
    每個關鍵字各建一列，都指向同一個 Skill_ID。
    """
    lex = pd.read_excel(path)

    rows = []
    for _, row in lex.iterrows():
        terms = []  # (比對詞, 是否為中文)

        # 優先用 Skill_Name_ZH（若有中文）
        zh = str(row.get("Skill_Name_ZH", "")).strip()
        if has_chinese(zh):
            terms.append((zh, True))
        else:
            # 沒有中文翻譯則用英文原名（小寫）
            en = str(row.get("Skill_Name", "")).strip().lower()
            if len(en) >= 2:
                terms.append((en, False))

        # 展開 Keywords 欄位（｜分隔的多個關鍵字）
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

        # 英文關鍵字額外加入詞幹形式
        # 例：'management' → 也加入 'manag'，讓 'managing' 能比對到
        extra = []
        for term, is_zh in terms:
            if not is_zh:
                stemmed = _stemmer.stem(term)
                if stemmed != term and len(stemmed) >= 3:
                    extra.append((stemmed, False))
        terms.extend(extra)

        # 去重，每個比對詞建一列
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
    """建立 Aho-Corasick 自動機，一次掃描即可比對所有關鍵字（比逐一比對快 100 倍以上）"""
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
    """判斷是否為 ASCII 英數字（不包含中文，因為中文的 isalnum() 也回傳 True）"""
    return c.isascii() and c.isalnum()


def _is_word_boundary(text: str, start: int, end: int) -> bool:
    """
    英文詞需確認前後是單詞邊界，避免 'C' 匹配到 'CPU' 中的 C。
    注意：中文字不視為邊界（例如 'ISO文件管理' 中 ISO 後面的「文」不阻擋比對）
    """
    before_ok = (start == 0) or (not _is_ascii_alnum(text[start - 1]) and text[start - 1] != "_")
    after_ok = (end == len(text)) or (not _is_ascii_alnum(text[end]) and text[end] != "_")
    return before_ok and after_ok


def match_skills(texts: list[str], lex: pd.DataFrame, automaton) -> list[dict]:
    """對三個欄位的文字內容進行技能比對，回傳該職缺的所有技能"""
    field_labels = ["職位描述", "工作技能", "電腦工具"]
    seen_skill_ids = set()
    results = []

    for label, raw_text in zip(field_labels, texts):
        if not isinstance(raw_text, str) or not raw_text.strip():
            continue
        text_lower = raw_text.lower()
        # 詞幹版本：讓 'managing' 能比對到 'management'（兩者詞幹同為 'manag'）
        text_stemmed = _stem_english_text(text_lower)
        # 若文字無英文，詞幹版與原文相同，只需掃一次
        scan_texts = [text_lower] if text_lower == text_stemmed else [text_lower, text_stemmed]
        # jieba 斷詞邊界（用於防止中文詞跨詞邊界誤抓，如「軟體操作」中的「體操」）
        zh_boundaries = _get_zh_boundaries(text_lower)

        for scan_text in scan_texts:
            for end_idx, indices in automaton.iter(scan_text):
                term = lex.loc[indices[0], "_match_term"]
                start_idx = end_idx - len(term) + 1
                has_zh = lex.loc[indices[0], "_has_chinese"]

                if has_zh:
                    # 中文詞：確認起止位置都是 jieba 斷詞邊界
                    # 只對原始文字做邊界檢查（詞幹版只處理英文，中文不變）
                    if scan_text is text_lower:
                        if start_idx not in zh_boundaries or (end_idx + 1) not in zh_boundaries:
                            continue
                else:
                    # 英文詞：確認前後不是英文字母（避免 C 比對到 CPU）
                    if not _is_word_boundary(scan_text, start_idx, end_idx + 1):
                        continue

                for idx in indices:
                    row = lex.loc[idx]
                    skill_id = row["Skill_ID"]
                    if skill_id in seen_skill_ids:
                        continue
                    seen_skill_ids.add(skill_id)

                    try:
                        cat_code = int(row["Category_Code"])
                    except (ValueError, TypeError):
                        cat_code = 0
                    is_software = cat_code in SOFTWARE_CATEGORIES

                    results.append({
                        "SKILL_ID": skill_id,
                        "SKILL_NAME": row["Skill_Name"],
                        "SKILL_NAME_ZH": row["Skill_Name_ZH"],
                        "SKILL_TYPE": row["Skill_Type"],
                        "SKILL_CATEGORY": str(row["Category_Code"]),
                        "SKILL_CATEGORY_NAME": row["Category_Name"],
                        "SKILL_SUBCATEGORY": str(row["Subcategory_Code"]),
                        "SKILL_SUBCATEGORY_NAME": row["Subcategory_Name"],
                        "IS_SOFTWARE": is_software,
                        "MATCHED_FROM": label,
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

    records = []
    for i, row in df.iterrows():
        job_id = str(row["工作編號"])
        texts = [
            str(row.get("職位描述", "")),
            str(row.get("工作技能", "")),
            str(row.get("電腦工具", "")),
        ]
        skills = match_skills(texts, lex, automaton)
        for s in skills:
            records.append({"ID": job_id, **s})

        if (i + 1) % 20 == 0:
            print(f"  已處理 {i + 1} / {len(df)} 筆，累計技能記錄：{len(records)}")

    result = pd.DataFrame(records)
    print(f"\n總技能記錄數：{len(result)}")
    if len(df) > 0:
        print(f"平均每職缺技能數：{len(result) / len(df):.1f}")

    per_job = result.groupby("ID").size()
    print(f"\n技能數分布：")
    print(f"  最少：{per_job.min()}")
    print(f"  最多：{per_job.max()}")
    print(f"  中位數：{per_job.median():.1f}")
    print(f"  ≥5 個技能的職缺比例：{(per_job >= 5).mean() * 100:.1f}%")

    result.to_parquet(output_path, index=False)
    print(f"\n【輸出完成】{output_path}")
    print("\n前 5 筆預覽：")
    print(result.head(5)[["ID", "SKILL_NAME", "SKILL_NAME_ZH", "SKILL_TYPE", "MATCHED_FROM"]].to_string())


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="104 職缺技能比對（長表格版）")
    parser.add_argument("--raw",     default=RAW_PATH,     help="104 職缺 Excel 路徑")
    parser.add_argument("--lexicon", default=LEXICON_PATH, help="技能詞庫 xlsx 路徑")
    parser.add_argument("--output",  default=OUTPUT_PATH,  help="輸出 parquet 路徑")
    parser.add_argument("--limit",   type=int, default=None, help="只處理前 N 筆（測試用）")
    args = parser.parse_args()

    process(args.raw, args.lexicon, args.output, args.limit)
