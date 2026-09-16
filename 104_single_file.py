"""
============================================================
獨立單檔比對腳本 — 不依賴 104_run_all_months.py，可以單獨複製這一個
檔案給團隊使用
============================================================

【用途】
  只吃一個指定的清理後 Excel 檔案（一個檔案通常對應一個縣市、一個月份），
  跑完直接輸出寬表格（wide，職缺 x 9 大類技能欄位），方便人工審閱。
  目前不需要一次跑全部 9 個月時用這支就好。

【執行方式】
  python3 104_single_file.py /path/to/cleaned_縣市_YYYYMM.xlsx

  不帶參數則用下面 INPUT_PATH 預設值。

【輸出】
  skills_output_all/_single/skills_{縣市}_{月份}_wide.xlsx

【⚠️ 重要：這是獨立檔案，跟 104_run_all_months.py 的比對邏輯是兩份分開的複本】
  這支腳本把 104_run_all_months.py 的比對邏輯完整複製了一份進來，
  刻意不 import 104_run_all_months.py，這樣團隊成員只要這一個檔案就能跑，
  不用連 104_run_all_months.py 一起發。
  代價是：以後如果 104_run_all_months.py 的比對邏輯又修正（例如詞庫路徑、
  誤判修正、比對規則），這裡不會自動同步，需要手動比對兩邊、複製過來。
  等確定不會再頻繁修正比對邏輯、要固定下來一次跑全部 9 個月時，
  建議改回統一從 104_run_all_months.py import（參考舊版 104_test_single_file.py
  的寫法），避免兩份邏輯長期並存、慢慢跑掉。
============================================================
"""

import re
import os
import sys
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

# ============================================================
# 【使用者設定區】
# ============================================================
INPUT_PATH   = "/Users/annie1127/Library/CloudStorage/OneDrive-個人/WHC2/104JobData/Data2/08_2026/清洗後檔案/cleaned_高雄市_202608.xlsx"
LEXICON_PATH = "/Users/annie1127/Library/CloudStorage/OneDrive-個人/ChenRA/AI_Job_project/詞庫/詞庫skill_lexicon_v13.xlsx"
OUTPUT_DIR   = "/Users/annie1127/Library/CloudStorage/OneDrive-個人/ChenRA/AI_Job_project/測試結果檔"
# ============================================================

# 台灣製造/品質領域重要 2 字詞，jieba 預設不認識會拆開，需強制加入
_JIEBA_2CHAR_WHITELIST = {
    "製程", "量試", "試作", "送樣", "首件", "稼働", "治具",
    "備料", "備件", "標案", "採購", "詢價", "驗收",
    "輪班", "物性", "化性", "採樣",
}

# 有意義的 2 字中文技能關鍵字白名單
# （一般規則：中文 < 3 字會被過濾；此白名單的 2 字詞例外放行）
_ZH_2CHAR_SKILL_ALLOWLIST = {
    # 「業務」拿掉：常泛指「職務/工作內容」（護理業務、行政業務⋯），不是只有業務銷售的意思，
    # 裸字留著會大量誤判；「銷售」本身跟其他複合詞（業務推廣、業務開拓⋯）已經有足夠涵蓋
    "銷售", "行銷", "推廣", "開發",
    "研發", "設計", "開發", "繪圖", "建模",
    "生產", "製造", "組裝", "加工", "備料", "安裝",
    "品管", "稽核", "檢驗", "管控",
    "採購", "倉儲", "物流", "配送",
    "財務", "會計", "人資", "薪資",
    "維修", "保養", "操作", "校正",
    "翻譯", "口譯",
    # 餐飲外場（2026-08-18 補：職缺常用這種 2 字詞列點，很少寫成「X服務」全稱）
    "跑單", "擺盤", "送餐", "帶位", "倒水", "點餐", "收銀", "結帳",
    # 照顧服務／清潔衛生（2026-08-18 補：同樣是條列式短語）
    "消毒", "更衣", "沐浴",
    # 保全（2026-08-29 補：巡邏本身是安全詞，跟原本已核准的裸字風險相當）
    "巡邏",
    # 基層領班／督導（2026-09-06 補：口語化裸字，語意明確、跟其他常用詞衝突風險低）
    "督導", "班長",
    # 美髮（2026-09-06 補：之前補「洗髮服務」時漏了把裸字一併放行，
    # 導致頓號展開產生的「剪髮/染髮/洗髮」實際上還是比對不到）
    "剪髮", "染髮", "洗髮", "燙髮", "護髮",
    # 餐飲清潔（2026-09-07 補：同義詞庫清單案例）
    "洗碗", "打掃",
    # 2026-09-07 全面稽核：補齊詞庫裡已經有登記、但漏放行的 2 字詞
    "傳票", "立帳", "分錄", "良率", "沖帳",       # 舊有缺口（v13 原始 TW_ 補充項目）
    "備菜", "打餐", "舌診", "脈診", "刷手", "煎肉", "試教",  # 這次新增技能時漏放行
    # 「跟刀」拿掉了：業務/銷售語境的「跟催、跟進」也會用到這個字，
    # 跟醫療上的「跟刀」（協助醫師開刀）撞在一起，改用「手術跟刀」較安全
    # 2026-09-07 發現「清潔」這個裸字從詞庫最初就沒被放行過（一直靠「環境清潔」
    # 這種 4 字複合詞頂著，直到職缺文字只單獨寫「清潔」才發現這個舊缺口）
    "清潔", "洗車", "推銷",
}

# ============================================================
# 同義詞字典（2026-09-07 正式落地：「方法二」）
# ============================================================
# 每一組是一組意思互通的字／詞。載入詞庫時，只要某個技能的中文詞條裡
# 包含這裡任一個字，就自動把該字換成同組其他字，生成新的詞條、掛在
# 同一個技能上——不用每次手動個別去補關鍵字，之後只要詞庫新增任何
# 含有這些字的技能，也會自動套用，不用重新手動抓一次。
# 目前的清單是把 CLAUDE_同義詞庫起點清單 裡驗證過的案例整理進來，
# 之後發現新的同義詞組，直接加進這裡就好。
_SYNONYM_GROUPS = [
    {"老師", "教師"},
    # 「材料/食材」拿掉了：材料是很廣泛的字（建材、原物料、工程材料都算），
    # 食材專指食物，兩個字概括程度不對等，硬當同義詞會把非食品業的「材料測量」
    # 之類的用法誤標成食品技能。「食材準備」已經有的「材料準備」是人工個別
    # 驗證過的關鍵字，不受這裡拿掉的影響，繼續保留在該筆技能自己的 Keywords 裡。
    {"飲料", "飲品"},
    {"會客登記", "登記換證", "訪客登記"},
    {"駕駛接送", "接送任務"},
    {"車輛加油", "加油作業", "油箱加油"},
    {"打掃", "清潔"},
    {"藥品調劑", "調劑處方", "藥事調劑"},
]


def expand_synonyms(term: str) -> set:
    """回傳這個詞套用同義詞字典後，能生成的所有變體（不含自己）。"""
    variants = set()
    for group in _SYNONYM_GROUPS:
        for m in group:
            if m in term:
                for m2 in group:
                    if m2 != m:
                        variants.add(term.replace(m, m2))
    variants.discard(term)
    return variants

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


def extract_month(filename: str) -> str:
    m = re.search(r"(\d{6})", os.path.basename(filename))
    return m.group(1) if m else "unknown"


def _stem_english_text(text: str) -> str:
    return re.sub(r"[a-z]+", lambda m: _stemmer.stem(m.group()), text)


# 2026-09-06 新增：列舉展開（共用字尾省略）
# 中文職缺常見寫法「剪、染、洗服務」= 剪服務、染服務、洗服務，子字串比對機制
# 天生抓不到被標點打斷、共用字尾省略的寫法。這裡用規則式展開：偵測「A、B、C+字尾」
# （每項限 1~2 字、至少 3 項，降低誤觸機率），把展開後的完整詞附加到文字尾端，
# 讓後面的比對邏輯可以額外掃到這些字面上沒有連續出現的詞。
# 2026-09-07 補：分隔符號除了頓號，也支援句點「.」（職缺常寫「加油.洗車.清潔」
# 這種用句點列舉的寫法）——項目字元排除數字，避免跟「1.」「2.」編號清單搞混。
_ENUM_SUFFIXES = ["服務", "作業", "管理", "處理", "工作", "技巧", "能力",
                  "維護", "操作", "製作", "訓練", "分析", "規劃", "執行", "經驗"]
_ENUM_SUFFIX_PATTERN = "|".join(_ENUM_SUFFIXES)
_ENUM_SEP = "、|\\."
_ENUM_PATTERN = re.compile(
    rf"([^、，,。.\s\d]{{1,2}}(?:(?:{_ENUM_SEP})[^、，,。.\s\d]{{1,2}}){{2,4}})({_ENUM_SUFFIX_PATTERN})"
)


def expand_enumeration(text: str) -> str:
    if not isinstance(text, str) or ("、" not in text and "." not in text):
        return text
    expansions = []
    for m in _ENUM_PATTERN.finditer(text):
        heads = re.split(_ENUM_SEP, m.group(1))
        suffix = m.group(2)
        for h in heads:
            if h and not h.endswith(suffix):
                expansions.append(h + suffix)

    # 美髮業特例：「剪、染、洗服務」省略的是中間的「髮」字，不是接在字尾，
    # 一般的列舉展開救不到，額外用這條規則處理
    for m in re.finditer(rf"((?:[剪染燙護洗造](?:{_ENUM_SEP})){{1,4}}[剪染燙護洗造])(?:服務|作業)?", text):
        for ch in re.split(_ENUM_SEP, m.group(1)):
            expansions.append(ch + "髮")

    if expansions:
        return text + " " + " ".join(expansions)
    return text


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

    for w in _JIEBA_2CHAR_WHITELIST:
        jieba.add_word(w, freq=1000)

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
                    if len(kw) < 2 or (len(kw) == 2 and kw not in _ZH_2CHAR_SKILL_ALLOWLIST):
                        continue
                    terms.append((kw, True))
                    jieba.add_word(kw, freq=1000); jcount += 1
                else:
                    if len(kw) < 2:
                        continue
                    terms.append((kw.lower(), False))

        # 同義詞字典展開：套用 _SYNONYM_GROUPS，自動生成同義變體詞條
        extra_syn = []
        for term, is_zh in terms:
            if not is_zh:
                continue
            for variant in expand_synonyms(term):
                if len(variant) >= 3 or variant in _ZH_2CHAR_SKILL_ALLOWLIST:
                    extra_syn.append((variant, True))
                    jieba.add_word(variant, freq=1000); jcount += 1
        terms.extend(extra_syn)

        # 英文詞加詞幹（門檻 >=7：短詞幹如 avail/optim/activ/intern/execut
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
    且兩者對應不同技能，捨棄短詞、只保留長詞。回傳 (保留的候選, 被抑制掉的字面文字集合)
    ——後者用於跨 text_lower / text_stemmed 兩次掃描共用抑制結果。"""
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


def build_skill_id_index(automaton) -> dict:
    """建立 Skill_ID -> 技能 dict 的查找表，供職稱消歧邏輯用。"""
    index = {}
    for term in automaton.keys():
        for term_len, skill, is_zh in automaton.get(term):
            index.setdefault(skill["SKILL_ID"], skill)
    return index


# 2026-08-29 新增：「開發／設計／操作／管控」這 4 個字太籠統，沒有掛在任何技能上
# （裸字比對不到）。只有在整篇職缺完全比對不到任何技能、且文字裡真的出現這個字時，
# 才用「104職位名稱」判斷屬於哪個領域去補一個相對精確的技能——不是無中生有幫沒寫的
# 職缺編技能，是文字裡已經有這個字、只是不知道是哪種意思。判斷不出來就維持不比對。
_TITLE_DISAMBIGUATION = {
    "開發": [
        (["軟體", "韌體", "程式", "資訊", "app", "APP", "IT"], "KS120L96KMYTDJ48NRSH"),  # 軟體開發
        (["業務", "銷售"], "KS1212B6QR5SK1LSD4S4"),                                        # 商業開發
        (["產品經理", "產品企劃", "研發"], "KS1270P6SLFCFT476Y3R"),                        # 新產品開發
    ],
    "設計": [
        (["機構"], "KS1269J6YQG4J3V5YGGW"),                                                # 機構設計
        (["電子", "電機", "硬體", "ic", "IC", "韌體", "類比"], "KS121X369RKT17LSJNZX"),    # 電路設計
        (["室內", "裝潢"], "KS122SP76CJCX9T70P6B"),                                        # 室內設計
        (["美編", "平面", "視覺", "ui", "UI", "ux", "UX", "廣告"], "KS124H46Q2PP8YC1WQ06"), # 平面設計
    ],
    "操作": [
        (["作業員", "技術員", "機台", "生產", "製造", "產線"], "TW_MFG_001"),  # 機台操作
    ],
    "管控": [
        (["品管", "品保", "QA", "QC", "稽核", "檢驗"], "KS1289C6QS0TSSB4PNGG"),  # 品質管理
        (["生產", "製造", "產線", "製程"], "KS1282X64QGHTFWQ8J2Y"),             # 生產管理
    ],
}


# 2026-09-09 新增：就近文意規則。職稱消歧是用「職位名稱」這種比較間接的訊號
# 去猜籠統字屬於哪個領域，但如果籠統字緊鄰的前後文本身就有更直接的線索
# （例如「新產品之開發」），文字本身的證據比職稱推論更準，應該優先採用。
# 容許中間夾雜常見連接詞（之/的），處理「插字打斷」的狀況（跟調劑健保處方同一種問題）。
_LOCAL_CONTEXT_RULES = {
    "開發": [
        (r"產品[之的]?開發", "KS1270P6SLFCFT476Y3R"),  # 新產品開發
    ],
}


def resolve_ambiguous_by_title(texts: list, job_title: str, automaton, skill_index: dict) -> list:
    """整篇職缺比對不到任何技能時，先看籠統字附近的文字有沒有更直接的線索（就近文意），
    沒有的話才退而求其次，用「104職位名稱」判斷屬於哪個領域。"""
    combined = "".join(t for t in texts if isinstance(t, str))
    title = job_title if isinstance(job_title, str) else ""
    resolved = []
    seen_ids = set()
    resolved_words = set()

    # 第一層：就近文意（比職稱推論更直接、更準，優先採用）
    for word, rules in _LOCAL_CONTEXT_RULES.items():
        if word not in combined:
            continue
        for pattern, skill_id in rules:
            if skill_id not in skill_index or not re.search(pattern, combined):
                continue
            if skill_id not in seen_ids:
                seen_ids.add(skill_id)
                resolved.append({**skill_index[skill_id], "MATCHED_FROM": "就近文意"})
            resolved_words.add(word)  # 這個籠統字已經有更直接的答案，職稱消歧不用再猜

    # 第二層：職稱消歧（就近文意沒處理過的籠統字，才用職稱輔助判斷）
    for word, rules in _TITLE_DISAMBIGUATION.items():
        if word not in combined or word in resolved_words:
            continue
        for title_keywords, skill_id in rules:
            if not skill_id or skill_id not in skill_index:
                continue
            if any(kw in title for kw in title_keywords):
                if skill_id in seen_ids:
                    continue
                seen_ids.add(skill_id)
                resolved.append({**skill_index[skill_id], "MATCHED_FROM": "職稱消歧"})
                break  # 同一個籠統字只採用第一條符合的規則
    return resolved


def match_skills(texts: list, automaton) -> list:
    field_labels = ["職位描述", "工作技能", "工具欄"]
    seen_ids      = set()
    seen_zh_terms = set()
    results = []

    for label, raw_text in zip(field_labels, texts):
        if not isinstance(raw_text, str) or not raw_text.strip():
            continue
        raw_text = expand_enumeration(raw_text)
        text_lower = raw_text.lower()
        text_stemmed = _stem_english_text(text_lower)
        scan_texts = [text_lower] if text_lower == text_stemmed else [text_lower, text_stemmed]
        zh_boundaries = _get_zh_boundaries(text_lower)
        suppressed_texts = set()

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
    title_col = "職位名稱" if "職位名稱" in df.columns else "104職位名稱"
    skill_index = build_skill_id_index(automaton)
    records = []
    for i, row in df.iterrows():
        job_id = str(row["工作編號"])
        texts = [
            str(row.get("職位描述", "")),
            str(row.get("工作技能", "")),
            str(row.get(tool_col, "")),
        ]
        matched = match_skills(texts, automaton)
        if not matched:
            matched = resolve_ambiguous_by_title(texts, str(row.get(title_col, "")), automaton, skill_index)
        for s in matched:
            records.append({"ID": job_id, "縣市": county, "月份": month, **s})
        if (i + 1) % 5000 == 0:
            print(f"    {i+1:,} / {len(df):,} 筆...")
    return pd.DataFrame(records)


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
    desc_cols = [c for c in desc_cols if c in raw_df.columns]
    desc = raw_df[desc_cols]
    return desc.merge(agg_df, on="ID", how="left")


def main():
    input_path = sys.argv[1] if len(sys.argv) > 1 else INPUT_PATH
    if not os.path.exists(input_path):
        print(f"⚠️ 找不到檔案：{input_path}")
        return

    county = extract_county(input_path)
    month = extract_month(input_path)
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    print("讀取詞庫並建立自動機...")
    t0 = time.time()
    automaton = load_automaton(LEXICON_PATH)
    print(f"  完成，耗時 {time.time()-t0:.1f}s\n")

    print(f"讀取資料：{input_path}")
    df = pd.read_excel(input_path)
    print(f"  {len(df):,} 筆職缺，縣市={county}，月份={month}\n")

    t0 = time.time()
    long_df = process_county(df, automaton, county, month)
    elapsed = time.time() - t0
    avg = len(long_df) / len(df) if len(df) else 0
    print(f"比對完成：{len(df):,} 筆 -> {len(long_df):,} 筆技能記錄，"
          f"耗時 {elapsed:.1f}s，平均 {avg:.2f} 個/職缺\n")

    wide_df = skills_to_wide(long_df, df)

    wide_path = os.path.join(OUTPUT_DIR, f"skills_{county}_{month}_wide.xlsx")
    wide_df.to_excel(wide_path, index=False)

    print(f"已輸出：\n  {wide_path}")


if __name__ == "__main__":
    main()
