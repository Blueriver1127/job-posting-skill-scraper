"""
批次翻譯技能詞庫：將英文技能名稱翻譯成繁體中文
- 輸入：skill_lexicon_v9.xlsx
- 輸出：skill_lexicon_v10.xlsx
- 使用 deep-translator（Google 翻譯，免費、不需要帳號）
"""

import re
import time
import argparse
import pandas as pd
from deep_translator import GoogleTranslator

BATCH_SIZE = 100
SAVE_EVERY = 20  # 每 20 批存一次進度
CHECKPOINT_PATH = "/Users/annie1127/Downloads/skill_lexicon_v10_checkpoint.xlsx"
OUTPUT_PATH = "/Users/annie1127/Downloads/skill_lexicon_v10.xlsx"
INPUT_PATH = "/Users/annie1127/Downloads/skill_lexicon_v9.xlsx"


def has_chinese(text: str) -> bool:
    if not isinstance(text, str):
        return False
    return bool(re.search(r"[一-鿿]", text))


def is_english_only(text: str) -> bool:
    if not isinstance(text, str) or not text.strip():
        return True
    return not has_chinese(text)


def load_lexicon(path: str) -> pd.DataFrame:
    print(f"讀取詞庫：{path}")
    lex = pd.read_excel(path)
    print(f"總技能數：{len(lex)}")
    return lex


def find_untranslated(lex: pd.DataFrame) -> pd.DataFrame:
    zh_english = lex["Skill_Name_ZH"].apply(is_english_only)
    kw_english = lex["Keywords"].apply(is_english_only)
    mask = zh_english & kw_english
    result = lex[mask].copy()
    print(f"需要翻譯的技能數：{len(result)}")
    return result


def translate_name(translator: GoogleTranslator, name: str) -> str:
    """翻譯技能名稱，若翻譯結果仍是英文則保留原文"""
    try:
        result = translator.translate(name)
        if result and has_chinese(result):
            return result.strip()
    except Exception:
        pass
    return name


def translate_keywords(translator: GoogleTranslator, name: str, zh_name: str) -> str:
    """根據技能名稱產生中文關鍵字"""
    kw_parts = []
    if has_chinese(zh_name):
        kw_parts.append(zh_name)
    # 嘗試翻譯 keywords（用名稱作為基礎）
    try:
        variant = translator.translate(name + " skills")
        if variant and has_chinese(variant):
            cleaned = variant.replace("技能", "").replace("能力", "").strip()
            if cleaned and cleaned != zh_name:
                kw_parts.append(cleaned)
    except Exception:
        pass
    return "｜".join(kw_parts) if kw_parts else ""


def process(input_path: str, output_path: str, checkpoint_path: str,
            start_idx: int = 0, limit: int = None):
    # 嘗試載入 checkpoint
    import os
    if os.path.exists(checkpoint_path):
        lex = load_lexicon(checkpoint_path)
        print(f"從 checkpoint 繼續")
    else:
        lex = load_lexicon(input_path)

    untranslated = find_untranslated(lex)
    if limit:
        untranslated = untranslated.head(limit)

    total = len(untranslated)
    print(f"共 {total} 個需翻譯，從第 {start_idx} 個開始")

    translator = GoogleTranslator(source="en", target="zh-TW")
    updated = 0
    errors = 0

    indices = untranslated.index.tolist()

    for i, lex_idx in enumerate(indices[start_idx:], start=start_idx):
        row = lex.loc[lex_idx]
        name = str(row["Skill_Name"])

        zh_name = translate_name(translator, name)
        zh_kw = ""
        if has_chinese(zh_name):
            zh_kw = zh_name  # 至少把名稱當關鍵字

        if has_chinese(zh_name):
            lex.at[lex_idx, "Skill_Name_ZH"] = zh_name
            existing_kw = str(row.get("Keywords", "")).strip()
            if existing_kw and existing_kw.lower() != "nan":
                lex.at[lex_idx, "Keywords"] = existing_kw + "｜" + zh_kw
            else:
                lex.at[lex_idx, "Keywords"] = zh_kw
            updated += 1
        else:
            errors += 1

        # 進度顯示
        if (i + 1) % 100 == 0:
            pct = (i + 1) / total * 100
            print(f"  {i + 1}/{total} ({pct:.1f}%)，成功 {updated}，跳過 {errors}")

        # 定期存檔
        if (i + 1) % (SAVE_EVERY * BATCH_SIZE) == 0:
            lex.to_excel(checkpoint_path, index=False)
            print(f"  已存 checkpoint（第 {i + 1} 個）")

        # 避免 Google 封鎖
        time.sleep(0.05)

    lex.to_excel(output_path, index=False)
    lex.to_excel(checkpoint_path, index=False)
    print(f"\n完成！成功翻譯：{updated}，無法翻譯（保留英文）：{errors}")
    print(f"已輸出：{output_path}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", default=INPUT_PATH)
    parser.add_argument("--output", default=OUTPUT_PATH)
    parser.add_argument("--checkpoint", default=CHECKPOINT_PATH)
    parser.add_argument("--start", type=int, default=0, help="從第幾個開始（斷點續傳）")
    parser.add_argument("--limit", type=int, default=None, help="只翻譯前 N 個（測試用）")
    args = parser.parse_args()

    process(args.input, args.output, args.checkpoint, args.start, args.limit)
