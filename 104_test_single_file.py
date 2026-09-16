"""
============================================================
單檔測試腳本 — 資料還沒整理完 / 想先驗證比對品質時用
不跑 104_run_all_months.py 的 MONTH_CONFIGS 九個月迴圈，
只吃一個指定的清理後 Excel 檔案，直接呼叫該檔案裡目前最新的
比對邏輯（含最長匹配優先、v13 詞庫所有修正），跑完立刻看結果。

【執行方式】
  python3 104_test_single_file.py /path/to/cleaned_縣市_YYYYMM.xlsx

  不帶參數則用下面 INPUT_PATH 預設值。

【輸出】
  skills_output_all/_test/skills_{縣市}_{月份}_wide.xlsx
  （只輸出寬表格方便人工審閱；長表格 long parquet 不落地，只在記憶體中用來組寬表格。
   放在 _test 子資料夾，不會混進正式的九個月整合檔）

【等資料抓確定後】
  改回用 104_run_all_months.py 的 MONTH_CONFIGS 一次跑全部，
  記得把這個月份的 raw_dir / glob pattern 設定同步更新。
============================================================
"""

import os
import re
import sys
import time
import importlib.util

import pandas as pd

_SPEC = importlib.util.spec_from_file_location(
    "m104", os.path.join(os.path.dirname(os.path.abspath(__file__)), "104_run_all_months.py")
)
m104 = importlib.util.module_from_spec(_SPEC)
_SPEC.loader.exec_module(m104)

# ============================================================
# 【使用者設定區】沒有從命令列帶路徑時，用這個預設值
# ============================================================
INPUT_PATH = "/Users/annie1127/Downloads/cleaned_新北市_202502.xlsx"
OUTPUT_DIR = "/Users/annie1127/Downloads/skills_output_all/_test"
# ============================================================


def extract_month(filename: str) -> str:
    m = re.search(r"(\d{6})", os.path.basename(filename))
    return m.group(1) if m else "unknown"


def main():
    input_path = sys.argv[1] if len(sys.argv) > 1 else INPUT_PATH
    if not os.path.exists(input_path):
        print(f"⚠️ 找不到檔案：{input_path}")
        return

    county = m104.extract_county(input_path)
    month = extract_month(input_path)
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    print(f"讀取詞庫並建立自動機...")
    t0 = time.time()
    automaton = m104.load_automaton(m104.LEXICON_PATH)
    print(f"  完成，耗時 {time.time()-t0:.1f}s\n")

    print(f"讀取資料：{input_path}")
    df = pd.read_excel(input_path)
    print(f"  {len(df):,} 筆職缺，縣市={county}，月份={month}")
    print(f"  欄位：{df.columns.tolist()}\n")

    t0 = time.time()
    long_df = m104.process_county(df, automaton, county, month)
    elapsed = time.time() - t0
    avg = len(long_df) / len(df) if len(df) else 0
    print(f"比對完成：{len(df):,} 筆 -> {len(long_df):,} 筆技能記錄，"
          f"耗時 {elapsed:.1f}s，平均 {avg:.2f} 個/職缺\n")

    wide_df = m104.skills_to_wide(long_df, df)

    wide_path = os.path.join(OUTPUT_DIR, f"skills_{county}_{month}_wide.xlsx")
    wide_df.to_excel(wide_path, index=False)

    print(f"已輸出：\n  {wide_path}")


if __name__ == "__main__":
    main()
