# 104 Job Skill Auto-Matching System

An [Aho-Corasick](https://en.wikipedia.org/wiki/Aho%E2%80%93Corasick_algorithm) multi-pattern string matching pipeline that automatically maps Chinese job posting text from Taiwan's 104 Job Bank to the [Lightcast Open Skills](https://lightcast.io/open-skills) taxonomy, producing a per-posting skill list (both a long-format table and a wide-format table organized by 9 top-level skill categories).

**Data source**: real job postings scraped from Taiwan's 104 Job Bank, across multiple counties and months. The skill lexicon is built on Lightcast's English skill taxonomy, extended with Traditional Chinese translations and locally-added skills (prefixed `TW_`) for occupations underrepresented in the original US-centric taxonomy — food service, security, elder care, childcare, occupational health and safety, construction site supervision, pharmacy, and more.

---

## What it does

1. Loads the skill lexicon (Excel) and builds an Aho-Corasick automaton from the `Skill_Name_ZH` / `Keywords` columns
2. Uses `jieba` for Chinese word segmentation (to find valid word boundaries) and a Porter Stemmer for English word-form variation
3. Scans three fields of each job posting: job description, required skills, and required tools
4. Applies a layered set of matching rules (see below) to resolve raw substring hits into clean, specific skill tags
5. Outputs a long table (posting × skill) and a wide table (posting × 9 skill-category columns)

## Matching logic (see `104_single_file.py`)

Naive substring matching on Chinese text runs into a surprising number of failure modes. This pipeline layers several rules on top of the base Aho-Corasick match to handle them:

- **Longest-match-wins**: when a short term and a longer term that fully contains it both match the same span but point to different skills, only the longer (more specific) match is kept — e.g. "development" alone is suppressed in favor of "software development" when both are present.
- **Word-boundary checks**: Chinese terms must align with `jieba` segmentation boundaries; English terms must not be flanked by alphanumeric characters, preventing false substring hits (e.g. "sql" inside "consultant").
- **English stemming threshold**: only stems of length ≥7 are indexed, to avoid common English words whose short stems collide with unrelated Lightcast entries (e.g. `availability` stemming to `avail`, which previously matched an unrelated healthcare software product named "Availity").
- **Enumeration expansion**: Chinese job postings often use a shared-suffix shorthand — "cut, dye, wash service" implicitly meaning "cut service, dye service, wash service" — that plain substring matching can't recover. A rule-based expander detects this pattern (supporting both Chinese enumeration commas and periods as separators) and expands it before matching.
- **Synonym dictionary** (`_SYNONYM_GROUPS`): declaring a synonym group (e.g. "teacher (formal)" / "teacher (colloquial)") automatically generates matching variants for every existing skill entry that contains either word — no need to manually patch keywords one at a time.
- **Title-based disambiguation** (`resolve_ambiguous_by_title`): a handful of Chinese words (roughly: "development", "design", "operation", "control") are too broad to safely tag to any single skill on their own. Only when a posting has zero matches from the main pipeline, and the posting title gives a strong domain signal, is one of these words resolved to a specific skill — the model never invents a skill for text that isn't there.
- **Local-context priority**: before falling back to title-based disambiguation, the pipeline first checks whether the ambiguous word has a more direct clue immediately adjacent in the text itself (tolerating inserted connective particles like "of"/"'s"). Direct textual evidence is preferred over inference from the job title whenever both are available.

## File structure

| File | Status | Description |
|---|---|---|
| `104_single_file.py` | ✅ **Primary, actively used** | Self-contained, no internal imports required. Takes one county/month's cleaned Excel file and outputs the wide-format skill match result |
| `104_run_all_months.py` | ✅ Main batch pipeline | Runs all counties across all months in one pass. Contains matching logic largely mirroring `104_single_file.py`, though the two currently require manual syncing |
| `104_test_single_file.py` | Reference only | An earlier single-file test wrapper that imports `104_run_all_months.py`; functionally overlapping with the two files above |
| `fix_existing_output.py` | Reference only | Post-hoc column-level patching of already-generated long-format parquet files (dedup, short-token removal) without re-running the match |
| `104_batch_monthly.py`, `104_skills_long.py`, `104_skills_wide.py` | ⚠️ **Deprecated** | Pre-dates the current lexicon version; wide-format output still uses the old "specialized/common skill" binary split rather than the current 9-category system. Logic has diverged from the active pipeline |
| `repair_202601_wide.py` | ⚠️ One-off script | Written to repair corrupted output for a specific month (Jan 2026); not a general-purpose tool |
| `translate_lexicon.py` | ⚠️ Historical | Used to machine-translate an earlier lexicon version (v9→v10); the lexicon has since progressed to v13 |

> **In practice, only `104_single_file.py` is needed.** The other files are kept for development history; the deprecated/one-off ones are candidates for future cleanup.

## Setup

```bash
pip install pandas openpyxl pyarrow ahocorasick-python nltk jieba
```

## Usage

```bash
python3 104_single_file.py /path/to/cleaned_county_YYYYMM.xlsx
```

If no argument is given, the script falls back to the `INPUT_PATH` default defined at the top of the file. Before running, update the three paths in the "user settings" section at the top:

```python
INPUT_PATH   = "..."   # cleaned job posting Excel file
LEXICON_PATH = "..."   # skill lexicon Excel file
OUTPUT_DIR   = "..."   # output directory
```

Output: `{OUTPUT_DIR}/skills_{county}_{month}_wide.xlsx` — a wide-format table (one row per posting, one column per skill category) for easy manual review.

## Skill lexicon

The lexicon is an Excel file (columns include `Skill_ID`, `Skill_Name`, `Skill_Name_ZH`, `Skill_Category`, `Keywords`, etc.), built on Lightcast's ~26,500 base skill entries, extended with:
- Traditional Chinese translations (`Skill_Name_ZH`, `Keywords` columns)
- Locally-added Taiwan-specific skills (`Skill_ID` prefixed `TW_`), covering domains the base taxonomy under-represents: occupational health and safety, food service, security, elder/childcare, education tutoring, construction site supervision, pharmacy, and more

The lexicon file itself is not included in this repository and must be supplied separately via `LEXICON_PATH`.

## Known limitations

- **Enumeration patterns beyond commas and periods, or with no punctuation at all**: the enumeration-expansion rule currently handles Chinese enumeration commas and periods; some real postings omit punctuation entirely and aren't caught.
- **Typos in the source data**: there is no fuzzy/typo-tolerant matching; literal spelling errors in postings cause missed matches.
- **9-category reclassification**: the current 9 top-level skill categories (`Skill_Category` column) don't always align well with what a skill's name actually describes. A redesigned category system has been finalized but not yet applied across the full lexicon.
- **Completeness has not been formally validated**: progress is currently tracked via a proxy metric (percentage of postings with zero matched skills). A gold-label human-annotated evaluation set has not yet been completed, so true precision/recall is not yet known.

## Project status

Actively iterating. The pipeline has been tested against real data from multiple Taiwan counties, with each new dataset surfacing additional lexicon gaps that get folded back into the matching rules and skill dictionary. Detailed change history is tracked in internal working documents not included in this repository.
