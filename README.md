# Content Match - Excel Column Comparison Tool

## Overview
This tool compares content between two Excel files (sheet1.xlsx and sheet2.xlsx) on a word-by-word basis, using **sheet1 as the 100% baseline**. The comparison is performed by matching records using composite keys (Basepack + Account) and analyzing text content across multiple columns.

## How the Matching Logic Works

### 1. File Structure
- **sheet1.xlsx** - Base file (reference, 100%)
- **sheet2.xlsx** - Comparison file
- Both files must contain: `Basepack`, `Account`, and comparison columns (Title, Feature bullets, Product description)

### 2. Composite Key Matching
Records are matched between files using: `Basepack + Account`
- Only matching records are compared
- Records not found in sheet2 are flagged as "NOT FOUND in Sheet2"

### 3. Word-by-Word Matching Algorithm

#### Core Principle
**Sheet1 column word count = 100%**

The matching is based on counting individual words, not character similarity.

#### How It Works

**Step 1: Normalize Text**
- **Remove special characters:**
  - Zero-width spaces: `\u200B`, `\u200C`, `\u200D`
  - Directional marks: `\u200E`, `\u200F`, `\u202A-\u202E`
  - Word joiners: `\u2060`, `\u2066-\u2069`
  - BOM (Byte Order Mark): `\uFEFF`
- **Decode Excel escapes:** `_x000D_` (carriage return), `_x000A_` (line feed), etc.
- **Normalize unicode characters:** NFKC normalization
- **Unify punctuation:**
  - Hyphens/dashes: `\u2010`, `\u2011`, `\u2012`, `\u2013`, `\u2014`, `\u2212` → `-`
  - Smart quotes: `\u2018`, `\u2019` → `'`
  - Smart double quotes: `\u201C`, `\u201D` → `"`
- **Collapse whitespace:** All unicode whitespace (`\u00A0`, `\u1680`, `\u2000-\u200A`, etc.) → single space
- **Normalize punctuation spacing:** Remove spaces before `,.;:!?` and add space after
- **Convert to lowercase**
- **Trim leading/trailing quotes and whitespace**

**Step 2: Extract Words**
- Use regex `\w+` to extract all words (alphanumeric sequences)
- Create word frequency counters for both columns

**Step 3: Calculate Match Percentage**
```
matched_words = sum of minimum counts for each word present in sheet1
base_word_count = total words in sheet1 column

match_percentage = (matched_words / base_word_count) × 100%
```

**Step 4: Identify Extra Content**
- If sheet2 has ALL words from sheet1 (100% match)
- AND sheet2 has additional words beyond sheet1
- Then mark that column name in "Columns with extra content"

#### Example

**Sheet1 - Title Column:**
```
"Premium Wireless Headphones Black"
Words: [premium, wireless, headphones, black]
Word count: 4 (this is 100%)
```

**Sheet2 - Title Column:**
```
"Premium Wireless Headphones Black Bluetooth"
Words: [premium, wireless, headphones, black, bluetooth]
Word count: 5
```

**Result:**
- Matched words: 4 (all sheet1 words found)
- Match percentage: (4 / 4) × 100 = **100%**
- Has extra: Yes (bluetooth)
- Column marked in: **"Columns with extra content: Title"**

### 4. Comparison Outputs

The tool generates `comparison_result.xlsx` with the following columns:

#### Key Columns
- **Basepack** - Product identifier
- **Account** - Account identifier
- **Remark** - Lists column names with mismatches
- **Missing attributes** - Columns that are empty in either sheet
- **MatchedColumns** - Columns with exact matches

#### Overall Metrics
- **MatchPercentWordLevel** - Overall match % across all columns
- **Feature bullet overall** - Match % for Feature bullet 1-6 combined
- **Columns with extra content** - Column names where sheet2 has 100% match + extra words

#### Per-Column Metrics
For each column (Title, Feature bullet 1-6, Product description):
- **{Column Name} MatchPercentWordLevel** - Individual column match %

### 5. Match Percentage Rules

#### 100% Match
- All words from sheet1 are found in sheet2 (same frequency)
- Sheet2 may have additional words (marked in "Columns with extra content")

#### Partial Match (0-99%)
- Some words from sheet1 are missing in sheet2
- Or word frequencies don't match

#### 0% Match
- None of the words from sheet1 are found in sheet2

#### Special Cases
- Both empty: **100%** (considered identical)
- Sheet1 empty, sheet2 has content: **0%**
- Sheet1 has content, sheet2 empty: **0%**

## Word Frequency Matching

The tool uses `Counter` to match word frequencies:

**Example:**
```
Sheet1: "premium premium quality"
Words: {premium: 2, quality: 1}

Sheet2: "premium quality product"
Words: {premium: 1, quality: 1, product: 1}

Matched:
- premium: min(2, 1) = 1
- quality: min(1, 1) = 1
Total matched: 2

Base count: 3 (sheet1 total words)
Match %: (2 / 3) × 100 = 66.67%
```

## Running the Tool

```bash
python content-match.py
```

**Prerequisites:**
- Python 3.x
- pandas library (`pip install pandas openpyxl`)
- sheet1.xlsx and sheet2.xlsx in the same directory

**Output:**
- `comparison_result.xlsx` - Detailed comparison report

## Important Notes

1. **Sheet1 is Always the Baseline** - Match percentages never exceed 100%
2. **Word Order Doesn't Matter** - "black headphones" = "headphones black"
3. **Case Insensitive** - "Premium" = "premium"
4. **Punctuation Ignored** - "high-quality" = "high quality"
5. **Duplicate Words Count** - "best best product" requires "best" twice for 100%
6. **Extra Content Tracking** - Only shows columns with 100% match + extra words

## Use Cases

- Product catalog comparison across retailers
- Content consistency verification
- Identifying enhanced product descriptions
- Quality assurance for content migration
- Detecting unauthorized content additions
