# Content Match - Excel Column Comparison Tool

## Overview
This tool compares content between two Excel files (ApprovedFile.xlsx and DeployedFile.xlsx) on a word-by-word basis, using **ApprovedFile as the 100% baseline**. The comparison is performed by matching records using composite keys (Basepack + Account) and analyzing text content across multiple columns.

## How the Matching Logic Works

### 1. File Structure
- **ApprovedFile.xlsx** - Base file (reference, 100%)
- **DeployedFile.xlsx** - Comparison file
- Both files must contain: `Basepack`, `Account`, and comparison columns (Title, Feature bullets, Product description)

### 2. Composite Key Matching
Records are matched between files using: `Basepack + Account`
- Only matching records are compared
- Records not found in DeployedFile are flagged as "NOT FOUND in DeployedFile"

### 3. Word-by-Word Matching Algorithm

#### Core Principle
**ApprovedFile column word count = 100%**

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
matched_words = sum of minimum counts for each word present in ApprovedFile
base_word_count = total words in ApprovedFile column

match_percentage = (matched_words / base_word_count) × 100%
```

**Step 4: Identify Extra Content**
- If DeployedFile has ALL words from ApprovedFile (100% match)
- AND DeployedFile has additional words beyond ApprovedFile
- Then mark that column name in "Columns with extra content"

#### Example

**ApprovedFile - Title Column:**
```
"Sunlight Colour Guard Detergent"
Words: [Sunlight, Colour, Guard, Detergent]
Word count: 4 (this is 100%)
```

**DeployedFile - Title Column:**
```
"Sunlight Colour Guard Detergent Powder"
Words: [Sunlight, Colour, Guard, Detergent, Powder]
Word count: 5
```

**Result:**
- Matched words: 4 (all ApprovedFile words found)
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
- **Columns with extra content** - Column names where DeployedFile has 100% match + extra words

#### Per-Column Metrics
For each column (Title, Feature bullet 1-6, Product description):
- **{Column Name} MatchPercentWordLevel** - Individual column match %

### 5. Match Percentage Rules

#### 100% Match
- All words from ApprovedFile are found in DeployedFile (same frequency)
- DeployedFile may have additional words (marked in "Columns with extra content")

#### Partial Match (0-99%)
- Some words from ApprovedFile are missing in DeployedFile
- Or word frequencies don't match

#### 0% Match
- None of the words from ApprovedFile are found in DeployedFile

#### Special Cases
- Both empty: **100%** (considered identical)
- ApprovedFile empty, DeployedFile has content: **0%**
- ApprovedFile has content, DeployedFile empty: **0%**

## Word Frequency Matching

The tool uses `Counter` to match word frequencies:

**Example:**
```
ApprovedFile: "premium premium quality"
Words: {premium: 2, quality: 1}

DeployedFile: "premium quality product"
Words: {premium: 1, quality: 1, product: 1}

Matched:
- premium: min(2, 1) = 1
- quality: min(1, 1) = 1
Total matched: 2

Base count: 3 (ApprovedFile total words)
Match %: (2 / 3) × 100 = 66.67%
```

## Running the Tool

```bash
python content-match.py
```

**Prerequisites:**
- Python 3.x
- pandas library (`pip install pandas openpyxl`)
- ApprovedFile.xlsx and DeployedFile.xlsx in the same directory

**Output:**
- `comparison_result.xlsx` - Detailed comparison report

## Important Notes

1. **ApprovedFile is Always the Baseline** - Match percentages never exceed 100%
2. **Word Order Doesn't Matter** - "Detergent Powder" = "Powder Detergent"
3. **Case Insensitive** - "Detergent" = "detergent"
4. **Punctuation Ignored** - "Detergent-Powder" = "Detergent Powder"
5. **Duplicate Words Count** - "detergent powder powder" requires "powder" twice for 100%
6. **Extra Content Tracking** - Only shows columns with 100% match + extra words

## Use Cases

- Product catalog comparison across retailers
- Content consistency verification
- Identifying enhanced product descriptions
- Quality assurance for content migration
- Detecting unauthorized content additions
