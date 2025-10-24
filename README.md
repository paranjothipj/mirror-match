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
- Use regex `\d+\.\d+|\w+` to extract:
  - Decimal numbers (e.g., "0.35", "12.5") as single tokens
  - Alphanumeric sequences (words, numbers)
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
- Has extra: Yes (Powder)
- Column marked in: **"Columns with extra content: Title"**

#### Example 2: Special Characters & Punctuation Handling

**ApprovedFile - Title Column:**
```
"Lakme 9 to 5 Flawless Matte Complexion Compact|| Almond|| 8 g"
```

**After Normalization:**
```
"lakme 9 to 5 flawless matte complexion compact almond 8 g"
Words: [lakme, 9, to, 5, flawless, matte, complexion, compact, almond, 8, g]
Word count: 11 (this is 100%)
```
*Note: "9 to 5" stays as separate words [9, to, 5] because there's no decimal point*

**DeployedFile - Title Column:**
```
"Lakme 9To5 Flawless Matte Complexion Compact (Almond)"
```

**After Normalization:**
```
"lakme 9to5 flawless matte complexion compact almond"
Words: [lakme, 9to5, flawless, matte, complexion, compact, almond]
Word count: 7
```

**Matching Process:**
- ✓ lakme (matched)
- ✗ 9 (not found - DeployedFile has "9to5" as single word)
- ✗ to (not found)
- ✗ 5 (not found - merged with 9 in DeployedFile)
- ✓ flawless (matched)
- ✓ matte (matched)
- ✓ complexion (matched)
- ✓ compact (matched)
- ✓ almond (matched)
- ✗ 8 (not found)
- ✗ g (not found)

**Result:**
- Matched words: 6 out of 11
- Match percentage: (6 / 11) × 100 = **54.55%**
- Has extra: No (missing words, not 100% match)
- Column marked in: *(not marked, less than 100%)*
- Remark: "Title" (mismatch detected)

#### Example 3: Product Variant Mismatch

**ApprovedFile - Title Column:**
```
"Elle 18 Eye Drama Kajal|| Super Black|| 0.35 g"
```

**After Normalization:**
```
"elle 18 eye drama kajal super black 0.35 g"
Words: [elle, 18, eye, drama, kajal, super, black, 0.35, g]
Word count: 9 (this is 100%)
```
*Note: "0.35" is kept as one word (decimal number)*

**DeployedFile - Title Column:**
```
"Elle 18 Eye Drama Kajal (Bold Black)"
```

**After Normalization:**
```
"elle 18 eye drama kajal bold black"
Words: [elle, 18, eye, drama, kajal, bold, black]
Word count: 7
```

**Matching Process:**
- ✓ elle (matched)
- ✓ 18 (matched)
- ✓ eye (matched)
- ✓ drama (matched)
- ✓ kajal (matched)
- ✗ super (not found - DeployedFile has "bold" instead)
- ✓ black (matched)
- ✗ 0.35 (not found)
- ✗ g (not found)

**Result:**
- Matched words: 6 out of 9
- Match percentage: (6 / 9) × 100 = **66.67%**
- Has extra: No (missing words, not 100% match)
- Column marked in: *(not marked, less than 100%)*
- Remark: "Title" (mismatch detected)

**Key Observations:**
- "Super Black" vs "Bold Black" - different variant names cause mismatch
- Missing weight "0.35 g" reduces match percentage
- `||` and `()` punctuation removed during normalization
- Decimal "0.35" kept as single word token

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
5. **Decimals Preserved** - "0.35" is one word (not split into "0" and "35")
6. **Duplicate Words Count** - "detergent powder powder" requires "powder" twice for 100%
7. **Extra Content Tracking** - Only shows columns with 100% match + extra words

## Use Cases

- Product catalog comparison across retailers
- Content consistency verification
- Identifying enhanced product descriptions
- Quality assurance for content migration
- Detecting unauthorized content additions
