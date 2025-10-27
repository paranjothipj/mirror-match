import pandas as pd
import re
from collections import Counter

columns_to_compare = [
    "Title",
    "Feature bullet 1",
    "Feature bullet 2",
    "Feature bullet 3",
    "Feature bullet 4",
    "Feature bullet 5",
    "Feature bullet 6",
    "Product description"
]


file1 = "sheet1.xlsx"  # Base sheet
file2 = "sheet2.xlsx"  # Comparison sheet


df1 = pd.read_excel(file1, dtype=str)
df2 = pd.read_excel(file2, dtype=str)


required_columns = columns_to_compare + ["Basepack", "Account"]
missing1 = [col for col in required_columns if col not in df1.columns]
missing2 = [col for col in required_columns if col not in df2.columns]
if missing1:
    raise ValueError(f"sheet1.xlsx is missing columns: {', '.join(missing1)}")
if missing2:
    raise ValueError(f"sheet2.xlsx is missing columns: {', '.join(missing2)}")


df1.fillna("", inplace=True)
df2.fillna("", inplace=True)

# Create composite key: Basepack + Account
df1['composite_key'] = df1['Basepack'] + '|' + df1['Account']
df2['composite_key'] = df2['Basepack'] + '|' + df2['Account']

df1.set_index("composite_key", inplace=True)
df2.set_index("composite_key", inplace=True)

results = []

for composite_key in df1.index:
    if composite_key in df2.index:
        row1 = df1.loc[composite_key]
        row2 = df2.loc[composite_key]
        basepack = row1['Basepack']
        account = row1['Account']
        
        def normalize(val):
            if len(val) < 100:
                print("-- ", val)
            # Robust normalization:
            # - Remove zero-width/BOM and directionality marks
            # - Decode Excel-style escapes like _x000D_
            # - Unicode normalize (NFKC)
            # - Unify hyphens/dashes and smart quotes
            # - Collapse all unicode whitespace (incl. NBSP) to single space
            # - Normalize punctuation spacing
            # - Trim, strip wrapping quotes, and lowercase
            s = str(val)
            s = re.sub(r"[\u200B-\u200F\u202A-\u202E\u2060\u2066-\u2069\uFEFF]", "", s)
            # Decode patterns like _x000D_, _x000A_, etc. to their unicode chars
            s = re.sub(r"_x([0-9A-Fa-f]{4})_", lambda m: chr(int(m.group(1), 16)), s)
            try:
                import unicodedata
                s = unicodedata.normalize("NFKC", s)
            except Exception:
                pass
            for ch in ["\u2010", "\u2011", "\u2012", "\u2013", "\u2014", "\u2212"]:
                s = s.replace(ch, "-")
            s = (s.replace("\u2018", "'").replace("\u2019", "'")
                   .replace("\u201C", '"').replace("\u201D", '"'))
            s = re.sub(r"[\s\u00A0\u1680\u180E\u2000-\u200A\u202F\u205F\u3000]+", " ", s)
            # Normalize punctuation spacing, but preserve decimal points
            s = re.sub(r"\s+([,;:?!])", r"\1", s)  # Remove space before punctuation (except period)
            s = re.sub(r"([,;:?!])\s*", r"\1 ", s)  # Add space after punctuation (except period)
            # Handle periods: only add space if not part of decimal number (e.g., keep 0.35 intact)
            s = re.sub(r"\.(?=\D|\s|$)", ". ", s)  # Add space after period if not followed by digit
            s = re.sub(r"\s+", " ", s)  # Clean up multiple spaces
            s = s.strip().strip('\"\'')
            s = s.lower()
            return s
        
        def word_match_percent_and_has_extra(a: str, b: str) -> tuple[float, bool]:
            # a = file1 (base, 100%), b = file2 (comparison)
            # Return percentage (capped at 100%) and whether file2 has extra words
            # Extract words including decimals (e.g., 0.35)
            w1 = re.findall(r"\d+\.\d+|\w+", a.lower())
            w2 = re.findall(r"\d+\.\d+|\w+", b.lower())

            if not w1 and not w2:
                return 100.0, False
            if not w1:  # file1 empty, file2 has words
                return 0.0, len(w2) > 0

            c1 = Counter(w1)
            c2 = Counter(w2)

            # Count matched words (minimum of both counters)
            matched = sum(min(c1[w], c2[w]) for w in c1)

            # Check if file2 has extra words (beyond what's in file1)
            has_extra = any(c2[word] > c1.get(word, 0) for word in c2)

            base_count = len(w1)  # file1 word count = 100%
            percentage = (matched / base_count) * 100.0

            return percentage, has_extra
        
        # Normalize values first
        norm1 = {col: normalize(row1.get(col, "")) for col in columns_to_compare}
        norm2 = {col: normalize(row2.get(col, "")) for col in columns_to_compare}
        
        mismatches = [col for col in columns_to_compare if norm1[col] != norm2[col]]
        matched_columns = [col for col in columns_to_compare if norm1[col] == norm2[col]]
        remark = "" if not mismatches else ", ".join(mismatches)

        # Identify missing attributes (empty in either sheet after normalization)
        missing_attributes = [col.lower() for col in columns_to_compare if (norm1[col] == "" or norm2[col] == "")]
        missing_attributes_str = ", ".join(missing_attributes)
        
        # Overall word-level match percentage (all attributes)
        # file1 as base (100%), capped at 100%
        words1 = re.findall(r"\d+\.\d+|\w+", " ".join(norm1[col] for col in columns_to_compare).lower())
        words2 = re.findall(r"\d+\.\d+|\w+", " ".join(norm2[col] for col in columns_to_compare).lower())

        if not words1 and not words2:
            overall_percent = 100.0
        elif not words1:
            overall_percent = 0.0
        else:
            c1_all = Counter(words1)
            c2_all = Counter(words2)
            matched_all = sum(min(c1_all[w], c2_all[w]) for w in c1_all)
            overall_percent = (matched_all / len(words1)) * 100.0

        # Feature bullet overall (Feature bullet 1-6 only, file1 as base, capped at 100%)
        bullet_cols = [f"Feature bullet {i}" for i in range(1, 7)]
        fb_text1 = " ".join(norm1[c] for c in bullet_cols)
        fb_text2 = " ".join(norm2[c] for c in bullet_cols)
        fb_tokens1 = re.findall(r"\d+\.\d+|\w+", fb_text1.lower())
        fb_tokens2 = re.findall(r"\d+\.\d+|\w+", fb_text2.lower())

        if not fb_tokens1 and not fb_tokens2:
            fb_overall_percent = 100.0
        elif not fb_tokens1:
            fb_overall_percent = 0.0
        else:
            c1 = Counter(fb_tokens1)
            c2 = Counter(fb_tokens2)
            matched_count = sum(min(c1[w], c2[w]) for w in c1)
            fb_overall_percent = (matched_count / len(fb_tokens1)) * 100.0
        
        # Find columns with extra words (sheet2 has more than sheet1)
        columns_with_extra = []
        for col in columns_to_compare:
            pct, has_extra = word_match_percent_and_has_extra(norm1[col], norm2[col])
            if has_extra and pct == 100.0:  # 100% match but has extra words
                columns_with_extra.append(col)

        row_result = {
            "Basepack": basepack,
            "Account": account,
            "Remark": remark,
            "Missing attributes": missing_attributes_str,
            "Feature bullet overall": f"{fb_overall_percent:.2f}%",
            "MatchedColumns": ", ".join(matched_columns),
            "MatchPercentWordLevel": f"{overall_percent:.2f}%",
            "Columns with extra content": ", ".join(columns_with_extra),
        }
        # Add per-header match percent
        for col in columns_to_compare:
            pct, _ = word_match_percent_and_has_extra(norm1[col], norm2[col])
            row_result[f"{col} MatchPercentWordLevel"] = f"{pct:.2f}%"
        results.append(row_result)
    else:
        # Not found in sheet2: still output per-header columns blank
        basepack = df1.loc[composite_key, 'Basepack']
        account = df1.loc[composite_key, 'Account']
        row_result = {
            "Basepack": basepack,
            "Account": account,
            "Remark": "NOT FOUND in Sheet2",
            "Missing attributes": "",
            "Feature bullet overall": "",
            "MatchedColumns": "",
            "MatchPercentWordLevel": "",
            "Columns with extra content": "",
        }
        for col in columns_to_compare:
            row_result[f"{col} MatchPercentWordLevel"] = ""
        results.append(row_result)

output_df = pd.DataFrame(results)
output_df.to_excel("comparison_result.xlsx", index=False)

print("Output saved to 'comparison_result.xlsx'")
