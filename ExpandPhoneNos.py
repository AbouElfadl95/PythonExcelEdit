import pandas as pd
import os
import re

# Read the input .xls file


# Function to decide if a row should be deleted:
def should_delete(cell):
    if pd.isna(cell):
        return True
    val = str(cell).strip()
    parts = re.split(r"[\/\-_]", val)
    parts = [p.strip() for p in parts if p.strip() != ""]
    if not parts:
        return False
    for p in parts:
        if not re.fullmatch(r"\d+", p):
            return False
        if len(p) >= 10:
            return False
    return True

# Keep rows that do NOT satisfy should_delete (skip header row)

# Clean column G
def clean_cell(cell):
    if pd.isna(cell):
        return ""
    normalized = re.sub(r"\s*([\/._-])\s*", r"\1", str(cell))
    parts = re.split(r"[\/._-]", normalized)
    cleaned_numbers = [p for p in parts if len(re.findall(r"\d", p)) >= 9]
    return "/".join(cleaned_numbers)


# Expand rows
def expand_rows_by_column_g(df):
    output_rows = []
    header = df.columns.tolist()
    output_rows.append(header)
    for idx, row in df.iloc[1:].iterrows():  # skip header
        cell_value = str(row.iloc[6]) if pd.notna(row.iloc[6]) else ""
        split_values = [v.strip() for v in re.split(r"[\/\-_]", cell_value) if v.strip()]
        if not split_values:
            output_rows.append(list(row))
            continue
        new_row = list(row)
        new_row[6] = split_values[0]
        output_rows.append(new_row)
        for val in split_values[1:]:
            blank_row = [""] * len(row)
            blank_row[6] = val
            output_rows.append(blank_row)
    return pd.DataFrame(output_rows[1:], columns=output_rows[0])


def get_newest_file(root_dir):
    newest_file = None
    newest_mtime = 0

    for dirpath, _, filenames in os.walk(root_dir):
        for fname in filenames:
            fpath = os.path.join(dirpath, fname)
            try:
                mtime = os.path.getmtime(fpath)
                if mtime > newest_mtime:
                    newest_mtime = mtime
                    newest_file = fname
            except Exception:
                pass  # skip files that can't be accessed

    return newest_file

# Apply final row expansion

# Save to .xlsx
input_file = input("Enter File Name: ")

print(f"Processing file: {input_file}")
df = pd.read_excel(input_file)


df_filtered = pd.concat([df.iloc[:1], df.iloc[1:].loc[~df.iloc[1:, 6].apply(should_delete)]])
df_filtered.iloc[1:, 6] = df_filtered.iloc[1:, 6].apply(clean_cell)
final_df = expand_rows_by_column_g(df_filtered)

final_df.to_excel("Output.xlsx", index=False, engine="openpyxl")

print("Done: output.xlsx created.")


