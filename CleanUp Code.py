# -*- coding: utf-8 -*-
"""
Created on Fri Aug 22 20:51:59 2025

@author: Admin
"""
# Importing Libraries
import pandas as pd

# Input/Output Files
csv_file = "C:\\Users\\Admin\\Documents\\Portfolio Case Studies\\Tableau\\FAF5 Data.csv" # CSV of Freight Data
lookup_file = "C:\\Users\\Admin\\Documents\\Portfolio Case Studies\\Tableau\\FAF5_metadata.xlsx" # Relational Lookup Data
region_file = "C:\\Users\\Admin\\Documents\\Portfolio Case Studies\\Tableau\\Region and Sub-Region.xlsx" #Region Lookup Data
output_file = "data_region_updated.xlsx" # New Output File Data

# Set the Data Frame for the CSV of Freight Data
df = pd.read_csv(csv_file)

# Load lookup sheets into dictionaries
lookup_sheets = pd.read_excel(lookup_file, sheet_name=None)
lookup_dicts = {sheet: dict(zip(data.iloc[:, 0], data.iloc[:, 1]))
                for sheet, data in lookup_sheets.items()}

# Mapping of CSV Freight Data columns to the lookup Relational Data columns
column_to_sheet = {
    "dms_origst": "State",
    "dms_destst": "State",
    "fr_orig": "FAF Zone (Foreign)",
    "fr_dest": "FAF Zone (Foreign)",
    "fr_inmode": "Mode",
    "fr_outmode": "Mode",
    "dms_mode": "Mode",
    "sctg2": "Commodity (SCTG2)",
    "trade_type": "Trade Type",
    "dist_band": "Distance Band"
}

# Helper for Name Mapping
def map_codes(series, mapping):
    out = series.map(mapping)
    if out.isna().any():
        mapping_str = {str(k): v for k, v in mapping.items()}
        out2 = series.astype(str).map(mapping_str)
        out = out.fillna(out2)
    return out.fillna(series)

# Replace codes in CSV with descriptive names from Relational Data
for col, sheet in column_to_sheet.items():
    if col in df.columns and sheet in lookup_dicts:
        df[col] = map_codes(df[col], lookup_dicts[sheet])

# Load Region/Subregion
region_df = pd.read_excel(region_file, sheet_name="Sheet1")

# Create mapping for State: (Region, Subregion)
region_dict = region_df.set_index(region_df.columns[0])[
    [region_df.columns[1], region_df.columns[2]]
].to_dict(orient="index")

# Helper to extract region/subregion by defining row and columns
def get_region_sub(state):
    row = region_dict.get(state)
    if row:
        return row[region_df.columns[1]], row[region_df.columns[2]]
    return None, None

# Add Region/Subregion columns next to origin/destination states in new data file
df[["dms_origreg", "dms_origsubreg"]] = df["dms_origst"].apply(lambda st: pd.Series(get_region_sub(st)))
df[["dms_destreg", "dms_destsubreg"]] = df["dms_destst"].apply(lambda st: pd.Series(get_region_sub(st)))

# Reorder new columns right after the state columns
def move_after(df, cols_to_move, after_col):
    cols = list(df.columns)
    for col in cols_to_move[::-1]:
        cols.insert(cols.index(after_col) + 1, cols.pop(cols.index(col)))
    return df[cols]

df = move_after(df, ["dms_origreg", "dms_origsubreg"], "dms_origst")
df = move_after(df, ["dms_destreg", "dms_destsubreg"], "dms_destst")

# Identify numeric columns and scale them
tons_cols = [c for c in df.columns if c.startswith("tons_")]
value_cols = [c for c in df.columns if c.startswith("value_")]
current_value_cols = [c for c in df.columns if c.startswith("current_value_")]

for col in tons_cols:
    if col in df.columns:
        df[col] = pd.to_numeric(df[col], errors="coerce") * 1000  # converts Tons to thousands

for col in value_cols:
    if col in df.columns:
        df[col] = pd.to_numeric(df[col], errors="coerce") * 1000  # converts value to millions $

for col in current_value_cols:
    if col in df.columns:
        df[col] = pd.to_numeric(df[col], errors="coerce") * 1000  # converts current value to millions $

# Save to Excel with xlsxwriter to maintain structure
with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
    df.to_excel(writer, index=False, sheet_name="Data")
    workbook = writer.book
    worksheet = writer.sheets["Data"]

    # Fefine formats
    tons_fmt = workbook.add_format({"num_format": "#,##0.00", "align": "right"})
    money_fmt = workbook.add_format({"num_format": "$#,##0.00", "align": "right"})

    # Apply formats to numeric columns
    for col_num, col_name in enumerate(df.columns):
        if col_name in tons_cols:
            worksheet.set_column(col_num, col_num, 14, tons_fmt)
        elif col_name in value_cols + current_value_cols:
            worksheet.set_column(col_num, col_num, 16, money_fmt)
        else:
            worksheet.set_column(col_num, col_num, 14)  # default width for other columns

print(f"✅ Excel file saved: {output_file}")