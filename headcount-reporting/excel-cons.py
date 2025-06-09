import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import numbers

# File paths
file_germany = r"C:\Users\zeidl\OneDrive\Documents\Coding\Headcount\Headcount Germany.xlsx"
file_us = r"C:\Users\zeidl\OneDrive\Documents\Coding\Headcount\US headcount.xlsx"
output_file = r"C:\Users\zeidl\OneDrive\Documents\Coding\Headcount\Merged_Headcount.xlsx"

# Read Excel files
df_germany = pd.read_excel(file_germany)
df_us = pd.read_excel(file_us)

# Rename columns for consistency
df_germany = df_germany.rename(columns={"Salary month": "Salary", "Office": "Location"})


# Clean function
def clean_data(df):
    df = df.dropna(how="all")
    df = df.dropna(axis=1, how="all")
    df = df.fillna("")
    return df


# Clean and convert Hire Date
def prepare_dataframe(df):
    df = clean_data(df)
    df["Hire Date"] = pd.to_datetime(df["Hire Date"], errors="coerce")  # Safely convert
    return df


df_germany = prepare_dataframe(df_germany)
df_us = prepare_dataframe(df_us)

# Write using openpyxl to control date formatting
with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
    for sheet_name, df in [("Germany", df_germany), ("US", df_us)]:
        df.to_excel(writer, sheet_name=sheet_name, index=False)

    workbook = writer.book
    for sheet_name in writer.sheets:
        sheet = writer.sheets[sheet_name]
        for cell in sheet[1]:  # header row
            if cell.value == "Hire Date":
                col_letter = cell.column_letter
                for row in sheet.iter_rows(min_row=2, min_col=cell.column, max_col=cell.column):
                    for c in row:
                        c.number_format = "DD/MM/YYYY"
