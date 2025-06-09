import os
import pandas as pd

# Customize this path to your Excel files
input_folder = "C:\Users\zeidl\OneDrive\Documents\Coding\Headcount"
output_file = "merged_countries.xlsx"

# Optional: Column standardization mapping (common headers)
column_mapping = {
    'Employee ID': 'ID',
    'Name': 'Name',
    'Country': 'Country',
    'Location': 'Office',
    'Gender': 'Gender',
    'Department': 'Department',
    'Job Title': 'Job'
    # Add more mappings as needed
}

def standardize_columns(df):
    new_columns = []
    for col in df.columns:
        col_clean = col.strip()
        new_col = column_mapping.get(col_clean, col_clean)
        new_columns.append(new_col)
    df.columns = new_columns
    return df

def clean_data(df):
    df = df.dropna(how='all')  # Drop completely empty rows
    df = df.dropna(axis=1, how='all')  # Drop completely empty columns
    df = df.fillna("")  # Optional: Replace NaNs with empty strings
    return df

# Collect all Excel files in the folder
excel_files = [f for f in os.listdir(input_folder) if f.endswith(('.xlsx', '.xls'))]

# Initialize Excel writer
with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    for file in excel_files:
        file_path = os.path.join(input_folder, file)
        try:
            df = pd.read_excel(file_path)
            df = standardize_columns(df)
            df = clean_data(df)

            # Use filename (without extension) as sheet name
            sheet_name = os.path.splitext(file)[0][:31]  # Max Excel sheet name length = 31
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            print(f"Processed: {file}")
        except Exception as e:
            print(f"Error processing {file}: {e}")

print(f"\nAll files merged into {output_file}")
