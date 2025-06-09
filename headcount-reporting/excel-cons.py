import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from openpyxl.styles import numbers
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl import load_workbook

# Optional column mapping to standardize names
column_mapping = {
    "Salary month": "Salary",
    "Office": "Location",
    "Offic": "Location",
    "Country Name": "Country",
    "Country": "Country",
    "Country_Code": "Country",
    # Add more mappings as needed
}


def clean_and_standardize(df):
    df = df.dropna(how="all")
    df = df.dropna(axis=1, how="all")
    df = df.fillna("")
    df.columns = [column_mapping.get(col.strip(), col.strip()) for col in df.columns]

    if "Hire Date" in df.columns:
        df["Hire Date"] = pd.to_datetime(df["Hire Date"], errors="coerce")
    return df


def process_excel_files(folder_path):
    output_file = os.path.join(folder_path, "Merged_Headcount.xlsx")
    excel_files = [f for f in os.listdir(folder_path) if f.endswith((".xlsx", ".xls"))]

    if not excel_files:
        messagebox.showwarning("No Excel Files", "No Excel files found in the selected folder.")
        return

    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        for file in excel_files:
            file_path = os.path.join(folder_path, file)
            try:
                df = pd.read_excel(file_path)
                df = clean_and_standardize(df)
                sheet_name = os.path.splitext(file)[0][:31]
                df.to_excel(writer, sheet_name=sheet_name, index=False)
            except Exception as e:
                messagebox.showerror("Error", f"Failed to process {file}:\n{e}")
                return

        # Format Hire Date in all sheets
        workbook = writer.book
        for sheet_name in writer.sheets:
            sheet = writer.sheets[sheet_name]
            header = [cell.value for cell in sheet[1]]
            if "Hire Date" in header:
                col_index = header.index("Hire Date") + 1
                for row in sheet.iter_rows(min_row=2, min_col=col_index, max_col=col_index):
                    for cell in row:
                        cell.number_format = "DD/MM/YYYY"

    messagebox.showinfo("Success", f"Excel files merged and saved as:\n{output_file}")


# GUI setup
def launch_gui():
    root = tk.Tk()
    root.title("Excel File Merger")
    root.geometry("400x200")

    def select_folder():
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            process_excel_files(folder_selected)

    label = tk.Label(root, text="Select a folder with Excel files to merge", pady=20)
    label.pack()

    button = tk.Button(root, text="Browse Folder", command=select_folder, width=20)
    button.pack(pady=10)

    root.mainloop()


if __name__ == "__main__":
    launch_gui()
