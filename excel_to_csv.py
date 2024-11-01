import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

def find_sheets_with_lat_lon(excel_file):
    sheets_with_lat_lon = []
    for sheet in excel_file.sheet_names:
        df = pd.read_excel(excel_file, sheet_name=sheet, nrows=5)
        if 'Latitude' in df.columns and 'Longitude' in df.columns:
            sheets_with_lat_lon.append(sheet)
    return sheets_with_lat_lon

def convert_excel_to_csv(excel_path, output_dir):
    # Load the Excel file
    excel_file = pd.ExcelFile(excel_path)
    
    # Find sheets with Latitude and Longitude columns
    sheets_to_convert = find_sheets_with_lat_lon(excel_file)
    
    # Iterate over the sheets and convert them to CSV
    for sheet in sheets_to_convert:
        df = pd.read_excel(excel_path, sheet_name=sheet)
        csv_path = os.path.join(output_dir, f"{sheet}.csv")
        df.to_csv(csv_path, index=False)
        print(f"Sheet {sheet} converted to {csv_path}")
    return sheets_to_convert

def select_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if file_path:
        output_dir = filedialog.askdirectory(title="Select Output Directory")
        if output_dir:
            sheets_converted = convert_excel_to_csv(file_path, output_dir)
            messagebox.showinfo("Success", f"Converted sheets: {', '.join(sheets_converted)}")

# Set up the GUI
root = tk.Tk()
root.title("Excel to CSV Converter")

frame = tk.Frame(root, width=400, height=200)
frame.pack_propagate(False)
frame.pack()

label = tk.Label(frame, text="Drag and drop your Excel file here", pady=20)
label.pack()

button = tk.Button(frame, text="Select File", command=select_file)
button.pack(pady=20)

root.mainloop()
