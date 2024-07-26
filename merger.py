import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import platform
from tabulate import tabulate
import subprocess

def is_gui():
    try:
        root = tk.Tk()
        root.withdraw()
        root.update_idletasks()
        return True
    except:
        return False

def merge_excel_files(file1, sheet1, file2, sheet2, output_file, include_header, remove_duplicates):
    try:
        # Read the specified sheets from the Excel files into DataFrames
        df1 = pd.read_excel(file1, sheet_name=sheet1)
        df2 = pd.read_excel(file2, sheet_name=sheet2)
        
        # Concatenate the DataFrames
        combined_df = pd.concat([df1, df2])
        
        # Remove duplicates if the option is selected
        if remove_duplicates:
            combined_df = combined_df.drop_duplicates()
        
        # Write the result to a new Excel file
        combined_df.to_excel(output_file, index=False, header=include_header)
        
        messagebox.showinfo("Success", f"Data from {file1} and {file2} merged and saved to {output_file}.")
        
        # Try to open the file if in a GUI environment
        if is_gui() and messagebox.askyesno("Open File", "Do you want to open the merged file?"):
            if not open_file(output_file):
                show_light_preview(combined_df)
        else:
            # If not in a GUI, render a light preview in terminal
            print("Rendering light preview in terminal:")
            print(tabulate_preview(combined_df))
    except Exception as e:
        messagebox.showerror("Error", str(e))

def open_file(file_path):
    try:
        if platform.system() == 'Windows':
            os.startfile(file_path)
            return True
        elif platform.system() == 'Darwin':  # macOS
            result = subprocess.run(['open', file_path], capture_output=True)
        else:  # Linux and other Unix-like systems
            result = subprocess.run(['xdg-open', file_path], capture_output=True)
        
        if result.returncode != 0:
            print(f"Error opening file: {result.stderr.decode().strip()}")
            return False
        return True
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
        return False

def tabulate_preview(df, max_rows=10):
    # Return a tabulated preview of the DataFrame (first max_rows rows)
    return tabulate(df.head(max_rows), headers='keys', tablefmt='grid')

def show_light_preview(df):
    preview_text = tabulate_preview(df)
    preview_window = tk.Toplevel()
    preview_window.title("Preview of Merged Data")
    
    text_widget = tk.Text(preview_window, wrap='none')
    text_widget.insert(tk.END, preview_text)
    text_widget.config(state=tk.DISABLED)
    
    text_widget.pack(expand=True, fill='both')

    scrollbar_y = tk.Scrollbar(preview_window, orient='vertical', command=text_widget.yview)
    scrollbar_y.pack(side='right', fill='y')
    text_widget.config(yscrollcommand=scrollbar_y.set)

    scrollbar_x = tk.Scrollbar(preview_window, orient='horizontal', command=text_widget.xview)
    scrollbar_x.pack(side='bottom', fill='x')
    text_widget.config(xscrollcommand=scrollbar_x.set)

def select_file(entry, sheet_combo):
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls *.xlsm *.xlsb *.ods *.csv *.xltx")])
    entry.delete(0, tk.END)
    entry.insert(0, file_path)
    
    # Update sheet names
    if file_path:
        try:
            sheets = pd.ExcelFile(file_path).sheet_names
            sheet_combo['values'] = sheets
            sheet_combo.current(0)
        except Exception as e:
            messagebox.showerror("Error", f"Could not read sheets: {e}")

def save_file(entry):
    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    entry.delete(0, tk.END)
    entry.insert(0, file_path)

def merge_files():
    file1 = entry_file1.get()
    sheet1 = combo_sheet1.get()
    file2 = entry_file2.get()
    sheet2 = combo_sheet2.get()
    output_file = entry_output.get()
    include_header = var_header.get()
    remove_duplicates = var_remove_duplicates.get()
    
    if not file1 or not file2 or not output_file:
        messagebox.showwarning("Input Error", "Please select both input files, sheets, and specify an output file.")
        return
    
    merge_excel_files(file1, sheet1, file2, sheet2, output_file, include_header, remove_duplicates)

# Create the main window
root = tk.Tk()
root.title("Excel Merger")

# Create and place the labels, entries, combos, and buttons
tk.Label(root, text="Select first Excel file:").grid(row=0, column=0, padx=10, pady=10)
entry_file1 = tk.Entry(root, width=50)
entry_file1.grid(row=0, column=1, padx=10, pady=10)
btn_file1 = tk.Button(root, text="Browse", command=lambda: select_file(entry_file1, combo_sheet1))
btn_file1.grid(row=0, column=2, padx=10, pady=10)

tk.Label(root, text="Select sheet:").grid(row=1, column=0, padx=10, pady=10)
combo_sheet1 = ttk.Combobox(root, width=47)
combo_sheet1.grid(row=1, column=1, padx=10, pady=10)

tk.Label(root, text="Select second Excel file:").grid(row=2, column=0, padx=10, pady=10)
entry_file2 = tk.Entry(root, width=50)
entry_file2.grid(row=2, column=1, padx=10, pady=10)
btn_file2 = tk.Button(root, text="Browse", command=lambda: select_file(entry_file2, combo_sheet2))
btn_file2.grid(row=2, column=2, padx=10, pady=10)

tk.Label(root, text="Select sheet:").grid(row=3, column=0, padx=10, pady=10)
combo_sheet2 = ttk.Combobox(root, width=47)
combo_sheet2.grid(row=3, column=1, padx=10, pady=10)

tk.Label(root, text="Output Excel file:").grid(row=4, column=0, padx=10, pady=10)
entry_output = tk.Entry(root, width=50)
entry_output.grid(row=4, column=1, padx=10, pady=10)
btn_output = tk.Button(root, text="Save As", command=lambda: save_file(entry_output))
btn_output.grid(row=4, column=2, padx=10, pady=10)

tk.Label(root, text="Include header row:").grid(row=5, column=0, padx=10, pady=10)
var_header = tk.BooleanVar(value=True)
chk_header = tk.Checkbutton(root, variable=var_header)
chk_header.grid(row=5, column=1, padx=10, pady=10, sticky='w')

tk.Label(root, text="Remove duplicates:").grid(row=6, column=0, padx=10, pady=10)
var_remove_duplicates = tk.BooleanVar(value=True)
chk_remove_duplicates = tk.Checkbutton(root, variable=var_remove_duplicates)
chk_remove_duplicates.grid(row=6, column=1, padx=10, pady=10, sticky='w')

btn_merge = tk.Button(root, text="Merge Files", command=merge_files)
btn_merge.grid(row=7, column=0, columnspan=3, pady=20)

# Run the main event loop
root.mainloop()
