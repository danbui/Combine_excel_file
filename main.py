import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog

def browse_folder():
    folder_path = filedialog.askdirectory()
    if folder_path:
        entry_folder_path.delete(0, tk.END)
        entry_folder_path.insert(tk.END, folder_path)

def combine_excels():
    folder_path = entry_folder_path.get()
    if folder_path:
        excel_files = [file for file in os.listdir(folder_path) if file.endswith('.xlsx')]
        if excel_files:
            combined_df = pd.DataFrame()
            for file in excel_files:
                df = pd.read_excel(os.path.join(folder_path, file))
                combined_df = pd.concat([combined_df, df], ignore_index=True)
            combined_df.to_excel(os.path.join(folder_path, 'combined.xlsx'), index=False)
            tk.messagebox.showinfo("Success", "Excel files combined successfully!")
        else:
            tk.messagebox.showwarning("Warning", "No Excel files found in the selected folder.")
    else:
        tk.messagebox.showwarning("Warning", "Please select a folder.")

# GUI Setup
root = tk.Tk()
root.title("Combine Excel Files")

label_folder = tk.Label(root, text="Select Folder:")
label_folder.grid(row=0, column=0, padx=5, pady=5, sticky="w")

entry_folder_path = tk.Entry(root, width=50)
entry_folder_path.grid(row=0, column=1, padx=5, pady=5)

button_browse = tk.Button(root, text="Browse", command=browse_folder)
button_browse.grid(row=0, column=2, padx=5, pady=5)

button_combine = tk.Button(root, text="Combine Excel Files", command=combine_excels)
button_combine.grid(row=1, column=0, columnspan=3, padx=5, pady=5)

root.mainloop()
