import tkinter as tk
from tkinter import filedialog


def select_xlsx_file():
    filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    entry_xlsx_file.delete(0, tk.END)
    entry_xlsx_file.insert(0, filename)


def select_csv_file():
    filename = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
    entry_csv_file.delete(0, tk.END)
    entry_csv_file.insert(0, filename)


def select_output_folder():
    foldername = filedialog.askdirectory()
    entry_output_folder.delete(0, tk.END)
    entry_output_folder.insert(0, foldername)


root = tk.Tk()

label_excel_file = tk.Label(root, text="Select Excel File:")
label_excel_file.grid(row=0, column=0, sticky=tk.W)
button_excel_file = tk.Button(root, text="Browse", command=select_xlsx_file)
button_excel_file.grid(row=0, column=1, sticky=tk.W)
entry_xlsx_file = tk.Entry(root, width=50)
entry_xlsx_file.grid(row=0, column=2, sticky=tk.W)

label_csv_file = tk.Label(root, text="Select CSV File:")
label_csv_file.grid(row=1, column=0, sticky=tk.W)
button_csv_file = tk.Button(root, text="Browse", command=select_csv_file)
button_csv_file.grid(row=1, column=1, sticky=tk.W)
entry_csv_file = tk.Entry(root, width=50)
entry_csv_file.grid(row=1, column=2, sticky=tk.W)

label_output_folder = tk.Label(root, text="Select Output Folder:")
label_output_folder.grid(row=2, column=0, sticky=tk.W)
button_output_folder = tk.Button(root, text="Browse", command=select_output_folder)
button_output_folder.grid(row=2, column=1, sticky=tk.W)
entry_output_folder = tk.Entry(root, width=50)
entry_output_folder.grid(row=2, column=2, sticky=tk.W)

button_process = tk.Button(root, text="Process Files", command=None)
button_process.grid(row=3, column=1, sticky=tk.W)

root.mainloop()
