import tkinter as tk  # import Tkinter package for GUI
from tkinter import filedialog  # import filedialog for file selection
import openpyxl  # import openpyxl for xlsx file manipulation
import csv  # import csv for csv file manipulation
from datetime import datetime  # import datetime for date manipulation


def select_xlsx_file():  # function to select xlsx file
    filename = filedialog.askopenfilename(
        filetypes=[("Excel files", "*.xlsx")])  # open file dialog and select xlsx file
    entry_xlsx_file.delete(0, tk.END)  # delete current entry in entry box
    entry_xlsx_file.insert(0, filename)  # insert selected file name into entry box


def select_csv_file():  # function to select csv file
    filename = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])  # open file dialog and select csv file
    entry_csv_file.delete(0, tk.END)  # delete current entry in entry box
    entry_csv_file.insert(0, filename)  # insert selected file name into entry box


def select_output_folder():  # function to select output folder
    foldername = filedialog.askdirectory()  # open file dialog and select output folder
    entry_output_folder.delete(0, tk.END)  # delete current entry in entry box
    entry_output_folder.insert(0, foldername)  # insert selected folder name into entry box


root = tk.Tk()  # create root window

label_excel_file = tk.Label(root, text="Select Excel File:")  # create label for excel file
label_excel_file.grid(row=0, column=0, sticky=tk.W)  # place label in root window
button_excel_file = tk.Button(root, text="Browse", command=select_xlsx_file)  # create button to select excel file
button_excel_file.grid(row=0, column=1, sticky=tk.W)  # place button in root window
entry_xlsx_file = tk.Entry(root, width=50)  # create entry box for excel file
entry_xlsx_file.grid(row=0, column=2, sticky=tk.W)  # place entry box in root window

label_csv_file = tk.Label(root, text="Select CSV File:")  # create label for csv file
label_csv_file.grid(row=1, column=0, sticky=tk.W)  # place label in root window
button_csv_file = tk.Button(root, text="Browse", command=select_csv_file)  # create button to select csv file
button_csv_file.grid(row=1, column=1, sticky=tk.W)  # place button in root window
entry_csv_file = tk.Entry(root, width=50)  # create entry box for csv file
entry_csv_file.grid(row=1, column=2, sticky=tk.W)  # place entry box in root window

label_output_folder = tk.Label(root, text="Select Output Folder:")  # create label for output folder
label_output_folder.grid(row=2, column=0, sticky=tk.W)  # place label in root window
button_output_folder = tk.Button(root, text="Browse",
                                 command=select_output_folder)  # create button to select output folder
button_output_folder.grid(row=2, column=1, sticky=tk.W)  # place button in root window
entry_output_folder = tk.Entry(root, width=50)  # create entry box for output folder
entry_output_folder.grid(row=2, column=2, sticky=tk.W)  # place entry box in root window

button_process = tk.Button(root, text="Process Files", command=None)  # create button to process files
button_process.grid(row=3, column=1, sticky=tk.W)  # place button in root window

root.mainloop()  # run root window


def process_files():
    # open xlsx file
    wb = openpyxl.load_workbook(entry_xlsx_file.get())  # load xlsx file

    # extract year and month from current file name (assumes file name is YYYY.MM WellsOne Expense Manager)
    file_name = entry_xlsx_file.get()  # get file name from entry box
    current_year_month = file_name.split(" ")[0]  # split file name by space and get first element
    current_year = current_year_month.split(".")[0]  # split current year and month by period and get first element
    current_month = current_year_month.split(".")[1]  # split current year and month by period and get second element

    # increment month by 1 and increment year by 1 if month is 13
    if current_month == "12":  # if current month is 12
        next_month = "01"  # set next month to 01
        next_year = str(int(current_year) + 1)  # increment year by 1
    else:  # if current month is not 12
        next_month = str(int(current_month) + 1)  # increment month by 1
        next_year = current_year  # set next year to current year

    # create new file name for next month
    next_year_month = next_year + "." + next_month  # create new file name for next month
    new_file_name = f"{next_year_month} WellsOne Expense Manager"  # create new file name for next month

    # save copy of wb as new file name in output folder
    wb.save(f"{entry_output_folder.get()}/{new_file_name}.xlsx")  # save copy of wb as new file name in output folder


button_process.config(command=process_files)  # set button_process command to process_files function
