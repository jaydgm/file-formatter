import json
import pprint
import tkinter as tk
from tkinter import filedialog
from openpyxl import load_workbook, Workbook

root = tk.Tk()
root.title('Excel combiner')
root.geometry('400x200')

file1 = None
file2 = None
output_path = r'C:\Users\rjayd\OneDrive\Documents\Excel files\cleaned_file.xlsx'
def select_file1():
    global file1

    file1 = filedialog.askopenfilename(
        title="Select First Excel File",
        filetypes=[("Excel files", "*.xlsx")]
    )

    if file1:
        file1_label.config(text=file1)

def select_file2():
    global file2

    file2 = filedialog.askopenfilename(
        title="Select Second Excel File",
        filetypes=[("Excel files", "*.xlsx")]
    )

    if file2:
        file2_label.config(text=file2)

def combine_files(file1, file2, output):
    if file1 is None and file2 is None:
        print('Waiting for file selection...')
        return

    wb1 = load_workbook(file1)
    wb2 = load_workbook(file2)
    cleaned_file = load_workbook(output_path)

    ws1 = wb1.active
    ws2 = wb2.active
    ws3 = cleaned_file.active

    data_dict1 = {}

    for row in ws1.iter_rows(min_row=5, max_row=18, min_col=6, max_col=51):
        for cell in row:
            if cell.value is not None:
                value = cell.value
                row_num = cell.row
                col_num = cell.column
                print(f'cell is at row {row_num} and at column {col_num}, value is {value} \n')
                data_dict1[value] = (row_num, col_num)
                print(data_dict1)

            

    cleaned_file.save(output_path)

    print('created excel file', output_path)

# file 1
tk.Button(
    root,
    text='Select File 1',
    command=select_file1
).pack(pady=10)

file1_label = tk.Label(root, text="No file selected")
file1_label.pack()

tk.Button(
    root,
    text='Select File 2',
    command=select_file2
).pack(pady=10)

file2_label = tk.Label(root, text="No file selected")
file2_label.pack()

# lambda runs combine_files once button is clicked
tk.Button(
    root,
    text='Combine Files',
    command= lambda: combine_files(file1, file2, output_path)
).pack(pady=15)

# start tkinter
root.mainloop()