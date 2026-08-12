import tkinter as tk
from tkinter import filedialog
from openpyxl import load_workbook, Workbook

root = tk.Tk()
root.title('Excel combiner')
root.geometry('400x200')

file1 = None
file2 = None

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

def combine_files(file1, file2):
    if file1 is None and file2 is None:
        print('Waiting for file selection...')
        return

    wb1 = load_workbook(file1)
    wb2 = load_workbook(file2)

    ws1 = wb1.active
    ws2 = wb2.active

    data_dict1 = {}

    # need to fix- not trying to get the value while keeping .iter_cols
    for cell in ws1.iter_cols(min_row=5, max_col=5):
        cell = ws1.cell(
            row=6,
            column=5
        ).value
        print(cell)

    cleaned_file = Workbook()

    file_path = r'C:\Users\rjayd\OneDrive\Documents\Excel files\cleaned_file.xlsx'
    cleaned_file.save(file_path)

    print('created excel file', file_path)


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
    command= lambda: combine_files(file1, file2)
).pack(pady=15)

# start tkinter
root.mainloop()