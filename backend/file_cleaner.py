import tkinter as tk
from tkinter import filedialog
from openpyxl import load_workbook, Workbook

def select_file():
    file_path = filedialog.askopenfilename(
        title = 'Select File',
        filetypes=[
            ('Excel files', '*.xlsx'),
            ('All files', '*.*')
        ]
    )

    if not file_path:
        return

    print(f'Selected file: {file_path}')

    wb = load_workbook(file_path)

    ws = wb.active

# create tkinter window
root = tk.Tk()
root.title('File Formatter')
root.geometry("400x300")

# create button
select_button = tk.Button(
    root,
    text='Select Excel File',
    command=select_file
)

select_button.pack(pady=100)

# start tkinter
root.mainloop()