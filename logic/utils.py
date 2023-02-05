import re

from tkinter import filedialog as fd

def make_path():
    file_path = fd.askopenfilename(
        filetypes=(('text files', '*.txt'),('All files', '*.*')),
        initialdir="C:/Downloads"
    )
    return file_path


def extract_txt_data(path):
    with open(path, "r") as file:
        for line in file:
            pass
    with open(r"C:\Users\ushak\YandexDisk\Удаленка\Общий диск\Отчеты FTOWER автоматизация\типовая.txt", "r", encoding="ANSI") as file:
        file_data = []
        for line in file:
            file_data.append(line.rstrip("\n"))

    # Foundation loads idexes extraction 
    for i in range(len(file_data)):
        file_data[i].rstrip("\n")
        if re.match("Foundation loads", file_data[i]):
            foundation_loads_start = i
        if re.match("Summary of foundation loads", file_data[i]):
            foundation_loads_end = i
            break
    foundation_loads_data = file_data[foundation_loads_start:foundation_loads_end]

    # Something else..