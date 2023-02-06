import pathlib
import pandas as pd
import re


from pandas import ExcelWriter
from tkinter import filedialog as fd


def make_path():
    file_path = fd.askopenfilename(
        filetypes=(('text files', '*.txt'),('All files', '*.*')),
        initialdir="C:/Downloads"
    )
    return file_path


def extract_txt_data(path=None):
    if path:
        with open(path, "r") as file:
            for line in file:
                pass
        with open(path, "r", encoding="ANSI") as file:
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

        # A part of "Case" and digit concatination 
        for i in range(len(foundation_loads_data)):
            if i not in [0, 2, len(foundation_loads_data)]:
                foundation_loads_data[i] = foundation_loads_data[i].split()
                if i == 1:
                    temporary_list = []
                    for j in range(len(foundation_loads_data[i])):
                        if foundation_loads_data[i][j]=="Case":
                            element = foundation_loads_data[i][j] + foundation_loads_data[i][j+1]
                            temporary_list.append(element)
                        elif foundation_loads_data[i][j] not in "123456789":
                            element = foundation_loads_data[i][j]
                            temporary_list.append(element)
                    foundation_loads_data[i] = temporary_list         
            else:
                foundation_loads_data[i] = [foundation_loads_data[i]]

        # Export data to xlsx
        writer = ExcelWriter("Foundation_loads_data.xlsx")
        df = pd.DataFrame(foundation_loads_data)
        df.to_excel(writer, "Sheet1")
        writer.close()

        return "Файл конвертирован."
    else:
        return "Для конвертации сначала выберите файл!"
