import base64
import datetime as dt
import os
import pathlib
import pandas as pd
import re
import sys


from docxtpl import DocxTemplate, InlineImage
from dotenv import load_dotenv
from docx.shared import Mm
from pandas import ExcelWriter
from sympy import symbols, latex
from tkinter import filedialog as fd
from tkinter import messagebox as mb

from core.exceptions import FilePathException


icondata = None
icon = os.getenv("icon")
if icon:
    icondata = base64.b64decode(icon)

if icondata:
    tempFile = "logo.ico"
    with open(tempFile, "wb") as iconfile:
        try:
            iconfile.write(icondata)
        except Exception as e:
            print(f"Error writing icon file: {e}")
else:
    print("No icon data found in environment variables")


load_dotenv()
# back_icon = base64.b64decode(os.getenv("ENCODED_BACK"))
# tempFile_back = os.path.abspath("back.png")

# open_icon = base64.b64decode(os.getenv("ENCODED_OPEN"))
# tempFile_open = os.path.abspath("open1.png")

# save_icon = base64.b64decode(os.getenv("ENCODED_SAVE"))
# tempFile_save = os.path.abspath("save1.png")

# lupa_icon = base64.b64decode(os.getenv("ENCODED_LUPA"))
# tempFile_lupa = os.path.abspath("lupa.png")

# with open(tempFile_back, "wb") as iconfileback:
#     iconfileback.write(back_icon)

# with open(tempFile_open, "wb") as iconfileopen:
#     iconfileopen.write(open_icon)

# with open(tempFile_save, "wb") as iconfilesave:
#     iconfilesave.write(save_icon)

# with open(tempFile_lupa, "wb") as iconfileplus:
#     iconfileplus.write(lupa_icon)


current_date = dt.datetime.today().strftime("%d.%m.%Y")
mm_yy = dt.datetime.today().strftime("%m.%Y")


def get_file_path(file):
    if getattr(sys, 'frozen', False):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.abspath(".")
    filepath = os.path.join(base_path, file)
    return filepath


def find_current_user():
    home_directory = os.path.expanduser("~")
    user_path_parts = os.path.normpath(home_directory).split(os.path.sep)
    current_user = user_path_parts[-1]
    return current_user


def make_path_txt(user=find_current_user()):
    file_path = fd.askopenfilename(
        filetypes=(('text files', '*.txt'),('All files', '*.*')),
        initialdir=f"C:/Users/{user}"
    )
    if "Общий диск" in file_path:
        return file_path
    if "КОЗУ (инженерная)" in file_path:
        return file_path
    else:
        mb.showinfo("ERROR", 'Выбранный вами файл должен находиться на "Общем диске".')
        return "Выберите файл с яндекс диска!"


def make_path_png(user=find_current_user()):
    file_path = fd.askopenfilename(
        filetypes=(('png files', '*.png'),('All files', '*.*')),
        initialdir=f"C:/Users/{user}"
    )
    if "Общий диск" in file_path:
        return file_path
    if "КОЗУ (инженерная)" in file_path:
        return file_path
    else:
        mb.showinfo("ERROR", 'Выбранный вами файл должен находиться на "Общем диске".')
        return "Выберите файл с яндекс диска!"


def make_path_xlsx(user=find_current_user()):
    file_path = fd.askopenfilename(
        filetypes=(('xlsx files', '*.xlsx'),('All files', '*.*')),
        initialdir=f"C:/Users/{user}"
    )
    if "Общий диск" in file_path:
        return file_path
    if "КОЗУ (инженерная)" in file_path:
        return file_path
    else:
        mb.showinfo("ERROR", 'Выбранный вами файл должен находиться на "Общем диске".')
        return "Выберите файл с яндекс диска!"


def make_multiple_path(user=find_current_user()):
    file_path = fd.askopenfilenames(
        filetypes=(('png files', '*.png'),('All files', '*.*')),
        initialdir=f"C:/Users/{user}"
    )
    flag = True
    for path in file_path:
        if "Общий диск"  or "КОЗУ (инженерная)" not in path:
            flag = False
    if flag:
        return file_path
    else:
        mb.showinfo("ERROR", 'Выбранный вами файл должен находиться на "Общем диске".')
        return "Выберите файл с яндекс диска!"


def extract_tables_1(path_to_txt_1):
#     try:
#         with open(path_to_txt_1, "r", encoding="ANSI") as file:
#             file_data = []
#             for line in file:
#                 file_data.append(line.rstrip("\n"))
#     except FilePathException as e:
#         print("!!!ERROR!!!", str(e))

#     for i in range(len(file_data)):
#         file_data[i].rstrip("\n")
#         if re.match("Summary of Joint Support Reactions For All Load Cases:", file_data[i]):
#             support_reaction_start = i
#         if re.match("Summary of Tip Deflections For All Load Cases:", file_data[i]):
#             support_reaction_end = i
    
#     support_reaction = file_data[support_reaction_start:support_reaction_end][6:]
    
#     return {
#         "support_reaction": support_reaction
#     }
    
    try:
        with open(path_to_txt_1, "r", encoding="ANSI") as file:
            file_data = []
            for line in file:
                file_data.append(line.rstrip("\n"))
    except FilePathException as e:
        print("!!!ERROR!!!", str(e))

    for i in range(len(file_data)):
        file_data[i].rstrip("\n")
        if re.match("Summary of Joint Support Reactions For All Load Cases:", file_data[i]):
            support_reaction_start = i
        if re.match("Summary of Tip Deflections For All Load Cases:", file_data[i]):
            support_reaction_end = i
        if re.match("Pole Deflection Usages For All Load Cases:", file_data[i]):
            pole_deflection_start = i
        if re.match("Tubes Summary:", file_data[i]):
            tubes_usage_start = i
        if re.match("\*\*\* Overall summary for all load cases - Usage = Maximum Stress / Allowable Stress", file_data[i]):
            tubes_usage_end = i
        if re.match("Summary of Tubular Davit Usages:", file_data[i]):
            davit_usage_start = i
        if re.match("\*\*\* Maximum Stress Summary for Each Load Case", file_data[i]):
            davit_usage_end = i
    
    support_reaction = file_data[support_reaction_start:support_reaction_end][6:]
    pole_deflection = file_data[pole_deflection_start:tubes_usage_start][6:]
    tubes_usage = file_data[tubes_usage_start:tubes_usage_end][6:]
    davit_usage = file_data[davit_usage_start:davit_usage_end][5:]
    
    return {
        "support_reaction": support_reaction,
        "pole_deflection": pole_deflection,
        "tubes_usage": tubes_usage,
        "davit_usage": davit_usage
    }


def extract_tables_data_1(
        path_to_txt_1
    ):
    tables = extract_tables_1(path_to_txt_1=path_to_txt_1)
    list_of_bending_moment = []
    list_of_vertical_force = []
    list_of_shear_force = []
    for row in tables["support_reaction"][:-1]:
        row = row.split()
        list_of_bending_moment.append(abs(float(row[-3])))
        list_of_vertical_force.append(abs(float(row[-7])))
        list_of_shear_force.append(abs(float(row[-6])))
    return {
        "bending_moment": max(list_of_bending_moment),
        "vertical_force": max(list_of_vertical_force),
        "shear_force": max(list_of_shear_force)
    }


def extract_tables_2(path_to_txt_2, is_stand):
    # try:
    #     with open(path_to_txt_2, "r", encoding="ANSI") as file:
    #         file_data = []
    #         for line in file:
    #             file_data.append(line.rstrip("\n"))
    # except FilePathException as e:
    #     print("!!!ERROR!!!", str(e))

    # for i in range(len(file_data)):
    #     file_data[i].rstrip("\n")
    #     if re.match("Steel Pole Properties:", file_data[i]):
    #         pole_properties_start = i
    #     if re.match("Steel Pole Connectivity:", file_data[i]):
    #         pole_connectivity_start = i
    # tubes_properties = file_data[pole_properties_start:pole_connectivity_start]
    # if is_stand == "Да":
    #     tubes_properties = tubes_properties[16:]
    # else:
    #     tubes_properties = tubes_properties[15:]
    # return {
    #     "tubes_properties": tubes_properties,
    # }
    try:
        with open(path_to_txt_2, "r", encoding="ANSI") as file:
            file_data = []
            for line in file:
                file_data.append(line.rstrip("\n"))
    except FilePathException as e:
        print("!!!ERROR!!!", str(e))

    for i in range(len(file_data)):
        file_data[i].rstrip("\n")
        if re.match("Steel Pole Properties:", file_data[i]):
            pole_properties_start = i
        if re.match("Steel Tubes Properties:", file_data[i]):
            pole_properties_end = i
        if re.match("Steel Pole Connectivity:", file_data[i]):
            pole_connectivity_start = i
        if re.match("Relative Attachment Labels for Steel Pole", file_data[i]):
            pole_attachments_start = i
        if re.match("Pole Steel Properties:", file_data[i]):
            pole_attachments_end = i
        if re.match("Joints Geometry:", file_data[i]):
            joints_start = i
    
    pole_properties = file_data[pole_properties_start:pole_properties_end]
    tubes_properties = file_data[pole_properties_start:pole_connectivity_start]
    if is_stand:
        joints_properties = file_data[joints_start:pole_properties_start][5:8]
        pole_properties = pole_properties[7:9]
        tubes_properties = tubes_properties[16:]
    else:
        pole_properties = pole_properties[7:8]
        tubes_properties = tubes_properties[15:]
        joints_properties = ""
    pole_connectivity = file_data[pole_connectivity_start:pole_attachments_start][6:]
    pole_attachments = file_data[pole_attachments_start:pole_attachments_end][6:]
        
    return {
        "pole_properties": pole_properties,
        "tubes_properties": tubes_properties,
        "joints_properties": joints_properties,
        "pole_connectivity": pole_connectivity,
        "pole_attachments": pole_attachments
    }


def extract_tables_data_2(
        path_to_txt_2,
        is_stand,
        is_plate,
        branches,
        ground_wire=None,
        ground_wire_attachment=None,
        wire_pos=None,
    ):
    tables = extract_tables_2(path_to_txt_2=path_to_txt_2, is_stand=is_stand)
    foundation_level = round(float(tables["joints_properties"][0].split()[4]), 2)\
    if is_stand else round(float(tables["pole_connectivity"][0].split()[3]), 2)
    pole_height = tables["pole_properties"][0].split()[-16] if not is_stand\
    else round(float(tables["pole_properties"][0].split()[-16]), 2) +\
    round(float(tables["pole_properties"][1].split()[-16]), 2)
    if is_plate:
       for i, s in enumerate(tables["tubes_properties"]):
           if re.match("Base Plate Properties:", s):
               index = i
               break
       tables["tubes_properties"] = tables["tubes_properties"][:i-1]
    else:
        tables["tubes_properties"] = tables["tubes_properties"][:-1]
    if is_stand:
        tables["tubes_properties"] = tables["tubes_properties"][:-8] +\
        [tables["tubes_properties"][-1]]
    bot_diameter = round(float(tables["tubes_properties"][-1].split()[-3]), 2) * 10
    top_diameter = round(float(tables["tubes_properties"][0].split()[-4]), 2) * 10
    davit_height_list = []
    if (wire_pos == "Вертикальное" or branches == "2") and ground_wire:
        for i in range(3):
            davit_height_list.append(
                str(round(float(tables["pole_attachments"][i].split()[-1]), 2) - foundation_level)
            )
        davit_height = ", ".join(davit_height_list)
        if ground_wire_attachment == "Ниже верха опоры":
            ground_wire_height = round(float(tables["pole_attachments"][3].split()[-1]), 2) - foundation_level
        else:
            ground_wire_height = pole_height
        if_ground_davit_height = f"Узел крепления троса располагается на высоте {ground_wire_height} м."
    elif wire_pos == "Горизонтальное" and ground_wire:
        davit_height = round(float(tables["pole_attachments"][0].split()[-1]), 2) - foundation_level
        if ground_wire_attachment == "Ниже верха опоры":
            ground_wire_height = round(float(tables["pole_attachments"][1].split()[-1]), 2) - foundation_level
        else:
            ground_wire_height = pole_height
        if_ground_davit_height = f"Узел крепления троса располагается на высоте {ground_wire_height} м."
    elif (wire_pos == "Вертикальное" or branches == "2") and not ground_wire:
        for i in range(3):
            davit_height_list.append(
                str(round(float(tables["pole_attachments"][i].split()[-1]), 2) - foundation_level)
            )
        davit_height = ", ".join(davit_height_list)
        if_ground_davit_height = ""
    elif wire_pos == "Горизонтальное" and not ground_wire:
        davit_height = round(float(tables["pole_attachments"][0].split()[-1]), 2) - foundation_level
        if_ground_davit_height = ""
    elif branches == "1" and ground_wire:
        for i in range(2):
            davit_height_list.append(
                str(round(float(tables["pole_attachments"][i].split()[-1]), 2) - foundation_level)
            )
        davit_height = ", ".join(davit_height_list)
        if ground_wire_attachment == "Ниже верха опоры":
            ground_wire_height = round(float(tables["pole_attachments"][2].split()[-1]), 2) - foundation_level
        else:
            ground_wire_height = pole_height
        if_ground_davit_height = f"Узел крепления троса располагается на высоте {ground_wire_height} м."
    elif branches == "1" and not ground_wire:
        for i in range(2):
            davit_height_list.append(
                str(round(float(tables["pole_attachments"][i].split()[-1]), 2) - foundation_level)
            )
        davit_height = ", ".join(davit_height_list)
        if_ground_davit_height = ""
    return {
        "bot_diameter": bot_diameter,
        "top_diameter": top_diameter,
        "pole_heigth": pole_height,
        "davit_heigth": davit_height,
        "ground_wire_heigth": ground_wire_height
    }


def extract_foundation_loads_and_diam(
        path_to_txt_1,
        path_to_txt_2,
        is_plate,
        is_stand,
        branches,
        ground_wire=None,
        ground_wire_attachment=None,
        wire_pos=None,
):
    tables_data_1 = extract_tables_data_1(
        path_to_txt_1=path_to_txt_1
    )

    tables_data_2 = extract_tables_data_2(
        path_to_txt_2=path_to_txt_2,
        is_stand=is_stand,
        is_plate=is_plate,
        branches=branches,
        ground_wire=ground_wire,
        ground_wire_attachment=ground_wire_attachment,
        wire_pos=wire_pos
    )
    return {
        "moment": tables_data_1["bending_moment"],
        "vert_force": tables_data_1["vertical_force"],
        "shear_force": tables_data_1["shear_force"],
        "bot_diam": tables_data_2["bot_diameter"],
        "top_diameter": tables_data_2["top_diameter"],
        "pole_heigth": tables_data_2["pole_heigth"],
        "davit_heigth": tables_data_2["davit_heigth"],
        "tros_heigth": tables_data_2["ground_wire_heigth"]
    }


def extract_tables(file_name):
    """
    This function finds the tables by slicing the txt
    and returns the list of str with data.
    """
    with open(file_name, "r", encoding="ANSI") as file:
        file_data = []
        for line in file:
            file_data.append(line.rstrip("\n"))

    for i in range(len(file_data)):
        file_data[i].rstrip("\n")
        if re.match("Body panel nr 1 ", file_data[i]):
            body_panel_start = i
        if re.match("Cross-arm nr 1    Local panel nr 1", file_data[i]):
            body_panel_end = i
        if re.match("Cross-arm nr 1    Local panel nr 1", file_data[i]):
            cross_arm_start = i
        if re.match("Max use factors for buckling in Body", file_data[i]):
            cross_arm_end = i
        if re.match("Load case nr 1 ", file_data[i]):
            deflection_start = i
        if re.match("Strength summary of lattice structure members", file_data[i]):
            deflection_end = i
    cross_arm_data = file_data[cross_arm_start:cross_arm_end]
    body_panel_data = file_data[body_panel_start:body_panel_end]
    deflection_data = file_data[deflection_start:deflection_end]

    return {
        "cross_arm_data": cross_arm_data,
        "body_panel_data": body_panel_data,
        "deflection_data": deflection_data
    }


def condigure_body_and_davit_tables(table_data):
    """
    This function gets rid of useless columns and returns the
    list of lists with data.
    """
    for i in range(len(table_data)):
        if len(table_data[i]) > 36:
            table_data[i] = table_data[i][5:8] + table_data[i][16:22]\
            + table_data[i][25:33] + table_data[i][61:66]\
            + table_data[i][66:71] + table_data[i][94:]
            if not table_data[i][18:22].strip():
                table_data[i] = table_data[i][:18] + "0.00" + table_data[i][22:]
            elif not table_data[i][23:27].strip():
                table_data[i] = table_data[i][:23] + "0.00" + table_data[i][27:]
            table_data[i] = table_data[i].split()
        else:
            table_data[i] = table_data[i].split()
    
    return table_data


def extract_body_and_davit_data(file_name):
    """
    This function makes DataFrame with body members and
    retuns DataFrame with max_use_factors.
    """
    tables_data = extract_tables(file_name=file_name)
    body_panel_data = condigure_body_and_davit_tables(tables_data["body_panel_data"])
    cross_arm_data = condigure_body_and_davit_tables(tables_data["cross_arm_data"])
    
    body_df = pd.DataFrame(body_panel_data)
    arm_df = pd.DataFrame(cross_arm_data)

    element = []
    l_buck = []
    f_buck = []
    c_buc = []
    c_ten = []
    profile = []
    leg_use_factor = []
    diagonal_use_factor = []
    horizontal_use_factor = []
    arm_use_factor = []

    for index, row in body_df.iterrows():
        if row[0]:
            if row[0] == "1":
                row[0] = "Пояс"
                leg_use_factor.append(float(row[3]))
                element.append(row[0])
                l_buck.append(abs(int(row[1])))
                f_buck.append(int(row[2]))
                c_buc.append(float(row[3]))
                c_ten.append(float(row[4]))
                profile.append(row[5])
            elif row[0] in "345":
                row[0] = "Раскос"
                element.append(row[0])
                l_buck.append(abs(int(row[1])))
                f_buck.append(int(row[2]))
                c_buc.append(float(row[3]))
                c_ten.append(float(row[4]))
                profile.append(row[5])
                diagonal_use_factor.append(float(row[3]))
            elif row[0] in "678":
                row[0] = "Распор"
                element.append(row[0])
                l_buck.append(abs(int(row[1])))
                f_buck.append(int(row[2]))
                c_buc.append(float(row[3]))
                c_ten.append(float(row[4]))
                profile.append(row[5])
                horizontal_use_factor.append(float(row[3]))
            else:
                body_df = body_df.drop(index)
        else:
            body_df = body_df.drop(index)

    for index, row in arm_df.iterrows():
        if row[0]:
            if row[0] in "1345678":
                arm_use_factor.append(float(row[3]))
            else:
                arm_df = arm_df.drop(index)
        else:
            arm_df = arm_df.drop(index)

    appendix_1 = pd.DataFrame(
        {
        "Элемент": element,
        "Расчетная длина, мм": l_buck,
        "Сжимающее усилие, Н": f_buck,
        "Процент использования несущей способности\
 при потери устойчивости": c_buc,
        "Процент использования несущей способности\
 при растяжении": c_ten,
        "Профиль": profile
        }
    )

    appendix_1.drop_duplicates(inplace=True)

    return {
        "leg_use_factor": leg_use_factor,
        "diagonal_use_factor": diagonal_use_factor,
        "horizontal_use_factor": horizontal_use_factor,
        "arm_use_factor": arm_use_factor,
        "appendix_1": appendix_1
    }


def extract_deflection_data(file_name,
                            lower=None,
                            middle=None,
                            upper=None,
                            ground=None):
    deflection_list = extract_tables(file_name=file_name)
    deflection_data = []
    
    for i in range(len(deflection_list["deflection_data"])):
        if deflection_list["deflection_data"][i][:4].strip().isdigit():
            deflection_data.append(deflection_list["deflection_data"][i].split())

    deflection_df = pd.DataFrame(
        deflection_data,
        columns=[chr(i) for i in range(65, 72)]
    )

    for idx, row in deflection_df.iterrows():
        row["B"] = abs(float(row["B"]))
        row["C"] = abs(float(row["C"]))
        row["D"] = abs(float(row["D"]))
        row["E"] = abs(float(row["E"]))

    deflection_max = deflection_df.groupby("B")[["C", "D", "E"]].transform("max")
    deflection_df = pd.concat([deflection_df["B"], deflection_max], axis=1)
    deflection_df["B"] = deflection_df["B"].astype(float)

    tower_deflection = deflection_df.loc[deflection_df["B"].nlargest(1).index[0], "C"]
    davit_lower_deflection = ""
    davit_middle_deflection = ""
    davit_upper_deflection = ""
    davit_ground_deflection = ""
    if lower is not None and lower[0]:
        davit_lower_deflection = deflection_df.loc[deflection_df["B"] == float(lower[0]), "E"].max()
    if middle is not None and middle[0]:
        davit_middle_deflection = deflection_df.loc[deflection_df["B"] == float(middle[0]), "E"].max()
    if upper is not None and upper[0]:
        davit_upper_deflection = deflection_df.loc[deflection_df["B"] == float(upper[0]), "E"].max()
    if ground is not None and ground[0]:
        davit_ground_deflection = deflection_df.loc[deflection_df["B"] == float(ground[0]), "E"].max()

    return {
        "tower": tower_deflection,
        "lower": davit_lower_deflection,
        "middle": davit_middle_deflection,
        "upper": davit_upper_deflection,
        "ground": davit_ground_deflection
    }


def put_data(
    project_name,
    project_code,
    pole_code,
    pole_type,
    developer,
    voltage,
    area,
    branches,
    wind_region,
    wind_pressure,
    ice_region,
    ice_thickness,
    ice_wind_pressure,
    year_average_temp,
    min_temp,
    max_temp,
    ice_temp,
    wind_temp,
    wind_reg_coef,
    ice_reg_coef,
    wire_hesitation,
    seismicity,
    wire,
    wire_tencion,
    ground_wire,
    sections,
    pole_height,
    pole_base,
    pole_top,
    fracture_1,
    fracture_2,
    height_davit_low,
    height_davit_mid,
    height_davit_up,
    length_davit_low_r,
    length_davit_low_l,
    length_davit_mid_r,
    length_davit_mid_l,
    length_davit_up_r,
    length_davit_up_l,
    length_davit_ground_r,
    length_davit_ground_l,
    wind_span,
    weight_span,
    pole,
    loads_str,
    path_to_txt
):
    """
    This fucntion generates .docx file and saves it.
    """
    if path_to_txt:
        doc = DocxTemplate("template.docx")

        if pole_type == "Анкерно-угловая":
            pole_type_variation = ["анкерно-угловых", "анкерно-угловой", "анкерно-угловая"]
        elif pole_type == "Концевая":
            pole_type_variation = ["концевых", "концевой", "концевая"]
        elif pole_type == "Отпаечная":
            pole_type_variation = ["отпаечных", "отпаечной", "отпаечная"]
        else:
            pole_type_variation = ["промежуточных", "промежуточной", "промежуточная"]
        wind_coef = "1" if branches=="1" else "1.1"
        ice_coef_1 = "1" if branches=="1" else "1.3"
        ice_coef_2 = "1.3" if ice_region in ["I", "II"] else "1.6"
        safety_coef = 1.1 if int(voltage) >= 330 else 1
        davit_dict = {
            "lower": (height_davit_low, length_davit_low_r, length_davit_low_l),
            "middle": (height_davit_mid, length_davit_mid_r, length_davit_mid_l),
            "upper": (height_davit_up, length_davit_up_r, length_davit_up_l)
        } if branches == "2" or height_davit_mid  else {
            "lower": (height_davit_low, length_davit_low_r, length_davit_low_l),
            "upper": (height_davit_up, length_davit_up_r, length_davit_up_l)
        }
        if not davit_dict["lower"][0]:
            del davit_dict["lower"]
        if not davit_dict["middle"][0]:
            del davit_dict["middle"]
        if not davit_dict["upper"][0]:
            del davit_dict["upper"]
        if length_davit_ground_r or length_davit_ground_l:
            davit_dict["ground"] = (pole_height, length_davit_ground_r, length_davit_ground_l)
        wire_factor = int(wire.split()[1].split("/")[0]) if wire else ""
        if loads_str:
            loads = [pic_dir.strip("}{") for pic_dir in loads_str.split("} {")]
        # need to check for loads length for exceptions
        if pole_type in ["Анкерно-угловая", "Отпаечная"] and branches == "1"\
        and wire_factor < 180 and ground_wire:
            loads_case_dict = {
                "max_wind": "Нормальный режим (провод и трос не оборваны"\
                " и свободны от гололеда). Максимальный ветер;",
                "max_wind_45": "Нормальный режим (провод и трос не оборваны"\
                " и свободны от гололеда). Максимальный ветер под углом 45°;",
                "ice_wind": "Нормальный режим (провод и трос не оборваны"\
                "и покрыты льдом). Ветер и гололед;",
                "ice_wind_45": "Нормальный режим (провод и трос не оборваны"\
                " и покрыты льдом). Ветер и гололед под углом 45°;",
                "upper_lower":\
                "Аварийный режим (обрыв верхней и нижней фазы провода);",
                "ground_wire": "Аварийный режим (обрыв троса);",
                "installation": "Монтажный режим"
            }
            if loads:
                loads_pic_dict = {
                    "max_wind_45": [
                        InlineImage(doc, image_descriptor=loads[0], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.1 Максимальный ветер",
                        InlineImage(doc, image_descriptor=loads[1], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.2 Максимальный ветер под углом 45°"
                    ],
                    "ice_wind_45": [
                        InlineImage(doc, image_descriptor=loads[2], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.3 Ветер и гололед",
                        InlineImage(doc, image_descriptor=loads[3], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.4 Ветер и гололед под углом 45°"
                    ],
                    "upper_lower": [
                        InlineImage(doc, image_descriptor=loads[4], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.5 Обрыв верхней и нижней фазы",
                        InlineImage(doc,image_descriptor=loads[5], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.6 Обрыв троса"
                    ],
                    "installation": [
                        InlineImage(doc, image_descriptor=loads[6], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.7 Монтажный режим"
                    ]
                }
        elif pole_type in ["Анкерно-угловая", "Отпаечная"] and branches == "1"\
        and wire_factor < 180 and not ground_wire:
            loads_case_dict = {
                "max_wind": "Нормальный режим (провод и трос не оборваны"\
                " и свободны от гололеда). Максимальный ветер;",
                "max_wind_45": "Нормальный режим (провод и трос не оборваны"\
                " и свободны от гололеда). Максимальный ветер под углом 45°;",
                "ice_wind": "Нормальный режим (провод и трос не оборваны"\
                "и покрыты голольдом). Ветер и гололед;",
                "ice_wind_45": "Нормальный режим (провод и трос не оборваны"\
                " и покрыты льдом). Ветер и гололед под углом 45°;",
                "upper_lower":\
                "Аварийный режим (обрыв верхней и нижней фазы провода);",
                "installation": "Монтажный режим"
            }
            if loads:
                loads_pic_dict = {
                    "max_wind_45": [
                        InlineImage(doc, image_descriptor=loads[0], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.1 Максимальный ветер",
                        InlineImage(doc, image_descriptor=loads[1], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.2 Максимальный ветер под углом 45°"
                    ],
                    "ice_wind_45": [
                        InlineImage(doc, image_descriptor=loads[2], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.3 Ветер и гололед",
                        InlineImage(doc, image_descriptor=loads[3], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.4 Ветер и гололед под углом 45°"
                    ],
                    "upper_lower_installation": [
                        InlineImage(doc, image_descriptor=loads[4], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.5 Обрыв верхней и нижней фазы",
                        InlineImage(doc, image_descriptor=loads[5], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.7 Монтажный режим"
                    ]
                }
        elif pole_type in ["Анкерно-угловая", "Отпаечная"] and branches == "1"\
        and wire_factor > 150 and ground_wire:
            loads_case_dict = {
                "max_wind": "Нормальный режим (провод и трос не оборваны"\
                " и свободны от гололеда). Максимальный ветер;",
                "max_wind_45": "Нормальный режим (провод и трос не оборваны"\
                " и свободны от гололеда). Максимальный ветер под углом 45°;",
                "ice_wind": "Нормальный режим (провод и трос не оборваны"\
                "и покрыты голольдом). Ветер и гололед;",
                "ice_wind_45": "Нормальный режим (провод и трос не оборваны"\
                " и покрыты льдом). Ветер и гололед под углом 45°;",
                "upper": "Аварийный режим (обрыв верхней фазы провода);",
                "lower": "Аварийный режим (обрыв нижней фазы провода);",
                "ground_wire": "Аварийный режим (обрыв троса);",
                "installation": "Монтажный режим"
            }
            if loads:
                loads_pic_dict = {
                    "max_wind_45": [
                        InlineImage(doc, image_descriptor=loads[0], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.1 Максимальный ветер",
                        InlineImage(doc, image_descriptor=loads[1], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.2 Максимальный ветер под углом 45°"
                    ],
                    "ice_wind_45": [
                        InlineImage(doc, image_descriptor=loads[2], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.3 Ветер и гололед",
                        InlineImage(doc, image_descriptor=loads[3], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.4 Ветер и гололед под углом 45°"
                    ],
                    "upper_lower": [
                        InlineImage(doc, image_descriptor=loads[4], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.5 Обрыв верхней фазы",
                        InlineImage(doc, image_descriptor=loads[5], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.5 Обрыв нижней фазы"
                    ],
                    "ground_wire_installation": [
                        InlineImage(doc,image_descriptor=loads[6], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.6 Обрыв троса",
                        InlineImage(doc, image_descriptor=loads[7], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.7 Монтажный режим"
                    ]
                }
        elif pole_type in ["Анкерно-угловая", "Отпаечная"] and branches == "1"\
        and wire_factor > 150 and not ground_wire:
            loads_case_dict = {
                "max_wind": "Нормальный режим (провод и трос не оборваны"\
                " и свободны от гололеда). Максимальный ветер;",
                "max_wind_45": "Нормальный режим (провод и трос не оборваны"\
                " и свободны от гололеда). Максимальный ветер под углом 45°;",
                "ice_wind": "Нормальный режим (провод и трос не оборваны"\
                "и покрыты голольдом). Ветер и гололед;",
                "ice_wind_45": "Нормальный режим (провод и трос не оборваны"\
                " и покрыты льдом). Ветер и гололед под углом 45°;",
                "upper": "Аварийный режим (обрыв верхней фазы провода);",
                "lower": "Аварийный режим (обрыв нижней фазы провода);",
                "installation": "Монтажный режим"
            }
            if loads:
                loads_pic_dict = {
                    "max_wind": [
                        InlineImage(doc, image_descriptor=loads[0], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.1 Максимальный ветер",
                        InlineImage(doc, image_descriptor=loads[1], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.2 Максимальный ветер под углом 45°"
                    ],
                    "ice_wind": [
                        InlineImage(doc, image_descriptor=loads[2], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.3 Ветер и гололед",
                        InlineImage(doc, image_descriptor=loads[3], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.4 Ветер и гололед под углом 45°"
                    ],
                    "upper": [
                        InlineImage(doc, image_descriptor=loads[4], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.5 Обрыв верхней фазы",
                        InlineImage(doc, image_descriptor=loads[5], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.5 Обрыв нижней фазы"
                    ],
                    "installation": [
                        InlineImage(doc, image_descriptor=loads[6], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.7 Монтажный режим"
                    ]
                }
        elif pole_type in ["Анкерно-угловая", "Отпаечная"] and branches == "2"\
        and wire_factor < 180 and ground_wire:
            loads_case_dict = {
                "max_wind": "Нормальный режим (провод и трос не оборваны"\
                " и свободны от гололеда). Максимальный ветер;",
                "max_wind_45": "Нормальный режим (провод и трос не оборваны"\
                " и свободны от гололеда). Максимальный ветер под углом 45°;",
                "ice_wind": "Нормальный режим (провод и трос не оборваны"\
                "и покрыты голольдом). Ветер и гололед;",
                "ice_wind_45": "Нормальный режим (провод и трос не оборваны"\
                " и покрыты льдом). Ветер и гололед под углом 45°;",
                "upper_middle":\
                "Аварийный режим (обрыв верхней и средней фазы провода);",
                "ground_wire": "Аварийный режим (обрыв троса);",
                "installation": "Монтажный режим"
            }
            if loads:
                loads_pic_dict = {
                    "max_wind": [
                        InlineImage(doc, image_descriptor=loads[0], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.1 Максимальный ветер",
                        InlineImage(doc, image_descriptor=loads[1], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.2 Максимальный ветер под углом 45°"
                    ],
                    "ice_wind": [
                        InlineImage(doc, image_descriptor=loads[2], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.3 Ветер и гололед",
                        InlineImage(doc, image_descriptor=loads[3], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.4 Ветер и гололед под углом 45°"
                    ],
                    "upper_middle": [
                        InlineImage(doc, image_descriptor=loads[4], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.5 Обрыв верхней и средней фазы",
                        InlineImage(doc,image_descriptor=loads[5], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.6 Обрыв троса"
                    ],
                    "installation": [
                        InlineImage(doc, image_descriptor=loads[6], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.7 Монтажный режим"
                    ]
                }
        elif pole_type in ["Анкерно-угловая", "Отпаечная"] and branches == "2"\
        and wire_factor < 180 and not ground_wire:
            loads_case_dict = {
                "max_wind": "Нормальный режим (провод и трос не оборваны"\
                " и свободны от гололеда). Максимальный ветер;",
                "max_wind_45": "Нормальный режим (провод и трос не оборваны"\
                " и свободны от гололеда). Максимальный ветер под углом 45°;",
                "ice_wind": "Нормальный режим (провод и трос не оборваны"\
                "и покрыты голольдом). Ветер и гололед;",
                "ice_wind_45": "Нормальный режим (провод и трос не оборваны"\
                " и покрыты льдом). Ветер и гололед под углом 45°;",
                "upper_middle":\
                "Аварийный режим (обрыв верхней и средней фазы провода);",
                "installation": "Монтажный режим"
            }
            if loads:
                loads_pic_dict = {
                    "max_wind": [
                        InlineImage(doc, image_descriptor=loads[0], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.1 Максимальный ветер",
                        InlineImage(doc, image_descriptor=loads[1], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.2 Максимальный ветер под углом 45°"
                    ],
                    "ice_wind": [
                        InlineImage(doc, image_descriptor=loads[2], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.3 Ветер и гололед",
                        InlineImage(doc, image_descriptor=loads[3], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.4 Ветер и гололед под углом 45°"
                    ],
                    "upper_middle": [
                        InlineImage(doc, image_descriptor=loads[4], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.5 Обрыв верхней и средней фазы",
                        InlineImage(doc, image_descriptor=loads[5], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.7 Монтажный режим"
                    ]
                }
        elif pole_type in ["Анкерно-угловая", "Отпаечная"] and branches == "2"\
        and wire_factor > 150 and ground_wire:
            loads_case_dict = {
                "max_wind": "Нормальный режим (провод и трос не оборваны"\
                " и свободны от гололеда). Максимальный ветер;",
                "max_wind_45": "Нормальный режим (провод и трос не оборваны"\
                " и свободны от гололеда). Максимальный ветер под углом 45°;",
                "ice_wind": "Нормальный режим (провод и трос не оборваны"\
                "и покрыты голольдом). Ветер и гололед;",
                "ice_wind_45": "Нормальный режим (провод и трос не оборваны"\
                " и покрыты льдом). Ветер и гололед под углом 45°;",
                "upper": "Аварийный режим (обрыв верхней фазы провода);",
                "middle": "Аварийный режим (обрыв средней фазы провода);",
                "ground_wire": "Аварийный режим (обрыв троса);",
                "installation": "Монтажный режим"
            }
            if loads:
                loads_pic_dict = {
                    "max_wind": [
                        InlineImage(doc, image_descriptor=loads[0], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.1 Максимальный ветер",
                        InlineImage(doc, image_descriptor=loads[1], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.2 Максимальный ветер под углом 45°"
                    ],
                    "ice_wind": [
                        InlineImage(doc, image_descriptor=loads[2], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.3 Ветер и гололед",
                        InlineImage(doc, image_descriptor=loads[3], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.4 Ветер и гололед под углом 45°"
                    ],
                    "upper": [
                        InlineImage(doc, image_descriptor=loads[4], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.5 Обрыв верхней фазы",
                        InlineImage(doc, image_descriptor=loads[5], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.5 Обрыв средней фазы"
                    ],
                    "ground_wire": [
                        InlineImage(doc,image_descriptor=loads[6], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.6 Обрыв троса",
                        InlineImage(doc, image_descriptor=loads[7], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.7 Монтажный режим"
                    ]
                }
        elif pole_type in ["Анкерно-угловая", "Отпаечная"] and branches == "2"\
        and wire_factor > 150 and not ground_wire:
            loads_case_dict = {
                "max_wind": "Нормальный режим (провод и трос не оборваны"\
                " и свободны от гололеда). Максимальный ветер;",
                "max_wind_45": "Нормальный режим (провод и трос не оборваны"\
                " и свободны от гололеда). Максимальный ветер под углом 45°;",
                "ice_wind": "Нормальный режим (провод и трос не оборваны"\
                "и покрыты голольдом). Ветер и гололед;",
                "ice_wind_45": "Нормальный режим (провод и трос не оборваны"\
                " и покрыты льдом). Ветер и гололед под углом 45°;",
                "upper": "Аварийный режим (обрыв верхней фазы провода);",
                "middle": "Аварийный режим (обрыв средней фазы провода);",
                "installation": "Монтажный режим"
            }
            if loads:
                loads_pic_dict = {
                    "max_wind": [
                        InlineImage(doc, image_descriptor=loads[0], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.1 Максимальный ветер",
                        InlineImage(doc, image_descriptor=loads[1], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.2 Максимальный ветер под углом 45°"
                    ],
                    "ice_wind": [
                        InlineImage(doc, image_descriptor=loads[2], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.3 Ветер и гололед",
                        InlineImage(doc, image_descriptor=loads[3], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.4 Ветер и гололед под углом 45°"
                    ],
                    "upper": [
                        InlineImage(doc, image_descriptor=loads[4], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.5 Обрыв верхней фазы",
                        InlineImage(doc, image_descriptor=loads[5], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.5 Обрыв средней фазы"
                    ],
                    "installation": [
                        InlineImage(doc, image_descriptor=loads[6], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.7 Монтажный режим"
                    ]
                }
        elif pole_type == "Промежуточная" and branches == "1"\
        and wire_factor < 180 and ground_wire:
            loads_case_dict = {
                "max_wind": "Нормальный режим (провод и трос не оборваны"\
                " и свободны от гололеда). Максимальный ветер;",
                "max_wind_45": "Нормальный режим (провод и трос не оборваны"\
                " и свободны от гололеда). Максимальный ветер под углом 45°;",
                "ice_wind": "Нормальный режим (провод и трос не оборваны"\
                "и покрыты льдом). Ветер и гололед;",
                "ice_wind_45": "Нормальный режим (провод и трос не оборваны"\
                " и покрыты льдом). Ветер и гололед под углом 45°;",
                "upper_lower":\
                "Аварийный режим (обрыв верхней и нижней фазы провода);",
                "ground_wire": "Аварийный режим (обрыв троса);"
            }
            if loads:
                loads_pic_dict = {
                    "max_wind_45": [
                        InlineImage(doc, image_descriptor=loads[0], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.1 Максимальный ветер",
                        InlineImage(doc, image_descriptor=loads[1], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.2 Максимальный ветер под углом 45°"
                    ],
                    "ice_wind_45": [
                        InlineImage(doc, image_descriptor=loads[2], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.3 Ветер и гололед",
                        InlineImage(doc, image_descriptor=loads[3], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.4 Ветер и гололед под углом 45°"
                    ],
                    "upper_lower": [
                        InlineImage(doc, image_descriptor=loads[4], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.5 Обрыв верхней и нижней фазы",
                        InlineImage(doc,image_descriptor=loads[5], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.6 Обрыв троса"
                    ],
                    "installation": [
                        InlineImage(doc, image_descriptor=loads[6], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.7 Монтажный режим"
                    ]
                }
        elif pole_type == "Промежуточная" and branches == "1"\
        and wire_factor < 180 and not ground_wire:
            loads_case_dict = {
                "max_wind": "Нормальный режим (провод и трос не оборваны"\
                " и свободны от гололеда). Максимальный ветер;",
                "max_wind_45": "Нормальный режим (провод и трос не оборваны"\
                " и свободны от гололеда). Максимальный ветер под углом 45°;",
                "ice_wind": "Нормальный режим (провод и трос не оборваны"\
                "и покрыты голольдом). Ветер и гололед;",
                "ice_wind_45": "Нормальный режим (провод и трос не оборваны"\
                " и покрыты льдом). Ветер и гололед под углом 45°;",
                "upper_lower":\
                "Аварийный режим (обрыв верхней и нижней фазы провода);"
            }
            if loads:
                loads_pic_dict = {
                    "max_wind": [
                        InlineImage(doc, image_descriptor=loads[0], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.1 Максимальный ветер",
                        InlineImage(doc, image_descriptor=loads[1], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.2 Максимальный ветер под углом 45°"
                    ],
                    "ice_wind": [
                        InlineImage(doc, image_descriptor=loads[2], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.3 Ветер и гололед",
                        InlineImage(doc, image_descriptor=loads[3], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.4 Ветер и гололед под углом 45°"
                    ],
                    "upper_lower": [
                        InlineImage(doc, image_descriptor=loads[4], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.5 Обрыв верхней и нижней фазы"
                    ]
                }
        elif pole_type == "Промежуточная" and branches == "1"\
        and wire_factor > 150 and ground_wire:
            loads_case_dict = {
                "max_wind": "Нормальный режим (провод и трос не оборваны"\
                " и свободны от гололеда). Максимальный ветер;",
                "max_wind_45": "Нормальный режим (провод и трос не оборваны"\
                " и свободны от гололеда). Максимальный ветер под углом 45°;",
                "ice_wind": "Нормальный режим (провод и трос не оборваны"\
                "и покрыты голольдом). Ветер и гололед;",
                "ice_wind_45": "Нормальный режим (провод и трос не оборваны"\
                " и покрыты льдом). Ветер и гололед под углом 45°;",
                "upper": "Аварийный режим (обрыв верхней фазы провода);",
                "lower": "Аварийный режим (обрыв нижней фазы провода);",
                "ground_wire": "Аварийный режим (обрыв троса);"
            }
            if loads:
                loads_pic_dict = {
                    "max_wind": [
                        InlineImage(doc, image_descriptor=loads[0], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.1 Максимальный ветер",
                        InlineImage(doc, image_descriptor=loads[1], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.2 Максимальный ветер под углом 45°"
                    ],
                    "ice_wind": [
                        InlineImage(doc, image_descriptor=loads[2], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.3 Ветер и гололед",
                        InlineImage(doc, image_descriptor=loads[3], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.4 Ветер и гололед под углом 45°"
                    ],
                    "upper": [
                        InlineImage(doc, image_descriptor=loads[4], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.5 Обрыв верхней фазы",
                        InlineImage(doc, image_descriptor=loads[5], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.5 Обрыв нижней фазы"
                    ],
                    "ground_wire": [
                        InlineImage(doc,image_descriptor=loads[6], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.6 Обрыв троса"
                    ]
                }
        elif pole_type == "Промежуточная" and branches == "1"\
        and wire_factor > 150 and not ground_wire:
            loads_case_dict = {
                "max_wind": "Нормальный режим (провод и трос не оборваны"\
                " и свободны от гололеда). Максимальный ветер;",
                "max_wind_45": "Нормальный режим (провод и трос не оборваны"\
                " и свободны от гололеда). Максимальный ветер под углом 45°;",
                "ice_wind": "Нормальный режим (провод и трос не оборваны"\
                "и покрыты голольдом). Ветер и гололед;",
                "ice_wind_45": "Нормальный режим (провод и трос не оборваны"\
                " и покрыты льдом). Ветер и гололед под углом 45°;",
                "upper": "Аварийный режим (обрыв верхней фазы провода);",
                "lower": "Аварийный режим (обрыв нижней фазы провода);"
            }
            if loads:
                loads_pic_dict = {
                    "max_wind": [
                        InlineImage(doc, image_descriptor=loads[0], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.1 Максимальный ветер",
                        InlineImage(doc, image_descriptor=loads[1], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.2 Максимальный ветер под углом 45°"
                    ],
                    "ice_wind": [
                        InlineImage(doc, image_descriptor=loads[2], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.3 Ветер и гололед",
                        InlineImage(doc, image_descriptor=loads[3], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.4 Ветер и гололед под углом 45°"
                    ],
                    "upper": [
                        InlineImage(doc, image_descriptor=loads[4], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.5 Обрыв верхней фазы",
                        InlineImage(doc, image_descriptor=loads[5], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.5 Обрыв нижней фазы"
                    ]
                }
        elif pole_type == "Промежуточная" and branches == "2"\
        and wire_factor < 180 and ground_wire:
            loads_case_dict = {
                "max_wind": "Нормальный режим (провод и трос не оборваны"\
                " и свободны от гололеда). Максимальный ветер;",
                "max_wind_45": "Нормальный режим (провод и трос не оборваны"\
                " и свободны от гололеда). Максимальный ветер под углом 45°;",
                "ice_wind": "Нормальный режим (провод и трос не оборваны"\
                "и покрыты голольдом). Ветер и гололед;",
                "ice_wind_45": "Нормальный режим (провод и трос не оборваны"\
                " и покрыты льдом). Ветер и гололед под углом 45°;",
                "upper_middle":\
                "Аварийный режим (обрыв верхней и средней фазы провода);",
                "ground_wire": "Аварийный режим (обрыв троса);",
            }
            if loads:
                loads_pic_dict = {
                    "max_wind": [
                        InlineImage(doc, image_descriptor=loads[0], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.1 Максимальный ветер",
                        InlineImage(doc, image_descriptor=loads[1], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.2 Максимальный ветер под углом 45°"
                    ],
                    "ice_wind": [
                        InlineImage(doc, image_descriptor=loads[2], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.3 Ветер и гололед",
                        InlineImage(doc, image_descriptor=loads[3], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.4 Ветер и гололед под углом 45°"
                    ],
                    "upper_middle": [
                        InlineImage(doc, image_descriptor=loads[4], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.5 Обрыв верхней и средней фазы",
                        InlineImage(doc,image_descriptor=loads[5], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.6 Обрыв троса"
                    ]
                }
        elif pole_type == "Промежуточная" and branches == "2"\
        and wire_factor < 180 and not ground_wire:
            loads_case_dict = {
                "max_wind": "Нормальный режим (провод и трос не оборваны"\
                " и свободны от гололеда). Максимальный ветер;",
                "max_wind_45": "Нормальный режим (провод и трос не оборваны"\
                " и свободны от гололеда). Максимальный ветер под углом 45°;",
                "ice_wind": "Нормальный режим (провод и трос не оборваны"\
                "и покрыты голольдом). Ветер и гололед;",
                "ice_wind_45": "Нормальный режим (провод и трос не оборваны"\
                " и покрыты льдом). Ветер и гололед под углом 45°;",
                "upper_middle":\
                "Аварийный режим (обрыв верхней и средней фазы провода);"
            }
            if loads:
                loads_pic_dict = {
                    "max_wind": [
                        InlineImage(doc, image_descriptor=loads[0], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.1 Максимальный ветер",
                        InlineImage(doc, image_descriptor=loads[1], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.2 Максимальный ветер под углом 45°"
                    ],
                    "ice_wind": [
                        InlineImage(doc, image_descriptor=loads[2], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.3 Ветер и гололед",
                        InlineImage(doc, image_descriptor=loads[3], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.4 Ветер и гололед под углом 45°"
                    ],
                    "upper_middle": [
                        InlineImage(doc, image_descriptor=loads[4], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.5 Обрыв верхней и средней фазы"
                    ]
                }
        elif pole_type == "Промежуточная" and branches == "2"\
        and wire_factor > 150 and ground_wire:
            loads_case_dict = {
                "max_wind": "Нормальный режим (провод и трос не оборваны"\
                " и свободны от гололеда). Максимальный ветер;",
                "max_wind_45": "Нормальный режим (провод и трос не оборваны"\
                " и свободны от гололеда). Максимальный ветер под углом 45°;",
                "ice_wind": "Нормальный режим (провод и трос не оборваны"\
                "и покрыты голольдом). Ветер и гололед;",
                "ice_wind_45": "Нормальный режим (провод и трос не оборваны"\
                " и покрыты льдом). Ветер и гололед под углом 45°;",
                "upper": "Аварийный режим (обрыв верхней фазы провода);",
                "middle": "Аварийный режим (обрыв средней фазы провода);",
                "ground_wire": "Аварийный режим (обрыв троса);"
            }
            if loads:
                loads_pic_dict = {
                    "max_wind": [
                        InlineImage(doc, image_descriptor=loads[0], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.1 Максимальный ветер",
                        InlineImage(doc, image_descriptor=loads[1], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.2 Максимальный ветер под углом 45°"
                    ],
                    "ice_wind": [
                        InlineImage(doc, image_descriptor=loads[2], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.3 Ветер и гололед",
                        InlineImage(doc, image_descriptor=loads[3], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.4 Ветер и гололед под углом 45°"
                    ],
                    "upper": [
                        InlineImage(doc, image_descriptor=loads[4], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.5 Обрыв верхней фазы",
                        InlineImage(doc, image_descriptor=loads[5], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.5 Обрыв средней фазы"
                    ],
                    "ground_wire": [
                        InlineImage(doc,image_descriptor=loads[6], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.6 Обрыв троса"
                    ]
                }
        elif pole_type == "Промежуточная" and branches == "2"\
        and wire_factor > 150 and not ground_wire:
            loads_case_dict = {
                "max_wind": "Нормальный режим (провод и трос не оборваны"\
                " и свободны от гололеда). Максимальный ветер;",
                "max_wind_45": "Нормальный режим (провод и трос не оборваны"\
                " и свободны от гололеда). Максимальный ветер под углом 45°;",
                "ice_wind": "Нормальный режим (провод и трос не оборваны"\
                "и покрыты голольдом). Ветер и гололед;",
                "ice_wind_45": "Нормальный режим (провод и трос не оборваны"\
                " и покрыты льдом). Ветер и гололед под углом 45°;",
                "upper": "Аварийный режим (обрыв верхней фазы провода);",
                "middle": "Аварийный режим (обрыв средней фазы провода);"
            }
            if loads:
                loads_pic_dict = {
                    "max_wind": [
                        InlineImage(doc, image_descriptor=loads[0], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.1 Максимальный ветер",
                        InlineImage(doc, image_descriptor=loads[1], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.2 Максимальный ветер под углом 45°"
                    ],
                    "ice_wind": [
                        InlineImage(doc, image_descriptor=loads[2], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.3 Ветер и гололед",
                        InlineImage(doc, image_descriptor=loads[3], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.4 Ветер и гололед под углом 45°"
                    ],
                    "upper": [
                        InlineImage(doc, image_descriptor=loads[4], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.5 Обрыв верхней фазы",
                        InlineImage(doc, image_descriptor=loads[5], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.5 Обрыв средней фазы"
                    ]
                }
        elif pole_type == "Концевая" and branches == "2"\
        and wire_factor > 150 and not ground_wire:
            loads_case_dict = {
                "max_wind": "Нормальный режим (провод и трос не оборваны"\
                " и свободны от гололеда). Максимальный ветер под углом 90°;",
                "max_wind_0": "Нормальный режим (провод и трос не оборваны"\
                " и свободны от гололеда). Максимальный ветер под углом 0°;",
                "ice_wind": "Нормальный режим (провод и трос не оборваны"\
                "и покрыты голольдом). Ветер под углом 90° и гололед;",
                "ice_wind_0": "Нормальный режим (провод и трос не оборваны"\
                "и покрыты голольдом). Ветер под углом 0° и гололед;"
            }
            if loads:
                loads_pic_dict = {
                    "max_wind": [
                        InlineImage(doc, image_descriptor=loads[0], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.1 Максимальный ветер под углом 90°",
                        InlineImage(doc, image_descriptor=loads[1], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.2 Максимальный ветер под углом 0°"
                    ],
                    "ice_wind": [
                        InlineImage(doc, image_descriptor=loads[2], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.3 Ветер под углом 90° и гололед",
                        InlineImage(doc, image_descriptor=loads[3], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.4 Ветер под углом 0° и гололед"
                    ]
                }
        else:
            loads_case_dict = ""
            loads_pic_dict = ""

        read_txt_dict = extract_body_and_davit_data(file_name=path_to_txt)
        leg_use = round(max(read_txt_dict["leg_use_factor"]) * 100, 1)
        diagonal_use = round(max(read_txt_dict["diagonal_use_factor"]) * 100, 1)
        horizontal_use = round(max(read_txt_dict["horizontal_use_factor"]) * 100, 1)
        arm_use = round(max(read_txt_dict["arm_use_factor"]) * 100, 1)
        for idx, row in read_txt_dict["appendix_1"].iterrows():
            if row[0] == "Пояс":
                leg_steel = row[5].split("/")[1]
            if row[0] == "Раскос":
                diagonal_steel = row[5].split("/")[1]
            if row[0] == "Распор":
                horizontal_steel = row[5].split("/")[1]

        deflection = extract_deflection_data(file_name=path_to_txt, **davit_dict)
        if float(pole_height) <= 60 and pole_type in ["Анкерно-угловая", "Концевая", "Отпаечная"]:
            normative_deflection = "анкерного или концевого типа высотой до 60 м - 1/120"
            normative_deflection_davit = "анкерного или концевого типа высотой до 60 м - 1/70"
            earth = 120
            max_deflection = round((float(pole_height) * 1000) / earth, 1)
            deflection_body = f"Высота опоры {pole_code} h = {pole_height} м.\n" +\
            f"Предельно допустимое отклонение верха опоры определяется по формуле:\n" +\
            f"h/{earth} = {round(float(pole_height)*1000)}/{earth} = {max_deflection} мм\n"
            earth_davit = 70
            if deflection["tower"] <= max_deflection:
                tower_deflection_result = deflection_body + f"{deflection['tower']} мм <= {max_deflection} мм"
                tower_deflection_flag = True
            else:
                tower_deflection_result = deflection_body + f"{deflection['tower']} мм > {max_deflection} мм"
                tower_deflection_flag = False
        elif float(pole_height) > 60:
            normative_deflection = "любого типа высотой выше 60 м - 1/140"
            normative_deflection_davit = "анкерного или концевого типа высотой выше 60 м - 1/70"
            earth = 140
            max_deflection = round((float(pole_height) * 1000) / earth, 1)
            deflection_body = f"Высота опоры {pole_code} h = {pole_height} м.\n" +\
            f"Предельно допустимое отклонение верха опоры определяется по формуле:\n" +\
            f"h/{earth} = {round(float(pole_height)*1000)}/{earth} = {max_deflection} мм\n"
            earth_davit = 70
            if deflection["tower"] <= max_deflection:
                tower_deflection_result = deflection_body + f"{deflection['tower']} мм <= {max_deflection} мм"
                tower_deflection_flag = True
            else:
                tower_deflection_result = deflection_body + f"{deflection['tower']} мм > {max_deflection} мм"
                tower_deflection_flag = False
        else:
            normative_deflection = "промежуточного типа не нормируется"
            normative_deflection_davit = "промежуточного типа - 1/50"
            earth_davit = 50
            tower_deflection_flag = True
            tower_deflection_result = ""

        davit_deflection_dict = dict(
            lower=[max(length_davit_low_l, length_davit_low_r), deflection["lower"]],
            middle=[max(length_davit_mid_l, length_davit_mid_r), deflection["middle"]],
            upper=[max(length_davit_up_l, length_davit_up_r), deflection["upper"]],
            ground=[max(length_davit_ground_l, length_davit_ground_r), deflection["ground"]]
        )
        if not davit_deflection_dict["lower"][0]:
            del davit_deflection_dict["lower"]
        if not davit_deflection_dict["middle"][0]:
            del davit_deflection_dict["middle"]
        if not davit_deflection_dict["upper"][0]:
            del davit_deflection_dict["upper"]
        if not davit_deflection_dict["ground"][0]:
            del davit_deflection_dict["ground"]

        davit_deflection_flag = True
        for key in davit_deflection_dict:
            davit_deflection_dict[key].append(
                round(float(davit_deflection_dict[key][0]) * 1000 / earth_davit, 1)
            )
            if davit_deflection_dict[key][1] <= davit_deflection_dict[key][2]:
                davit_deflection_dict[key].append(
                    f"{davit_deflection_dict[key][1]} <= {davit_deflection_dict[key][2]}"
                )
            else:
                davit_deflection_dict[key].append(
                    f"{davit_deflection_dict[key][1]} > {davit_deflection_dict[key][2]}"
                )
                davit_deflection_flag = False

        element_flag = True if leg_use <=100 and diagonal_use <= 100\
        and horizontal_use <= 100 else False

        result = "соответствует" if tower_deflection_flag and davit_deflection_flag\
        and element_flag else "не соответствует"

        context = {
            "project_name": project_name,
            "project_code": project_code,
            "year": dt.date.today().year,
            "pole_code": pole_code,
            "pole_type_1": pole_type_variation[0],
            "pole_type_2": pole_type_variation[1],
            "pole_type_3": pole_type_variation[2],
            "developer": developer,
            "voltage": voltage,
            "area": area,
            "branches": branches,
            "wind_region": wind_region,
            "wind_pressure": wind_pressure,
            "ice_region": ice_region,
            "ice_thickness": ice_thickness,
            "ice_wind_pressure": ice_wind_pressure,
            "year_average_temp": year_average_temp,
            "min_temp": min_temp,
            "max_temp": max_temp,
            "ice_temp": ice_temp,
            "wind_temp": wind_temp,
            "wind_reg_coef": wind_reg_coef,
            "ice_reg_coef": ice_reg_coef,
            "wire_hesitation": wire_hesitation,
            "seismicity": seismicity,
            "wire": wire,
            "wire_tencion": wire_tencion,
            "ground_wire": ground_wire,
            "wind_coef": wind_coef,
            "ice_coef_1": ice_coef_1,
            "ice_coef_2": ice_coef_2,
            "sections": sections,
            "pole_height": pole_height,
            "pole_base": pole_base,
            "pole_top": pole_top,
            "fracture_1": fracture_1,
            "fracture_2": fracture_2,
            "safety_coef": safety_coef,
            "davit_dict": davit_dict,
            "wind_span": wind_span,
            "weight_span": weight_span,
            "loads_case_dict": loads_case_dict,
            "pole_pic": InlineImage(doc, image_descriptor=pole, width=Mm(80), height=Mm(150)),
            "load_pic_dict": loads_pic_dict,
            "leg_use": leg_use,
            "diagonal_use": diagonal_use,
            "horizontal_use": horizontal_use,
            "arm_use": arm_use,
            "tower_deflection": deflection["tower"],
            "normative_deflection": normative_deflection,
            # "max_deflection": max_deflection,
            "tower_deflection_result": tower_deflection_result,
            "davit_deflection_dict": davit_deflection_dict,
            "normative_deflection_davit": normative_deflection_davit,
            "leg_steel": leg_steel,
            "diagonal_steel": diagonal_steel,
            "horizontal_steel": horizontal_steel,
            "result": result
        }

        dir_name = fd.asksaveasfilename(
                filetypes=[("docx file", ".docx")],
                defaultextension=".docx"
            )
        if dir_name:
            doc.render(context)
            doc.save(dir_name)


def generate_appendix(path_to_txt):
    """
    This function generates appendix_1.
    """
    if path_to_txt:
        read_txt_dict = extract_body_and_davit_data(file_name=path_to_txt)
        dir_name = fd.asksaveasfilename(
                filetypes=[("xlsx file", ".xlsx")],
                defaultextension=".xlsx"
            )
        if dir_name:
            with ExcelWriter(dir_name) as writer:
                read_txt_dict["appendix_1"].to_excel(writer, "Sheet1", index=False)


class Kompas_work:
    def __init__(self) -> None:
        pass        

    def concatenate_pdf_files(self,title,folder, output):
        from pdfrw import PdfReader, PdfWriter, IndirectPdfDict
        writer = PdfWriter()
        paths = self.scan_folder(folder+"\\PDF",'*.pdf')
        for path in paths:
            reader = PdfReader(folder+'\\PDF\\'+path)
            writer.addpages(reader.pages)
        writer.trailer.Info = IndirectPdfDict(
            Title=title,
            Author='Amast',
            Subject='PDF Combinations',
            Creator='The Concatenator'
        )
        writer.write(output)

    def scan_folder(self,folder,mask):
        import glob
        temp = [thing.split('\\')[-1] for thing in glob.glob(f"{folder}\{mask}")]
        return temp


class KompasAPI:
    def __init__(self):
        self.connect_api()

    def connect_api(self):
        import pythoncom
        from win32com.client import Dispatch, gencache
        from core import MiscellaneousHelpers as MH
        #  Подключим константы API Компас
        self.kompas6_constants = gencache.EnsureModule("{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0).constants
        self.kompas6_constants_3d = gencache.EnsureModule("{2CAF168C-7961-4B90-9DA2-701419BEEFE3}", 0, 1, 0).constants
        self.kompas6_api5_module,self.kompas_object = self.api5_connect()
        self.kompas_api7_module,self.application = self.api7_connect()
        self.hide_messages_in_kompas()
        self.application.Visible=False

    def api5_connect(self):
        #  Подключим описание интерфейсов API5
        from win32com.client import Dispatch, gencache
        import pythoncom
        from core import MiscellaneousHelpers as MH
        kompas6_api5_module = gencache.EnsureModule("{0422828C-F174-495E-AC5D-D31014DBBE87}", 0, 1, 0)
        kompas_object = kompas6_api5_module.KompasObject(Dispatch("Kompas.Application.5").
                _oleobj_.QueryInterface(kompas6_api5_module.KompasObject.CLSID,pythoncom.IID_IDispatch))
        MH.iKompasObject = kompas_object
        return kompas6_api5_module,kompas_object

    def api7_connect(self):
        #  Подключим описание интерфейсов API7
        from win32com.client import Dispatch, gencache
        import pythoncom
        from core import MiscellaneousHelpers as MH
        kompas_api7_module = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0)
        application = kompas_api7_module.IApplication(Dispatch("Kompas.Application.7").
                _oleobj_.QueryInterface(kompas_api7_module.IApplication.CLSID,pythoncom.IID_IDispatch))
        MH.iApplication = application
        return kompas_api7_module,application

    def visible(self,state:bool):
        self.application.Visible=state

    def open_2D_file(self, path):    
        kompas_document_2d = self.kompas_api7_module.IKompasDocument2D(
                self.application.Documents.Open(PathName=path, Visible=True, ReadOnly=True)
        )
        iDocument2D = self.kompas_object.ActiveDocument2D()
        return iDocument2D

    def hide_messages_in_kompas(self):
        self.application.HideMessage = 1 #скрывает всплывающие сообщения, соглашаясь на них. показывать сообщения - 0, отвечать нет - 2

    def show_messages_in_kompas(self):
        self.application.HideMessage = 0

    def get_2D_file(self):
        kompas_document_2d = self.kompas_api7_module.IKompasDocument2D(
                self.application.ActiveDocument
        )
        iDocument2D = self.kompas_object.ActiveDocument2D()
        return iDocument2D

    def close_2D_file(self):
        self.application.ActiveDocument.Close(False)

    def open_3D_file(self, path):
        kompas_document_3d = self.kompas_api7_module.IKompasDocument3D(
                self.application.Documents.Open(path, True, False)
        )
        iDocument3D = self.kompas_object.ActiveDocument3D()
        return iDocument3D

    def get_3D_file(self):
        kompas_document_3d = self.kompas_api7_module.IKompasDocument3D(
                self.application.ActiveDocument
        )
        iDocument3D = self.kompas_object.ActiveDocument3D()
        return iDocument3D

    def close_assmebly(self):
        self.application.ActiveDocument.Close(False)


class AssemblyAPI:
    def __init__(self,kompas:KompasAPI) -> None:
        self.kompas=kompas

    def change_external_variables(self, dic):
        from core import LDefin3D
        # Присваивание переменных для фундаментов
        # Получаем интерфейс компонента и обновляем коллекцию внешних переменных
        iPart = self.kompas.get_3D_file().GetPart(LDefin3D.pTop_Part)
        VariableCollection = iPart.VariableCollection()
        VariableCollection.refresh()
        # Получаем интерфейс переменной по её имени
        Nb = VariableCollection.GetByName('Nb', True, True)  # 'L22' - имя
        dob = VariableCollection.GetByName('dob', True, True)  # 'L22' - имя
        M = VariableCollection.GetByName('M', True, True)  # 'L22' - имя
        Dfpk = VariableCollection.GetByName('Dfpk', True, True)
        Dn1 = VariableCollection.GetByName('Dn1', True, True)  # 'L1' - имя
        if 'dlina_svai' in dic: L1 = VariableCollection.GetByName('L1', True, True)  # 'L1' - имя
        if 'dk' in dic: dk = VariableCollection.GetByName('dk', True, True)  # 'L1' - имя
        if 'dap' in dic: dap = VariableCollection.GetByName('dap', True, True)  # 'L1' - имя
        if 'doap' in dic: doap = VariableCollection.GetByName('doap', True, True)  # 'L1' - имя
        if 'dok' in dic: dok = VariableCollection.GetByName('dok', True, True)  # 'L1' - имя
        if 'Nap' in dic: Nap = VariableCollection.GetByName('Nap', True, True)  # 'L1' - имя
        if 'Nk' in dic: Nk = VariableCollection.GetByName('Nk', True, True)  # 'L1' - имя
        if 'x' in dic: x = VariableCollection.GetByName('x', True, True)  # 'L1' - имя
        if 'y' in dic: y = VariableCollection.GetByName('y', True, True)  # 'L1' - имя

        # Задаём новое значение переменной
        try:
            Nb.value = dic['bolts_quantity']
            dob.value = dic['hole_diameter']
            M.value = dic['wall_distance']
            Dfpk.value = dic['diameter_flanca']
            Dn1.value = dic['diameter_truby']
        except: pass
        if 'dlina_svai' in dic: L1.value = dic['dlina_svai']
        if 'dk' in dic: dk.value = dic['dk']
        if 'dap' in dic: dap.value = dic['dap']
        if 'doap' in dic: doap.value = dic['doap']
        if 'dok' in dic: dok.value = dic['dok']
        if 'Nap' in dic: Nap.value = dic['Nap']
        if 'Nk' in dic: Nk.value = dic['Nk']
        if 'x' in dic: x.value = dic['x']
        if 'y' in dic: y.value = dic['y']
        # Перестраиваем модель и сохраняем
        iPart.RebuildModel()

    def change_pole_variables(self, dic):
        from core import LDefin3D
        # Присваивание переменных для фундаментов
        # Получаем интерфейс компонента и обновляем коллекцию внешних переменных
        iPart = self.kompas.get_3D_file().GetPart(LDefin3D.pTop_Part)
        VariableCollection = iPart.VariableCollection()
        VariableCollection.refresh()
        # Получаем интерфейс переменной по её имени
        D = VariableCollection.GetByName('Nb', True, True)
        Da = VariableCollection.GetByName('dob', True, True)
        Db = VariableCollection.GetByName('M', True, True)
        Dc = VariableCollection.GetByName('Dc', True, True)
        Dd = VariableCollection.GetByName('Dd', True, True)
        d = VariableCollection.GetByName('d', True, True)
        La = VariableCollection.GetByName('La', True, True)
        Lb = VariableCollection.GetByName('Lb', True, True)
        Lc = VariableCollection.GetByName('Lc', True, True)
        Ld = VariableCollection.GetByName('Ld', True, True)
        Le = VariableCollection.GetByName('Le', True, True)
        Sa = VariableCollection.GetByName('Sa', True, True)
        Sb = VariableCollection.GetByName('Sb', True, True)
        Sc = VariableCollection.GetByName('Sc', True, True)
        Sd = VariableCollection.GetByName('Sd', True, True)
        Se = VariableCollection.GetByName('Se', True, True)
        N1 = VariableCollection.GetByName('N1', True, True)
        N2 = VariableCollection.GetByName('N2', True, True)
        N3 = VariableCollection.GetByName('N3', True, True)
        N4 = VariableCollection.GetByName('N4', True, True)
        N5 = VariableCollection.GetByName('N5', True, True)
        D1 = VariableCollection.GetByName('D1', True, True)
        D2 = VariableCollection.GetByName('D2', True, True)
        D3 = VariableCollection.GetByName('D3', True, True)
        D4 = VariableCollection.GetByName('D4', True, True)
        D5 = VariableCollection.GetByName('D5', True, True)
        S1 = VariableCollection.GetByName('S1', True, True)
        S2 = VariableCollection.GetByName('S2', True, True)
        S3 = VariableCollection.GetByName('S3', True, True)
        S4 = VariableCollection.GetByName('S4', True, True)
        S5 = VariableCollection.GetByName('S5', True, True)
        Ht1 = VariableCollection.GetByName('Ht1', True, True)
        Ht2 = VariableCollection.GetByName('Ht2', True, True)
        Ht3 = VariableCollection.GetByName('Ht3', True, True)
        Ht4 = VariableCollection.GetByName('Ht4', True, True)
        Ht5 = VariableCollection.GetByName('Ht5', True, True)
        Ht6 = VariableCollection.GetByName('Ht6', True, True)
        Ht7 = VariableCollection.GetByName('Ht7', True, True)
        Dt1 = VariableCollection.GetByName('Dt1', True, True)
        dt1 = VariableCollection.GetByName('dt1', True, True)
        Dt2 = VariableCollection.GetByName('Dt2', True, True)
        dt2 = VariableCollection.GetByName('dt2', True, True)
        Dt3 = VariableCollection.GetByName('Dt3', True, True)
        dt3 = VariableCollection.GetByName('dt3', True, True)
        Dt4 = VariableCollection.GetByName('Dt4', True, True)
        dt4 = VariableCollection.GetByName('dt4', True, True)
        Dt5 = VariableCollection.GetByName('Dt5', True, True)
        dt5 = VariableCollection.GetByName('dt5', True, True)
        Dt6 = VariableCollection.GetByName('Dt6', True, True)
        dt6 = VariableCollection.GetByName('dt6', True, True)
        Dt7 = VariableCollection.GetByName('Dt7', True, True)
        dt7 = VariableCollection.GetByName('dt7', True, True)
        C1 = VariableCollection.GetByName('C1', True, True)
        C2 = VariableCollection.GetByName('C2', True, True)
        C3 = VariableCollection.GetByName('C3', True, True)
        C4 = VariableCollection.GetByName('C4', True, True)
        C5 = VariableCollection.GetByName('C5', True, True)
        C6 = VariableCollection.GetByName('C6', True, True)
        C7 = VariableCollection.GetByName('C7', True, True)
        Tt1 = VariableCollection.GetByName('Tt1', True, True)
        Tt2 = VariableCollection.GetByName('Tt2', True, True)
        Tt3 = VariableCollection.GetByName('Tt3', True, True)
        Tt4 = VariableCollection.GetByName('Tt4', True, True)
        Tt5 = VariableCollection.GetByName('Tt5', True, True)
        Tt6 = VariableCollection.GetByName('Tt6', True, True)
        Tt7 = VariableCollection.GetByName('Tt7', True, True)
        Nt1 = VariableCollection.GetByName('Nt1', True, True)
        Nt2 = VariableCollection.GetByName('Nt2', True, True)
        Nt3 = VariableCollection.GetByName('Nt3', True, True)
        Nt4 = VariableCollection.GetByName('Nt4', True, True)
        Nt5 = VariableCollection.GetByName('Nt5', True, True)
        Nt6 = VariableCollection.GetByName('Nt6', True, True)
        Nt7 = VariableCollection.GetByName('Nt7', True, True)
        Dbt1 = VariableCollection.GetByName('Dbt1', True, True)
        Dbt2 = VariableCollection.GetByName('Dbt2', True, True)
        Dbt3 = VariableCollection.GetByName('Dbt3', True, True)
        Dbt4 = VariableCollection.GetByName('Dbt4', True, True)
        Dbt5 = VariableCollection.GetByName('Dbt5', True, True)
        Dbt6 = VariableCollection.GetByName('Dbt6', True, True)
        Dbt7 = VariableCollection.GetByName('Dbt7', True, True)
        L1 = VariableCollection.GetByName('L1', True, True)
        L2 = VariableCollection.GetByName('L2', True, True)
        L3 = VariableCollection.GetByName('L3', True, True)
        L4 = VariableCollection.GetByName('L4', True, True)
        L5 = VariableCollection.GetByName('L5', True, True)
        L6 = VariableCollection.GetByName('L6', True, True)
        L7 = VariableCollection.GetByName('L7', True, True)
        L8 = VariableCollection.GetByName('L8', True, True)
        L9 = VariableCollection.GetByName('L9', True, True)
        L10 = VariableCollection.GetByName('L10', True, True)
        L11 = VariableCollection.GetByName('L11', True, True)
        L12 = VariableCollection.GetByName('L12', True, True)
        L13 = VariableCollection.GetByName('L13', True, True)
        L14 = VariableCollection.GetByName('L14', True, True)
        T1 = VariableCollection.GetByName('T1', True, True)
        T2 = VariableCollection.GetByName('T2', True, True)
        T3 = VariableCollection.GetByName('T3', True, True)
        T4 = VariableCollection.GetByName('T4', True, True)
        T5 = VariableCollection.GetByName('T5', True, True)
        T6 = VariableCollection.GetByName('T6', True, True)
        T7 = VariableCollection.GetByName('T7', True, True)
        Yz1 = VariableCollection.GetByName('Yz1', True, True)
        Yz2 = VariableCollection.GetByName('Yz2', True, True)
        Yz3 = VariableCollection.GetByName('Yz3', True, True)
        Yz4 = VariableCollection.GetByName('Yz4', True, True)
        Yz5 = VariableCollection.GetByName('Yz5', True, True)
        Yz6 = VariableCollection.GetByName('Yz6', True, True)
        Yz7 = VariableCollection.GetByName('Yz7', True, True)
        Yz8 = VariableCollection.GetByName('Yz8', True, True)
        Yz9 = VariableCollection.GetByName('Yz9', True, True)
        Yz10 = VariableCollection.GetByName('Yz10', True, True)
        Yz11 = VariableCollection.GetByName('Yz11', True, True)
        Yz12 = VariableCollection.GetByName('Yz12', True, True)
        Yz13 = VariableCollection.GetByName('Yz13', True, True)
        Yz14 = VariableCollection.GetByName('Yz14', True, True)
        Yz15 = VariableCollection.GetByName('Yz15', True, True)
        Yz16 = VariableCollection.GetByName('Yz16', True, True)
        Yz17 = VariableCollection.GetByName('Yz17', True, True)
        Yz18 = VariableCollection.GetByName('Yz18', True, True)
        Yz19 = VariableCollection.GetByName('Yz19', True, True)
        Yz20 = VariableCollection.GetByName('Yz20', True, True)
        Yz21 = VariableCollection.GetByName('Yz21', True, True)
        Yz22 = VariableCollection.GetByName('Yz22', True, True)
        Yz23 = VariableCollection.GetByName('Yz23', True, True)
        Yz24 = VariableCollection.GetByName('Yz24', True, True)
        Yz25 = VariableCollection.GetByName('Yz25', True, True)
        Yz26 = VariableCollection.GetByName('Yz26', True, True)
        Yz27 = VariableCollection.GetByName('Yz27', True, True)
        Yz28 = VariableCollection.GetByName('Yz28', True, True)
        Yz29 = VariableCollection.GetByName('Yz29', True, True)
        Yz30 = VariableCollection.GetByName('Yz30', True, True)
        Yz31 = VariableCollection.GetByName('Yz31', True, True)
        Yz32 = VariableCollection.GetByName('Yz32', True, True)
        Yz36 = VariableCollection.GetByName('Yz36', True, True)
        Yz37 = VariableCollection.GetByName('Yz37', True, True)
        Yz41 = VariableCollection.GetByName('Yz41', True, True)
        Yz42 = VariableCollection.GetByName('Yz42', True, True)
        Yz46 = VariableCollection.GetByName('Yz46', True, True)
        Yz47 = VariableCollection.GetByName('Yz47', True, True)
        Yz51 = VariableCollection.GetByName('Yz51', True, True)
        Yz52 = VariableCollection.GetByName('Yz52', True, True)
        Yz56 = VariableCollection.GetByName('Yz56', True, True)
        Yz57 = VariableCollection.GetByName('Yz57', True, True)
        Yz61 = VariableCollection.GetByName('Yz61', True, True)
        Yz62 = VariableCollection.GetByName('Yz62', True, True)
        Yz66 = VariableCollection.GetByName('Yz66', True, True)
        Yz67 = VariableCollection.GetByName('Yz67', True, True)
        Pz1 = VariableCollection.GetByName('Yz1', True, True)
        Pz2 = VariableCollection.GetByName('Yz2', True, True)
        Pz3 = VariableCollection.GetByName('Yz3', True, True)
        Pz4 = VariableCollection.GetByName('Yz4', True, True)
        Pz5 = VariableCollection.GetByName('Yz5', True, True)
        Pz6 = VariableCollection.GetByName('Yz6', True, True)
        Pz7 = VariableCollection.GetByName('Yz7', True, True)
        Pz8 = VariableCollection.GetByName('Yz8', True, True)
        Pz9 = VariableCollection.GetByName('Yz9', True, True)
        Pz10 = VariableCollection.GetByName('Yz10', True, True)
        Pz11 = VariableCollection.GetByName('Yz11', True, True)
        Pz12 = VariableCollection.GetByName('Yz12', True, True)
        Pz13 = VariableCollection.GetByName('Yz13', True, True)
        Pz14 = VariableCollection.GetByName('Yz14', True, True)
        Pz15 = VariableCollection.GetByName('Yz15', True, True)
        Pz16 = VariableCollection.GetByName('Yz16', True, True)
        Pz17 = VariableCollection.GetByName('Yz17', True, True)
        Pz18 = VariableCollection.GetByName('Yz18', True, True)
        Pz19 = VariableCollection.GetByName('Yz19', True, True)
        Pz20 = VariableCollection.GetByName('Yz20', True, True)
        Pz21 = VariableCollection.GetByName('Yz21', True, True)
        Pz22 = VariableCollection.GetByName('Yz22', True, True)
        Pz23 = VariableCollection.GetByName('Yz23', True, True)
        Pz24 = VariableCollection.GetByName('Yz24', True, True)
        Pz25 = VariableCollection.GetByName('Yz25', True, True)
        Pz26 = VariableCollection.GetByName('Yz26', True, True)
        Pz27 = VariableCollection.GetByName('Yz27', True, True)
        Pz28 = VariableCollection.GetByName('Yz28', True, True)
        Pz29 = VariableCollection.GetByName('Yz29', True, True)
        Pz30 = VariableCollection.GetByName('Yz30', True, True)
        Pz31 = VariableCollection.GetByName('Yz31', True, True)
        Pz32 = VariableCollection.GetByName('Yz32', True, True)
        Pz36 = VariableCollection.GetByName('Yz36', True, True)
        Pz37 = VariableCollection.GetByName('Yz37', True, True)
        Pz41 = VariableCollection.GetByName('Yz41', True, True)
        Pz42 = VariableCollection.GetByName('Yz42', True, True)
        Pz46 = VariableCollection.GetByName('Yz46', True, True)
        Pz47 = VariableCollection.GetByName('Yz47', True, True)
        Pz51 = VariableCollection.GetByName('Yz51', True, True)
        Pz52 = VariableCollection.GetByName('Yz52', True, True)
        Pz56 = VariableCollection.GetByName('Yz56', True, True)
        Pz57 = VariableCollection.GetByName('Yz57', True, True)
        Pz61 = VariableCollection.GetByName('Yz61', True, True)
        Pz62 = VariableCollection.GetByName('Yz62', True, True)
        Pz66 = VariableCollection.GetByName('Yz66', True, True)
        Pz67 = VariableCollection.GetByName('Yz67', True, True)
        Lfund = VariableCollection.GetByName('Lfund', True, True)
        Sfund = VariableCollection.GetByName('Sfund', True, True)
        Ut1 = VariableCollection.GetByName('Ut1', True, True)
        Ut2 = VariableCollection.GetByName('Ut2', True, True)
        Ut3 = VariableCollection.GetByName('Ut3', True, True)
        Ut4 = VariableCollection.GetByName('Ut4', True, True)
        Ut5 = VariableCollection.GetByName('Ut5', True, True)
        Ut6 = VariableCollection.GetByName('Ut6', True, True)
        Ut7 = VariableCollection.GetByName('Ut7', True, True)
        H1 = VariableCollection.GetByName('H1', True, True)
        YC1 = VariableCollection.GetByName('YC1', True, True)
        YC2 = VariableCollection.GetByName('YC2', True, True)
        YC3 = VariableCollection.GetByName('YC3', True, True)
        YC4 = VariableCollection.GetByName('YC4', True, True)
        YC5 = VariableCollection.GetByName('YC5', True, True)
        YC6 = VariableCollection.GetByName('YC6', True, True)
        YC7 = VariableCollection.GetByName('YC7', True, True)
        Ltlp1 = VariableCollection.GetByName('Ltlp1', True, True)
        Ltlp2 = VariableCollection.GetByName('Ltlp2', True, True)
        Ltlp3 = VariableCollection.GetByName('Ltlp3', True, True)
        Ltlp4 = VariableCollection.GetByName('Ltlp4', True, True)
        PG = VariableCollection.GetByName('PG', True, True)
        try:
            D.value = dic['Nb']
            Da.value = dic['dob']
            Db.value = dic['M']
            Dc.value = dic['Dc']
            Dd.value = dic['Dd']
            d.value = dic['d']
            La.value = dic['La']
            Lb.value = dic['Lb']
            Lc.value = dic['Lc']
            Ld.value = dic['Ld']
            Le.value = dic['Le']
            Sa.value = dic['Sa']
            Sb.value = dic['Sb']
            Sc.value = dic['Sc']
            Sd.value = dic['Sd']
            Se.value = dic['Se']
            N1.value = dic['N1']
            N2.value = dic['N2']
            N3.value = dic['N3']
            N4.value = dic['N4']
            N5.value = dic['N5']
            D1.value = dic['D1']
            D2.value = dic['D2']
            D3.value = dic['D3']
            D4.value = dic['D4']
            D5.value = dic['D5']
            S1.value = dic['S1']
            S2.value = dic['S2']
            S3.value = dic['S3']
            S4.value = dic['S4']
            S5.value = dic['S5']
            Ht1.value = dic['Ht1']
            Ht2.value = dic['Ht2']
            Ht3.value = dic['Ht3']
            Ht4.value = dic['Ht4']
            Ht5.value = dic['Ht5']
            Ht6.value = dic['Ht6']
            Ht7.value = dic['Ht7']
            Dt1.value = dic['Dt1']
            dt1.value = dic['dt1']
            Dt2.value = dic['Dt2']
            dt2.value = dic['dt2']
            Dt3.value = dic['Dt3']
            dt3.value = dic['dt3']
            Dt4.value = dic['Dt4']
            dt4.value = dic['dt4']
            Dt5.value = dic['Dt5']
            dt5.value = dic['dt5']
            Dt6.value = dic['Dt6']
            dt6.value = dic['dt6']
            Dt7.value = dic['Dt7']
            dt7.value = dic['dt7']
            C1.value = dic['C1']
            C2.value = dic['C2']
            C3.value = dic['C3']
            C4.value = dic['C4']
            C5.value = dic['C5']
            C6.value = dic['C6']
            C7.value = dic['C7']
            Tt1.value = dic['Tt1']
            Tt2.value = dic['Tt2']
            Tt3.value = dic['Tt3']
            Tt4.value = dic['Tt4']
            Tt5.value = dic['Tt5']
            Tt6.value = dic['Tt6']
            Tt7.value = dic['Tt7']
            Nt1.value = dic['Nt1']
            Nt2.value = dic['Nt2']
            Nt3.value = dic['Nt3']
            Nt4.value = dic['Nt4']
            Nt5.value = dic['Nt5']
            Nt6.value = dic['Nt6']
            Nt7.value = dic['Nt7']
            Dbt1.value = dic['Dbt1']
            Dbt2.value = dic['Dbt2']
            Dbt3.value = dic['Dbt3']
            Dbt4.value = dic['Dbt4']
            Dbt5.value = dic['Dbt5']
            Dbt6.value = dic['Dbt6']
            Dbt7.value = dic['Dbt7']
            L1.value = dic['L1']
            L2.value = dic['L2']
            L3.value = dic['L3']
            L4.value = dic['L4']
            L5.value = dic['L5']
            L6.value = dic['L6']
            L7.value = dic['L7']
            L8.value = dic['L8']
            L9.value = dic['L9']
            L10.value = dic['L10']
            L11.value = dic['L11']
            L12.value = dic['L12']
            L13.value = dic['L13']
            L14.value = dic['L14']
            T1.value = dic['T1']
            T2.value = dic['T2']
            T3.value = dic['T3']
            T4.value = dic['T4']
            T5.value = dic['T5']
            T6.value = dic['T6']
            T7.value = dic['T7']
            Yz1.value = dic['Yz1']
            Yz2.value = dic['Yz2']
            Yz3.value = dic['Yz3']
            Yz4.value = dic['Yz4']
            Yz5.value = dic['Yz5']
            Yz6.value = dic['Yz6']
            Yz7.value = dic['Yz7']
            Yz8.value = dic['Yz8']
            Yz9.value = dic['Yz9']
            Yz10.value = dic['Yz10']
            Yz11.value = dic['Yz11']
            Yz12.value = dic['Yz12']
            Yz13.value = dic['Yz13']
            Yz14.value = dic['Yz14']
            Yz15.value = dic['Yz15']
            Yz16.value = dic['Yz16']
            Yz17.value = dic['Yz17']
            Yz18.value = dic['Yz18']
            Yz19.value = dic['Yz19']
            Yz20.value = dic['Yz20']
            Yz21.value = dic['Yz21']
            Yz22.value = dic['Yz22']
            Yz23.value = dic['Yz23']
            Yz24.value = dic['Yz24']
            Yz25.value = dic['Yz25']
            Yz26.value = dic['Yz26']
            Yz27.value = dic['Yz27']
            Yz28.value = dic['Yz28']
            Yz29.value = dic['Yz29']
            Yz30.value = dic['Yz30']
            Yz31.value = dic['Yz31']
            Yz32.value = dic['Yz32']
            Yz36.value = dic['Yz36']
            Yz37.value = dic['Yz37']
            Yz41.value = dic['Yz41']
            Yz42.value = dic['Yz42']
            Yz46.value = dic['Yz46']
            Yz47.value = dic['Yz47']
            Yz51.value = dic['Yz51']
            Yz52.value = dic['Yz52']
            Yz56.value = dic['Yz56']
            Yz57.value = dic['Yz57']
            Yz61.value = dic['Yz61']
            Yz62.value = dic['Yz62']
            Yz66.value = dic['Yz66']
            Yz67.value = dic['Yz67']
            Pz1.value = dic['Yz1']
            Pz2.value = dic['Yz2']
            Pz3.value = dic['Yz3']
            Pz4.value = dic['Yz4']
            Pz5.value = dic['Yz5']
            Pz6.value = dic['Yz6']
            Pz7.value = dic['Yz7']
            Pz8.value = dic['Yz8']
            Pz9.value = dic['Yz9']
            Pz10.value = dic['Yz10']
            Pz11.value = dic['Yz11']
            Pz12.value = dic['Yz12']
            Pz13.value = dic['Yz13']
            Pz14.value = dic['Yz14']
            Pz15.value = dic['Yz15']
            Pz16.value = dic['Yz16']
            Pz17.value = dic['Yz17']
            Pz18.value = dic['Yz18']
            Pz19.value = dic['Yz19']
            Pz20.value = dic['Yz20']
            Pz21.value = dic['Yz21']
            Pz22.value = dic['Yz22']
            Pz23.value = dic['Yz23']
            Pz24.value = dic['Yz24']
            Pz25.value = dic['Yz25']
            Pz26.value = dic['Yz26']
            Pz27.value = dic['Yz27']
            Pz28.value = dic['Yz28']
            Pz29.value = dic['Yz29']
            Pz30.value = dic['Yz30']
            Pz31.value = dic['Yz31']
            Pz32.value = dic['Yz32']
            Pz36.value = dic['Yz36']
            Pz37.value = dic['Yz37']
            Pz41.value = dic['Yz41']
            Pz42.value = dic['Yz42']
            Pz46.value = dic['Yz46']
            Pz47.value = dic['Yz47']
            Pz51.value = dic['Yz51']
            Pz52.value = dic['Yz52']
            Pz56.value = dic['Yz56']
            Pz57.value = dic['Yz57']
            Pz61.value = dic['Yz61']
            Pz62.value = dic['Yz62']
            Pz66.value = dic['Yz66']
            Pz67.value = dic['Yz67']
            Lfund.value = dic['Lfund']
            Sfund.value = dic['Sfund']
            Ut1.value = dic['Ut1']
            Ut2.value = dic['Ut2']
            Ut3.value = dic['Ut3']
            Ut4.value = dic['Ut4']
            Ut5.value = dic['Ut5']
            Ut6.value = dic['Ut6']
            Ut7.value = dic['Ut7']
            H1.value = dic['H1']
            YC1.value = dic['YC1']
            YC2.value = dic['YC2']
            YC3.value = dic['YC3']
            YC4.value = dic['YC4']
            YC5.value = dic['YC5']
            YC6.value = dic['YC6']
            YC7.value = dic['YC7']
            Ltlp1.value = dic['Ltlp1']
            Ltlp2.value = dic['Ltlp2']
            Ltlp3.value = dic['Ltlp3']
            Ltlp4.value = dic['Ltlp4']
            PG.value = dic['PG']
        except: pass
        # Перестраиваем модель и сохраняем
        iPart.RebuildModel()

    def save_assembly(self):
        kompas_document = self.kompas.application.ActiveDocument
        kompas_document.Save()

    def save_as_file(self):
        dir_name = fd.asksaveasfilename(
                filetypes=[("Сборка компас", ".a3d")],
                defaultextension=".a3d"
            )
        if dir_name:
            self.kompas.application.ActiveDocument.SaveAs(dir_name)

class DrawingsAPI:
    def __init__(self,kompas:KompasAPI):
        self.kompas=kompas

    def save_as_pdf(self, drawingpath):
        iConverter = self.kompas.application.Converter(self.kompas.kompas_object.ksSystemPath(5) + "\Pdf2d.dll")
        if self.kompas.get_2D_file():
            dir_name = fd.asksaveasfilename(
                    filetypes=[("pdf file", ".pdf")],
                    defaultextension=".pdf"
                )
        if dir_name:    
            iConverter.Convert(drawingpath, dir_name, 0, False)
        return dir_name
    
    def save_as_Kompas(self):
        dir_name = fd.asksaveasfilename(
                filetypes=[("Чертеж компас", ".cdw")],
                defaultextension=".cdw"
            )
        if dir_name:
            self.kompas.application.ActiveDocument.SaveAs(dir_name)

    def __define_stamp_settings(self,data:str,type:str):
        standard_font_height=7
        standard_font_scale = 1
        string = data
        if type=='date':
            standard_font_height=3.5
            standard_font_scale = 1
            string=data
        elif type=='name':
            if 30<data.__len__()<40:
                standard_font_height=6
                standard_font_scale = 0.75
                string = data
            elif 40<=data.__len__()<100:
                standard_font_height=5
                standard_font_scale = 0.75
                string = data
            elif 100<=data.__len__()<160:
                standard_font_height=5
                standard_font_scale = 0.75
                string = data
            elif 160<=data.__len__()<190:
                standard_font_height=5
                standard_font_scale = 0.75
                string = data    
            elif data.__len__()>=160:
                standard_font_height=3
                standard_font_scale = 0.75
                string = data
            else:
                standard_font_height=7
                standard_font_scale = 0.75
                string = data
            if data.__contains__(" ") and data.__len__()>49:
                temp = data.split(" ")
                first_string=temp.pop(0)
                string=""
                while temp.__len__() >= 1:
                    first_string+=" " + temp.pop(0)
                    if first_string.__len__()>data.__len__()*0.5 or first_string.__len__()>90:
                        if temp.__len__() != 0:
                            if temp[0] != "":
                                string += first_string + "\n"
                                first_string=""
                string += first_string
        elif type=="number":
            if 30<data.__len__()<40:
                standard_font_height=6
                standard_font_scale = 0.75
                string = data
            elif data.__len__()>=40:
                standard_font_height=4
                standard_font_scale = 0.75
                string = data
        return standard_font_height, standard_font_scale, string

    def __stamp(self,column:int,args:tuple):#column:int,standard_font_height, standard_font_scale, string):
        from core import LDefin2D
        iStamp = self.kompas.get_2D_file().GetStamp()
        iStamp.ksOpenStamp()
        iStamp.ksColumnNumber(column)
        
        for item in args[2].split("\n"):
            iTextLineParam = self.kompas.kompas6_api5_module.ksTextLineParam(
                    self.kompas.kompas_object.GetParamStruct(self.kompas.kompas6_constants.ko_TextLineParam))
            iTextLineParam.Init()
            iTextLineParam.style = 32769 #32768из компас макро
            iTextItemArray = self.kompas.kompas_object.GetDynamicArray(LDefin2D.TEXT_ITEM_ARR)
            iTextItemParam = self.kompas.kompas6_api5_module.ksTextItemParam(
                    self.kompas.kompas_object.GetParamStruct(self.kompas.kompas6_constants.ko_TextItemParam))
            iTextItemParam.Init()
            iTextItemParam.iSNumb = 0
            iTextItemParam.s = item
            iTextItemParam.type = 0
            iTextItemFont = self.kompas.kompas6_api5_module.ksTextItemFont(iTextItemParam.GetItemFont())
            iTextItemFont.Init()
            iTextItemFont.bitVector = 4096
            iTextItemFont.color = 0
            iTextItemFont.fontName = "GOST type A"
            iTextItemFont.height = args[0]
            iTextItemFont.ksu = args[1]
            iTextItemArray.ksAddArrayItem(-1, iTextItemParam)

            if args[2].split("\n").__len__()>0:
                iTextItemParam = self.kompas.kompas6_api5_module.ksTextItemParam(
                        self.kompas.kompas_object.GetParamStruct(self.kompas.kompas6_constants.ko_TextItemParam))
                iTextItemParam.Init()
                iTextItemParam.iSNumb = 0
                iTextItemParam.s = ""
                iTextItemParam.type = 0
                iTextItemFont = self.kompas.kompas6_api5_module.ksTextItemFont(iTextItemParam.GetItemFont())
                iTextItemFont.Init()
                iTextItemFont.bitVector = 2624
                iTextItemFont.color = 0
                iTextItemFont.fontName = "GOST type A"
                iTextItemFont.height = 3.5
                iTextItemFont.ksu = 0.75
                iTextItemArray.ksAddArrayItem(-1, iTextItemParam)

                iTextItemParam = self.kompas.kompas6_api5_module.ksTextItemParam(
                        self.kompas.kompas_object.GetParamStruct(self.kompas.kompas6_constants.ko_TextItemParam))
                iTextItemParam.Init()
                iTextItemParam.iSNumb = 0
                iTextItemParam.s = ""
                iTextItemParam.type = 0
                iTextItemFont = self.kompas.kompas6_api5_module.ksTextItemFont(iTextItemParam.GetItemFont())
                iTextItemFont.Init()
                iTextItemFont.bitVector = 0
                iTextItemFont.color = 0
                iTextItemFont.fontName = "GOST type A"
                iTextItemFont.height = 2.5
                iTextItemFont.ksu = 0.75
                iTextItemArray.ksAddArrayItem(-1, iTextItemParam)
            
            iTextLineParam.SetTextItemArr(iTextItemArray)
            iStamp.ksTextLine(iTextLineParam)
        iStamp.ksCloseStamp()

    def change_stamp(self,date,drw_number,cap_object,proveril,razrabotal,gip,type=0):
        if type==0:
            self.__stamp(110,self.__define_stamp_settings(razrabotal,"date"))
            self.__stamp(111,self.__define_stamp_settings(proveril,"date"))
            self.__stamp(115,self.__define_stamp_settings(gip,"date"))
            self.__stamp(130,self.__define_stamp_settings(date,"date"))
            self.__stamp(131,self.__define_stamp_settings(date,"date"))
            self.__stamp(135,self.__define_stamp_settings(date,"date"))
            self.__stamp(2,self.__define_stamp_settings(drw_number,"number"))
            self.__stamp(302,self.__define_stamp_settings(cap_object,"name"))
        elif type==1: #Для чертежей узлов - только чертежи номер 74,75
            self.__stamp(110,self.__define_stamp_settings(razrabotal,"date"))
            self.__stamp(111,self.__define_stamp_settings(proveril,"date"))
            self.__stamp(115,self.__define_stamp_settings(gip,"date"))
            self.__stamp(130,self.__define_stamp_settings(date,"date"))
            self.__stamp(131,self.__define_stamp_settings(date,"date"))
            self.__stamp(135,self.__define_stamp_settings(date,"date"))
            self.__stamp(1,self.__define_stamp_settings(drw_number,"number"))
            self.__stamp(2,self.__define_stamp_settings(cap_object,"name"))
        elif type==2: #Для ПЗ общих данных
            self.__stamp(110,self.__define_stamp_settings(razrabotal,"date"))
            self.__stamp(111,self.__define_stamp_settings(proveril,"date"))
            self.__stamp(130,self.__define_stamp_settings(date,"date"))
            self.__stamp(131,self.__define_stamp_settings(date,"date"))
            self.__stamp(1,self.__define_stamp_settings(drw_number,"number"))
            

    def create_text_on_drawing(self,x,y, content:str,height_of_text=5):
        self.kompas.kompas_object.ActiveDocument2D().ksOpenView(0)
        from core import LDefin2D
        iParagraphParam = self.kompas.kompas6_api5_module.ksParagraphParam(
                self.kompas.kompas_object.GetParamStruct(
                        self.kompas.kompas6_constants.ko_ParagraphParam
                )
        )
        iParagraphParam.Init()
        iParagraphParam.x = x
        iParagraphParam.y = y
        iDocument2D = self.kompas.get_2D_file()
        iDocument2D.ksParagraph(iParagraphParam)

        for item in content.split("\n"):
            iTextLineParam = self.kompas.kompas6_api5_module.ksTextLineParam(
                    self.kompas.kompas_object.GetParamStruct(
                            self.kompas.kompas6_constants.ko_TextLineParam
                    )
            )
            iTextLineParam.Init()
            iTextItemArray = self.kompas.kompas_object.GetDynamicArray(
                    LDefin2D.TEXT_ITEM_ARR)
            iTextItemParam = self.kompas.kompas6_api5_module.ksTextItemParam(
                    self.kompas.kompas_object.GetParamStruct(
                            self.kompas.kompas6_constants.ko_TextItemParam))
            iTextItemParam.Init()
            iTextItemParam.s = item
            iTextItemFont = self.kompas.kompas6_api5_module.ksTextItemFont(
                    iTextItemParam.GetItemFont())
            iTextItemFont.Init()
            iTextItemFont.bitVector = 4096
            iTextItemFont.height = height_of_text
            iTextItemFont.ksu = 1
            iTextItemArray.ksAddArrayItem(-1, iTextItemParam)
            iTextLineParam.SetTextItemArr(iTextItemArray)
            iDocument2D.ksTextLine(iTextLineParam)
        obj = iDocument2D.ksEndObj()


class Pole(Kompas_work):
    def do_events(self, thisdict):
        if thisdict == None: return

        kompas = KompasAPI()    
        self.assembly_work(kompas, thisdict)
        kompas.application.Visible=True

    def assembly_work(self, kompas, thisdict):
        import os.path
        path_pole_assembly = os.path.abspath(
            "core\\static\\multifaceted_pole\\Тестовая_сборка\\Тестовая_сборка.a3d"
        )
        if not os.path.isfile(path_pole_assembly):
            print(f"ОШИБКА!!\n\nФайл сборки {path_pole_assembly} не найден")

        kompas.open_3D_file(path_pole_assembly)
        ass = AssemblyAPI(kompas)
        
        params_for_change_variables_dic = self.get_params_variables(thisdict)
        ass.change_external_variables(params_for_change_variables_dic)
        
        self.drawing_work(kompas, thisdict)

        ass.save_as_file()
        kompas.close_assmebly()

    def drawing_work(self,kompas,thisdict):
        drw = DrawingsAPI(kompas)
        path_svai_drawing = os.path.abspath("core\\static\\свая.cdw")
        kompas.open_2D_file(path_svai_drawing)

        current_date = dt.datetime.now()
        mm_yy = current_date.strftime("%m.%Y")
        drw.change_stamp(
                mm_yy,
                thisdict['project_code'],
                thisdict['project_name'],
                "Родчихин",
                thisdict['developer'],
                "Родчихин",
        )
        drw.save_as_pdf(path_svai_drawing)
        drw.save_as_Kompas()
        kompas.close_2D_file()

    def get_params_variables(self,thisdict):
        params_for_change_variables_dic = {
                'D': thisdict['D'],
                'Da': thisdict['Da'],
                'Db': thisdict['Db'],
                'Dc': thisdict['Dc'],
                'Dd': thisdict['Dd'],
                'd': thisdict['d'],
                'La': thisdict['La'],
                'Lb': thisdict['Lb'],
                'Lc': thisdict['Lc'],
                'Ld': thisdict['Ld'],
                'Le': thisdict['Le'],
                'Sa': thisdict['Sa'],
                'Sb': thisdict['Sb'],
                'Sc': thisdict['Sc'],
                'Sd': thisdict['Sd'],
                'Se': thisdict['Se'],
                'N1': thisdict['N1'],
                'N2': thisdict['N2'],
                'N3': thisdict['N3'],
                'N4': thisdict['N4'],
                'N5': thisdict['N5'],
                'D1': thisdict['D1'],
                'D2': thisdict['D2'],
                'D3': thisdict['D3'],
                'D4': thisdict['D4'],
                'D5': thisdict['D5'],
                'S1': thisdict['S1'],
                'S2': thisdict['S2'],
                'S3': thisdict['S3'],
                'S4': thisdict['S4'],
                'S5': thisdict['S5'],
                'Ht1': thisdict['Ht1'],
                'Ht2': thisdict['Ht2'],
                'Ht3': thisdict['Ht3'],
                'Ht4': thisdict['Ht4'],
                'Ht5': thisdict['Ht5'],
                'Ht6': thisdict['Ht6'],
                'Ht7': thisdict['Ht7'],
                'Dt1': thisdict['Dt1'],
                'dt1': thisdict['dt1'],
                'Dt2': thisdict['Dt2'],
                'dt2': thisdict['dt2'],
                'Dt3': thisdict['Dt3'],
                'dt3': thisdict['dt3'],
                'Dt4': thisdict['Dt4'],
                'dt4': thisdict['dt4'],
                'Dt5': thisdict['Dt5'],
                'dt5': thisdict['dt5'],
                'Dt6': thisdict['Dt6'],
                'dt6': thisdict['dt6'],
                'Dt7': thisdict['Dt7'],
                'dt7': thisdict['dt7'],
                'C1': thisdict['C1'],
                'C2': thisdict['C2'],
                'C3': thisdict['C3'],
                'C4': thisdict['C4'],
                'C5': thisdict['C5'],
                'C6': thisdict['C6'],
                'C7': thisdict['C7'],
                'Tt1': thisdict['Tt1'],
                'Tt2': thisdict['Tt2'],
                'Tt3': thisdict['Tt3'],
                'Tt4': thisdict['Tt4'],
                'Tt5': thisdict['Tt5'],
                'Tt6': thisdict['Tt6'],
                'Tt7': thisdict['Tt7'],
                'Nt1': thisdict['Nt1'],
                'Nt2': thisdict['Nt2'],
                'Nt3': thisdict['Nt3'],
                'Nt4': thisdict['Nt4'],
                'Nt5': thisdict['Nt5'],
                'Nt6': thisdict['Nt6'],
                'Nt7': thisdict['Nt7'],
                'Dbt1': thisdict['Dbt1'],
                'Dbt2': thisdict['Dbt2'],
                'Dbt3': thisdict['Dbt3'],
                'Dbt4': thisdict['Dbt4'],
                'Dbt5': thisdict['Dbt5'],
                'Dbt6': thisdict['Dbt6'],
                'Dbt7': thisdict['Dbt7'],
                'L1': thisdict['L1'],
                'L2': thisdict['L2'],
                'L3': thisdict['L3'],
                'L4': thisdict['L4'],
                'L5': thisdict['L5'],
                'L6': thisdict['L6'],
                'L7': thisdict['L7'],
                'L8': thisdict['L8'],
                'L9': thisdict['L9'],
                'L10': thisdict['L10'],
                'L11': thisdict['L11'],
                'L12': thisdict['L12'],
                'L13': thisdict['L13'],
                'L14': thisdict['L14'],
                'T1': thisdict['T1'],
                'T2': thisdict['T2'],
                'T3': thisdict['T3'],
                'T4': thisdict['T4'],
                'T5': thisdict['T5'],
                'T6': thisdict['T6'],
                'T7': thisdict['T7'],
                'Yz1': thisdict['Yz1'],
                'Yz2': thisdict['Yz2'],
                'Yz3': thisdict['Yz3'],
                'Yz4': thisdict['Yz4'],
                'Yz5': thisdict['Yz5'],
                'Yz6': thisdict['Yz6'],
                'Yz7': thisdict['Yz7'],
                'Yz8': thisdict['Yz8'],
                'Yz9': thisdict['Yz9'],
                'Yz10': thisdict['Yz10'],
                'Yz11': thisdict['Yz11'],
                'Yz12': thisdict['Yz12'],
                'Yz13': thisdict['Yz13'],
                'Yz14': thisdict['Yz14'],
                'Yz15': thisdict['Yz15'],
                'Yz16': thisdict['Yz16'],
                'Yz17': thisdict['Yz17'],
                'Yz18': thisdict['Yz18'],
                'Yz19': thisdict['Yz19'],
                'Yz20': thisdict['Yz20'],
                'Yz21': thisdict['Yz21'],
                'Yz22': thisdict['Yz22'],
                'Yz23': thisdict['Yz23'],
                'Yz24': thisdict['Yz24'],
                'Yz25': thisdict['Yz25'],
                'Yz26': thisdict['Yz26'],
                'Yz27': thisdict['Yz27'],
                'Yz28': thisdict['Yz28'],
                'Yz29': thisdict['Yz29'],
                'Yz30': thisdict['Yz30'],
                'Yz31': thisdict['Yz31'],
                'Yz32': thisdict['Yz32'],
                'Yz36': thisdict['Yz36'],
                'Yz37': thisdict['Yz37'],
                'Yz41': thisdict['Yz41'],
                'Yz42': thisdict['Yz42'],
                'Yz46': thisdict['Yz46'],
                'Yz47': thisdict['Yz47'],
                'Yz51': thisdict['Yz51'],
                'Yz52': thisdict['Yz52'],
                'Yz56': thisdict['Yz56'],
                'Yz57': thisdict['Yz57'],
                'Yz61': thisdict['Yz61'],
                'Yz62': thisdict['Yz62'],
                'Yz66': thisdict['Yz66'],
                'Yz67': thisdict['Yz67'],
                'Pz1': thisdict['Pz1'],
                'Pz2': thisdict['Pz2'],
                'Pz3': thisdict['Pz3'],
                'Pz4': thisdict['Pz4'],
                'Pz5': thisdict['Pz5'],
                'Pz6': thisdict['Pz6'],
                'Pz7': thisdict['Pz7'],
                'Pz8': thisdict['Pz8'],
                'Pz9': thisdict['Pz9'],
                'Pz10': thisdict['Pz10'],
                'Pz11': thisdict['Pz11'],
                'Pz12': thisdict['Pz12'],
                'Pz13': thisdict['Pz13'],
                'Pz14': thisdict['Pz14'],
                'Pz15': thisdict['Pz15'],
                'Pz16': thisdict['Pz16'],
                'Pz17': thisdict['Pz17'],
                'Pz18': thisdict['Pz18'],
                'Pz19': thisdict['Pz19'],
                'Pz20': thisdict['Pz20'],
                'Pz21': thisdict['Pz21'],
                'Pz22': thisdict['Pz22'],
                'Pz23': thisdict['Pz23'],
                'Pz24': thisdict['Pz24'],
                'Pz25': thisdict['Pz25'],
                'Pz26': thisdict['Pz26'],
                'Pz27': thisdict['Pz27'],
                'Pz28': thisdict['Pz28'],
                'Pz29': thisdict['Pz29'],
                'Pz30': thisdict['Pz30'],
                'Pz31': thisdict['Pz31'],
                'Pz32': thisdict['Pz32'],
                'Pz36': thisdict['Pz36'],
                'Pz37': thisdict['Pz37'],
                'Pz41': thisdict['Pz41'],
                'Pz42': thisdict['Pz42'],
                'Pz46': thisdict['Pz46'],
                'Pz47': thisdict['Pz47'],
                'Pz51': thisdict['Pz51'],
                'Pz52': thisdict['Pz52'],
                'Pz56': thisdict['Pz56'],
                'Pz57': thisdict['Pz57'],
                'Pz61': thisdict['Pz61'],
                'Pz62': thisdict['Pz62'],
                'Pz66': thisdict['Pz66'],
                'Pz67': thisdict['Pz67'],
                'Lfund': thisdict['Lfund'],
                'Sfund': thisdict['Sfund'],
                'Ut1': thisdict['Ut1'],
                'Ut2': thisdict['Ut2'],
                'Ut3': thisdict['Ut3'],
                'Ut4': thisdict['Ut4'],
                'Ut5': thisdict['Ut5'],
                'Ut6': thisdict['Ut6'],
                'Ut7': thisdict['Ut7'],
                'H1': thisdict['H1'],
                'YC1': thisdict['YC1'],
                'YC2': thisdict['YC2'],
                'YC3': thisdict['YC3'],
                'YC4': thisdict['YC4'],
                'YC5': thisdict['YC5'],
                'YC6': thisdict['YC6'],
                'YC7': thisdict['YC7'],
                'Ltlp1': thisdict['Ltlp1'],
                'Ltlp2': thisdict['Ltlp2'],
                'Ltlp3': thisdict['Ltlp3'],
                'Ltlp4': thisdict['Ltlp4'],
                'PG': thisdict['PG']
        }
        return params_for_change_variables_dic


def do_magic(dict, myclass):
    temp_obj = myclass()
    schema_pdf_path = temp_obj.do_events(dict)
    return schema_pdf_path
