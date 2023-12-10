import base64
import datetime as dt
import os
import pathlib
import pandas as pd
import re


from docxtpl import DocxTemplate, InlineImage
from dotenv import load_dotenv
from docx.shared import Mm
from pandas import ExcelWriter
from sympy import symbols, latex
from tkinter import filedialog as fd

from core.exceptions import FilePathException


# load_dotenv()
# icondata= base64.b64decode(os.getenv("ICON"))
# tempFile= "logo.ico"

# with open(tempFile, "wb") as iconfile:
#     iconfile.write(icondata)

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
back_icon = base64.b64decode(os.getenv("ENCODED_BACK"))
tempFile_back = os.path.abspath("back.png")

open_icon = base64.b64decode(os.getenv("ENCODED_OPEN"))
tempFile_open = os.path.abspath("open1.png")

save_icon = base64.b64decode(os.getenv("ENCODED_SAVE"))
tempFile_save = os.path.abspath("save1.png")

lupa_icon = base64.b64decode(os.getenv("ENCODED_LUPA"))
tempFile_lupa = os.path.abspath("lupa.png")

with open(tempFile_back, "wb") as iconfileback:
    iconfileback.write(back_icon)

with open(tempFile_open, "wb") as iconfileopen:
    iconfileopen.write(open_icon)

with open(tempFile_save, "wb") as iconfilesave:
    iconfilesave.write(save_icon)

with open(tempFile_lupa, "wb") as iconfileplus:
    iconfileplus.write(lupa_icon)


current_date = dt.datetime.now()
mm_yy = current_date.strftime("%m.%Y")


def make_path_txt():
    file_path = fd.askopenfilename(
        filetypes=(('text files', '*.txt'),('All files', '*.*')),
        initialdir="C:/Downloads"
    )
    return file_path


def make_path_png():
    file_path = fd.askopenfilename(
        filetypes=(('png files', '*.png'),('All files', '*.*')),
        initialdir="C:/Downloads"
    )
    return file_path


def make_path_xlsx():
    file_path = fd.askopenfilename(
        filetypes=(('xlsx files', '*.xlsx'),('All files', '*.*')),
        initialdir="C:/Downloads"
    )
    return file_path


def make_multiple_path():
    file_path = fd.askopenfilenames(
        filetypes=(('png files', '*.png'),('All files', '*.*')),
        initialdir="C:/Downloads"
    )
    return file_path


def extract_tables_1(path_to_txt_1):
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
    
    support_reaction = file_data[support_reaction_start:support_reaction_end][6:]
    
    return {
        "support_reaction": support_reaction
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
        if re.match("Steel Pole Connectivity:", file_data[i]):
            pole_connectivity_start = i
    tubes_properties = file_data[pole_properties_start:pole_connectivity_start]
    if is_stand == "Да":
        tubes_properties = tubes_properties[16:]
    else:
        tubes_properties = tubes_properties[15:]
    return {
        "tubes_properties": tubes_properties,
    }


def extract_tables_data_2(
        path_to_txt_2,
        is_stand,
        is_plate
    ):
    tables = extract_tables_2(path_to_txt_2=path_to_txt_2, is_stand=is_stand)
    if is_plate == "Да":
       for i, s in enumerate(tables["tubes_properties"]):
           if re.match("Base Plate Properties:", s):
               index = i
               break
       tables["tubes_properties"] = tables["tubes_properties"][:i-1]
    else:
        tables["tubes_properties"] = tables["tubes_properties"][:-1]
    if is_stand == "Да":
        tables["tubes_properties"] = tables["tubes_properties"][:-8] +\
        [tables["tubes_properties"][-1]]
    bot_diameter = round(float(tables["tubes_properties"][-1].split()[-3]), 2) * 10
    return {
        "bot_diameter": bot_diameter
    }


def extract_foundation_loads_and_diam(
        path_to_txt_1,
        path_to_txt_2,
        is_plate,
        is_stand
):
    tables_data_1 = extract_tables_data_1(
        path_to_txt_1=path_to_txt_1
    )

    tables_data_2 = extract_tables_data_2(
        path_to_txt_2=path_to_txt_2,
        is_stand=is_stand,
        is_plate=is_plate
    )
    return {
        "moment": tables_data_1["bending_moment"],
        "vert_force": tables_data_1["vertical_force"],
        "shear_force": tables_data_1["shear_force"],
        "bot_diam": tables_data_2["bot_diameter"]
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

    def save_assembly(self):
        kompas_document = self.kompas.application.ActiveDocument
        kompas_document.Save()

    def save_as_file(self):
        # import os
        # directory = '%s' % (pdf_file_path)
        # if not os.path.exists(directory):
        #     os.makedirs(directory)
        dir_name = fd.asksaveasfilename(
                filetypes=[("Сборка компас", ".a3d")],
                defaultextension=".a3d"
            )
        if dir_name:
            self.kompas.application.ActiveDocument.SaveAs(dir_name)


class DrawingsAPI:
    def __init__(self,kompas:KompasAPI):
        self.kompas=kompas

    def save_as_pdf(self,drawingpath,pdf_file_path,filename):
        # import os
        iConverter = self.kompas.application.Converter(self.kompas.kompas_object.ksSystemPath(5) + "\Pdf2d.dll")
        if self.kompas.get_2D_file():
            # directory = '%s' % (pdf_file_path)
            # if not os.path.exists(directory):
            #     os.makedirs(directory)
            dir_name = fd.asksaveasfilename(
                    filetypes=[("pdf file", ".pdf")],
                    defaultextension=".pdf"
                )
        if dir_name:    
            iConverter.Convert(drawingpath, dir_name, 0, False)
    
    def save_as_Kompas(self,pdf_file_path,filename):
        # import os
        # directory = '%s' % (pdf_file_path)
        # if not os.path.exists(directory):
        #     os.makedirs(directory)
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
        #doc2D = 
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


def do_magic(dict, myclass):
    temp_obj = myclass()
    temp_obj.do_events(dict)
    # temp_obj.concatenate_pdf_files(
    #         dict['project_name'],
    #         dict['default_path_result'],
    #         dict['default_path_result']+"\\"+dict['project_name']+".pdf"
    # )