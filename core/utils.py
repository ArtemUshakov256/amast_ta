import base64
import datetime as dt
import os
import pathlib
import pandas as pd
import re


from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
from pandas import ExcelWriter
from sympy import symbols, latex
from tkinter import filedialog as fd


# icondata= base64.b64decode(os.getenv("icon"))
# tempFile= "logo.ico"
# iconfile= open(tempFile,"wb")
# iconfile.write(icondata)
# iconfile.close()

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


def make_multiple_path():
    file_path = fd.askopenfilenames(
        filetypes=(('png files', '*.png'),('All files', '*.*')),
        initialdir="C:/Downloads"
    )
    return file_path


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

