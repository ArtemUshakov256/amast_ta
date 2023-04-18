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

from logic.constants import *


# icondata= base64.b64decode(os.getenv("icon"))
# tempFile= "logo.ico"
# iconfile= open(tempFile,"wb")
# iconfile.write(icondata)
# iconfile.close()


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


def extract_tables(path_to_txt_2, is_stand):
    with open(path_to_txt_2, "r", encoding="ANSI") as file:
        file_data = []
        for line in file:
            file_data.append(line.rstrip("\n"))

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


def extract_tables_data(
        path_to_txt_2,
        branches,
        is_stand=False,
        is_plate=False,
        ground_wire=None,
        ground_wire_attachment=None,
        wire_pos=None,
    ):
    tables = extract_tables(path_to_txt_2, is_stand)

    foundation_level = round(float(tables["joints_properties"][0].split()[4]), 2) if is_stand\
    else round(float(tables["pole_connectivity"][0].split()[3]), 2)
    pole_height = tables["pole_properties"][0].split()[-16] if not is_stand\
    else round(float(tables["pole_properties"][0].split()[-16]), 2) +\
    round(float(tables["pole_properties"][1].split()[-16]), 2)
    number = tables["pole_properties"][0].split()[-13][:-1]
    face_count = f"{number}-гранного"
    
    if is_plate:
       tables["tubes_properties"] = tables["tubes_properties"][:-25]
    else:
        tables["tubes_properties"] = tables["tubes_properties"][:-1]
    if is_stand:
        tables["tubes_properties"] = tables["tubes_properties"][:-8] +\
        tables["tubes_properties"][-1:]
    sections = len(tables["tubes_properties"])

    sections_length_list = []
    for row in tables["tubes_properties"]:
        sections_length_list.append(row.split()[-14])
    sections_length = ", ".join(sections_length_list)

    connection_type = "телескопического" if float(tables["tubes_properties"][0].split()[-1]) > 0\
    else "фланцевого"
    if_telescope_connection = IS_TELESCOPE
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
            ground_wire_height = round(float(tables["pole_attachments"][3].split()[-1]), 2) - foundation_level
        else:
            ground_wire_height = pole_height
        if_ground_davit_height = f"Узел крепления троса располагается на высоте {ground_wire_height} м."
    elif branches == "1" and ground_wire:
        for i in range(2):
            davit_height_list.append(
                str(round(float(tables["pole_attachments"][i].split()[-1]), 2) - foundation_level)
            )
        davit_height = ", ".join(davit_height_list)
        if ground_wire_attachment == "Ниже верха опоры":
            ground_wire_height = round(float(tables["pole_attachments"][3].split()[-1]), 2) - foundation_level
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
    elif branches == "1" and not ground_wire:
        for i in range(2):
            davit_height_list.append(
                str(round(float(tables["pole_attachments"][i].split()[-1]), 2) - foundation_level)
            )
        davit_height = ", ".join(davit_height_list)
        if_ground_davit_height = ""
    
    return {
        "foundation_level": foundation_level,
        "pole_height": pole_height,
        "sections": sections,
        "face_count": face_count,
        "sections_length": sections_length,
        "connection_type": connection_type,
        "if_telescope_connection": if_telescope_connection,
        "bot_diameter": bot_diameter,
        "top_diameter": top_diameter,
        "davit_height": davit_height,
        "if_ground_davit_height": if_ground_davit_height
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
    wire,
    wire_tencion,
    ground_wire,
    oksn,
    wind_span,
    weight_span,
    is_stand,
    is_plate,
    is_mont_schema,
    wire_pos,
    ground_wire_attachment,
    quantity_of_ground_wire,
    pole,
    loads_str,
    path_to_txt_2
):
    """
    This fucntion generates .docx file and saves it.
    """
    if path_to_txt_2:
        doc = DocxTemplate("multifaceted_template.docx")

        wind_coef = "1" if branches=="1" else "1.1"
        ice_coef_1 = "1" if branches=="1" else "1.3"
        ice_coef_2 = "1.3" if ice_region in ["I", "II"] else "1.6"

        wire_factor = int(wire.split()[1].split("/")[0]) if wire else ""
        if loads_str:
            loads = [pic_dir.strip("}{") for pic_dir in loads_str.split("} {")]
        # need to check for loads length for exceptions
        if pole_type in ["Анкерно-угловая", "Отпаечная"] and branches == "1"\
        and wire_factor < 180 and ground_wire:
            loads_case_dict = {
                "max_wind": "Нормальный режим (провод и трос не оборваны"\
                " и свободны от гололеда). Максимальный ветер;",
                "ice_wind": "Нормальный режим (провод и трос не оборваны"\
                "и покрыты льдом). Ветер и гололед;",
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
                        "Рис.3.1.2 Ветер и гололед",
                    ],
                    "upper_lower": [
                        InlineImage(doc, image_descriptor=loads[2], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.3 Обрыв верхней и нижней фазы",
                        InlineImage(doc,image_descriptor=loads[3], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.4 Обрыв троса"
                    ],
                    "installation": [
                        InlineImage(doc, image_descriptor=loads[4], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.5 Монтажный режим"
                    ]
                }
        elif pole_type in ["Анкерно-угловая", "Отпаечная"] and branches == "1"\
        and wire_factor < 180 and not ground_wire:
            loads_case_dict = {
                "max_wind": "Нормальный режим (провод и трос не оборваны"\
                " и свободны от гололеда). Максимальный ветер;",
                "ice_wind": "Нормальный режим (провод и трос не оборваны"\
                "и покрыты голольдом). Ветер и гололед;",
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
                        "Рис.3.1.2 Ветер и гололед",
                    ],
                    "upper_lower_installation": [
                        InlineImage(doc, image_descriptor=loads[2], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.3 Обрыв верхней и нижней фазы",
                        InlineImage(doc, image_descriptor=loads[3], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.4 Монтажный режим"
                    ]
                }
        elif pole_type in ["Анкерно-угловая", "Отпаечная"] and branches == "1"\
        and wire_factor > 150 and ground_wire:
            loads_case_dict = {
                "max_wind": "Нормальный режим (провод и трос не оборваны"\
                " и свободны от гололеда). Максимальный ветер;",
                "ice_wind": "Нормальный режим (провод и трос не оборваны"\
                "и покрыты голольдом). Ветер и гололед;",
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
                        "Рис.3.1.2 Ветер и гололед",
                    ],
                    "upper_lower": [
                        InlineImage(doc, image_descriptor=loads[2], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.3 Обрыв верхней фазы",
                        InlineImage(doc, image_descriptor=loads[3], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.4 Обрыв нижней фазы"
                    ],
                    "ground_wire_installation": [
                        InlineImage(doc,image_descriptor=loads[4], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.5 Обрыв троса",
                        InlineImage(doc, image_descriptor=loads[5], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.6 Монтажный режим"
                    ]
                }
        elif pole_type in ["Анкерно-угловая", "Отпаечная"] and branches == "1"\
        and wire_factor > 150 and not ground_wire:
            loads_case_dict = {
                "max_wind": "Нормальный режим (провод и трос не оборваны"\
                " и свободны от гололеда). Максимальный ветер;",
                "ice_wind": "Нормальный режим (провод и трос не оборваны"\
                "и покрыты голольдом). Ветер и гололед;",
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
                        "Рис.3.1.2 Ветер и гололед",
                    ],
                    "upper": [
                        InlineImage(doc, image_descriptor=loads[2], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.3 Обрыв верхней фазы",
                        InlineImage(doc, image_descriptor=loads[3], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.4 Обрыв нижней фазы"
                    ],
                    "installation": [
                        InlineImage(doc, image_descriptor=loads[4], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.5 Монтажный режим"
                    ]
                }
        elif pole_type in ["Анкерно-угловая", "Отпаечная"] and branches == "2"\
        and wire_factor < 180 and ground_wire:
            loads_case_dict = {
                "max_wind": "Нормальный режим (провод и трос не оборваны"\
                " и свободны от гололеда). Максимальный ветер;",
                "ice_wind": "Нормальный режим (провод и трос не оборваны"\
                "и покрыты голольдом). Ветер и гололед;",
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
                        "Рис.3.1.2 Ветер и гололед",
                    ],
                    "upper_middle": [
                        InlineImage(doc, image_descriptor=loads[2], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.3 Обрыв верхней и средней фазы",
                        InlineImage(doc,image_descriptor=loads[3], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.4 Обрыв троса"
                    ],
                    "installation": [
                        InlineImage(doc, image_descriptor=loads[4], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.5 Монтажный режим"
                    ]
                }
        elif pole_type in ["Анкерно-угловая", "Отпаечная"] and branches == "2"\
        and wire_factor < 180 and not ground_wire:
            loads_case_dict = {
                "max_wind": "Нормальный режим (провод и трос не оборваны"\
                " и свободны от гололеда). Максимальный ветер;",
                "ice_wind": "Нормальный режим (провод и трос не оборваны"\
                "и покрыты голольдом). Ветер и гололед;",
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
                        "Рис.3.1.2 Ветер и гололед",
                    ],
                    "upper_middle": [
                        InlineImage(doc, image_descriptor=loads[2], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.3 Обрыв верхней и средней фазы",
                        InlineImage(doc, image_descriptor=loads[3], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.4 Монтажный режим"
                    ]
                }
        elif pole_type in ["Анкерно-угловая", "Отпаечная"] and branches == "2"\
        and wire_factor > 150 and ground_wire:
            loads_case_dict = {
                "max_wind": "Нормальный режим (провод и трос не оборваны"\
                " и свободны от гололеда). Максимальный ветер;",
                "ice_wind": "Нормальный режим (провод и трос не оборваны"\
                "и покрыты голольдом). Ветер и гололед;",
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
                        "Рис.3.1.2 Ветер и гололед",
                    ],
                    "upper": [
                        InlineImage(doc, image_descriptor=loads[2], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.3 Обрыв верхней фазы",
                        InlineImage(doc, image_descriptor=loads[3], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.4 Обрыв средней фазы"
                    ],
                    "ground_wire": [
                        InlineImage(doc,image_descriptor=loads[4], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.5 Обрыв троса",
                        InlineImage(doc, image_descriptor=loads[5], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.6 Монтажный режим"
                    ]
                }
        elif pole_type in ["Анкерно-угловая", "Отпаечная"] and branches == "2"\
        and wire_factor > 150 and not ground_wire:
            loads_case_dict = {
                "max_wind": "Нормальный режим (провод и трос не оборваны"\
                " и свободны от гололеда). Максимальный ветер;",
                "ice_wind": "Нормальный режим (провод и трос не оборваны"\
                "и покрыты голольдом). Ветер и гололед;",
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
                        "Рис.3.1.2 Ветер и гололед",
                    ],
                    "upper": [
                        InlineImage(doc, image_descriptor=loads[2], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.3 Обрыв верхней фазы",
                        InlineImage(doc, image_descriptor=loads[3], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.4 Обрыв средней фазы"
                    ],
                    "installation": [
                        InlineImage(doc, image_descriptor=loads[4], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.5 Монтажный режим"
                    ]
                }
        elif pole_type == "Промежуточная" and branches == "1"\
        and wire_factor < 180 and ground_wire:
            loads_case_dict = {
                "max_wind": "Нормальный режим (провод и трос не оборваны"\
                " и свободны от гололеда). Максимальный ветер;",
                "ice_wind": "Нормальный режим (провод и трос не оборваны"\
                "и покрыты льдом). Ветер и гололед;",
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
                        "Рис.3.1.2 Ветер и гололед",
                    ],
                    "upper_lower": [
                        InlineImage(doc, image_descriptor=loads[2], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.3 Обрыв верхней и нижней фазы",
                        InlineImage(doc,image_descriptor=loads[3], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.4 Обрыв троса"
                    ],
                    "installation": [
                        InlineImage(doc, image_descriptor=loads[4], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.5 Монтажный режим"
                    ]
                }
        elif pole_type == "Промежуточная" and branches == "1"\
        and wire_factor < 180 and not ground_wire:
            loads_case_dict = {
                "max_wind": "Нормальный режим (провод и трос не оборваны"\
                " и свободны от гололеда). Максимальный ветер;",
                "ice_wind": "Нормальный режим (провод и трос не оборваны"\
                "и покрыты голольдом). Ветер и гололед;",
                "upper_lower":\
                "Аварийный режим (обрыв верхней и нижней фазы провода);"
            }
            if loads:
                loads_pic_dict = {
                    "max_wind": [
                        InlineImage(doc, image_descriptor=loads[0], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.1 Максимальный ветер",
                        InlineImage(doc, image_descriptor=loads[1], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.2 Ветер и гололед",
                    ],
                    "upper_lower": [
                        InlineImage(doc, image_descriptor=loads[2], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.3 Обрыв верхней и нижней фазы"
                    ]
                }
        elif pole_type == "Промежуточная" and branches == "1"\
        and wire_factor > 150 and ground_wire:
            loads_case_dict = {
                "max_wind": "Нормальный режим (провод и трос не оборваны"\
                " и свободны от гололеда). Максимальный ветер;",
                "ice_wind": "Нормальный режим (провод и трос не оборваны"\
                "и покрыты голольдом). Ветер и гололед;",
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
                        "Рис.3.1.2 Ветер и гололед",
                    ],
                    "upper": [
                        InlineImage(doc, image_descriptor=loads[2], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.3 Обрыв верхней фазы",
                        InlineImage(doc, image_descriptor=loads[3], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.4 Обрыв нижней фазы"
                    ],
                    "ground_wire": [
                        InlineImage(doc,image_descriptor=loads[4], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.5 Обрыв троса"
                    ]
                }
        elif pole_type == "Промежуточная" and branches == "1"\
        and wire_factor > 150 and not ground_wire:
            loads_case_dict = {
                "max_wind": "Нормальный режим (провод и трос не оборваны"\
                " и свободны от гололеда). Максимальный ветер;",
                "ice_wind": "Нормальный режим (провод и трос не оборваны"\
                "и покрыты голольдом). Ветер и гололед;",
                "upper": "Аварийный режим (обрыв верхней фазы провода);",
                "lower": "Аварийный режим (обрыв нижней фазы провода);"
            }
            if loads:
                loads_pic_dict = {
                    "max_wind": [
                        InlineImage(doc, image_descriptor=loads[0], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.1 Максимальный ветер",
                        InlineImage(doc, image_descriptor=loads[1], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.2 Ветер и гололед",
                    ],
                    "upper": [
                        InlineImage(doc, image_descriptor=loads[2], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.3 Обрыв верхней фазы",
                        InlineImage(doc, image_descriptor=loads[3], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.4 Обрыв нижней фазы"
                    ]
                }
        elif pole_type == "Промежуточная" and branches == "2"\
        and wire_factor < 180 and ground_wire:
            loads_case_dict = {
                "max_wind": "Нормальный режим (провод и трос не оборваны"\
                " и свободны от гололеда). Максимальный ветер;",
                "ice_wind": "Нормальный режим (провод и трос не оборваны"\
                "и покрыты голольдом). Ветер и гололед;",
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
                        "Рис.3.1.2 Ветер и гололед",
                    ],
                    "upper_middle": [
                        InlineImage(doc, image_descriptor=loads[2], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.3 Обрыв верхней и средней фазы",
                        InlineImage(doc,image_descriptor=loads[3], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.4 Обрыв троса"
                    ]
                }
        elif pole_type == "Промежуточная" and branches == "2"\
        and wire_factor < 180 and not ground_wire:
            loads_case_dict = {
                "max_wind": "Нормальный режим (провод и трос не оборваны"\
                " и свободны от гололеда). Максимальный ветер;",
                "ice_wind": "Нормальный режим (провод и трос не оборваны"\
                "и покрыты голольдом). Ветер и гололед;",
                "upper_middle":\
                "Аварийный режим (обрыв верхней и средней фазы провода);"
            }
            if loads:
                loads_pic_dict = {
                    "max_wind": [
                        InlineImage(doc, image_descriptor=loads[0], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.1 Максимальный ветер",
                        InlineImage(doc, image_descriptor=loads[1], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.2 Ветер и гололед",
                    ],
                    "upper_middle": [
                        InlineImage(doc, image_descriptor=loads[2], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.3 Обрыв верхней и средней фазы"
                    ]
                }
        elif pole_type == "Промежуточная" and branches == "2"\
        and wire_factor > 150 and ground_wire:
            loads_case_dict = {
                "max_wind": "Нормальный режим (провод и трос не оборваны"\
                " и свободны от гололеда). Максимальный ветер;",
                "ice_wind": "Нормальный режим (провод и трос не оборваны"\
                "и покрыты голольдом). Ветер и гололед;",
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
                        "Рис.3.1.2 Ветер и гололед",
                    ],
                    "upper": [
                        InlineImage(doc, image_descriptor=loads[2], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.3 Обрыв верхней фазы",
                        InlineImage(doc, image_descriptor=loads[3], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.4 Обрыв средней фазы"
                    ],
                    "ground_wire": [
                        InlineImage(doc,image_descriptor=loads[4], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.5 Обрыв троса"
                    ]
                }
        elif pole_type == "Промежуточная" and branches == "2"\
        and wire_factor > 150 and not ground_wire:
            loads_case_dict = {
                "max_wind": "Нормальный режим (провод и трос не оборваны"\
                " и свободны от гололеда). Максимальный ветер;",
                "ice_wind": "Нормальный режим (провод и трос не оборваны"\
                "и покрыты голольдом). Ветер и гололед;",
                "upper": "Аварийный режим (обрыв верхней фазы провода);",
                "middle": "Аварийный режим (обрыв средней фазы провода);"
            }
            if loads:
                loads_pic_dict = {
                    "max_wind": [
                        InlineImage(doc, image_descriptor=loads[0], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.1 Максимальный ветер",
                        InlineImage(doc, image_descriptor=loads[1], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.2 Ветер и гололед",
                    ],
                    "upper": [
                        InlineImage(doc, image_descriptor=loads[2], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.3 Обрыв верхней фазы",
                        InlineImage(doc, image_descriptor=loads[3], width=Mm(66), height=Mm(66)),
                        "Рис.3.1.4 Обрыв средней фазы"
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

        tables_data = extract_tables_data(
            path_to_txt_2=path_to_txt_2,
            is_plate=is_plate,
            is_stand=is_stand,
            branches=branches,
            wire_pos=wire_pos,
            ground_wire=ground_wire,
            ground_wire_attachment=ground_wire_attachment
        )

        gr_wire_q = "одного" if quantity_of_ground_wire=="1" else "двух"
        gr_wire_w = "грозозащитного троса" if quantity_of_ground_wire=="1" else "грозозащитных тросов"
        if_oksn_w = f" и ОКСН марки {oksn}." if oksn else "."
        if_ground_wire_w = f", {gr_wire_q} {gr_wire_w} марки {ground_wire}"
        if_ground_wire_and_quantity = if_ground_wire_w + if_oksn_w
        if_ground_wire = " и троса"
        if_mont_schema = "Монтажная схема привидена в приложении Б" if is_mont_schema else ""

        context = {
            "project_name": project_name,
            "project_code": project_code,
            "year": dt.date.today().year,
            "pole_code": pole_code,
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
            "wire": wire,
            "wire_tencion": wire_tencion,
            "ground_wire": ground_wire,
            "wind_coef": wind_coef,
            "ice_coef_1": ice_coef_1,
            "ice_coef_2": ice_coef_2,
            "wind_span": wind_span,
            "weight_span": weight_span,
            "foundation_level": tables_data["foundation_level"],
            "pole_height": tables_data["pole_height"],
            "sections": tables_data["sections"],
            "face_count": tables_data["face_count"],
            "sections_length": tables_data["sections_length"],
            "connection_type": tables_data["connection_type"],
            "if_telescope_connection": tables_data["if_telescope_connection"],
            "bot_diameter": tables_data["bot_diameter"],
            "top_diameter": tables_data["top_diameter"],
            "davit_height": tables_data["davit_height"],
            "if_ground_davit_height": tables_data["if_ground_davit_height"],
            "if_ground_wire": if_ground_wire,
            "if_ground_wire_and_quantity": if_ground_wire_and_quantity,
            "if_mont_schema": if_mont_schema,
            "loads_case_dict": loads_case_dict,
            "pole_pic": InlineImage(doc, image_descriptor=pole, width=Mm(80), height=Mm(150)),
            "load_pic_dict": loads_pic_dict
        }

        dir_name = fd.asksaveasfilename(
                filetypes=[("docx file", ".docx")],
                defaultextension=".docx"
            )
        if dir_name:
            doc.render(context)
            doc.save(dir_name)