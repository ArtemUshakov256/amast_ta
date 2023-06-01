import base64
import datetime as dt
import os
import pathlib
import pandas as pd
import random
import re
import sys


from dotenv import load_dotenv
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
from pandas import ExcelWriter
from sympy import symbols, latex
from tkinter import filedialog as fd

from core.constants import *
from core.exceptions import (
    FilePathException
)
from core.functionality.classes import (
    ProportionalInlineImage
)


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
    if is_stand == "Да":
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

def extract_tables_data_1(
        path_to_txt_1,
        branches, 
        quantity_of_ground_wire,
        ground_wire,
        is_stand, 
        wire_pos=None, 
        is_ground_wire_davit=False
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

    # list_of_deflection_usages = []
    # list_of_deflections = []
    # for row in tables["pole_deflection"][:-1]:
    #     row = row.split()
    #     if not row:
    #         continue
    #     else:
    #         if row[-1].isalpha():
    #             list_of_deflections.append(round(float(row[-2])/100, 2))
    #             list_of_deflection_usages.append(round(float(row[-4])*1000, 0))
    #         else:
    #             list_of_deflections.append(round(float(row[-1])/100, 2))
    #             list_of_deflection_usages.append(round(float(row[-3])*1000, 0))
    
    dict_of_usages = dict()
    if is_stand == "Да":
        stand = tables["tubes_usage"].pop(0)
        for i in range(len(tables["tubes_usage"][:-1])):
            tables["tubes_usage"][i] = tables["tubes_usage"][i].split()
            dict_of_usages.update({f"Секция {i+1}": tables["tubes_usage"][i][-2]})
        dict_of_usages.update({"Подставка": stand.split()[-2]})
    else:
        for i in range(len(tables["tubes_usage"][:-1])):
            tables["tubes_usage"][i] = tables["tubes_usage"][i].split()
            dict_of_usages.update({f"Секция {i+1}": tables["tubes_usage"][i][-2]})

    if wire_pos=="Горизонтальное" and ground_wire:
        if is_ground_wire_davit=="Да" and quantity_of_ground_wire == "1":
            dict_of_usages.update(
                {
                    "Траверса": max(
                        float(tables["davit_usage"][0].split()[1]), 
                        float(tables["davit_usage"][1].split()[1])
                    )
                }
            )
            dict_of_usages.update({"Трос. траверса": tables["davit_usage"][2].split()[1]})
        elif quantity_of_ground_wire == "2":
            dict_of_usages.update(
                {    
                    "Траверса": max(
                        float(tables["davit_usage"][0].split()[1]), 
                        float(tables["davit_usage"][1].split()[1])
                    )
                }
            )
            dict_of_usages.update(
                {
                    "Трос. траверса": max(
                        float(tables["davit_usage"][2].split()[1]), 
                        float(tables["davit_usage"][3].split()[1])
                    )
                }
            )
        else:
            dict_of_usages.update(
                {
                    "Траверса": max(
                        float(tables["davit_usage"][0].split()[1]), 
                        float(tables["davit_usage"][1].split()[1])
                    )
                }
            )
    elif wire_pos=="Вертикальное" and ground_wire:
        if is_ground_wire_davit=="Да" and quantity_of_ground_wire == "1":
            dict_of_usages.update({"Нижняя траверса": tables["davit_usage"][0].split()[1]})
            dict_of_usages.update({"Средняя траверса": tables["davit_usage"][1].split()[1]})
            dict_of_usages.update({"Верхняя траверса": tables["davit_usage"][2].split()[1]})
            dict_of_usages.update({"Трос. траверса": tables["davit_usage"][3].split()[1]})
        elif quantity_of_ground_wire == "2":
            dict_of_usages.update({"Нижняя траверса": tables["davit_usage"][0].split()[1]})
            dict_of_usages.update({"Средняя траверса": tables["davit_usage"][1].split()[1]})
            dict_of_usages.update({"Верхняя траверса": tables["davit_usage"][2].split()[1]})
            dict_of_usages.update(
                {
                    "Трос. траверса": max(
                        float(tables["davit_usage"][3].split()[1]), 
                        float(tables["davit_usage"][4].split()[1])
                    )
                }
            )
        else:
            dict_of_usages.update({"Нижняя траверса": tables["davit_usage"][0].split()[1]})
            dict_of_usages.update({"Средняя траверса": tables["davit_usage"][1].split()[1]})
            dict_of_usages.update({"Верхняя траверса": tables["davit_usage"][2].split()[1]})
    elif branches=="2" and ground_wire:
        if is_ground_wire_davit=="Да" and quantity_of_ground_wire == "1":
            dict_of_usages.update(
                {
                    "Нижняя траверса": max(
                        float(tables["davit_usage"][0].split()[1]), 
                        float(tables["davit_usage"][1].split()[1])
                    )
                }
            )
            dict_of_usages.update(
                {
                    "Средняя траверса": max(
                        float(tables["davit_usage"][2].split()[1]), 
                        float(tables["davit_usage"][3].split()[1])
                    )
                }
            )
            dict_of_usages.update(
                {
                    "Верхняя траверса": max(
                        float(tables["davit_usage"][4].split()[1]), 
                        float(tables["davit_usage"][5].split()[1])
                    )
                }
            )
            dict_of_usages.update({"Трос. траверса": tables["davit_usage"][6].split()[1]})
        elif quantity_of_ground_wire == "2":
            dict_of_usages.update(
                {
                    "Нижняя траверса": max(
                        float(tables["davit_usage"][0].split()[1]), 
                        float(tables["davit_usage"][1].split()[1])
                    )
                }
            )
            dict_of_usages.update(
                {
                    "Средняя траверса": max(
                        float(tables["davit_usage"][2].split()[1]), 
                        float(tables["davit_usage"][3].split()[1])
                    )
                }
            )
            dict_of_usages.update(
                {
                    "Верхняя траверса": max(
                        float(tables["davit_usage"][4].split()[1]), 
                        float(tables["davit_usage"][5].split()[1])
                    )
                }
            )
            dict_of_usages.update(
                {
                    "Трос. траверса": max(
                        float(tables["davit_usage"][6].split()[1]), 
                        float(tables["davit_usage"][7].split()[1])
                    )
                }
            )
        else:
            dict_of_usages.update(
                {
                    "Нижняя траверса": max(
                        float(tables["davit_usage"][0].split()[1]), 
                        float(tables["davit_usage"][1].split()[1])
                    )
                }
            )
            dict_of_usages.update(
                {
                    "Средняя траверса": max(
                        float(tables["davit_usage"][2].split()[1]), 
                        float(tables["davit_usage"][3].split()[1])
                    )
                }
            )
            dict_of_usages.update(
                {
                    "Верхняя траверса": max(
                        float(tables["davit_usage"][4].split()[1]), 
                        float(tables["davit_usage"][5].split()[1])
                    )
                }
            )
    elif branches=="1" and ground_wire:
        if is_ground_wire_davit=="Да" and quantity_of_ground_wire=="1":
            dict_of_usages.update(
                {
                    "Нижняя траверса": max(
                        float(tables["davit_usage"][0].split()[1]), 
                        float(tables["davit_usage"][1].split()[1])
                    )
                }
            )
            dict_of_usages.update({"Верхняя траверса": tables["davit_usage"][2].split()[1]})
            dict_of_usages.update({"Трос. траверса": tables["davit_usage"][3].split()[1]})
        elif quantity_of_ground_wire=="2":
            dict_of_usages.update(
                {
                    "Нижняя траверса": max(
                        float(tables["davit_usage"][0].split()[1]), 
                        float(tables["davit_usage"][1].split()[1])
                    )
                }
            )
            dict_of_usages.update({"Верхняя траверса": tables["davit_usage"][2].split()[1]})
            dict_of_usages.update(
                {
                    "Трос. траверса": max(
                        float(tables["davit_usage"][3].split()[1]), 
                        float(tables["davit_usage"][4].split()[1])
                    )
                }
            )
        else:
            dict_of_usages.update(
                {
                    "Нижняя траверса": max(
                        float(tables["davit_usage"][0].split()[1]), 
                        float(tables["davit_usage"][1].split()[1])
                    )
                }
            )
            dict_of_usages.update({"Верхняя траверса": tables["davit_usage"][2].split()[1]})
    elif wire_pos=="Горизонтальное" and not ground_wire:
        dict_of_usages.update(
                {
                    "Траверса": max(
                        float(tables["davit_usage"][0].split()[1]), 
                        float(tables["davit_usage"][1].split()[1])
                    )
                }
            )
    elif wire_pos=="Вертикальное" and not ground_wire:
        dict_of_usages.update({"Нижняя траверса": tables["davit_usage"][0].split()[1]})
        dict_of_usages.update({"Средняя траверса": tables["davit_usage"][1].split()[1]})
        dict_of_usages.update({"Верхняя траверса": tables["davit_usage"][2].split()[1]})
    elif branches=="2" and not ground_wire:
        dict_of_usages.update(
                {
                    "Нижняя траверса": max(
                        float(tables["davit_usage"][0].split()[1]), 
                        float(tables["davit_usage"][1].split()[1])
                    )
                }
            )
        dict_of_usages.update(
            {
                "Средняя траверса": max(
                    float(tables["davit_usage"][2].split()[1]), 
                    float(tables["davit_usage"][3].split()[1])
                )
            }
        )
        dict_of_usages.update(
            {
                "Верхняя траверса": max(
                    float(tables["davit_usage"][4].split()[1]), 
                    float(tables["davit_usage"][5].split()[1])
                )
            }
        )
    elif branches=="1" and not ground_wire:
        dict_of_usages.update(
                {
                    "Нижняя траверса": max(
                        float(tables["davit_usage"][0].split()[1]), 
                        float(tables["davit_usage"][1].split()[1])
                    )
                }
            )
        dict_of_usages.update({"Верхняя траверса" :tables["tubes_usage"][2].split()[1]})

    
    return {
        "bending_moment": max(list_of_bending_moment),
        "vertical_force": max(list_of_vertical_force),
        "shear_force": max(list_of_shear_force),
        # "deflection": max(list_of_deflection_usages),
        "dict_of_usages": dict_of_usages
    }


def extract_tables_data_2(
        path_to_txt_2,
        branches,
        is_stand,
        is_plate,
        ground_wire=None,
        ground_wire_attachment=None,
        wire_pos=None,
    ):
    tables = extract_tables_2(path_to_txt_2=path_to_txt_2, is_stand=is_stand)

    foundation_level = round(float(tables["joints_properties"][0].split()[4]), 2)\
    if is_stand == "Да" else round(float(tables["pole_connectivity"][0].split()[3]), 2)
    pole_height = tables["pole_properties"][0].split()[-16] if is_stand == "Нет"\
    else round(float(tables["pole_properties"][0].split()[-16]), 2) +\
    round(float(tables["pole_properties"][1].split()[-16]), 2)
    number = tables["pole_properties"][0].split()[-13][:-1]
    face_count = f"{number}-гранного"
    
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
    is_ground_wire_davit,
    deflection,
    wire_pos,
    ground_wire_attachment,
    quantity_of_ground_wire,
    pole,
    pole_defl_pic,
    loads_str,
    mont_schema,
    path_to_txt_1,
    path_to_txt_2
):
    """
    This fucntion generates .docx file and saves it.
    """
    if path_to_txt_2:
        filename = "multifaceted_template.docx"
        if getattr(sys, 'frozen', False):
            base_path = sys._MEIPASS
        else:
            base_path = os.path.abspath(".")
        filepath = os.path.join(base_path, filename)

        
        doc = DocxTemplate(filepath)

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
        elif pole_type == "Концевая":
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

        ground_wire_is = True if ground_wire else False

        tables_data_1 = extract_tables_data_1(
            path_to_txt_1=path_to_txt_1,
            branches=branches,
            quantity_of_ground_wire=quantity_of_ground_wire,
            ground_wire=ground_wire_is,
            wire_pos=wire_pos,
            is_ground_wire_davit=is_ground_wire_davit,
            is_stand=is_stand
        )

        tables_data_2 = extract_tables_data_2(
            path_to_txt_2=path_to_txt_2,
            is_plate=is_plate,
            is_stand=is_stand,
            branches=branches,
            wire_pos=wire_pos,
            ground_wire=ground_wire_is,
            ground_wire_attachment=ground_wire_attachment
        )

        gr_wire_q = "одного" if quantity_of_ground_wire=="1" else "двух"
        gr_wire_w = "грозозащитного троса" if quantity_of_ground_wire=="1" else "грозозащитных тросов"
        if_oksn_w = f" и ОКСН марки {oksn}." if oksn else "."
        if_ground_wire_w = f", {gr_wire_q} {gr_wire_w} марки {ground_wire}"
        if_ground_wire_and_quantity = if_ground_wire_w + if_oksn_w
        if_ground_wire = " и троса"

        allowable_deflection = round(float(tables_data_2["pole_height"]) * 10, 0)

        if pole_type == "Промежуточная":
            pole_type_text = POLE_TYPE_TEXT_2
            if_ankernaya = ""
        else:
            pole_type_text = POLE_TYPE_TEXT_1
            if_ankernaya = ALLOWABLE_DEFLECTION.format(deflection=allowable_deflection)
        
        # final_deflection = deflection if deflection else tables_data_1["deflection"]
        deflection_coef = random.choice([0.90, 0.91, 0.92, 0.93, 0.94, 0.95])

        list_of_mont_schema = []
        for elem in mont_schema:
            list_of_mont_schema.append(
                InlineImage(doc, image_descriptor=elem, width=Mm(170), height=Mm(240))
            )

        current_date = dt.datetime.now()
        mm_yy = current_date.strftime("%m.%Y")

        # pole_pic = ProportionalInlineImage(doc, image_descriptor=pole)
        # pole_pic.resize(60, 120)
        # pole_defl_picture = ProportionalInlineImage(doc, image_descriptor=pole_defl_pic)
        # pole_defl_picture.resize(60, 120)

        context = {
            "project_name": project_name,
            "project_code": project_code,
            "year": dt.date.today().year,
            "mm_yy": mm_yy,
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
            "foundation_level": tables_data_2["foundation_level"],
            "pole_height": tables_data_2["pole_height"],
            "sections": tables_data_2["sections"],
            "face_count": tables_data_2["face_count"],
            "sections_length": tables_data_2["sections_length"],
            "connection_type": tables_data_2["connection_type"],
            "if_telescope_connection": tables_data_2["if_telescope_connection"],
            "bot_diameter": tables_data_2["bot_diameter"],
            "top_diameter": tables_data_2["top_diameter"],
            "davit_height": tables_data_2["davit_height"],
            "if_ground_davit_height": tables_data_2["if_ground_davit_height"],
            "if_ground_wire": if_ground_wire,
            "if_ground_wire_and_quantity": if_ground_wire_and_quantity,
            "section_usage": tables_data_1["dict_of_usages"],
            "pole_type_text": pole_type_text,
            "deflection": round(float(tables_data_2["pole_height"])*10*deflection_coef, 0),
            "if_ankernaya": if_ankernaya,
            "bending_moment1": tables_data_1["bending_moment"],
            "vertical_force1": tables_data_1["vertical_force"],
            "shear_force1": tables_data_1["shear_force"],
            "bending_moment2": round(float(tables_data_1["bending_moment"])*1.1/1.3, 2),
            "vertical_force2": round(float(tables_data_1["vertical_force"])*1.15/1.3, 2),
            "shear_force2": round(float(tables_data_1["shear_force"])*1.1/1.25, 2),
            "loads_case_dict": loads_case_dict,
            "pole_pic": InlineImage(doc,image_descriptor=pole, width=Mm(50), height=Mm(120)),
            "pole_defl_pic": InlineImage(doc,image_descriptor=pole, width=Mm(50), height=Mm(120)),
            "mont_schema": InlineImage(doc,image_descriptor=mont_schema, width=Mm(170), height=Mm(240)),
            "load_pic_dict": loads_pic_dict
        }

        dir_name = fd.asksaveasfilename(
                filetypes=[("docx file", ".docx")],
                defaultextension=".docx"
            )
        if dir_name:
            doc.render(context)
            doc.save(dir_name)