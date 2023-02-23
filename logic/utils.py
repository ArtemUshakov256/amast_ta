import datetime as dt
import pathlib
import pandas as pd
import re


from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
from pandas import ExcelWriter
from tkinter import filedialog as fd


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
        filetypes=(('jpeg files', '*.jpg'), ('png files', '*.png')),
        initialdir="C:/Downloads"
    )
    return file_path


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
    wind_span,
    weight_span,
    pole,
    loads_str
):
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
    safety_coef = 1.1 if voltage >= 330 else 1
    davit_dict = {
        "lower": (height_davit_low, length_davit_low_r, length_davit_low_l),
        "middle": (height_davit_mid, length_davit_mid_r, length_davit_mid_l),
        "upper": (height_davit_up, length_davit_up_r, length_davit_up_l)
    } if branches == "2" or height_davit_mid  else {
        "lower": (height_davit_low, length_davit_low_r, length_davit_low_l),
        "upper": (height_davit_up, length_davit_up_r, length_davit_up_l)
    }
    wire_factor = int(wire.split()[1].split("/")[0]) if wire else ""
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
        "load_pic_dict": loads_pic_dict
    }
    doc.render(context)
    doc.save("generated_doc.docx")


def extract_txt_data(path=None):
    try:
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

        # Ask for saving path
        dir_name = fd.asksaveasfilename(
            filetypes=[("xlsx file", ".xlsx")],
            defaultextension=".xlsx"
        )

        # Export data to xlsx
        if dir_name:
            writer = ExcelWriter(dir_name)
            df = pd.DataFrame(foundation_loads_data)
            df.to_excel(writer, sheet_name="Лист1", index=False, header=False)
            writer.close()
        else:
            result = "Файл не сохранен."

    except Exception as E:
        result = "Что-то пошло не так, конвертация не удалась =("
    
    
