import datetime as dt
import openpyxl
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


def calculate_foundation(
        flanec_diam,
        thickness_svai,
        deepness_svai,
        height_svai,
        typical_ground,
        is_init_data,
        ground_water_lvl,
        quantity_of_ige, 
        nomer_ige1,
        nomer_ige2,
        nomer_ige3,
        nomer_ige4,
        nomer_ige5,
        ground_type1,
        ground_type2,
        ground_type3,
        ground_type4,
        ground_type5,
        ground_name1,
        ground_name2,
        ground_name3,
        ground_name4,
        ground_name5,
        verh_sloy1,
        verh_sloy2,
        verh_sloy3,
        verh_sloy4,
        verh_sloy5,
        nijn_sloy1,
        nijn_sloy2,
        nijn_sloy3,
        nijn_sloy4,
        nijn_sloy5,
        coef_poristosti1,
        coef_poristosti2,
        coef_poristosti3,
        coef_poristosti4,
        coef_poristosti5,
        udel_scep1,
        udel_scep2,
        udel_scep3,
        udel_scep4,
        udel_scep5,
        ugol_vn_tr1,
        ugol_vn_tr2,
        ugol_vn_tr3,
        ugol_vn_tr4,
        ugol_vn_tr5,
        ves_gr_prir1,
        ves_gr_prir2,
        ves_gr_prir3,
        ves_gr_prir4,
        ves_gr_prir5,
        def_mod1,
        def_mod2,
        def_mod3,
        def_mod4,
        def_mod5,
        pole_type,
        ugol_vntr_tr,
        udel_sceplenie,
        ves_grunta,
        deform_module
):
    calculation_file = os.path.abspath("core/static/Фундамент/Расчет_сваи.xlsx")
    workbook = openpyxl.load_workbook(calculation_file)

    interface_sheet = workbook["Интерфейс"]
    typical_ground_sheet = workbook["Типовые грунты"]
    zadanie_gruntov_sheet = workbook["Задание грунтов"]
    calculation_sheet = workbook["Расчет сваи"]
    flanec_calculation_sheet = workbook["Расчет массы фланца"]

    flanec_calculation_sheet["C3"] = flanec_diam
    flanec_calculation_sheet["C4"] = thickness_svai
    flanec_calculation_sheet["C5"] = deepness_svai
    calculation_sheet["I7"] = height_svai
    zadanie_gruntov_sheet["B23"] = ground_water_lvl
    result = dict()
    if is_init_data:
        interface_sheet["B8"] = "Исходные данные есть"
        if quantity_of_ige == "1":
            zadanie_gruntov_sheet["D6"] = nomer_ige1
            zadanie_gruntov_sheet["D7"] = ground_type1
            zadanie_gruntov_sheet["D8"] = ground_name1
            zadanie_gruntov_sheet["D9"] = verh_sloy1
            zadanie_gruntov_sheet["D10"] = nijn_sloy1
            zadanie_gruntov_sheet["D13"] = coef_poristosti1
            zadanie_gruntov_sheet["D14"] = udel_scep1
            zadanie_gruntov_sheet["D15"] = ugol_vn_tr1
            zadanie_gruntov_sheet["D16"] = ves_gr_prir1
            zadanie_gruntov_sheet["D18"] = def_mod1
        elif quantity_of_ige == "2":
            zadanie_gruntov_sheet["D6"] = nomer_ige1
            zadanie_gruntov_sheet["E6"] = nomer_ige2
            zadanie_gruntov_sheet["D7"] = ground_type1
            zadanie_gruntov_sheet["E7"] = ground_type2
            zadanie_gruntov_sheet["D8"] = ground_name1
            zadanie_gruntov_sheet["E8"] = ground_name2
            zadanie_gruntov_sheet["D9"] = verh_sloy1
            zadanie_gruntov_sheet["E9"] = verh_sloy2
            zadanie_gruntov_sheet["D10"] = nijn_sloy1
            zadanie_gruntov_sheet["E10"] = nijn_sloy2
            zadanie_gruntov_sheet["D13"] = coef_poristosti1
            zadanie_gruntov_sheet["E13"] = coef_poristosti2
            zadanie_gruntov_sheet["D14"] = udel_scep1
            zadanie_gruntov_sheet["E14"] = udel_scep2
            zadanie_gruntov_sheet["D15"] = ugol_vn_tr1
            zadanie_gruntov_sheet["E15"] = ugol_vn_tr2
            zadanie_gruntov_sheet["D16"] = ves_gr_prir1
            zadanie_gruntov_sheet["E16"] = ves_gr_prir2
            zadanie_gruntov_sheet["D18"] = def_mod1
            zadanie_gruntov_sheet["E18"] = def_mod2
        elif quantity_of_ige == "3":
            zadanie_gruntov_sheet["D6"] = nomer_ige1
            zadanie_gruntov_sheet["E6"] = nomer_ige2
            zadanie_gruntov_sheet["F6"] = nomer_ige3
            zadanie_gruntov_sheet["D7"] = ground_type1
            zadanie_gruntov_sheet["E7"] = ground_type2
            zadanie_gruntov_sheet["F7"] = ground_type3
            zadanie_gruntov_sheet["D8"] = ground_name1
            zadanie_gruntov_sheet["E8"] = ground_name2
            zadanie_gruntov_sheet["F8"] = ground_name3
            zadanie_gruntov_sheet["D9"] = verh_sloy1
            zadanie_gruntov_sheet["E9"] = verh_sloy2
            zadanie_gruntov_sheet["F9"] = verh_sloy3
            zadanie_gruntov_sheet["D10"] = nijn_sloy1
            zadanie_gruntov_sheet["E10"] = nijn_sloy2
            zadanie_gruntov_sheet["F10"] = nijn_sloy3
            zadanie_gruntov_sheet["D13"] = coef_poristosti1
            zadanie_gruntov_sheet["E13"] = coef_poristosti2
            zadanie_gruntov_sheet["F13"] = coef_poristosti3
            zadanie_gruntov_sheet["D14"] = udel_scep1
            zadanie_gruntov_sheet["E14"] = udel_scep2
            zadanie_gruntov_sheet["F14"] = udel_scep3
            zadanie_gruntov_sheet["D15"] = ugol_vn_tr1
            zadanie_gruntov_sheet["E15"] = ugol_vn_tr2
            zadanie_gruntov_sheet["F15"] = ugol_vn_tr3
            zadanie_gruntov_sheet["D16"] = ves_gr_prir1
            zadanie_gruntov_sheet["E16"] = ves_gr_prir2
            zadanie_gruntov_sheet["F16"] = ves_gr_prir3
            zadanie_gruntov_sheet["D18"] = def_mod1
            zadanie_gruntov_sheet["E18"] = def_mod2
            zadanie_gruntov_sheet["F18"] = def_mod3
        elif quantity_of_ige == "4":
            zadanie_gruntov_sheet["D6"] = nomer_ige1
            zadanie_gruntov_sheet["E6"] = nomer_ige2
            zadanie_gruntov_sheet["F6"] = nomer_ige3
            zadanie_gruntov_sheet["G6"] = nomer_ige4
            zadanie_gruntov_sheet["D7"] = ground_type1
            zadanie_gruntov_sheet["E7"] = ground_type2
            zadanie_gruntov_sheet["F7"] = ground_type3
            zadanie_gruntov_sheet["G7"] = ground_type4
            zadanie_gruntov_sheet["D8"] = ground_name1
            zadanie_gruntov_sheet["E8"] = ground_name2
            zadanie_gruntov_sheet["F8"] = ground_name3
            zadanie_gruntov_sheet["G8"] = ground_name4
            zadanie_gruntov_sheet["D9"] = verh_sloy1
            zadanie_gruntov_sheet["E9"] = verh_sloy2
            zadanie_gruntov_sheet["F9"] = verh_sloy3
            zadanie_gruntov_sheet["G9"] = verh_sloy4
            zadanie_gruntov_sheet["D10"] = nijn_sloy1
            zadanie_gruntov_sheet["E10"] = nijn_sloy2
            zadanie_gruntov_sheet["F10"] = nijn_sloy3
            zadanie_gruntov_sheet["G10"] = nijn_sloy4
            zadanie_gruntov_sheet["D13"] = coef_poristosti1
            zadanie_gruntov_sheet["E13"] = coef_poristosti2
            zadanie_gruntov_sheet["F13"] = coef_poristosti3
            zadanie_gruntov_sheet["G13"] = coef_poristosti4
            zadanie_gruntov_sheet["D14"] = udel_scep1
            zadanie_gruntov_sheet["E14"] = udel_scep2
            zadanie_gruntov_sheet["F14"] = udel_scep3
            zadanie_gruntov_sheet["G14"] = udel_scep4
            zadanie_gruntov_sheet["D15"] = ugol_vn_tr1
            zadanie_gruntov_sheet["E15"] = ugol_vn_tr2
            zadanie_gruntov_sheet["F15"] = ugol_vn_tr3
            zadanie_gruntov_sheet["G15"] = ugol_vn_tr4
            zadanie_gruntov_sheet["D16"] = ves_gr_prir1
            zadanie_gruntov_sheet["E16"] = ves_gr_prir2
            zadanie_gruntov_sheet["F16"] = ves_gr_prir3
            zadanie_gruntov_sheet["G16"] = ves_gr_prir4
            zadanie_gruntov_sheet["D18"] = def_mod1
            zadanie_gruntov_sheet["E18"] = def_mod2
            zadanie_gruntov_sheet["F18"] = def_mod3
            zadanie_gruntov_sheet["G18"] = def_mod4
        elif quantity_of_ige == "5":
            zadanie_gruntov_sheet["D6"] = nomer_ige1
            zadanie_gruntov_sheet["E6"] = nomer_ige2
            zadanie_gruntov_sheet["F6"] = nomer_ige3
            zadanie_gruntov_sheet["G6"] = nomer_ige4
            zadanie_gruntov_sheet["H6"] = nomer_ige5
            zadanie_gruntov_sheet["D7"] = ground_type1
            zadanie_gruntov_sheet["E7"] = ground_type2
            zadanie_gruntov_sheet["F7"] = ground_type3
            zadanie_gruntov_sheet["G7"] = ground_type4
            zadanie_gruntov_sheet["H7"] = ground_type5
            zadanie_gruntov_sheet["D8"] = ground_name1
            zadanie_gruntov_sheet["E8"] = ground_name2
            zadanie_gruntov_sheet["F8"] = ground_name3
            zadanie_gruntov_sheet["G8"] = ground_name4
            zadanie_gruntov_sheet["H8"] = ground_name5
            zadanie_gruntov_sheet["D9"] = verh_sloy1
            zadanie_gruntov_sheet["E9"] = verh_sloy2
            zadanie_gruntov_sheet["F9"] = verh_sloy3
            zadanie_gruntov_sheet["G9"] = verh_sloy4
            zadanie_gruntov_sheet["H9"] = verh_sloy5
            zadanie_gruntov_sheet["D10"] = nijn_sloy1
            zadanie_gruntov_sheet["E10"] = nijn_sloy2
            zadanie_gruntov_sheet["F10"] = nijn_sloy3
            zadanie_gruntov_sheet["G10"] = nijn_sloy4
            zadanie_gruntov_sheet["H10"] = nijn_sloy5
            zadanie_gruntov_sheet["D13"] = coef_poristosti1
            zadanie_gruntov_sheet["E13"] = coef_poristosti2
            zadanie_gruntov_sheet["F13"] = coef_poristosti3
            zadanie_gruntov_sheet["G13"] = coef_poristosti4
            zadanie_gruntov_sheet["H13"] = coef_poristosti5
            zadanie_gruntov_sheet["D14"] = udel_scep1
            zadanie_gruntov_sheet["E14"] = udel_scep2
            zadanie_gruntov_sheet["F14"] = udel_scep3
            zadanie_gruntov_sheet["G14"] = udel_scep4
            zadanie_gruntov_sheet["H14"] = udel_scep5
            zadanie_gruntov_sheet["D15"] = ugol_vn_tr1
            zadanie_gruntov_sheet["E15"] = ugol_vn_tr2
            zadanie_gruntov_sheet["F15"] = ugol_vn_tr3
            zadanie_gruntov_sheet["G15"] = ugol_vn_tr4
            zadanie_gruntov_sheet["H15"] = ugol_vn_tr5
            zadanie_gruntov_sheet["D16"] = ves_gr_prir1
            zadanie_gruntov_sheet["E16"] = ves_gr_prir2
            zadanie_gruntov_sheet["F16"] = ves_gr_prir3
            zadanie_gruntov_sheet["G16"] = ves_gr_prir4
            zadanie_gruntov_sheet["H16"] = ves_gr_prir5
            zadanie_gruntov_sheet["D18"] = def_mod1
            zadanie_gruntov_sheet["E18"] = def_mod2
            zadanie_gruntov_sheet["F18"] = def_mod3
            zadanie_gruntov_sheet["G18"] = def_mod4
            zadanie_gruntov_sheet["H18"] = def_mod5
    else:
        interface_sheet["B8"] = "Исходных данных нет"
        typical_ground_sheet["B24"] = typical_ground
        typical_ground_sheet["D24"] = ugol_vntr_tr
        typical_ground_sheet["E24"] = udel_sceplenie
        typical_ground_sheet["F24"] = ves_grunta
        typical_ground_sheet["G24"] = deform_module
    
    zadanie_gruntov_sheet["B26"] = pole_type

    workbook.save(calculation_file)
    workbook.close()

    wb = openpyxl.load_workbook(calculation_file)

    if is_init_data:
        result.update({
            "sr_udel_scep": wb["Задание грунтов"]["K14"].value,
            "sr_ugol_vn_tr": wb["Задание грунтов"]["K15"].value,
            "sr_ves_gr_ras": wb["Задание грунтов"]["K17"].value,
            "sr_def_mod": wb["Задание грунтов"]["K18"].value
        })

    result.update(
        {
            "coef_isp_s245": wb["Расчет сваи"]["H26"].value,
            "coef_isp_s345": wb["Расчет сваи"]["H27"].value,
            "coef_isp_gor": wb["Расчет сваи"]["D98"].value,
            "ugol_pov": wb["Расчет сваи"]["D107"].value
        }
    )

    wb.save(calculation_file)
    wb.close()
    print(result)

    return result