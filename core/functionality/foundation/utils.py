import datetime as dt
import openpyxl
import os
import pathlib
import pandas as pd
import random
import re
import sys
import xlwings as xw

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
        deform_module,
        coef_nadej,
        coef_usl_rab
):
    calculation_file = os.path.abspath("core/static/Фундамент/Расчет_сваи.xlsx")
    # workbook = openpyxl.load_workbook(calculation_file)
    workbook = xw.Book(calculation_file)

    interface_sheet = workbook.sheets["Интерфейс"]
    typical_ground_sheet = workbook.sheets["Типовые грунты"]
    zadanie_gruntov_sheet = workbook.sheets["Задание грунтов"]
    calculation_sheet = workbook.sheets["Расчет сваи"]
    flanec_calculation_sheet = workbook.sheets["Расчет массы фланца"]

    flanec_calculation_sheet["C3"].value = flanec_diam
    flanec_calculation_sheet["C4"].value = thickness_svai
    flanec_calculation_sheet["C5"].value = deepness_svai
    calculation_sheet["I7"].value = height_svai
    result = dict()
    if is_init_data:
        interface_sheet["B8"].value = "Исходные данные есть"
        zadanie_gruntov_sheet["B23"].value = ground_water_lvl
        zadanie_gruntov_sheet["C26"].value = coef_nadej
        zadanie_gruntov_sheet["C28"].value = coef_usl_rab
        if "Песок" in [ground_type1, ground_type2, ground_type3, ground_type4, ground_type5]:
            zadanie_gruntov_sheet["C28"].value = coef_usl_rab_dict["Песок"]
        elif "Супесь" in [ground_type1, ground_type2, ground_type3, ground_type4, ground_type5]:
            zadanie_gruntov_sheet["C28"].value = coef_usl_rab_dict["Супесь"]
        elif "Суглинок" in [ground_type1, ground_type2, ground_type3, ground_type4, ground_type5]:
            zadanie_gruntov_sheet["C28"].value = coef_usl_rab_dict["Суглинок"]
        else:
            zadanie_gruntov_sheet["C28"].value = coef_usl_rab_dict["Глина"]
        if quantity_of_ige == "1":
            zadanie_gruntov_sheet["D6"].value = nomer_ige1
            zadanie_gruntov_sheet["D7"].value = ground_type1
            zadanie_gruntov_sheet["D8"].value = ground_name1
            zadanie_gruntov_sheet["D9"].value = verh_sloy1
            zadanie_gruntov_sheet["D10"].value = nijn_sloy1
            zadanie_gruntov_sheet["D13"].value = coef_poristosti1
            zadanie_gruntov_sheet["D14"].value = udel_scep1
            zadanie_gruntov_sheet["D15"].value = ugol_vn_tr1
            zadanie_gruntov_sheet["D16"].value = ves_gr_prir1
            zadanie_gruntov_sheet["D18"].value = def_mod1
        elif quantity_of_ige == "2":
            zadanie_gruntov_sheet["D6"].value = nomer_ige1
            zadanie_gruntov_sheet["E6"].value = nomer_ige2
            zadanie_gruntov_sheet["D7"].value = ground_type1
            zadanie_gruntov_sheet["E7"].value = ground_type2
            zadanie_gruntov_sheet["D8"].value = ground_name1
            zadanie_gruntov_sheet["E8"].value = ground_name2
            zadanie_gruntov_sheet["D9"].value = verh_sloy1
            zadanie_gruntov_sheet["E9"].value = verh_sloy2
            zadanie_gruntov_sheet["D10"].value = nijn_sloy1
            zadanie_gruntov_sheet["E10"].value = nijn_sloy2
            zadanie_gruntov_sheet["D13"].value = coef_poristosti1
            zadanie_gruntov_sheet["E13"].value = coef_poristosti2
            zadanie_gruntov_sheet["D14"].value = udel_scep1
            zadanie_gruntov_sheet["E14"].value = udel_scep2
            zadanie_gruntov_sheet["D15"].value = ugol_vn_tr1
            zadanie_gruntov_sheet["E15"].value = ugol_vn_tr2
            zadanie_gruntov_sheet["D16"].value = ves_gr_prir1
            zadanie_gruntov_sheet["E16"].value = ves_gr_prir2
            zadanie_gruntov_sheet["D18"].value = def_mod1
            zadanie_gruntov_sheet["E18"].value = def_mod2
        elif quantity_of_ige == "3":
            zadanie_gruntov_sheet["D6"].value = nomer_ige1
            zadanie_gruntov_sheet["E6"].value = nomer_ige2
            zadanie_gruntov_sheet["F6"].value = nomer_ige3
            zadanie_gruntov_sheet["D7"].value = ground_type1
            zadanie_gruntov_sheet["E7"].value = ground_type2
            zadanie_gruntov_sheet["F7"].value = ground_type3
            zadanie_gruntov_sheet["D8"].value = ground_name1
            zadanie_gruntov_sheet["E8"].value = ground_name2
            zadanie_gruntov_sheet["F8"].value = ground_name3
            zadanie_gruntov_sheet["D9"].value = verh_sloy1
            zadanie_gruntov_sheet["E9"].value = verh_sloy2
            zadanie_gruntov_sheet["F9"].value = verh_sloy3
            zadanie_gruntov_sheet["D10"].value = nijn_sloy1
            zadanie_gruntov_sheet["E10"].value = nijn_sloy2
            zadanie_gruntov_sheet["F10"].value = nijn_sloy3
            zadanie_gruntov_sheet["D13"].value = coef_poristosti1
            zadanie_gruntov_sheet["E13"].value = coef_poristosti2
            zadanie_gruntov_sheet["F13"].value = coef_poristosti3
            zadanie_gruntov_sheet["D14"].value = udel_scep1
            zadanie_gruntov_sheet["E14"].value = udel_scep2
            zadanie_gruntov_sheet["F14"].value = udel_scep3
            zadanie_gruntov_sheet["D15"].value = ugol_vn_tr1
            zadanie_gruntov_sheet["E15"].value = ugol_vn_tr2
            zadanie_gruntov_sheet["F15"].value = ugol_vn_tr3
            zadanie_gruntov_sheet["D16"].value = ves_gr_prir1
            zadanie_gruntov_sheet["E16"].value = ves_gr_prir2
            zadanie_gruntov_sheet["F16"].value = ves_gr_prir3
            zadanie_gruntov_sheet["D18"].value = def_mod1
            zadanie_gruntov_sheet["E18"].value = def_mod2
            zadanie_gruntov_sheet["F18"].value = def_mod3
        elif quantity_of_ige == "4":
            zadanie_gruntov_sheet["D6"].value = nomer_ige1
            zadanie_gruntov_sheet["E6"].value = nomer_ige2
            zadanie_gruntov_sheet["F6"].value = nomer_ige3
            zadanie_gruntov_sheet["G6"].value = nomer_ige4
            zadanie_gruntov_sheet["D7"].value = ground_type1
            zadanie_gruntov_sheet["E7"].value = ground_type2
            zadanie_gruntov_sheet["F7"].value = ground_type3
            zadanie_gruntov_sheet["G7"].value = ground_type4
            zadanie_gruntov_sheet["D8"].value = ground_name1
            zadanie_gruntov_sheet["E8"].value = ground_name2
            zadanie_gruntov_sheet["F8"].value = ground_name3
            zadanie_gruntov_sheet["G8"].value = ground_name4
            zadanie_gruntov_sheet["D9"].value = verh_sloy1
            zadanie_gruntov_sheet["E9"].value = verh_sloy2
            zadanie_gruntov_sheet["F9"].value = verh_sloy3
            zadanie_gruntov_sheet["G9"].value = verh_sloy4
            zadanie_gruntov_sheet["D10"].value = nijn_sloy1
            zadanie_gruntov_sheet["E10"].value = nijn_sloy2
            zadanie_gruntov_sheet["F10"].value = nijn_sloy3
            zadanie_gruntov_sheet["G10"].value = nijn_sloy4
            zadanie_gruntov_sheet["D13"].value = coef_poristosti1
            zadanie_gruntov_sheet["E13"].value = coef_poristosti2
            zadanie_gruntov_sheet["F13"].value = coef_poristosti3
            zadanie_gruntov_sheet["G13"].value = coef_poristosti4
            zadanie_gruntov_sheet["D14"].value = udel_scep1
            zadanie_gruntov_sheet["E14"].value = udel_scep2
            zadanie_gruntov_sheet["F14"].value = udel_scep3
            zadanie_gruntov_sheet["G14"].value = udel_scep4
            zadanie_gruntov_sheet["D15"].value = ugol_vn_tr1
            zadanie_gruntov_sheet["E15"].value = ugol_vn_tr2
            zadanie_gruntov_sheet["F15"].value = ugol_vn_tr3
            zadanie_gruntov_sheet["G15"].value = ugol_vn_tr4
            zadanie_gruntov_sheet["D16"].value = ves_gr_prir1
            zadanie_gruntov_sheet["E16"].value = ves_gr_prir2
            zadanie_gruntov_sheet["F16"].value = ves_gr_prir3
            zadanie_gruntov_sheet["G16"].value = ves_gr_prir4
            zadanie_gruntov_sheet["D18"].value = def_mod1
            zadanie_gruntov_sheet["E18"].value = def_mod2
            zadanie_gruntov_sheet["F18"].value = def_mod3
            zadanie_gruntov_sheet["G18"].value = def_mod4
        elif quantity_of_ige == "5":
            zadanie_gruntov_sheet["D6"].value = nomer_ige1
            zadanie_gruntov_sheet["E6"].value = nomer_ige2
            zadanie_gruntov_sheet["F6"].value = nomer_ige3
            zadanie_gruntov_sheet["G6"].value = nomer_ige4
            zadanie_gruntov_sheet["H6"].value = nomer_ige5
            zadanie_gruntov_sheet["D7"].value = ground_type1
            zadanie_gruntov_sheet["E7"].value = ground_type2
            zadanie_gruntov_sheet["F7"].value = ground_type3
            zadanie_gruntov_sheet["G7"].value = ground_type4
            zadanie_gruntov_sheet["H7"].value = ground_type5
            zadanie_gruntov_sheet["D8"].value = ground_name1
            zadanie_gruntov_sheet["E8"].value = ground_name2
            zadanie_gruntov_sheet["F8"].value = ground_name3
            zadanie_gruntov_sheet["G8"].value = ground_name4
            zadanie_gruntov_sheet["H8"].value = ground_name5
            zadanie_gruntov_sheet["D9"].value = verh_sloy1
            zadanie_gruntov_sheet["E9"].value = verh_sloy2
            zadanie_gruntov_sheet["F9"].value = verh_sloy3
            zadanie_gruntov_sheet["G9"].value = verh_sloy4
            zadanie_gruntov_sheet["H9"].value = verh_sloy5
            zadanie_gruntov_sheet["D10"].value = nijn_sloy1
            zadanie_gruntov_sheet["E10"].value = nijn_sloy2
            zadanie_gruntov_sheet["F10"].value = nijn_sloy3
            zadanie_gruntov_sheet["G10"].value = nijn_sloy4
            zadanie_gruntov_sheet["H10"].value = nijn_sloy5
            zadanie_gruntov_sheet["D13"].value = coef_poristosti1
            zadanie_gruntov_sheet["E13"].value = coef_poristosti2
            zadanie_gruntov_sheet["F13"].value = coef_poristosti3
            zadanie_gruntov_sheet["G13"].value = coef_poristosti4
            zadanie_gruntov_sheet["H13"].value = coef_poristosti5
            zadanie_gruntov_sheet["D14"].value = udel_scep1
            zadanie_gruntov_sheet["E14"].value = udel_scep2
            zadanie_gruntov_sheet["F14"].value = udel_scep3
            zadanie_gruntov_sheet["G14"].value = udel_scep4
            zadanie_gruntov_sheet["H14"].value = udel_scep5
            zadanie_gruntov_sheet["D15"].value = ugol_vn_tr1
            zadanie_gruntov_sheet["E15"].value = ugol_vn_tr2
            zadanie_gruntov_sheet["F15"].value = ugol_vn_tr3
            zadanie_gruntov_sheet["G15"].value = ugol_vn_tr4
            zadanie_gruntov_sheet["H15"].value = ugol_vn_tr5
            zadanie_gruntov_sheet["D16"].value = ves_gr_prir1
            zadanie_gruntov_sheet["E16"].value = ves_gr_prir2
            zadanie_gruntov_sheet["F16"].value = ves_gr_prir3
            zadanie_gruntov_sheet["G16"].value = ves_gr_prir4
            zadanie_gruntov_sheet["H16"].value = ves_gr_prir5
            zadanie_gruntov_sheet["D18"].value = def_mod1
            zadanie_gruntov_sheet["E18"].value = def_mod2
            zadanie_gruntov_sheet["F18"].value = def_mod3
            zadanie_gruntov_sheet["G18"].value = def_mod4
            zadanie_gruntov_sheet["H18"].value = def_mod5
    else:
        interface_sheet["B8"].value = "Исходных данных нет"
        typical_ground_sheet["B24"].value = typical_ground
        typical_ground_sheet["D24"].value = ugol_vntr_tr
        typical_ground_sheet["E24"].value = udel_sceplenie
        typical_ground_sheet["F24"].value = ves_grunta
        typical_ground_sheet["G24"].value = deform_module
        typical_ground_sheet["H24"].value = coef_usl_rab
        typical_ground_sheet["I24"].value = coef_nadej
    
    zadanie_gruntov_sheet["B26"].value = pole_type

    # workbook.save(calculation_file)
    # workbook.close()

    # wb = openpyxl.load_workbook(calculation_file, data_only=True)

    # zad_gr_sheet = wb["Задание грунтов"]
    # ras_svai_sheet = wb["Расчет сваи"]

    if is_init_data:
        result.update({
            "sr_udel_scep": zadanie_gruntov_sheet["K14"].value,
            "sr_ugol_vn_tr": zadanie_gruntov_sheet["K15"].value,
            "sr_ves_gr_ras": zadanie_gruntov_sheet["K17"].value,
            "sr_def_mod": zadanie_gruntov_sheet["K18"].value
        })

    coef_isp_s245 = round(float(calculation_sheet["H26"].value), 2)
    coef_isp_s345 = round(float(calculation_sheet["H27"].value), 2)
    coef_isp_gor = round(float(calculation_sheet["D98"].value), 2)
    ugol_pov = float(calculation_sheet["D107"].value)

    result.update(
        {
            "coef_isp_s245": coef_isp_s245,
            "coef_isp_s345": coef_isp_s345,
            "coef_isp_gor": coef_isp_gor,
            "ugol_pov": ugol_pov
        }
    )

    if coef_isp_s245 <= 0.5:
        result.update({"coef_isp_s245_bg": "#11ED02"})
    elif 0.50 < coef_isp_s245 <= 0.65:
        result.update({"coef_isp_s245_bg": "#F9FF05"})
    elif 0.65 < coef_isp_s245 <= 0.8:
        result.update({"coef_isp_s245_bg": "#FFB905"})
    elif 0.8 < coef_isp_s245 <= 0.95:
        result.update({"coef_isp_s245_bg": "#A16030"})
    else:
        result.update({"coef_isp_s245_bg": "#FA0D00"})
    if coef_isp_s345 <= 0.5:
        result.update({"coef_isp_s345_bg": "#11ED02"})
    elif 0.50 < coef_isp_s345 <= 0.65:
        result.update({"coef_isp_s345_bg": "#F9FF05"})
    elif 0.65 < coef_isp_s345 <= 0.8:
        result.update({"coef_isp_s345_bg": "#FFB905"})
    elif 0.8 < coef_isp_s345 <= 0.95:
        result.update({"coef_isp_s345_bg": "#A16030"})
    else:
        result.update({"coef_isp_s345_bg": "#FA0D00"})
    if coef_isp_gor <= 0.5:
        result.update({"coef_isp_gor_bg": "#11ED02"})
    elif 0.50 < coef_isp_gor <= 0.65:
        result.update({"coef_isp_gor_bg": "#F9FF05"})
    elif 0.65 < coef_isp_gor <= 0.8:
        result.update({"coef_isp_gor_bg": "#FFB905"})
    elif 0.8 < coef_isp_gor <= 0.95:
        result.update({"coef_isp_gor_bg": "#A16030"})
    else:
        result.update({"coef_isp_gor_bg": "#FA0D00"})
    if ugol_pov <= 0.0005:
        result.update({"ugol_pov_bg": "#11ED02"})
    elif 0.00005 < ugol_pov <= 0.00065:
        result.update({"ugol_pov_bg": "#F9FF05"})
    elif 0.00065 < ugol_pov <= 0.0008:
        result.update({"ugol_pov_bg": "#FFB905"})
    elif 0.0008 < ugol_pov <= 0.00095:
        result.update({"ugol_pov_bg": "#A16030"})
    else:
        result.update({"ugol_pov_bg": "#FA0D00"})

    workbook.save()
    workbook.close()

    return result