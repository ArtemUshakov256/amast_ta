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
from core.utils import (
    mm_yy,
    do_magic,
    Kompas_work
)
from math import atan, tan, pi


def calculate_foundation(
        moment,
        vert_force,
        shear_force,
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
        coef_nadej
):
    calculation_file = os.path.abspath("core\static\Расчет_сваи.xlsx")
    app = xw.App(visible=False)
    workbook = app.books.open(calculation_file)

    interface_sheet = workbook.sheets["Интерфейс"]
    typical_ground_sheet = workbook.sheets["Типовые грунты"]
    zadanie_gruntov_sheet = workbook.sheets["Задание грунтов"]
    calculation_sheet = workbook.sheets["Расчет сваи"]
    flanec_calculation_sheet = workbook.sheets["Расчет массы фланца"]

    flanec_calculation_sheet["C3"].value = flanec_diam
    flanec_calculation_sheet["C4"].value = thickness_svai
    flanec_calculation_sheet["C5"].value = deepness_svai
    calculation_sheet["I7"].value = int(height_svai) / 1000
    calculation_sheet["I14"].value = moment
    calculation_sheet["I16"].value = shear_force
    calculation_sheet["I18"].value = vert_force
    result = dict()
    if is_init_data:
        interface_sheet["B8"].value = "Исходные данные есть"
        zadanie_gruntov_sheet["B23"].value = ground_water_lvl
        zadanie_gruntov_sheet["C26"].value = coef_nadej
        coef_usl_rab_list = []
        actual_coef_nadej_dict = {
           nijn_sloy1: ground_type1,
           nijn_sloy2: ground_type2,
           nijn_sloy3: ground_type3,
           nijn_sloy4: ground_type4,
           nijn_sloy5: ground_type5 
        }
        nijn_sloy_list = [nijn_sloy1, nijn_sloy2, nijn_sloy3, nijn_sloy4, nijn_sloy5]
        nijn_sloy_list = [elem for elem in nijn_sloy_list if elem.strip()]
        for elem in nijn_sloy_list:
            coef = actual_coef_nadej_dict[f"{elem}"]
            if len(nijn_sloy_list) == 1:
                coef_usl_rab_list.append(coef_usl_rab_dict[f"{coef}"])
                break
            if ground_water_lvl > elem and elem:
                coef_usl_rab_list.append(coef_usl_rab_dict[f"{coef}"])

        zadanie_gruntov_sheet["C28"].value = min(coef_usl_rab_list)
        if quantity_of_ige == "1":
            for i in range(6, 19):
                if i in [11, 12, 17]:
                    continue
                else:
                    zadanie_gruntov_sheet[f"E{i}"].value = None
                    zadanie_gruntov_sheet[f"F{i}"].value = None
                    zadanie_gruntov_sheet[f"G{i}"].value = None
                    zadanie_gruntov_sheet[f"H{i}"].value = None
            zadanie_gruntov_sheet["D6"].value = nomer_ige1
            zadanie_gruntov_sheet["D7"].value = ground_type1
            zadanie_gruntov_sheet["D8"].value = ground_name1
            zadanie_gruntov_sheet["D9"].value = verh_sloy1
            zadanie_gruntov_sheet["D10"].value = nijn_sloy1
            zadanie_gruntov_sheet["D13"].value = coef_poristosti1
            zadanie_gruntov_sheet["D14"].value = udel_scep1
            zadanie_gruntov_sheet["D15"].value = ugol_vn_tr1
            zadanie_gruntov_sheet["D16"].value = ves_gr_prir1
            zadanie_gruntov_sheet["D18"].value = float(def_mod1.replace(",", ".")) * 1000
        elif quantity_of_ige == "2":
            for i in range(6, 19):
                if i in [11, 12, 17]:
                    continue
                else:
                    zadanie_gruntov_sheet[f"F{i}"].value = None
                    zadanie_gruntov_sheet[f"G{i}"].value = None
                    zadanie_gruntov_sheet[f"H{i}"].value = None
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
            zadanie_gruntov_sheet["D18"].value = float(def_mod1.replace(",", ".")) * 1000
            zadanie_gruntov_sheet["E18"].value = float(def_mod2.replace(",", ".")) * 1000
        elif quantity_of_ige == "3":
            for i in range(6, 19):
                if i in [11, 12, 17]:
                    continue
                else:
                    zadanie_gruntov_sheet[f"G{i}"].value = None
                    zadanie_gruntov_sheet[f"H{i}"].value = None
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
            zadanie_gruntov_sheet["D18"].value = float(def_mod1.replace(",", ".")) * 1000
            zadanie_gruntov_sheet["E18"].value = float(def_mod2.replace(",", ".")) * 1000
            zadanie_gruntov_sheet["F18"].value = float(def_mod3.replace(",", ".")) * 1000
        elif quantity_of_ige == "4":
            for i in range(6, 19):
                if i in [11, 12, 17]:
                    continue
                else:
                    zadanie_gruntov_sheet[f"H{i}"].value = None
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
            zadanie_gruntov_sheet["D18"].value = float(def_mod1.replace(",", ".")) * 1000
            zadanie_gruntov_sheet["E18"].value = float(def_mod2.replace(",", ".")) * 1000
            zadanie_gruntov_sheet["F18"].value = float(def_mod3.replace(",", ".")) * 1000
            zadanie_gruntov_sheet["G18"].value = float(def_mod4.replace(",", ".")) * 1000
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
            zadanie_gruntov_sheet["D18"].value = float(def_mod1.replace(",", ".")) * 1000
            zadanie_gruntov_sheet["E18"].value = float(def_mod2.replace(",", ".")) * 1000
            zadanie_gruntov_sheet["F18"].value = float(def_mod3.replace(",", ".")) * 1000
            zadanie_gruntov_sheet["G18"].value = float(def_mod4.replace(",", ".")) * 1000
            zadanie_gruntov_sheet["H18"].value = float(def_mod5.replace(",", ".")) * 1000
    else:
        interface_sheet["B8"].value = "Исходных данных нет"
        typical_ground_sheet["B24"].value = typical_ground
        typical_ground_sheet["D24"].value = ugol_vntr_tr
        typical_ground_sheet["E24"].value = udel_sceplenie
        typical_ground_sheet["F24"].value = ves_grunta
        typical_ground_sheet["G24"].value = deform_module
        typical_ground_key = typical_ground.split(" ")[0]
        typical_ground_sheet["H24"].value = coef_usl_rab_dict[f"{typical_ground_key}"]
        typical_ground_sheet["I24"].value = coef_nadej
    
    zadanie_gruntov_sheet["B26"].value = pole_type

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
    app.quit()

    return result


def save_xlsx():
    calculation_file = os.path.abspath(
        "core\static\Расчет_сваи.xlsx"
    )
    app = xw.App(visible=False)
    workbook = app.books.open(calculation_file)
    dir_name = fd.asksaveasfilename(
                filetypes=[("xlsx file", ".xlsx")],
                defaultextension=".xlsx"
            )
    if dir_name:
        workbook.save(dir_name)
    workbook.close()
    app.quit()


def make_rpzf(
        project_name,
        project_code,
        pole_code,
        developer,
        diam_svai,
        deepness_svai,
        height_svai,
        moment1,
        vert_force1,
        shear_force1,
        ige_name,
        building_adress,
        razrez_skvajin,
        picture1,
        picture2,
        xlsx_svai
):
    filepath = os.path.abspath("core\\static\\rpzf_template.docx")
    doc = DocxTemplate(filepath)

    length_svai = (float(deepness_svai) + float(height_svai)) / 1000
    moment2 = round(float(moment1) / 9.8 * 1.1 / 1.3, 2)
    vert_force2 = round(float(vert_force1) / 9.8 * 1.2 / 1.3, 2)
    shear_force2 = round(float(shear_force1) / 9.8 * 1.1 / 1.2, 2)
    height_shear_force = round(float(moment1) / float(shear_force1), 2)
    
    app = xw.App(visible=False)
    workbook = app.books.open(xlsx_svai)
    calculation_sheet = workbook.sheets["Расчет сваи"]
    sr_ugol_vn_tr = calculation_sheet["D9"].value
    sr_udel_scep = calculation_sheet["D7"].value
    sr_ves_gr_ras = float(calculation_sheet["D13"].value) * 1000
    sr_def_mod = calculation_sheet["D15"].value
    
    ugol_sdviga = round(atan(tan(sr_ugol_vn_tr)\
         + (sr_udel_scep / 10)) * (180 / pi), 2)
    coef_tr_met_po_gr = round(0.9 * (tan(sr_ugol_vn_tr) - 0.1), 2)
    coef_ep_dav = round(1 - 0.03 * float(sr_udel_scep), 2)

    context = {
            "project_name": project_name,
            "project_code": project_code,
            "year": dt.date.today().year,
            "mm_yy": mm_yy,
            "pole_code": pole_code,
            "developer": developer,
            "diam_svai": float(diam_svai) / 1000,
            "length_svai": length_svai,
            "deepness_svai": float(deepness_svai) / 1000,
            "moment1": round(moment1 / 9.8, 2),
            "vert_force1": round(vert_force1 / 9.8, 2),
            "shear_force1": round(shear_force1 / 9.8, 2),
            "moment2": moment2,
            "vert_force2": vert_force2,
            "shear_force2": shear_force2,
            "ige_name": ige_name,
            "building_adress": building_adress,
            "razrez_skvajin": razrez_skvajin,
            "picture1": InlineImage(
                doc,image_descriptor=picture1, width=Mm(172), height=Mm(146)
            ),
            "sr_ugol_vn_tr": sr_ugol_vn_tr,
            "sr_udel_scep": sr_udel_scep,
            "sr_ves_gr_ras": sr_ves_gr_ras,
            "sr_def_mod": sr_def_mod,
            "height_shear_force": height_shear_force,
            "ugol_sdviga": ugol_sdviga,
            "coef_tr_met_po_gr": coef_tr_met_po_gr,
            "coef_ep_dav": coef_ep_dav,
            "D85": round(float(str(calculation_sheet["D85"].value).replace(",", ".")), 2),
            "D86": round(float(str(calculation_sheet["D86"].value).replace(",", ".")), 2),
            "D87": round(float(str(calculation_sheet["D87"].value).replace(",", ".")), 2),
            "D92": round(float(str(calculation_sheet["D92"].value).replace(",", ".")), 2),
            "D93": round(float(str(calculation_sheet["D93"].value).replace(",", ".")), 2),
            "D94": round(float(str(calculation_sheet["D94"].value).replace(",", ".")), 2),
            "D96": round(float(str(calculation_sheet["D96"].value).replace(",", ".")), 2),
            "D45": round(float(str(calculation_sheet["D45"].value).replace(",", ".")), 2),
            "D29": round(float(str(calculation_sheet["D29"].value).replace(",", ".")), 2),
            "D50": round(float(str(calculation_sheet["D50"].value).replace(",", ".")), 2),
            "D64": round(float(str(calculation_sheet["D64"].value).replace(",", ".")), 2),
            "D104": round(float(str(calculation_sheet["D104"].value).replace(",", ".")), 2),
            "D107": round(float(str(calculation_sheet["D107"].value).replace(",", ".")), 2),
            "picture2": InlineImage(
                doc,image_descriptor=picture2, width=Mm(165), height=Mm(123)
            )
    }

    for i in range(64, 80):
        context.update({f"D{i}": round(float(str(calculation_sheet[f"D{i}"].value).replace(",", ".")), 2)})
    
    workbook.close()
    app.quit()

    dir_name = fd.asksaveasfilename(
                filetypes=[("docx file", ".docx")],
                defaultextension=".docx"
            )
    doc.render(context)
    if dir_name:
        doc.save(dir_name)


class Flanec(Kompas_work):
    def do_events(self, thisdict):
        if thisdict == None: return

        from core.utils import KompasAPI
        kompas = KompasAPI()
        self.assembly_work(kompas,thisdict)
        kompas.application.Visible=True

    def assembly_work(self,kompas,thisdict):
        from core.utils import AssemblyAPI
        
        kompas.open_3D_file(thisdict['path_flanec_assembly'])
        ass = AssemblyAPI(kompas)
        
        params_for_change_variables_dic = self.get_params_variables(thisdict)
        ass.change_external_variables(params_for_change_variables_dic)
        
        self.drawing_work(kompas,thisdict)

        ass.save_as_file(thisdict['default_path_result']+"\\KOMPAS",'Фланец-сборка')
        kompas.close_assmebly()

    def drawing_work(self,kompas,thisdict):
        from core.utils import DrawingsAPI
        drw = DrawingsAPI(kompas)
        kompas.open_2D_file(thisdict['path_flanec_drawing'])
        
        drw.change_stamp(
                thisdict['stamp_date'],
                thisdict['stamp_number'].replace(chr(92),"\\"),
                thisdict['stamp_caption'],
                thisdict['stamp_proveril'],
                thisdict['stamp_razrabotal'],
                thisdict['stamp_gip'],
        )
        drw.save_as_pdf(
                thisdict['path_flanec_drawing'],
                thisdict['default_path_result']+"\\PDF",
                "51 - Фланец чертеж"
        )
        drw.save_as_Kompas(
                thisdict['default_path_result']+"\\KOMPAS",
                'Фланец-чертеж'
        )
        kompas.close_2D_file()

    def get_params_variables(self,thisdict):
        params_for_change_variables_dic = {
                'bolts_quantity': thisdict['flanec_bolts_quantity'],
                'hole_diameter': thisdict['flanec_hole_diameter'],
                'wall_distance': int(round(float(thisdict['flanec_wall_distance'].replace(",",".")),0)),
                'diameter_flanca': int(round(float(thisdict['flanec_diameter_flanca'].replace(",",".")),0)),
                'diameter_truby': int(round(float(thisdict['flanec_diameter_truby'].replace(",",".")),0))
        }
        return params_for_change_variables_dic


class Svai(Kompas_work):
    def do_events(self, thisdict):
        if thisdict == None: return

        from core.utils import KompasAPI
        kompas = KompasAPI()    
        self.assembly_work(kompas, thisdict)
        kompas.application.Visible=True

    def assembly_work(self, kompas, thisdict):
        import os.path
        path_svai_assembly = os.path.abspath("core\\static\\свая.a3d")
        if not os.path.isfile(path_svai_assembly):
            print(f"ОШИБКА!!\n\nФайл сборки {path_svai_assembly} не найден")
            
        from core.utils import AssemblyAPI
        kompas.open_3D_file(path_svai_assembly)
        ass = AssemblyAPI(kompas)
        
        params_for_change_variables_dic = self.get_params_variables(thisdict)
        ass.change_external_variables(params_for_change_variables_dic)
        
        self.drawing_work(kompas, thisdict)

        ass.save_as_file()
        kompas.close_assmebly()

    def drawing_work(self,kompas,thisdict):
        from core.utils import DrawingsAPI
        drw = DrawingsAPI(kompas)
        path_svai_drawing = os.path.abspath("core\\static\\свая.cdw")
        kompas.open_2D_file(path_svai_drawing)

        import datetime as dt
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
                'bolts_quantity': thisdict['kol_boltov'],
                'hole_diameter': thisdict['hole_diam'],
                'wall_distance': int(thisdict['rast_m']),
                'diameter_flanca': int(thisdict['flanec_diam']),
                'diameter_truby': int(thisdict['diam_svai']),
                'dlina_svai': thisdict['dlina_svai']
        }
        return params_for_change_variables_dic


class Anker(Kompas_work):
    def do_events(self, thisdict):
        if thisdict == None: return
        
        from core.utils import KompasAPI
        kompas = KompasAPI()
        self.assembly_work(kompas, thisdict)
        kompas.application.Visible=True

    def assembly_work(self, kompas, thisdict):
        import os.path
        path_anker_assembly = os.path.abspath("core\\static\\anker.a3d")
        if not os.path.isfile(path_anker_assembly):
            print(f"ОШИБКА!!\n\nФайл сборки {path_anker_assembly} не найден")

        from core.utils import AssemblyAPI
        kompas.open_3D_file(path_anker_assembly)
        ass = AssemblyAPI(kompas)
        
        ass.change_external_variables(self.get_params_variables(thisdict))

        self.drawing_work(kompas, thisdict)

        ass.save_as_file()
        kompas.close_assmebly()

    def drawing_work(self,kompas,thisdict):
        from core.utils import DrawingsAPI
        drw = DrawingsAPI(kompas)
        path_anker_drawing = os.path.abspath("core\\static\\anker.cdw")
        kompas.open_2D_file(path_anker_drawing)
        
        import datetime as dt
        current_date = dt.datetime.now()
        mm_yy = current_date.strftime("%m.%Y")
        drw.change_stamp(
                mm_yy,
                thisdict['project_code'],
                thisdict['project_name'],
                "Родчихин",
                thisdict['developer'],
                "Родчихин",
                1
        )
        drw.save_as_pdf(path_anker_drawing)
        drw.save_as_Kompas()
        kompas.close_2D_file()

    def get_params_variables(self,thisdict):
        params_for_change_variables_dic = {
                'dk': int(thisdict['diam_okr_bolt']),
                'dap': int(thisdict['diam_okr_bolt']),
                'doap': int(thisdict['hole_diam']),
                'dok': int(thisdict['hole_diam']),
                'Nap': thisdict['kol_boltov'],
                'Nk': thisdict['kol_boltov'],
                'x': int(thisdict['x']),
                'y': int(thisdict['y'])
        }
        return params_for_change_variables_dic

def make_schema(thisdict, num):
    if num == 0:
        do_magic(thisdict, Svai)
    if num == 1:
        do_magic(thisdict, Anker)
    