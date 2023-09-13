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
from core.utils import mm_yy


def calculate_bolt(
        bend_moment,
        vert_force,
        pole_diam,
        bolt,
        kol_boltov,
        bolt_class
):
    calculation_file = os.path.abspath(
        "core/static/Анкерные_закладные/Расчет_анкерных_болтов.xlsx"
    )
    app = xw.App(visible=False)
    workbook = app.books.open(calculation_file)
    sheet = workbook.sheets["Фланец"]

    result = dict()
    sheet["A4"].value = pole_diam
    sheet["H3"].value = bolt
    sheet["E18"].value = bend_moment
    sheet["E19"].value = vert_force
    sheet["E23"].value = kol_boltov
    sheet["U3"].value = bolt_class
    diam_okr_bolt = round(float(str(sheet["A11"].value).replace(",", ".")))
    diam_flanca = round(float(str(sheet["A14"].value).replace(",", ".")))
    coef_isp_m24 = round(float(str(sheet["F43"].value).replace(",", ".")), 2)
    coef_isp_m30 = round(float(str(sheet["F44"].value).replace(",", ".")), 2)
    coef_isp_m36 = round(float(str(sheet["F45"].value).replace(",", ".")), 2)
    coef_isp_m42 = round(float(str(sheet["F46"].value).replace(",", ".")), 2)
    coef_isp_m48 = round(float(str(sheet["F47"].value).replace(",", ".")), 2)
    coef_isp_m52 = round(float(str(sheet["F48"].value).replace(",", ".")), 2)
    coef_isp_m56 = round(float(str(sheet["F49"].value).replace(",", ".")), 2)

    result.update({
        "diam_okr_bolt": diam_okr_bolt,
        "diam_flanca": diam_flanca,
        "coef_isp_m24": coef_isp_m24,
        "coef_isp_m30": coef_isp_m30,
        "coef_isp_m36": coef_isp_m36,
        "coef_isp_m42": coef_isp_m42,
        "coef_isp_m48": coef_isp_m48,
        "coef_isp_m52": coef_isp_m52,
        "coef_isp_m56": coef_isp_m56
    })

    if coef_isp_m24 <= 0.6:
        result.update({"coef_isp_m24_bg": "#11ED02"})
    elif 0.60 < coef_isp_m24 <= 0.8:
        result.update({"coef_isp_m24_bg": "#F9FF05"})
    elif 0.8 < coef_isp_m24 <= 0.95:
        result.update({"coef_isp_m24_bg": "#FFB905"})
    else:
        result.update({"coef_isp_m24_bg": "#FA0D00"})
    if coef_isp_m30 <= 0.6:
        result.update({"coef_isp_m30_bg": "#11ED02"})
    elif 0.60 < coef_isp_m30 <= 0.8:
        result.update({"coef_isp_m30_bg": "#F9FF05"})
    elif 0.8 < coef_isp_m30 <= 0.95:
        result.update({"coef_isp_m30_bg": "#FFB905"})
    else:
        result.update({"coef_isp_m30_bg": "#FA0D00"})
    if coef_isp_m36 <= 0.6:
        result.update({"coef_isp_m36_bg": "#11ED02"})
    elif 0.60 < coef_isp_m36 <= 0.8:
        result.update({"coef_isp_m36_bg": "#F9FF05"})
    elif 0.8 < coef_isp_m36 <= 0.95:
        result.update({"coef_isp_m36_bg": "#FFB905"})
    else:
        result.update({"coef_isp_m36_bg": "#FA0D00"})
    if coef_isp_m42 <= 0.6:
        result.update({"coef_isp_m42_bg": "#11ED02"})
    elif 0.60 < coef_isp_m42 <= 0.8:
        result.update({"coef_isp_m42_bg": "#F9FF05"})
    elif 0.8 < coef_isp_m42 <= 0.95:
        result.update({"coef_isp_m42_bg": "#FFB905"})
    else:
        result.update({"coef_isp_m42_bg": "#FA0D00"})
    if coef_isp_m48 <= 0.6:
        result.update({"coef_isp_m48_bg": "#11ED02"})
    elif 0.60 < coef_isp_m48 <= 0.8:
        result.update({"coef_isp_m48_bg": "#F9FF05"})
    elif 0.8 < coef_isp_m48 <= 0.95:
        result.update({"coef_isp_m48_bg": "#FFB905"})
    else:
        result.update({"coef_isp_m48_bg": "#FA0D00"})
    if coef_isp_m52 <= 0.6:
        result.update({"coef_isp_m52_bg": "#11ED02"})
    elif 0.60 < coef_isp_m52 <= 0.8:
        result.update({"coef_isp_m52_bg": "#F9FF05"})
    elif 0.8 < coef_isp_m52 <= 0.95:
        result.update({"coef_isp_m52_bg": "#FFB905"})
    else:
        result.update({"coef_isp_m52_bg": "#FA0D00"})
    if coef_isp_m56 <= 0.6:
        result.update({"coef_isp_m56_bg": "#11ED02"})
    elif 0.60 < coef_isp_m56 <= 0.8:
        result.update({"coef_isp_m56_bg": "#F9FF05"})
    elif 0.8 < coef_isp_m56 <= 0.95:
        result.update({"coef_isp_m56_bg": "#FFB905"})
    else:
        result.update({"coef_isp_m56_bg": "#FA0D00"})

    workbook.save()
    workbook.close()
    app.quit()

    return result


def save_xlsx():
    calculation_file = os.path.abspath(
        "core/static/Анкерные_закладные/Расчет_анкерных_болтов.xlsx"
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


def make_rpzaz(
        project_name,
        project_code,
        developer,
        moment,
        vert_force,
        shear_force,
        bolt_xlsx_path
    ):
    filename = "core/static/Анкерные_закладные/rpzaz_template.docx"
    if getattr(sys, 'frozen', False):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.abspath(".")
    filepath = os.path.join(base_path, filename)
    
    doc1 = DocxTemplate(filepath)

    app = xw.App(visible=False)
    workbook = app.books.open(bolt_xlsx_path)
    sheet = workbook.sheets["Фланец"]

    ploshad_sech = bolt_dict[str(sheet["H3"].value)][2]
    d = 35 * int(str(sheet["H3"].value)[-2:])

    context = {
        "project_name": project_name,
        "project_code": project_code,
        "year": dt.date.today().year,
        "developer": developer,
        "mm_yy": mm_yy,
        "moment": moment,
        "vert_force": vert_force,
        "shear_force": shear_force,
        "E38": round(float(str(sheet["E38"].value).replace(",", ".")), 2),
        "T41": round(float(str(sheet["T41"].value).replace(",", ".")), 2),
        "ploshad_sech": ploshad_sech,
        "D": d,
        "bolt": sheet["H3"].value,
        "class_prochnosti": sheet["U3"].value
    }

    workbook.close()
    app.quit()

    dir_name = fd.asksaveasfilename(
                filetypes=[("docx file", ".docx")],
                defaultextension=".docx"
            )
    if dir_name:
        doc1.render(context)
        doc1.save(dir_name)


def make_pasport(
    project_name,
    project_code,
    pole_code,
    bolt_xlsx_path,
    picture1_path
):
    filename = "core/static/Анкерные_закладные/Паспорт_закладной_детали.docx"
    if getattr(sys, 'frozen', False):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.abspath(".")
    filepath = os.path.join(base_path, filename)
    
    doc = DocxTemplate(filepath)

    app = xw.App(visible=False)
    workbook = app.books.open(bolt_xlsx_path)
    sheet = workbook.sheets["Фланец"]
    dlina_bolta = 35 * int(str(sheet["H3"].value)[-2:]) + 250
    x = float(str(bolt_dict[sheet["H3"].value][3]).replace(",", "."))\
        + round(float(str(sheet["A11"].value).replace(",", ".")))
    y = float(str(bolt_dict[sheet["H3"].value][4]).replace(",", "."))\
        + round(float(str(sheet["A11"].value).replace(",", ".")))

    context = {
        "project_name": project_name,
        "project_code": project_code,
        "pole_code": pole_code,
        "year": dt.date.today().year,
        "kol_bolt": int(sheet["E23"].value),
        "bolt": sheet["H3"].value,
        "dlina_bolta": dlina_bolta,
        "diam_okr_bolt": round(float(str(sheet["A11"].value).replace(",", "."))),
        "x": x,
        "y": y,
        "bolt_x_dlina": str(sheet["H3"].value) + " x " + str(dlina_bolta),
        "bolt_bez_m": str(sheet["H3"].value)[-2:],
        "kol_gaika": 5 * int(sheet["E23"].value),
        "kol_shaiba": 4 * int(sheet["E23"].value),
        "picture1": InlineImage(
                doc,image_descriptor=picture1_path, width=Mm(140), height=Mm(200)
            )
    }

    workbook.close()
    app.quit()

    dir_name = fd.asksaveasfilename(
                filetypes=[("docx file", ".docx")],
                defaultextension=".docx"
            )
    if dir_name:
        doc.render(context)
        doc.save(dir_name)