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
from core.utils import mm_yy, current_date


def make_tkr(
    project_name,
    project_code,
    developer,
    general_info,
    klimat,
    relief,
    geologia,
    flora,
    gidrologia,
    diam_osn,
    diam_verha,
    h,
    teor_massa_metala,
    ploschad_uchastka,
    territoria_raspoloj,
    god_vvoda_v_ekspl,
    wind_region,
    wind_pressure,
    area,
    ice_region,
    ice_thickness,
    ice_wind_pressure,
    year_average_temp,
    min_temp,
    wind_temp,
    ice_temp,
    max_temp,
    wind_reg_coef,
    ice_reg_coef,
    seism_rayon
):
    filename = "core\\static\\kozu_tkr_template.docx"
    if getattr(sys, 'frozen', False):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.abspath(".")
    filepath = os.path.join(base_path, filename)
    
    doc1 = DocxTemplate(filepath)

    context = {
        "project_name": project_name,
        "project_code": project_code,
        "year": dt.date.today().year,
        "developer": developer,
        "mm_yy": mm_yy,
        "general_info": general_info,
        "klimat": klimat,
        "relief": relief,
        "geologia": geologia,
        "flora": flora,
        "gidrologia": gidrologia,
        "diam_osn": diam_osn,
        "diam_verha": diam_verha,
        "h": h,
        "teor_massa_metala": teor_massa_metala,
        "ploschad_uchastka": ploschad_uchastka,
        "territoria_raspoloj": territoria_raspoloj,
        "god_vvoda_v_ekspl": god_vvoda_v_ekspl,
        "area": area,
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
        "seism_rayon": seism_rayon,
        "current_date": current_date
    }

    dir_name = fd.asksaveasfilename(
                filetypes=[("docx file", ".docx")],
                defaultextension=".docx"
            )
    if dir_name:
        doc1.render(context)
        doc1.save(dir_name)


def make_pz(
    project_code,
    developer
):
    filename = "core\\static\\kozu_pz_template.docx"
    if getattr(sys, 'frozen', False):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.abspath(".")
    filepath = os.path.join(base_path, filename)
    
    doc1 = DocxTemplate(filepath)

    context = {
        "project_code": project_code,
        "year": dt.date.today().year,
        "developer": developer,
        "mm_yy": mm_yy,
        "current_date": current_date
    }

    dir_name = fd.asksaveasfilename(
                filetypes=[("docx file", ".docx")],
                defaultextension=".docx"
            )
    if dir_name:
        doc1.render(context)
        doc1.save(dir_name)