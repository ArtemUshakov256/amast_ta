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
from docx2pdf import convert
from pandas import ExcelWriter
from pypdf import PdfMerger
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
    sp_wind_region,
    wind_nagr,
    sp_ice_region,
    snow_nagr,
    golol_rayon,
    rvs,
    diam_osn,
    diam_verha,
    h,
    teor_massa_metala,
    ploschad_uchastka,
    territoria_raspoloj,
    god_vvoda_v_ekspl,
    min_temp,
    max_temp,
    speca,
    speca_pz
):
    filename = "core\\static\\kozu_tkr_template.docx"
    if getattr(sys, 'frozen', False):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.abspath(".")
    filepath = os.path.join(base_path, filename)
    
    doc_tkr = DocxTemplate(filepath)

    context_tkr = {
        "project_name": project_name,
        "project_code": project_code,
        "year": dt.date.today().year,
        "developer": developer,
        "mm_yy": mm_yy,
        "sp_wind_region": sp_wind_region,
        "wind_nagr": wind_nagr,
        "sp_ice_region": sp_ice_region,
        "snow_nagr": snow_nagr,
        "golol_rayon": golol_rayon,
        "rvs": rvs,
        "diam_osn": diam_osn,
        "diam_verha": diam_verha,
        "h": h,
        "teor_massa_metala": teor_massa_metala,
        "ploschad_uchastka": ploschad_uchastka,
        "territoria_raspoloj": territoria_raspoloj,
        "god_vvoda_v_ekspl": god_vvoda_v_ekspl,
        "min_temp": min_temp,
        "max_temp": max_temp,
        "current_date": current_date,
        "speca": InlineImage(doc_tkr,image_descriptor=speca, width=Mm(100), height=Mm(170))
    }

    dir_name_tkr = fd.asksaveasfilename(
                filetypes=[("docx file", ".docx")],
                defaultextension=".docx"
            )
    if dir_name_tkr:
        doc_tkr.render(context_tkr)
        doc_tkr.save(dir_name_tkr)
        tkr_pdf = dir_name_tkr[:dir_name_tkr.rindex(".")] + ".pdf"
        convert(dir_name_tkr, tkr_pdf)

    filename_pzg = "core\\static\\kozu_pzg_template.docx"
    if getattr(sys, 'frozen', False):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.abspath(".")
    filepath_pzg = os.path.join(base_path, filename_pzg)
    
    doc_pzg = DocxTemplate(filepath_pzg)

    context_pzg = {
        "project_code": project_code,
        "year": dt.date.today().year,
        "developer": developer,
        "mm_yy": mm_yy,
        "current_date": current_date
    }

    dir_name_pzg = fd.asksaveasfilename(
                filetypes=[("docx file", ".docx")],
                defaultextension=".docx"
            )
    if dir_name_pzg:
        doc_pzg.render(context_pzg)
        doc_pzg.save(dir_name_pzg)
        pzg_pdf = dir_name_pzg[:dir_name_pzg.rindex(".")] + ".pdf"
        convert(dir_name_pzg, pzg_pdf)

    filename_pz = "core\\static\\kozu_pz_template.docx"
    if getattr(sys, 'frozen', False):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.abspath(".")
    filepath_pz = os.path.join(base_path, filename_pz)
    
    doc_pz = DocxTemplate(filepath_pz)

    context_pz = {
        "speca_pz": InlineImage(doc_pz,image_descriptor=speca_pz, width=Mm(120), height=Mm(190)),
        "project_name": project_name,
        "project_code": project_code,
        "year": dt.date.today().year,
        "developer": developer,
        "mm_yy": mm_yy,
        "sp_wind_region": sp_wind_region,
        "wind_nagr": wind_nagr,
        "sp_ice_region": sp_ice_region,
        "snow_nagr": snow_nagr,
        "golol_rayon": golol_rayon,
        "rvs": rvs,
        "diam_osn": diam_osn,
        "diam_verha": diam_verha,
        "h": h,
        "teor_massa_metala": teor_massa_metala,
        "min_temp": min_temp,
        "max_temp": max_temp,
        "current_date": current_date
    }

    dir_name_pz = fd.asksaveasfilename(
                filetypes=[("docx file", ".docx")],
                defaultextension=".docx"
            )
    if dir_name_pz:
        doc_pz.render(context_pz)
        doc_pz.save(dir_name_pz)
        pz_pdf = dir_name_pz[:dir_name_pz.rindex(".")] + ".pdf"
        convert(dir_name_pz, pz_pdf)
    pdfs = [tkr_pdf, pzg_pdf, pz_pdf]
    merger = PdfMerger()
    for pdf in pdfs:
        merger.append(pdf)
    compilated_pdf = fd.asksaveasfilename(
                filetypes=[("pdf file", ".pdf")],
                defaultextension=".pdf"
            )
    if compilated_pdf:
        merger.write(compilated_pdf)
        merger.close()