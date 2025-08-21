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
from core.utils import (
    mm_yy,
    current_date,
    Kompas_work,
    do_magic,
    get_file_path,
    make_kozu_schema,
    make_word
)


def make_tkr(
    project_name,
    project_code,
    developer,
    min_temp,
    max_temp,
    sp_wind_region,
    wind_nagr,
    golol_rayon,
    v_m_bpla,
    zasch_obj,
    steel,
    anal_steel,
    konst,
    kp_bolt,
    fundament,
    fund_osn,
    rayon_str,
    kol_obj,
    massa_obsch,
    territoria_raspoloj,
    mont_vremya,
    expl_god,
    shag_yach,
    post_nagr,
    strela,
    natyajenie,
    klass_betona,
    morozostoikost,
    vodonepronicaemost,
    quantity_of_obj,
    dlina_elem,
    zaschichaemyi_obj1,
    zaschichaemyi_obj2,
    zaschichaemyi_obj3,
    zaschichaemyi_obj4,
    length_kozup1,
    length_kozup2,
    length_kozup3,
    length_kozup4,
    width_kozup1,
    width_kozup2,
    width_kozup3,
    width_kozup4,
    h1,
    h2,
    h3,
    h4,
    massa_kozup1,
    massa_kozup2,
    massa_kozup3,
    massa_kozup4,
    speca,
    vid_kozu_p,
    table3,
    table4,
    table5,
    table6,
    raschet_model,
    usiliya1,
    usiliya2,
    usiliya3,
    usiliya4,
    usiliya5
):
    fund_elem = fund[f"{fundament}"]
    # ground1 = ground

    filepath = get_file_path("core\\static\\kozu-p_tkr_template.docx")
    
    doc_tkr = DocxTemplate(filepath)

    developers = ["Беляева", "Горохов", "Горшенев",
                  "Денисенко", "Кокорев", "Мельситов",
                  "Миронов", "Перелыгин", "Ушаков"]
    podp = get_file_path(f"core\\static\\{developer}.png") if developer in developers\
    else get_file_path("core\\static\\Ушаков.png")

    context_tkr = {
        "project_name": project_name,
        "project_code": project_code,
        "year": dt.date.today().year,
        "bartal_code": bartal_code,
        "dlina_stoiki": dlina_stoiki,
        "dlina_rigelya_1": dlina_rigelya_1,
        "dlina_rigelya_2": dlina_rigelya_2,
        "set": seti,
        "ten": ten,
        "fund_or_ferma_usil": fund_or_ferma_usil,
        "object_titul": object_titul,
        "zasch_obj_list": zasch_obj_list,
        "vid_kozu_p": InlineImage(doc_tkr,image_descriptor=vid_kozu_p, width=Mm(152), height=Mm(121)),
        "rayon_str": rayon_str,
        "str_klim_zone": str_klim_zona,
        "vid_klim": vid_klim,
        "sp_wind_reg": sp_wind_region,
        "wind_nagr": wind_nagr,
        "golol_rayon": golol_rayon,
        "golol_thick": golol_thick,
        "seism": seism,
        "sbros": sbros,
        "fundament": fundament,
        "fund_osn": fund_osn,
        "kozu_pz": kozu_pz,
        "ground0": ground0,
        "grounding_initial_data": grounding_initial_data,
        "r1": r1,
        "h": h,
        "r2": r2,
        "ground1": ground1,
        "kozup_height": kozup_height,
    }

    dir_name_tkr = fd.asksaveasfilename(
                filetypes=[("docx file", ".docx")],
                defaultextension=".docx"
            )
    if dir_name_tkr:
        doc_tkr.render(context_tkr)
        doc_tkr.save(dir_name_tkr)
        # tkr_pdf = dir_name_tkr[:dir_name_tkr.rindex(".")] + ".pdf"
        # convert(dir_name_tkr, tkr_pdf)
    
    kozu_pz = []
    kozu_parameters = [
        (zaschichaemyi_obj1, length_kozup1, width_kozup1, h1, massa_kozup1),
        (zaschichaemyi_obj2, length_kozup2, width_kozup2, h2, massa_kozup2),
        (zaschichaemyi_obj3, length_kozup3, width_kozup3, h3, massa_kozup3),
        (zaschichaemyi_obj4, length_kozup4, width_kozup4, h4, massa_kozup4)
    ]
    for i in range(int(quantity_of_obj)):
        kozu_pz.extend([
            f"Технические характеристики защитного сооружения {kozu_parameters[i][0]}",
            f"1) Длина в осях - {kozu_parameters[i][1]}, мм",
            f"2) Ширина в осях - {kozu_parameters[i][2]}, мм",
            f"3) Высота - {kozu_parameters[i][3]}, мм",
            f"4) Металлоемкость металлокаркаса - {kozu_parameters[i][4]}, т"
        ])
    
    vid_kozu_paths = [pic_dir.strip("}{") for pic_dir in vid_kozu_p.split("} {")]
        
    filepath_pz = get_file_path("core\\static\\kozu-p_pz_template.docx")
    
    doc_pz = DocxTemplate(filepath_pz)
    
    vid_kozu_obj = []
    for path in vid_kozu_paths:
        vid_kozu_obj.append(InlineImage(doc_pz, image_descriptor=path, width=Mm(10), height=Mm(10)))
    
    context_pz = {
        "project_name": project_name,
        "project_code": project_code,
        "year": dt.date.today().year,
        "developer": developer,
        "mm_yy": mm_yy,
        "sp_wind_reg": sp_wind_region,
        "wind_nagr": wind_nagr,
        "golol_rayon": golol_rayon,
        "zasch_obj": zasch_obj,
        "min_temp": min_temp,
        "max_temp": max_temp,
        "current_date": current_date,
        "v_m_bpla": v_m_bpla,
        "zasch_obj":zasch_obj,
        "konst": konst,
        "vid_kozu_p": vid_kozu_obj,
        "steel": steel,
        "anal_steel": anal_steel,
        "kozu_pz": kozu_pz,
        "fundament": fundament,
        "fund_elem": fund_elem,
        "fund_osn": fund_osn,
        "post_nagr": post_nagr,
        "strela": strela,
        "natyajenie": natyajenie,
        "ish_schema": InlineImage(doc_pz, image_descriptor=table3, width=Mm(121), height=Mm(110)),
        "rasch_model_sverhu": InlineImage(doc_pz, image_descriptor=table4, width=Mm(121), height=Mm(110)),
        "rasch_model1": InlineImage(doc_pz, image_descriptor=table5, width=Mm(121), height=Mm(110)),
        "coef_isp": InlineImage(doc_pz, image_descriptor=table6, width=Mm(121), height=Mm(110)),
        "perem_x": InlineImage(doc_pz, image_descriptor=raschet_model, width=Mm(121), height=Mm(110)),
        "perem_y": InlineImage(doc_pz, image_descriptor=usiliya1, width=Mm(121), height=Mm(110)),
        "perem_z": InlineImage(doc_pz, image_descriptor=usiliya2, width=Mm(121), height=Mm(110)),
        "usiliya3": InlineImage(doc_pz, image_descriptor=usiliya3, width=Mm(153), height=Mm(104)),
        "usiliya4": InlineImage(doc_pz, image_descriptor=usiliya4, width=Mm(153), height=Mm(104)),
        "usiliya5": InlineImage(doc_pz, image_descriptor=usiliya5, width=Mm(153), height=Mm(104)),
        "podp": InlineImage(doc_pz,image_descriptor=podp, width=Mm(6), height=Mm(5))
    }

    dir_name_pz = fd.asksaveasfilename(
                filetypes=[("docx file", ".docx")],
                defaultextension=".docx"
            )
    if dir_name_pz:
        doc_pz.render(context_pz)
        doc_pz.save(dir_name_pz)

    filepath_pzf = get_file_path("core\\static\\kozu_pzo_template.docx")
    
    doc_pzf = DocxTemplate(filepath_pzf)

    context_pzf = {
        "project_code": project_code,
        "project_name": project_name,
        "year": dt.date.today().year,
        "developer": developer,
        "mm_yy": mm_yy,
        "current_date": current_date,
        "klass_betona": klass_betona,
        "morozostoikost": morozostoikost,
        "vodonepronicaemost": vodonepronicaemost,
        "dlina_elem": dlina_elem,
        "podp": InlineImage(doc_pzf,image_descriptor=podp, width=Mm(6), height=Mm(5))
    }

    dir_name_pzf = fd.asksaveasfilename(
                filetypes=[("docx file", ".docx")],
                defaultextension=".docx"
            )
    if dir_name_pzf:
        doc_pzf.render(context_pzf)
        doc_pzf.save(dir_name_pzf)
        # pzf_pdf = dir_name_pzf[:dir_name_pzf.rindex(".")] + ".pdf"
        # convert(dir_name_pzf, pzf_pdf)