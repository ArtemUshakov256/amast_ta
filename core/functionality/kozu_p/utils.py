import datetime as dt
import openpyxl
import os
import pathlib
import pandas as pd
import random
import re
import sys
import xlwings as xw
import fitz

from dotenv import load_dotenv
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
from docx2pdf import convert
from pandas import ExcelWriter
from pypdf import PdfMerger
from sympy import symbols, latex
from tkinter import filedialog as fd
from io import BytesIO
from docx.shared import Inches
from typing import List, Union

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
    voltage,
    temp_max,
    temp_min,
    temp_avg,
    temp_max95,
    temp_min98,
    is_svaya,
    is_tros,
    is_ferma,
    is_existing_ground,
    sp_wind_region,
    wind_nagr,
    golol_rayon,
    golol_thick,
    str_klim,
    vid_klim,
    seism,
    bartal_code,
    object_titul,
    seti,
    fundament,
    rayon_str,
    sbros,
    grounding_initial_data,
    r1,
    r2,
    h,
    isol_rast,
    raspr_nagr,
    quantity_of_obj,
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
    # dlina_rigelya1_1,
    # dlina_rigelya1_2,
    # dlina_rigelya1_3,
    # dlina_rigelya1_4,
    # dlina_rigelya2_1,
    # dlina_rigelya2_2,
    # dlina_rigelya2_3,
    # dlina_rigelya2_4,
    # dlina_rigelya3_1,
    # dlina_rigelya3_2,
    # dlina_rigelya3_3,
    # dlina_rigelya3_4,
    # dlina_rigelya4_1,
    # dlina_rigelya4_2,
    # dlina_rigelya4_3,
    # dlina_rigelya4_4,
    # dlina_rigelya5_1,
    # dlina_rigelya5_2,
    # dlina_rigelya5_3,
    # dlina_rigelya5_4,
    # dlina_stoiki1,
    # dlina_stoiki2,
    # dlina_stoiki3,
    # dlina_stoiki4,
    titul_list,
    vid_kozu_p,
    ish_schema,
    rasch_model_sverhu,
    rasch_model1,
    rasch_model2,
    coef_isp,
    perem_x,
    perem_y,
    perem_z,
    prodolnoe_usil,
    m_y,
    m_z,
    q_z,
    q_y,
    nagr_v_rigel,
    sum_peremesch_v_rigel,
    nagr_v_uzel,
    sum_peremesch_uzel,
    nagr_g_rigel,
    sum_peremesch_g_rigel,
    usil_osn,
    usil_n,
    usil_m,
    usil_q,
):
    filepath = get_file_path("core\\static\\kozu_p_pz_template.docx")
    
    doc_tkr = DocxTemplate(filepath)

    # developers = ["Беляева", "Горохов", "Горшенев",
    #               "Денисенко", "Кокорев", "Мельситов",
    #               "Миронов", "Перелыгин", "Ушаков"]
    # podp = get_file_path(f"core\\static\\{developer}.png") if developer in developers\
    # else get_file_path("core\\static\\Ушаков.png")

    titul = pdf_to_png_bytes(titul_list)
    titul = [InlineImage(doc_tkr, image_descriptor=BytesIO(image), width=Mm(210), height=Mm(297))\
             for image in titul]

    # stoiki_list = [dlina_stoiki1, dlina_stoiki2, dlina_stoiki3, dlina_stoiki4] 
    # rigeli1_list = [dlina_rigelya1_1, dlina_rigelya1_2, dlina_rigelya1_3, dlina_rigelya1_4]
    # rigeli2_list = [dlina_rigelya2_1, dlina_rigelya2_2, dlina_rigelya2_3, dlina_rigelya2_4]
    # dlina_stoiki = ", ".join([stoika for stoika in stoiki_list if stoika])
    # dlina_rigelya1 = ", ".join([rigel1 for rigel1 in rigeli1_list if rigel1])
    # dlina_rigelya2 = ", ".join([rigel2 for rigel2 in rigeli2_list if rigel2])

    if is_svaya and is_ferma:
        ustr_svai = USTR_SVAI
        ferma_usil = "Ферма усиления"
        ten = "10"
        eleven = "11"
    elif is_svaya:
        ustr_svai = USTR_SVAI
        ten = "10"
        eleven = ""
    elif is_ferma:
        ferma_usil = "Ферма усиления"
        eleven = "11"
        ten = ""
    else:
        ustr_svai = ""
        ferma_usil = ""
        ten = ""
        eleven = ""

    zasch_obj_list = []
    kozu_parameters = [
        (zaschichaemyi_obj1, length_kozup1, width_kozup1, h1, massa_kozup1),
        (zaschichaemyi_obj2, length_kozup2, width_kozup2, h2, massa_kozup2),
        (zaschichaemyi_obj3, length_kozup3, width_kozup3, h3, massa_kozup3),
        (zaschichaemyi_obj4, length_kozup4, width_kozup4, h4, massa_kozup4)
    ]
    for i in range(int(quantity_of_obj)):
        zasch_obj_list.extend([
            f"Технические характеристики защитного сооружения {kozu_parameters[i][0]}",
            f"1) Длина в осях - {kozu_parameters[i][1]}, мм",
            f"2) Ширина в осях - {kozu_parameters[i][2]}, мм",
            f"3) Высота - {kozu_parameters[i][3]}, мм",
            f"4) Металлоемкость металлокаркаса - {kozu_parameters[i][4]}, т"
        ])

    vid_kozu_paths = [pic_dir.strip("}{") for pic_dir in vid_kozu_p.split("} {")]
    vid_kozu_obj = []
    for path in vid_kozu_paths:
        vid_kozu_obj.append(InlineImage(doc_tkr, image_descriptor=path, width=Mm(10), height=Mm(10)))

    ferma = FERMA if is_ferma else ""
    tros = TROSOVAYA_FERMA if is_tros else ""
    
    ferma_list = [ferm for ferm in [ferma, tros] if ferm]

    ground0 = GROUND[str(is_existing_ground)][0]
    ground1 = GROUND[str(is_existing_ground)][1]
    
    context_tkr = {
        "titul_list": titul,
        "project_name": project_name,
        "project_code": project_code,
        "year": dt.date.today().year,
        "bartal_code": bartal_code,
        # "dlina_stoiki": dlina_stoiki,
        # "dlina_rigelya_1": dlina_rigelya1,
        # "dlina_rigelya_2": dlina_rigelya2,
        "set": seti,
        "ten": ten,
        "eleven": eleven,
        "ustr_svai": ustr_svai,
        "ferma_usil": ferma_usil,
        "object_titul": object_titul,
        "zasch_obj_list": zasch_obj_list,
        "vid_kozu_p": vid_kozu_obj,
        "rayon_str": rayon_str,
        "str_klim_zone": str_klim,
        "vid_klim": vid_klim,
        "sp_wind_reg": sp_wind_region,
        "wind_nagr": wind_nagr,
        "temp_max": temp_max,
        "temp_min": temp_min,
        "temp_avg": temp_avg,
        "temp_max95": temp_max95,
        "temp_min98":temp_min98,
        "golol_rayon": golol_rayon,
        "golol_thick": golol_thick,
        "seism": seism,
        "ferma_list": ferma_list,
        "sbros": SBROS[sbros],
        "fundament": FUND_DICT[fundament],
        "fund_osn": fund[fundament],
        "ground0": ground0,
        "grounding_initial_data": grounding_initial_data,
        "r1": r1,
        "h": h,
        "r2": r2,
        "ground1": ground1,
        "kozup_height": max(h1, h2, h3, h4),
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
    
    filepath_pz = get_file_path("core\\static\\kozu_p_pril_template.docx")
    
    doc_pz = DocxTemplate(filepath_pz)
    
    context_pz = {
        "project_name": project_name,
        "project_code": project_code,
        "object_titul": object_titul,
        "year": dt.date.today().year,
        "isol_rast": isol_rast,
        "voltage": voltage,
        "raspr_nagr": raspr_nagr,
        "sp_wind_reg": sp_wind_region,
        "wind_nagr": wind_nagr,
        "golol_rayon": golol_rayon,
        "golol_thick": golol_thick,
        "ish_schema": InlineImage(doc_pz, image_descriptor=ish_schema, width=Mm(121), height=Mm(110)),
        "rasch_model_sverhu": InlineImage(doc_pz, image_descriptor=rasch_model_sverhu, width=Mm(121), height=Mm(110)),
        "rasch_model1": InlineImage(doc_pz, image_descriptor=rasch_model1, width=Mm(121), height=Mm(110)),
        "rasch_model2": InlineImage(doc_pz, image_descriptor=rasch_model2, width=Mm(121), height=Mm(110)),
        "coef_isp": InlineImage(doc_pz, image_descriptor=coef_isp, width=Mm(121), height=Mm(110)),
        "perem_x": InlineImage(doc_pz, image_descriptor=perem_x, width=Mm(121), height=Mm(110)),
        "perem_y": InlineImage(doc_pz, image_descriptor=perem_y, width=Mm(121), height=Mm(110)),
        "perem_z": InlineImage(doc_pz, image_descriptor=perem_z, width=Mm(121), height=Mm(110)),
        "prodolnoe_usil": InlineImage(doc_pz, image_descriptor=prodolnoe_usil, width=Mm(121), height=Mm(110)),
        "m_y": InlineImage(doc_pz, image_descriptor=m_y, width=Mm(121), height=Mm(110)),
        "m_z": InlineImage(doc_pz, image_descriptor=m_z, width=Mm(121), height=Mm(110)),
        "q_z": InlineImage(doc_pz, image_descriptor=q_z, width=Mm(121), height=Mm(110)),
        "q_y": InlineImage(doc_pz, image_descriptor=q_y, width=Mm(121), height=Mm(110)),
        "nagr_v_rigel": InlineImage(doc_pz, image_descriptor=nagr_v_rigel, width=Mm(121), height=Mm(110)),
        "nagr_v_uzel": InlineImage(doc_pz, image_descriptor=nagr_v_uzel, width=Mm(121), height=Mm(110)),
        "sum_peremesch_uzel": InlineImage(doc_pz, image_descriptor=sum_peremesch_uzel, width=Mm(121), height=Mm(110)),
        "sum_peremesch_v_rigel": InlineImage(doc_pz, image_descriptor=sum_peremesch_v_rigel, width=Mm(121), height=Mm(110)),
        "nagr_g_rigel": InlineImage(doc_pz, image_descriptor=nagr_g_rigel, width=Mm(121), height=Mm(110)),
        "sum_peremesch_g_rigel": InlineImage(doc_pz, image_descriptor=sum_peremesch_g_rigel, width=Mm(121), height=Mm(110)),
        "usil_osn": InlineImage(doc_pz, image_descriptor=usil_osn, width=Mm(121), height=Mm(110)),
        "usil_n": InlineImage(doc_pz, image_descriptor=usil_n, width=Mm(121), height=Mm(110)),
        "usil_m": InlineImage(doc_pz, image_descriptor=usil_m, width=Mm(121), height=Mm(110)),
        "usil_q": InlineImage(doc_pz, image_descriptor=usil_q, width=Mm(121), height=Mm(110)),
    }

    dir_name_pz = fd.asksaveasfilename(
                filetypes=[("docx file", ".docx")],
                defaultextension=".docx"
            )
    if dir_name_pz:
        doc_pz.render(context_pz)
        doc_pz.save(dir_name_pz)


def pdf_to_png_bytes(
    pdf_source: Union[str, pathlib.Path, bytes],
    dpi: int = 200,
) -> List[bytes]:
    """
    Возвращает список PNG в виде bytes для каждой страницы PDF (0..N-1).
    pdf_source: путь к PDF или bytes содержимое PDF.
    dpi: 150-220 обычно достаточно; выше — больше размер.
    """
    if isinstance(pdf_source, (str, pathlib.Path)):
        doc = fitz.open(str(pdf_source))
    elif isinstance(pdf_source, (bytes, bytearray, memoryview)):
        doc = fitz.open(stream=pdf_source, filetype="pdf")
    else:
        raise TypeError("pdf_source must be path or bytes")

    zoom = dpi / 72.0
    mat = fitz.Matrix(zoom, zoom)

    png_list: List[bytes] = []
    for page in doc:
        pix = page.get_pixmap(matrix=mat, alpha=False)  # alpha=False -> RGB
        png_list.append(pix.tobytes("png"))             # готовый PNG как bytes
    return png_list


def calculate_pril_i(n, m, q, fi, ci, l, b):

    # fv = float(float(n) + float(l) * float(b) * 0.5 * 2.5 * 2.5, 1)
    # eb = float(m / fv, 3)
    # bi = float(b) - 2 * eb
    # nyu = float(float(l) / float(b), 2)
    # ey = float(1 - 0.25 / nyu, 2)
    # eq = float(1 + 1.5 / nyu, 2)
    # ec = float(1 + 0.3 / nyu, 2)
    # nu = bi * float(l) * ("" * ey * bi * )

    # return {
    #     "Fv": fv,
    #     "e": eb,
    #     "b": bi,
    #     "nyu": nyu,
    #     "Ey": ey,
    #     "Eq": eq,
    #     "Ec": ec,
    #     "Nu": nu,
    #     "F": f
    # }
    pass