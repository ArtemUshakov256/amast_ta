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
):
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
            
    filepath = get_file_path("core\\static\\kozu-p_tkr_template.docx")
    
    doc_tkr = DocxTemplate(filepath)

    context_tkr = {
        "project_name": project_name,
        "project_code": project_code,
        "year": dt.date.today().year,
        "developer": developer,
        "mm_yy": mm_yy,
        "rayon_str": rayon_str,
        "sp_wind_region": sp_wind_region,
        "wind_nagr": wind_nagr,
        "sp_ice_region": sp_ice_region,
        "snow_nagr": snow_nagr,
        "golol_rayon": golol_rayon,
        "kozu_tkr": kozu_tkr,
        "rvs": rvs,
        "ploschad_uchastka": ploschad_uchastka,
        "territoria_raspoloj": territoria_raspoloj,
        "min_temp": min_temp,
        "max_temp": max_temp,
        "current_date": current_date,
        "speca": InlineImage(doc_tkr,image_descriptor=speca, width=Mm(100), height=Mm(170)),
        "pesch_gr": pesch_gr
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
    
    context_pz = {
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
        "zasch_obj": zasch_obj,
        "min_temp": min_temp,
        "max_temp": max_temp,
        "current_date": current_date,
        "pesch_gr": pesch_gr
    }
    
    pz_list, vor_list = [], []
    kozu_schema_dict = {
        "100_3k": "",
        "5k_50k": ""
    }

    if kozu_pz_100_3k:
        filepath_pz_100_3k = get_file_path("core\\static\\kozu_pz_template_100_3k.docx")
        doc_pz = DocxTemplate(filepath_pz_100_3k)
        context_pz["rvs"] = rvs1
        context_pz["kozu_pz"] = kozu_pz_100_3k
        context_pz["speca_pz"] = InlineImage(doc_pz,image_descriptor=speca_pz1, width=Mm(120), height=Mm(140))
        context_pz["vid_kozu"] = InlineImage(doc_pz,image_descriptor=vid_kozu1, width=Mm(120), height=Mm(140))
        kozu_schema_path_100_3k = get_file_path("core\\static\\kozu_schema_100_3k.cdw")
        kozu_schema_dict["100_3k"] = kozu_schema_path_100_3k
        vor_list.append(make_vor)
        dir_name_pz_100_3k = fd.asksaveasfilename(
            filetypes=[("docx file", ".docx")],
            defaultextension=".docx"
        )
        if dir_name_pz_100_3k:
            doc_pz.render(context_pz)
            doc_pz.save(dir_name_pz_100_3k)
            pz_pdf_100_3k = dir_name_pz_100_3k[:dir_name_pz_100_3k.rindex(".")] + ".pdf"
            convert(dir_name_pz_100_3k, pz_pdf_100_3k)
            pz_list.append(pz_pdf_100_3k)
    if kozu_pz_5k_10k:
        filepath_pz_5k_10k = get_file_path("core\\static\\kozu_pz_template_5k_10k.docx")
        doc_pz = DocxTemplate(filepath_pz_5k_10k)
        context_pz["rvs"] = rvs2
        context_pz["kozu_pz"] = kozu_pz_5k_10k
        context_pz["speca_pz"] = InlineImage(doc_pz,image_descriptor=speca_pz2, width=Mm(120), height=Mm(140))
        context_pz["vid_kozu"] = InlineImage(doc_pz,image_descriptor=vid_kozu2, width=Mm(120), height=Mm(140))
        kozu_schema_path_5k_50k = get_file_path("core\\static\\kozu_schema_5k_50k.cdw")
        kozu_schema_dict["5k_50k"] = kozu_schema_path_5k_50k
        dir_name_pz_5k_10k = fd.asksaveasfilename(
            filetypes=[("docx file", ".docx")],
            defaultextension=".docx"
        )
        if dir_name_pz_5k_10k:
            doc_pz.render(context_pz)
            doc_pz.save(dir_name_pz_5k_10k)
            pz_pdf_5k_10k = dir_name_pz_5k_10k[:dir_name_pz_5k_10k.rindex(".")] + ".pdf"
            convert(dir_name_pz_5k_10k, pz_pdf_5k_10k)
            pz_list.append(pz_pdf_5k_10k)
    if kozu_pz_20k_30k:
        filepath_pz_20k_30k = get_file_path("core\\static\\kozu_pz_template_20k_30k.docx")
        doc_pz = DocxTemplate(filepath_pz_20k_30k)
        context_pz["rvs"] = rvs3
        context_pz["kozu_pz"] = kozu_pz_20k_30k
        context_pz["speca_pz"] = InlineImage(doc_pz,image_descriptor=speca_pz3, width=Mm(120), height=Mm(140))
        context_pz["vid_kozu"] = InlineImage(doc_pz,image_descriptor=vid_kozu3, width=Mm(120), height=Mm(140))
        kozu_schema_path_5k_50k = get_file_path("core\\static\\kozu_schema_5k_50k.cdw")
        kozu_schema_dict["5k_50k"] = kozu_schema_path_5k_50k
        dir_name_pz_20k_30k = fd.asksaveasfilename(
            filetypes=[("docx file", ".docx")],
            defaultextension=".docx"
        )
        if dir_name_pz_20k_30k:
            doc_pz.render(context_pz)
            doc_pz.save(dir_name_pz_20k_30k)
            pz_pdf_20k_30k = dir_name_pz_20k_30k[:dir_name_pz_20k_30k.rindex(".")] + ".pdf"
            convert(dir_name_pz_20k_30k, pz_pdf_20k_30k)
            pz_list.append(pz_pdf_20k_30k)
    if kozu_pz_40k_50k:
        filepath_pz_40k_50k = get_file_path("core\\static\\kozu_pz_template_40k_50k.docx")
        doc_pz = DocxTemplate(filepath_pz_40k_50k)
        context_pz["rvs"] = rvs4
        context_pz["kozu_pz"] = kozu_pz_40k_50k
        context_pz["speca_pz"] = InlineImage(doc_pz,image_descriptor=speca_pz4, width=Mm(120), height=Mm(140))
        context_pz["vid_kozu"] = InlineImage(doc_pz,image_descriptor=vid_kozu4, width=Mm(120), height=Mm(140))
        kozu_schema_path_5k_50k = get_file_path("core\\static\\kozu_schema_5k_50k.cdw")
        kozu_schema_dict["5k_50k"] = kozu_schema_path_5k_50k
        dir_name_pz_40k_50k = fd.asksaveasfilename(
            filetypes=[("docx file", ".docx")],
            defaultextension=".docx"
        )
        if dir_name_pz_40k_50k:
            doc_pz.render(context_pz)
            doc_pz.save(dir_name_pz_40k_50k)
            pz_pdf_40k_50k = dir_name_pz_40k_50k[:dir_name_pz_40k_50k.rindex(".")] + ".pdf"
            convert(dir_name_pz_40k_50k, pz_pdf_40k_50k)
            pz_list.append(pz_pdf_40k_50k)

    filepath_pzo = get_file_path("core\\static\\kozu_pzo_template.docx")
    
    doc_pzo = DocxTemplate(filepath_pzo)

    context_pzo = {
        "project_code": project_code,
        "project_name": project_name,
        "year": dt.date.today().year,
        "developer": developer,
        "mm_yy": mm_yy,
        "current_date": current_date
    }

    dir_name_pzo = fd.asksaveasfilename(
                filetypes=[("docx file", ".docx")],
                defaultextension=".docx"
            )
    if dir_name_pzo:
        doc_pzo.render(context_pzo)
        doc_pzo.save(dir_name_pzo)
        pzo_pdf = dir_name_pzo[:dir_name_pzo.rindex(".")] + ".pdf"
        convert(dir_name_pzo, pzo_pdf)

    stamp_data = {
        "project_code": project_code,
        "project_name": project_name,
        "developer": developer
    }

    schema_pdf = []

    if kozu_schema_dict["100_3k"]:
        schema_pdf_path_100_3k = make_kozu_schema(stamp_data, Kozu, kozu_schema_dict["100_3k"])
        schema_pdf.append(schema_pdf_path_100_3k)
    if kozu_schema_dict["5k_50k"]:
        schema_pdf_path_5k_50k = make_kozu_schema(stamp_data, Kozu, kozu_schema_dict["5k_50k"])
        schema_pdf.append(schema_pdf_path_5k_50k)
    certificates_pdf_path = get_file_path("core\\static\\kozu_certificates.pdf")

    titul_vor_dict = {
        "project_code": project_code,
        "project_name": project_name,
        "year": dt.date.today().year,
        "current_date": current_date
    }
    titul_vor = make_word(
        "core\\static\\titul_vor.docx",
        titul_vor_dict
    )
    titul_vor_pdf = titul_vor[:titul_vor.rindex(".")] + ".pdf"
    convert(titul_vor, titul_vor_pdf)

    rvs_pdf_list = []
    for rvs in kozu_parameters:
        if rvs[0]:
            print(rvs[0])
            rvs_pdf = make_vor(
                n=rvs[5],
                h=int(rvs[3])/1000,
                d_nijn=int(rvs[1])/1000,
                d_verh=int(rvs[2])/1000,
                m=rvs[4],
                is_gabion=is_gabion,
                rvs=rvs[0]
            )
            rvs_pdf_list.append(rvs_pdf)

    pdfs = [list_sogl] + [tkr_pdf] + [eskiz_kozu] + pz_list + [pzo_pdf] + schema_pdf + \
        [kont_zazel, mont_schema] + [titul_vor_pdf] + rvs_pdf_list + [certificates_pdf_path]
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


class Kozu(Kompas_work):
    def do_events(self, thisdict, path):
        if thisdict == None: return

        from core.utils import KompasAPI
        kompas = KompasAPI()    
        schema_pdf_path = self.drawing_work(kompas, thisdict, path)
        kompas.application.Visible=True
        return schema_pdf_path

    def drawing_work(self,kompas,thisdict, path):
        from core.utils import DrawingsAPI
        drw = DrawingsAPI(kompas)
        path_kozu_schema = path
        kompas.open_2D_file(path_kozu_schema)
        project_code = thisdict["project_code"] + "-КВП-ПЧ"
        drw.change_stamp(
                mm_yy,
                project_code,
                thisdict['project_name'],
                "Беляева",
                thisdict['developer'],
                "Беляева",
        )
        schema_pdf_path = drw.save_as_pdf(path_kozu_schema)
        drw.save_as_Kompas()
        kompas.close_2D_file()
        return schema_pdf_path