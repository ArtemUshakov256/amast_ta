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
    make_kozu_schema
    )


def make_vor(
    n,
    h,
    d_nijn,
    d_verh,
    m,
    is_gabion
):
    calculation_file = get_file_path("core\static\kozu_vor.xlsx")
    app = xw.App(visible=False)
    workbook = app.books.open(calculation_file)

    initial_data_sheet = workbook.sheets["Исходные данные"]
    bez_gabiona_sheet = workbook.sheets["ВОР без габиона"]
    s_gabionom_sheet = workbook.sheets["ВОР с габионом"]

    initial_data_sheet["D2"].value = n
    initial_data_sheet["C3"].value = h
    initial_data_sheet["B4"].value = d_nijn
    initial_data_sheet["B5"].value = d_verh
    initial_data_sheet["F10"].value = m

    if is_gabion:
        vor = s_gabionom_sheet.to_pdf()
        print(vor)
        return vor
    else:
        vor = bez_gabiona_sheet.to_pdf()
    
    workbook.save()
    workbook.close()
    app.quit()


def make_tkr(
    project_name,
    project_code,
    developer,
    sp_wind_region,
    wind_nagr,
    sp_ice_region,
    snow_nagr,
    golol_rayon,
    rvs1,
    rvs2,
    rvs3,
    rvs4,
    diam_osn1,
    diam_osn2,
    diam_osn3,
    diam_osn4,
    diam_verha1,
    diam_verha2,
    diam_verha3,
    diam_verha4,
    h1,
    h2,
    h3,
    h4,
    massa_rvs1,
    massa_rvs2,
    massa_rvs3,
    massa_rvs4,
    kol_rvs1,
    kol_rvs2,
    kol_rvs3,
    kol_rvs4,
    ploschad_uchastka1,
    ploschad_uchastka2,
    ploschad_uchastka3,
    ploschad_uchastka4,
    zasch_obj,
    territoria_raspoloj,
    god_vvoda_v_ekspl,
    min_temp,
    max_temp,
    speca,
    speca_pz1,
    speca_pz2,
    speca_pz3,
    speca_pz4,
    vid_kozu1,
    vid_kozu2,
    vid_kozu3,
    vid_kozu4,
    rayon_str,
    eskiz_kozu,
    quantity_of_rvs,
    is_gabion,
    list_sogl,
    kont_zazel,
    mont_schema
):
    rvs_list = [rvs2, rvs3, rvs4]
    ploschad_uchastka_list = [ploschad_uchastka2, ploschad_uchastka3, ploschad_uchastka4]
    rvs = f"РВС-{rvs1}"
    ploschad_uchastka = ploschad_uchastka1
    for i in range(3):
        if rvs_list[i] and ploschad_uchastka_list[i]:
            rvs += f", РВС-{rvs_list[i]}"
            ploschad_uchastka += f", {ploschad_uchastka_list[i]}"

    rvs1 = int(rvs1) if rvs1 else 0
    rvs2 = int(rvs2) if rvs2 else 0
    rvs3 = int(rvs3) if rvs3 else 0
    rvs4 = int(rvs4) if rvs4 else 0
    kozu_tkr, kozu_pz_100_3k, kozu_pz_5k_10k, kozu_pz_20k_30k, kozu_pz_40k_50k = [], [], [], [], []
    kozu_parameters = [
        (rvs1, diam_osn1, diam_verha1, h1, massa_rvs1, kol_rvs1),
        (rvs2, diam_osn2, diam_verha2, h2, massa_rvs2, kol_rvs2),
        (rvs3, diam_osn3, diam_verha3, h3, massa_rvs3, kol_rvs3),
        (rvs4, diam_osn4, diam_verha4, h4, massa_rvs4, kol_rvs4)
        ]
    for i in range(int(quantity_of_rvs)):
        kozu_tkr.extend([
            f"Технические характеристики защитного сооружения для РВС-{kozu_parameters[i][0]}, на 1 шт:",
            f"1) Диаметр в основании - {kozu_parameters[i][1]}, мм",
            f"2) Диаметр верхней части - {kozu_parameters[i][2]}, мм",
            f"3) Высота от основания - {kozu_parameters[i][3]}, мм",
            f"4) Металлоескость металлокаркаса - {kozu_parameters[i][4]}, т"
        ])
        if 0 < kozu_parameters[i][0] <= 3000:
            obsch_massa1 = round(float(kozu_parameters[i][4]) * int(kozu_parameters[i][5]), 1)
            kozu_pz_100_3k.extend([
                "Технические характеристики защитного сооружения",
                f"1) Диаметр в основании - {kozu_parameters[i][1]}, мм",
                f"2) Диаметр верхней части - {kozu_parameters[i][2]}, мм",
                f"3) Высота от основания - {kozu_parameters[i][3]}, мм",
                f"4) Общая металлоескость металлокаркаса - {obsch_massa1}, т"
            ])
        elif 3000 < kozu_parameters[i][0] <= 10000:
            obsch_massa2 = round(float(kozu_parameters[i][4]) * int(kozu_parameters[i][5]), 1)
            kozu_pz_5k_10k.extend([
                "Технические характеристики защитного сооружения",
                f"1) Диаметр в основании - {kozu_parameters[i][1]}, мм",
                f"2) Диаметр верхней части - {kozu_parameters[i][2]}, мм",
                f"3) Высота от основания - {kozu_parameters[i][3]}, мм",
                f"4) Общая металлоескость металлокаркаса - {obsch_massa2}, т"
            ])
        elif 10000 < kozu_parameters[i][0] <= 30000:
            obsch_massa3 = round(float(kozu_parameters[i][4]) * int(kozu_parameters[i][5]), 1)
            kozu_pz_20k_30k.extend([
                "Технические характеристики защитного сооружения",
                f"1) Диаметр в основании - {kozu_parameters[i][1]}, мм",
                f"2) Диаметр верхней части - {kozu_parameters[i][2]}, мм",
                f"3) Высота от основания - {kozu_parameters[i][3]}, мм",
                f"4) Общая металлоескость металлокаркаса - {obsch_massa3}, т"
            ])
        elif 30000 < kozu_parameters[i][0] <= 50000:
            obsch_massa4 = round(float(kozu_parameters[i][4]) * int(kozu_parameters[i][5]), 1)
            kozu_pz_40k_50k.extend([
                "Технические характеристики защитного сооружения",
                f"1) Диаметр в основании - {kozu_parameters[i][1]}, мм",
                f"2) Диаметр верхней части - {kozu_parameters[i][2]}, мм",
                f"3) Высота от основания - {kozu_parameters[i][3]}, мм",
                f"4) Общая металлоескость металлокаркаса - {obsch_massa4}, т"
            ])
            
    filepath = get_file_path("core\\static\\kozu_tkr_template.docx")
    
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
        "god_vvoda_v_ekspl": god_vvoda_v_ekspl,
        "min_temp": min_temp,
        "max_temp": max_temp,
        "current_date": current_date,
        "speca": InlineImage(doc_tkr,image_descriptor=speca, width=Mm(100), height=Mm(170)),
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
        "current_date": current_date
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
        # vor_list.append(make_vor)
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

    pdfs = [list_sogl] + [tkr_pdf] + [eskiz_kozu] + pz_list + [pzo_pdf] +\
        schema_pdf + [kont_zazel, mont_schema, certificates_pdf_path]
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
        project_code = thisdict["project_code"] + "-КВП"
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