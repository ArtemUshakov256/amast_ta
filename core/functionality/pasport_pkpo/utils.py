import datetime as dt
import os
import pathlib
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
)

def make_passport(
    pole_code,
    voltage,
    wind_region,
    ice_region,
    max_temp,
    min_temp,
    kol_km,
    kol_ap_zaj_km,
    kol_ap_zaj_opn,
    kol_otv_zaj,
    kol_opn,
    kol_kab_krep,
    konc_kor,
    kol_konc_kor,
    ppsa,
    ppsa_length,
    kol_skoba_bol,
    kol_skoba_mal,
    kol_styajka,
    razed,
    syst_telemekh,
    izmerit_ustr,
    panel_rel_zasch,
    syst_sobstv_nujd,
    syst_temp_monit,
    obor_antiterror,
    rezerv_kabelya,
    vch_vols,
    pole_heigth,
    tr_heigth,
    akep,
    tros_heigth,
    bot_diam,
    top_diam,
    sech_jili,
    sech_ekr,
    sech_al_prov,
    naib_rab_voltage,
    naib_dlit_dop_rab_voltage,
    oprosniy_list
):
    filepath = os.path.abspath("core\\static\\pasport_pkpo.docx")
    doc = DocxTemplate(filepath)

    tr_heigth = tr_heigth.replace(", ", "\n")
    oprosniy_list = [pic_dir.strip("}{") for pic_dir in oprosniy_list.split("} {")]
    oprosniy_list = [InlineImage(doc, image_descriptor=pic, width=Mm(170), height=Mm(240)) for pic in oprosniy_list]
    razed = "1 комплект" if razed else "по запросу"
    syst_telemekh = "1 комплект" if syst_telemekh else "по запросу"
    izmerit_ustr = "1 комплект" if izmerit_ustr else "по запросу"
    panel_rel_zasch = "1 комплект" if panel_rel_zasch else "по запросу"
    syst_sobstv_nujd = "1 комплект" if syst_sobstv_nujd else "по запросу"
    syst_temp_monit = "1 комплект" if syst_temp_monit else "по запросу"
    obor_antiterror = "1 комплект" if obor_antiterror else "по запросу"
    rezerv_kabelya = "1 комплект" if rezerv_kabelya else "по запросу"
    vch_vols = "1 комплект" if vch_vols else "по запросу"

    context = {
        "pole_code": pole_code,
        "voltage": voltage,
        "wind_region": wind_region,
        "ice_region": ice_region,
        "max_temp": max_temp,
        "min_temp": min_temp,
        "kol_km": kol_km,
        "kol_ap_zaj_km": kol_ap_zaj_km,
        "kol_ap_zaj_opn": kol_ap_zaj_opn,
        "kol_otv_zaj": kol_otv_zaj,
        "kol_opn": kol_opn,
        "kol_kab_krep": kol_kab_krep,
        "konc_kor": konc_kor,
        "kol_konc_kor": kol_konc_kor,
        "ppsa": ppsa,
        "ppsa_length": ppsa_length,
        "kol_skoba_bol": kol_skoba_bol,
        "kol_skoba_mal": kol_skoba_mal,
        "kol_styajka": kol_styajka,
        "razed": razed,
        "syst_telemekh": syst_telemekh,
        "izmerit_ustr": izmerit_ustr,
        "panel_rel_zasch": panel_rel_zasch,
        "syst_sobstv_nujd": syst_sobstv_nujd,
        "syst_temp_monit": syst_temp_monit,
        "obor_antiterror": obor_antiterror,
        "rezerv_kabelya": rezerv_kabelya,
        "vch_vols": vch_vols,
        "pole_heigth": pole_heigth,
        "tr_heigth": tr_heigth,
        "akep": akep,
        "tros_heigth": tros_heigth,
        "bot_diam": bot_diam,
        "top_diam": top_diam,
        "sech_jili": sech_jili,
        "sech_ekr": sech_ekr,
        "sech_al_prov": sech_al_prov,
        "naib_rab_voltage": naib_rab_voltage,
        "naib_dlit_dop_rab_voltage": naib_dlit_dop_rab_voltage,
        "oprosniy_list": oprosniy_list,
        "year": dt.date.today().year,
        "mm_yy": mm_yy,
    }

    dir_name = fd.asksaveasfilename(
                filetypes=[("docx file", ".docx")],
                defaultextension=".docx"
            )
    doc.render(context)
    if dir_name:
        doc.save(dir_name)
