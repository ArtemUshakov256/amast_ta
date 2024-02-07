import tkinter as tk


from PIL import ImageTk
from core.db.db_connector import Database
from core.exceptions import CheckCalculationData
from tkinter.ttk import Combobox
from tkinter import messagebox as mb


from core.utils import (
    # tempFile_back,
    AssemblyAPI,
    make_multiple_path
)
from core.functionality.pasport_pkpo.utils import (
    make_passport
)


class PasportPkpo(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self.title("Паспорт ПКПО")
        self.geometry("650x605+400+5")
        self.resizable(False, False)
        self.config(bg="#FFFFFF")
        self.db = Database()

        # self.back_icon = ImageTk.PhotoImage(
        #     file=tempFile_back
        # )

        self.module_bg = tk.Frame(
            self,
            width=630,
            height=595,
            borderwidth=2,
            relief="sunken"
        )

        self.back_to_main_window_button = tk.Button(
            self,
            # image=self.back_icon,
            text="Назад",
            command=self.back_to_main_window
        )

        self.kol_km_label = tk.Label(
            self,
            text='Концевая муфта, шт',
            width=28,
            anchor="e"
        )
        self.kol_km_entry = tk.Entry(
            self,
            width=10,
            relief="sunken",
            bd=2
        )

        self.osn_komp_pkpo_label = tk.Label(
            self,
            text='Основная комплектация ПКПО:',
            width=28,
            anchor="e",
            font=("default", 10, "bold")
        )

        self.kol_ap_zaj_opn_label = tk.Label(
            self,
            text='Аппаратный зажим ОПН, шт',
            width=28,
            anchor="e"
        )
        self.kol_ap_zaj_opn_entry = tk.Entry(
            self,
            width=10,
            relief="sunken",
            bd=2
        )

        self.kol_ap_zaj_km_label = tk.Label(
            self,
            text='Аппаратный зажим КМ, шт',
            width=28,
            anchor="e"
        )
        self.kol_ap_zaj_km_entry = tk.Entry(
            self,
            width=10,
            relief="sunken",
            bd=2
        )

        self.kol_otv_zaj_label = tk.Label(
            self,
            text='Ответвительный зажим, шт',
            width=28,
            anchor="e"
        )
        self.kol_otv_zaj_entry = tk.Entry(
            self,
            width=10,
            relief="sunken",
            bd=2
        )

        self.kol_opn_label = tk.Label(
            self,
            text='ОПН, шт',
            width=28,
            anchor="e"
        )
        self.kol_opn_entry = tk.Entry(
            self,
            width=10,
            relief="sunken",
            bd=2
        )

        self.kol_kab_krep_label = tk.Label(
            self,
            text='Кабельные крепления, шт',
            width=28,
            anchor="e"
        )
        self.kol_kab_krep_entry = tk.Entry(
            self,
            width=10,
            relief="sunken",
            bd=2
        )

        self.konc_kor_label = tk.Label(
            self,
            text='Марка концевой коробки',
            width=28,
            anchor="e"
        )
        self.konc_kor_entry = tk.Entry(
            self,
            width=10,
            relief="sunken",
            bd=2
        )

        self.kol_konc_kor_label = tk.Label(
            self,
            text='Концевая коробка, шт',
            width=28,
            anchor="e"
        )
        self.kol_konc_kor_entry = tk.Entry(
            self,
            width=10,
            relief="sunken",
            bd=2
        )

        self.ppsa_label = tk.Label(
            self,
            text='Марка ППСа',
            width=28,
            anchor="e"
        )
        self.ppsa_entry = tk.Entry(
            self,
            width=10,
            relief="sunken",
            bd=2
        )

        self.ppsa_length_label = tk.Label(
            self,
            text='Длина ППСа, м',
            width=28,
            anchor="e"
        )
        self.ppsa_length_entry = tk.Entry(
            self,
            width=10,
            relief="sunken",
            bd=2
        )

        self.kol_skoba_bol_label = tk.Label(
            self,
            text='Большая скоба, шт',
            width=28,
            anchor="e"
        )
        self.kol_skoba_bol_entry = tk.Entry(
            self,
            width=10,
            relief="sunken",
            bd=2
        )

        self.kol_skoba_mal_label = tk.Label(
            self,
            text='Малая скоба, шт',
            width=28,
            anchor="e"
        )
        self.kol_skoba_mal_entry = tk.Entry(
            self,
            width=10,
            relief="sunken",
            bd=2
        )

        self.kol_styajka_label = tk.Label(
            self,
            text='Стяжка, шт',
            width=28,
            anchor="e"
        )
        self.kol_styajka_entry = tk.Entry(
            self,
            width=10,
            relief="sunken",
            bd=2
        )

        self.dop_kompl_pkpo_label = tk.Label(
            self,
            text='Дополнительная комплектация ПКПО:',
            width=35,
            anchor="e",
            font=("default", 10, "bold")
        )

        self.is_razed_var = tk.IntVar()
        self.is_razed_checkbutton = tk.Checkbutton(
            self,
            text="Разъединитель",
            variable=self.is_razed_var
        )

        self.is_syst_telemekh_var = tk.IntVar()
        self.is_syst_telemekh_checkbutton = tk.Checkbutton(
            self,
            text="Система телемеханики",
            variable=self.is_syst_telemekh_var
        )

        self.is_izmerit_ustr_var = tk.IntVar()
        self.is_izmerit_ustr_checkbutton = tk.Checkbutton(
            self,
            text="Измерительные устройства",
            variable=self.is_izmerit_ustr_var
        )

        self.is_panel_rel_zasch_var = tk.IntVar()
        self.is_panel_rel_zasch_checkbutton = tk.Checkbutton(
            self,
            text="Панели релейно защиты и автоматики",
            variable=self.is_panel_rel_zasch_var
        )

        self.is_syst_sobstv_nujd_var = tk.IntVar()
        self.is_syst_sobstv_nujd_checkbutton = tk.Checkbutton(
            self,
            text="Система собственных нужд",
            variable=self.is_syst_sobstv_nujd_var
        )

        self.is_syst_temp_monit_var = tk.IntVar()
        self.is_syst_temp_monit_checkbutton = tk.Checkbutton(
            self,
            text="Система температурного мониторинга",
            variable=self.is_syst_temp_monit_var
        )

        self.is_obor_antiterror_var = tk.IntVar()
        self.is_obor_antiterror_checkbutton = tk.Checkbutton(
            self,
            text='Оборудование "Антитеррор"',
            variable=self.is_obor_antiterror_var
        )

        self.is_rezerv_kabelya_var = tk.IntVar()
        self.is_rezerv_kabelya_checkbutton = tk.Checkbutton(
            self,
            text='Резерв кабеля на опоре',
            variable=self.is_rezerv_kabelya_var
        )

        self.is_vch_vols_var = tk.IntVar()
        self.is_vch_vols_checkbutton = tk.Checkbutton(
            self,
            text='ВЧ-оборудование, ВОЛС',
            variable=self.is_vch_vols_var
        )

        self.akep_label = tk.Label(
            self,
            text='Расположение АКЭП (если нет "-"), м',
            width=35,
            anchor="e"
        )
        self.akep_entry = tk.Entry(
            self,
            width=10,
            relief="sunken",
            bd=2
        )

        self.har_pkpo_label = tk.Label(
            self,
            text='Характеристики ПКПО:',
            width=28,
            anchor="e",
            font=("default", 10, "bold")
        )

        self.sech_jili_label = tk.Label(
            self,
            text='Сечение жилы, мм2',
            width=50,
            anchor="e"
        )
        self.sech_jili_entry = tk.Entry(
            self,
            width=10,
            relief="sunken",
            bd=2
        )

        self.sech_ekr_label = tk.Label(
            self,
            text='Сечение экрана, мм2',
            width=50,
            anchor="e"
        )
        self.sech_ekr_entry = tk.Entry(
            self,
            width=10,
            relief="sunken",
            bd=2
        )

        self.sech_al_prov_label = tk.Label(
            self,
            text='Сечение алюминиевой части провода, мм2',
            width=50,
            anchor="e"
        )
        self.sech_al_prov_entry = tk.Entry(
            self,
            width=10,
            relief="sunken",
            bd=2
        )

        self.naib_rab_voltage_label = tk.Label(
            self,
            text='Наибольшее рабочее напряжение, кВ',
            width=50,
            anchor="e"
        )
        self.naib_rab_voltage_entry = tk.Entry(
            self,
            width=10,
            relief="sunken",
            bd=2
        )

        self.naib_dlit_dop_rab_voltage_label = tk.Label(
            self,
            text='Наибольшее длительно допустимое рабочее напряжение, кВ',
            width=50,
            anchor="e"
        )
        self.naib_dlit_dop_rab_voltage_entry = tk.Entry(
            self,
            width=10,
            relief="sunken",
            bd=2
        )

        self.media_label = tk.Label(
            self,
            text='Ссылки на опросный лист:',
            width=28,
            anchor="e",
            font=("default", 10, "bold")
        )

        self.path_to_oprosniy_list_label = tk.Label(
            self,
            text='Опросный лист ".png"',
            anchor="e",
            width=21
        )
        self.path_to_oprosniy_list_entry = tk.Entry(
            self,
            width=30,
            justify="left",
            relief="sunken",
            bd=2
        )
        self.browse_oprosniy_list_button = tk.Button(
            self,
            text="Обзор",
            command=self.go_to_make_multiple_path
        )

        self.make_pasport_button = tk.Button(
            self,
            text="Создать паспорт",
            command=self.go_to_make_passport
        )

    def run(self):
        self.draw_widgets()
        self.paste_passport_pkpo_data()
        self.mainloop()

    def paste_passport_pkpo_data(self):
        try:
            self.kol_km_entry.delete(0, "end")
            self.kol_km_entry.insert(0, self.parent.kol_km)
            self.kol_ap_zaj_opn_entry.delete(0, "end")
            self.kol_ap_zaj_opn_entry.insert(0, self.parent.kol_ap_zaj_opn)
            self.kol_ap_zaj_km_entry.delete(0, "end")
            self.kol_ap_zaj_km_entry.insert(0, self.parent.kol_ap_zaj_km)
            self.kol_otv_zaj_entry.delete(0, "end")
            self.kol_otv_zaj_entry.insert(0, self.parent.kol_otv_zaj)
            self.kol_opn_entry.delete(0, "end")
            self.kol_opn_entry.insert(0, self.parent.kol_opn)
            self.kol_kab_krep_entry.delete(0, "end")
            self.kol_kab_krep_entry.insert(0, self.parent.kol_kab_krep)
            self.konc_kor_entry.delete(0, "end")
            self.konc_kor_entry.insert(0, self.parent.konc_kor)
            self.kol_konc_kor_entry.delete(0, "end")
            self.kol_konc_kor_entry.insert(0, self.parent.kol_konc_kor)
            self.ppsa_entry.delete(0, "end")
            self.ppsa_entry.insert(0, self.parent.ppsa)
            self.ppsa_length_entry.delete(0, "end")
            self.ppsa_length_entry.insert(0, self.parent.ppsa_length)
            self.kol_skoba_bol_entry.delete(0, "end")
            self.kol_skoba_bol_entry.insert(0, self.parent.kol_skoba_bol)
            self.kol_skoba_mal_entry.delete(0, "end")
            self.kol_skoba_mal_entry.insert(0, self.parent.kol_skoba_mal)
            self.kol_styajka_entry.delete(0, "end")
            self.kol_styajka_entry.insert(0, self.parent.kol_styajka)
            # self.is_razed_var = self.parent.razed
            if self.parent.razed:
                self.is_razed_checkbutton.select()
            # self.is_syst_telemekh_var = self.parent.syst_telemekh
            if self.parent.syst_telemekh:
                self.is_syst_telemekh_checkbutton.select()
            # self.is_izmerit_ustr_var = self.parent.izmerit_ustr
            if self.parent.izmerit_ustr:
                self.is_izmerit_ustr_checkbutton.select()
            # self.is_panel_rel_zasch_var = self.parent.panel_rel_zasch
            if self.parent.panel_rel_zasch:
                self.is_panel_rel_zasch_checkbutton.select()
            # self.is_syst_sobstv_nujd_var = self.parent.syst_sobstv_nujd
            if self.parent.syst_sobstv_nujd:
                self.is_syst_sobstv_nujd_checkbutton.select()
            # self.is_syst_temp_monit_var = self.parent.syst_temp_monit
            if self.parent.syst_temp_monit:
                self.is_syst_temp_monit_checkbutton.select()
            # self.is_obor_antiterror_var = self.parent.obor_antiterror
            if self.parent.obor_antiterror:
                self.is_obor_antiterror_checkbutton.select()
            # self.is_rezerv_kabelya_var = self.parent.rezerv_kabelya
            if self.parent.rezerv_kabelya:
                self.is_rezerv_kabelya_checkbutton.select()
            # self.is_vch_vols_var = self.parent.vch_vols
            if self.parent.vch_vols:
                self.is_vch_vols_checkbutton.select()
            self.akep_entry.delete(0, "end")
            self.akep_entry.insert(0, self.parent.akep)
            self.sech_jili_entry.delete(0, "end")
            self.sech_jili_entry.insert(0, self.parent.sech_jili)  
            self.sech_ekr_entry.delete(0, "end")
            self.sech_ekr_entry.insert(0, self.parent.sech_ekr)  
            self.sech_al_prov_entry.delete(0, "end")
            self.sech_al_prov_entry.insert(0, self.parent.sech_al_prov)  
            self.naib_rab_voltage_entry.delete(0, "end")
            self.naib_rab_voltage_entry.insert(0, self.parent.naib_rab_voltage)  
            self.naib_dlit_dop_rab_voltage_entry.delete(0, "end")
            self.naib_dlit_dop_rab_voltage_entry.insert(0, self.parent.naib_dlit_dop_rab_voltage)
            opr_list = ["{" + pic + "}" for pic in self.parent.oprosniy_list]
            self.path_to_oprosniy_list_entry.delete(0, "end")
            self.path_to_oprosniy_list_entry.insert(0, " ".join(opr_list))            

        except Exception as e:
            print("!INFO!: Сохраненные данные отсутствуют.")
    
    def draw_widgets(self):
        self.module_bg.place(x=10, y=0)
        self.back_to_main_window_button.place(x=15, y=3)
        self.osn_komp_pkpo_label.place(x=50, y=29)
        self.kol_km_label.place(x=15, y=52)
        self.kol_km_entry.place(x=220, y=52)
        self.kol_ap_zaj_km_label.place(x=15, y=75)
        self.kol_ap_zaj_km_entry.place(x=220, y=75)
        self.kol_ap_zaj_opn_label.place(x=15, y=98)
        self.kol_ap_zaj_opn_entry.place(x=220, y=98)
        self.kol_otv_zaj_label.place(x=15, y=121)
        self.kol_otv_zaj_entry.place(x=220, y=121)
        self.kol_opn_label.place(x=15, y=144)
        self.kol_opn_entry.place(x=220, y=144)
        self.kol_kab_krep_label.place(x=15, y=167)
        self.kol_kab_krep_entry.place(x=220, y=167)
        self.konc_kor_label.place(x=15, y=190)
        self.konc_kor_entry.place(x=220, y=190)
        self.kol_konc_kor_label.place(x=15, y=213)
        self.kol_konc_kor_entry.place(x=220, y=213)
        self.ppsa_label.place(x=15, y=236)
        self.ppsa_entry.place(x=220, y=236)
        self.ppsa_length_label.place(x=15, y=259)
        self.ppsa_length_entry.place(x=220, y=259)
        self.kol_skoba_bol_label.place(x=15,y=282)
        self.kol_skoba_bol_entry.place(x=220, y=282)
        self.kol_skoba_mal_label.place(x=15, y=305)
        self.kol_skoba_mal_entry.place(x=220, y=305)
        self.kol_styajka_label.place(x=15, y=328)
        self.kol_styajka_entry.place(x=220, y=328)
        self.dop_kompl_pkpo_label.place(x=305, y=29)
        self.is_razed_checkbutton.place(x=330, y=52)
        self.is_syst_telemekh_checkbutton.place(x=330, y=75)
        self.is_izmerit_ustr_checkbutton.place(x=330, y=98)
        self.is_panel_rel_zasch_checkbutton.place(x=330, y=121)
        self.is_syst_sobstv_nujd_checkbutton.place(x=330, y=144)
        self.is_syst_temp_monit_checkbutton.place(x=330, y=167)
        self.is_obor_antiterror_checkbutton.place(x=330, y=190)
        self.is_rezerv_kabelya_checkbutton.place(x=330, y=213)
        self.is_vch_vols_checkbutton.place(x=330, y=236)
        self.akep_label.place(x=295, y=259)
        self.akep_entry.place(x=550, y=259)
        self.har_pkpo_label.place(x=185, y=360)
        self.sech_jili_label.place(x=50, y=383)
        self.sech_jili_entry.place(x=410, y=383)
        self.sech_ekr_label.place(x=50, y=406)
        self.sech_ekr_entry.place(x=410, y=406)
        self.sech_al_prov_label.place(x=50, y=429)
        self.sech_al_prov_entry.place(x=410, y=429)
        self.naib_rab_voltage_label.place(x=50, y=452)
        self.naib_rab_voltage_entry.place(x=410, y=452)
        self.naib_dlit_dop_rab_voltage_label.place(x=50, y=475)
        self.naib_dlit_dop_rab_voltage_entry.place(x=410, y=475)
        self.media_label.place(x=200, y=510)
        self.path_to_oprosniy_list_label.place(x=120, y=535)
        self.path_to_oprosniy_list_entry.place(x=272, y=535)
        self.browse_oprosniy_list_button.place(x=459, y=533)
        self.make_pasport_button.place(x=290, y=560)
    
    def go_to_make_passport(self):
        make_passport(
            pole_code=self.parent.pole_code,
            voltage=self.parent.voltage,
            wind_region=self.parent.wind_region,
            ice_region=self.parent.ice_region,
            max_temp=self.parent.max_temp,
            min_temp=self.parent.min_temp,
            kol_km=self.kol_km_entry.get(),
            kol_ap_zaj_km=self.kol_ap_zaj_km_entry.get(),
            kol_ap_zaj_opn=self.kol_ap_zaj_opn_entry.get(),
            kol_otv_zaj=self.kol_otv_zaj_entry.get(),
            kol_opn=self.kol_opn_entry.get(),
            kol_kab_krep=self.kol_kab_krep_entry.get(),
            konc_kor=self.konc_kor_entry.get(),
            kol_konc_kor=self.kol_konc_kor_entry.get(),
            ppsa=self.ppsa_entry.get(),
            ppsa_length=self.ppsa_length_entry.get(),
            kol_skoba_bol=self.kol_skoba_bol_entry.get(),
            kol_skoba_mal=self.kol_skoba_mal_entry.get(),
            kol_styajka=self.kol_styajka_entry.get(),
            razed=self.is_razed_var.get(),
            syst_telemekh=self.is_syst_telemekh_var.get(),
            izmerit_ustr=self.is_izmerit_ustr_var.get(),
            panel_rel_zasch=self.is_panel_rel_zasch_var.get(),
            syst_sobstv_nujd=self.is_syst_sobstv_nujd_var.get(),
            syst_temp_monit=self.is_syst_temp_monit_var.get(),
            obor_antiterror=self.is_obor_antiterror_var.get(),
            rezerv_kabelya=self.is_rezerv_kabelya_var.get(),
            vch_vols=self.is_vch_vols_var.get(),
            pole_heigth=self.parent.pls_pole_data["pole_heigth"],
            tr_heigth=self.parent.pls_pole_data["davit_heigth"],
            akep=self.akep_entry.get(),
            tros_heigth=self.parent.pls_pole_data["tros_heigth"],
            bot_diam=self.parent.pls_pole_data["bot_diam"],
            top_diam=self.parent.pls_pole_data["top_diameter"],
            sech_jili=self.sech_jili_entry.get(),
            sech_ekr=self.sech_ekr_entry.get(),
            sech_al_prov=self.sech_al_prov_entry.get(),
            naib_rab_voltage=self.naib_rab_voltage_entry.get(),
            naib_dlit_dop_rab_voltage=self.naib_dlit_dop_rab_voltage_entry.get(),
            oprosniy_list=self.path_to_oprosniy_list_entry.get()
        )
        opros_list = self.path_to_oprosniy_list_entry.get()
        opros_list = [pic_dir.strip("}{") for pic_dir in opros_list.split("} {")]
        for i in range(3):
            if len(opros_list) < 3:
                opros_list.append("")
            if opros_list[i]:
                opros_list[i] = opros_list[i].split("Удаленка")[1]
        oprosniy = ":".join(opros_list)
        self.db.add_passport_pkpo_data(
            initial_data_id=self.parent.initial_data_id,
            kol_km=self.kol_km_entry.get(),
            kol_ap_zaj_km=self.kol_ap_zaj_km_entry.get(),
            kol_ap_zaj_opn=self.kol_ap_zaj_opn_entry.get(),
            kol_otv_zaj=self.kol_otv_zaj_entry.get(),
            kol_opn=self.kol_opn_entry.get(),
            kol_kab_krep=self.kol_kab_krep_entry.get(),
            konc_kor=self.konc_kor_entry.get(),
            kol_konc_kor=self.kol_konc_kor_entry.get(),
            ppsa=self.ppsa_entry.get(),
            ppsa_length=self.ppsa_length_entry.get(),
            kol_skoba_bol=self.kol_skoba_bol_entry.get(),
            kol_skoba_mal=self.kol_skoba_mal_entry.get(),
            kol_styajka=self.kol_styajka_entry.get(),
            razed=self.is_razed_var.get(),
            syst_telemekh=self.is_syst_telemekh_var.get(),
            izmerit_ustr=self.is_izmerit_ustr_var.get(),
            panel_rel_zasch=self.is_panel_rel_zasch_var.get(),
            syst_sobstv_nujd=self.is_syst_sobstv_nujd_var.get(),
            syst_temp_monit=self.is_syst_temp_monit_var.get(),
            obor_antiterror=self.is_obor_antiterror_var.get(),
            rezerv_kabelya=self.is_rezerv_kabelya_var.get(),
            vch_vols=self.is_vch_vols_var.get(),
            akep=self.akep_entry.get(),
            sech_jili=self.sech_jili_entry.get(),
            sech_ekr=self.sech_ekr_entry.get(),
            sech_al_prov=self.sech_al_prov_entry.get(),
            naib_rab_voltage=self.naib_rab_voltage_entry.get(),
            naib_dlit_dop_rab_voltage=self.naib_dlit_dop_rab_voltage_entry.get(),
            oprosniy_list=oprosniy
        )

    def go_to_make_multiple_path(self):
        self.file_path = make_multiple_path()
        self.path_to_oprosniy_list_entry.delete("0", "end") 
        self.path_to_oprosniy_list_entry.insert("insert", self.file_path)

    def back_to_main_window(self):
        self.destroy()
        self.parent.deiconify()

    def validate_sloy_quantity(self, value):
        if value in ["1", "2", "3", "4", "5"]:
            return True
        return False
    
    def validate_sloy_type(self, value):
        if value in ["Песок", "Супесь", "Суглинок", "Глина"]:
            return True
        return False


if __name__ == "__main__":
    main_page = PasportPkpo()
    main_page.run()