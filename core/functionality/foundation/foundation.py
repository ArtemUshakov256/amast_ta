import os
import re
import tkinter as tk


from PIL import Image, ImageTk
from tkinter.ttk import Combobox, Style, Checkbutton


from core.utils import (
    tempFile_back,
    tempFile_open,
    tempFile_save,
    tempFile_plus,
    tempFile_minus,
    make_path_txt,
    make_path_png,
    make_multiple_path
)
from core.functionality.foundation.utils import *


class FoundationCalculation(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self.title("Расчет фундамента")
        self.geometry("730x704+400+10")
        self.resizable(False, False)
        self.config(bg="#FFFFFF")

        self.back_icon = ImageTk.PhotoImage(
            file=tempFile_back
        )
        self.open_icon = ImageTk.PhotoImage(
            file=tempFile_open
        )
        self.save_icon = ImageTk.PhotoImage(
            file=tempFile_save
        )
        self.plus_icon = ImageTk.PhotoImage(
            file=tempFile_plus
        )

        self.minus_icon = ImageTk.PhotoImage(
            file=tempFile_minus
        )

        self.module_bg = tk.Frame(
            self,
            width=710,
            height=693,
            borderwidth=2,
            relief="sunken"
        )

        self.back_to_main_window_button = tk.Button(
            self,
            image=self.back_icon,
            command=self.destroy
        )

        self.open_button = tk.Button(
            self,
            image=self.open_icon,
            command=self.open_data
        )
        
        self.save_button = tk.Button(
            self,
            image=self.save_icon,
            command=self.save_data
        )

        self.diam_svai_label = tk.Label(
            self,
            text='Диаметр сваи "под ключ", мм',
            width=28,
            anchor="e"
        )
        self.diam_svai_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2
        )

        self.thickness_svai_label = tk.Label(
            self,
            text='Толщина сваи, мм',
            width=28,
            anchor="e"
        )
        self.thickness_svai_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2
        )

        self.deepness_svai_label = tk.Label(
            self,
            text='Глубина заложения сваи, мм',
            width=28,
            anchor="e"
        )
        self.deepness_svai_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2
        )

        self.height_svai_label = tk.Label(
            self,
            text='Высота головы сваи, м',
            width=28,
            anchor="e"
        )
        self.height_svai_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2
        )

        self.is_initial_data_var = tk.IntVar()
        self.is_initial_data_checkbutton = tk.Checkbutton(
            self,
            text="Есть исходные данные грунтов",
            variable=self.is_initial_data_var,
            command=self.toggle_state
        )

        self.typical_ground_label = tk.Label(
            self,
            text='Выбранный "типовой грунт"',
            width=28,
            anchor="e"
        )
        self.typical_ground_variable = tk.StringVar()
        self.typical_ground_combobox = Combobox(
            self,
            width=26,
            values=(
                "Песок Крупный (e = 0.45)",
                "Песок Крупный (e = 0.65)",
                "Песок Мелкий (e = 0.45)",
                "Песок Мелкий (e = 0.75)",
                "Песок Пылеватый (e = 0.45)",
                "Песок Пылеватый (e = 0.75)",
                "Супесь (e = 0.45)",
                "Супесь (e = 0.65)",
                "Супесь (e = 0.85)",
                "Суглинок (e = 0.45)",
                "Суглинок (e = 0.65)",
                "Суглинок (e = 0.85)",
                "Суглинок (e = 1.05)",
                "Глина (e = 0.55)",
                "Глина (e = 0.75)",
                "Глина (e = 0.95)",
                "Глина (e = 1.05)",
            ),
            textvariable=self.typical_ground_variable
        )
        self.typical_ground_combobox.bind("<<ComboboxSelected>>", self.paste_typical_ground_data)

        self.udel_sceplenie_label = tk.Label(
            self,
            text='Удельное сцепление, С, кПа',
            width=28,
            anchor="e"
        )
        self.udel_sceplenie_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2
        )

        self.ugol_vntr_trenia_label = tk.Label(
            self,
            text='Угол внутреннего трения ф, град',
            width=28,
            anchor="e"
        )
        self.ugol_vntr_trenia_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2
        )

        self.ves_grunta_label = tk.Label(
            self,
            text='Вес грунта, т/м3',
            width=28,
            anchor="e"
        )
        self.ves_grunta_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2
        )

        self.deform_module_label = tk.Label(
            self,
            text='Модуль деформации Е, кПа',
            width=28,
            anchor="e"
        )
        self.deform_module_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2
        )

        self.ground_water_lvl_label = tk.Label(
            self,
            text='Уровень грунтовых вод',
            width=28,
            anchor="e"
        )
        self.ground_water_lvl_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2
        )

        # self.plus_button = tk.Button(
        #     self,
        #     image=self.plus_icon,
        #     command=self.add_sloy
        # )
        # self.minus_button = tk.Button(
        #     self,
        #     image=self.minus_icon,
        #     command=self.del_sloy
        # )

        self.quantity_of_sloy_label = tk.Label(
            self,
            text='Количество слоев',
            width=20,
            anchor="e"
        )
        self.quantity_of_sloy_combobox = Combobox(
            self,
            width=12,
            values=(
                "1",
                "2",
                "3",
                "4",
                "5"
            )
        )

        self.sloy_nomer1_label = tk.Label(
            self,
            text='Слой №1',
            width=7,
            anchor="e"
        )

        self.sloy_nomer2_label = tk.Label(
            self,
            text='Слой №2',
            width=7,
            anchor="e"
        )

        self.sloy_nomer3_label = tk.Label(
            self,
            text='Слой №4',
            width=7,
            anchor="e"
        )

        self.sloy_nomer4_label = tk.Label(
            self,
            text='Слой №4',
            width=7,
            anchor="e"
        )

        self.sloy_nomer5_label = tk.Label(
            self,
            text='Слой №5',
            width=7,
            anchor="e"
        )
        
        self.nomer_ige_label = tk.Label(
            self,
            text='Номер ИГЭ',
            width=20,
            anchor="e"
        )
        self.nomer_ige1_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2
        )
        self.nomer_ige2_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.nomer_ige3_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.nomer_ige4_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.nomer_ige5_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            state="disabled"
        )

        self.ground_type_label = tk.Label(
            self,
            text='Тип грунта',
            width=20,
            anchor="e"
        )
        self.ground_type1_combobox = Combobox(
            self,
            width=12,
            values=(
                "Песок",
                "Супесь",
                "Суглинок",
                "Глина"
            )
        )
        self.ground_type2_combobox = Combobox(
            self,
            width=12,
            values=(
                "Песок",
                "Супесь",
                "Суглинок",
                "Глина"
            ),
            state="disabled"
        )
        self.ground_type3_combobox = Combobox(
            self,
            width=12,
            values=(
                "Песок",
                "Супесь",
                "Суглинок",
                "Глина"
            ),
            state="disabled"
        )
        self.ground_type4_combobox = Combobox(
            self,
            width=12,
            values=(
                "Песок",
                "Супесь",
                "Суглинок",
                "Глина"
            ),
            state="disabled"
        )
        self.ground_type5_combobox = Combobox(
            self,
            width=12,
            values=(
                "Песок",
                "Супесь",
                "Суглинок",
                "Глина"
            ),
            state="disabled"
        )

        self.ground_name_label = tk.Label(
            self,
            text='Наименование грунта',
            width=20,
            anchor="e"
        )
        self.ground_name1_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2
        )
        self.ground_name2_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.ground_name3_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.ground_name4_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.ground_name5_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            state="disabled"
        )

        self.verh_sloy_label = tk.Label(
            self,
            text='Отметка верха слоя, м',
            width=20,
            anchor="e"
        )
        self.verh_sloy1_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2
        )
        self.verh_sloy2_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.verh_sloy3_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.verh_sloy4_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.verh_sloy5_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            state="disabled"
        )

        self.nijn_sloy_label = tk.Label(
            self,
            text='Отметка низа слоя, м',
            width=20,
            anchor="e"
        )
        self.nijn_sloy1_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2
        )
        self.nijn_sloy2_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.nijn_sloy3_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.nijn_sloy4_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.nijn_sloy5_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            state="disabled"
        )

        self.mochn_sloya_label = tk.Label(
            self,
            text='Мощность слоя, м',
            width=20,
            anchor="e"
        )
        self.mochn_sloya1_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2
        )
        self.mochn_sloya2_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.mochn_sloya3_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.mochn_sloya4_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.mochn_sloya5_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            state="disabled"
        )

        self.tech_param_label = tk.Label(
            self,
            text='Технические характеристики:',
            width=20,
            anchor="e"
        )

        self.coef_poristosti_label = tk.Label(
            self,
            text='Коэффициент пористости, e',
            width=20,
            anchor="e"
        )
        self.coef_poristosti1_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2
        )
        self.coef_poristosti2_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.coef_poristosti3_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.coef_poristosti4_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.coef_poristosti5_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            state="disabled"
        )

        self.udel_scep_label = tk.Label(
            self,
            text='Удельное сцепление cI, кПа',
            width=20,
            anchor="e"
        )
        self.udel_scep1_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2
        )
        self.udel_scep2_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.udel_scep3_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.udel_scep4_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.udel_scep5_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            state="disabled"
        )

        self.ugol_vn_tr_label = tk.Label(
            self,
            text='Угол внутреннего трения фI, град',
            width=20,
            anchor="e"
        )
        self.ugol_vn_tr1_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2
        )
        self.ugol_vn_tr2_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.ugol_vn_tr3_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.ugol_vn_tr4_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.ugol_vn_tr5_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            state="disabled"
        )

        self.ves_gr_prir_label = tk.Label(
            self,
            text='Вес грунта природный, т/м3',
            width=20,
            anchor="e"
        )
        self.ves_gr_prir1_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2
        )
        self.ves_gr_prir2_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.ves_gr_prir3_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.ves_gr_prir4_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.ves_gr_prir5_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            state="disabled"
        )

        self.ves_gr_ras_label = tk.Label(
            self,
            text='Вес грунта расчетный, т/м3',
            width=20,
            anchor="e"
        )
        self.ves_gr_ras1_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2
        )
        self.ves_gr_ras2_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.ves_gr_ras3_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.ves_gr_ras4_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.ves_gr_ras5_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            state="disabled"
        )

        self.def_mod_label = tk.Label(
            self,
            text='Модуль деформации E, кПа',
            width=20,
            anchor="e"
        )
        self.def_mod1_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2
        )
        self.def_mod2_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.def_mod3_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.def_mod4_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.def_mod5_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            state="disabled"
        )

    def run(self):
        self.draw_widgets()
        self.mainloop()

    def draw_widgets(self):
        self.module_bg.place(x=10, y=0)

        self.back_to_main_window_button.place(x=15, y=2)
        self.open_button.place(x=41, y=2)
        self.save_button.place(x=67, y=2)

        self.diam_svai_label.place(x=15, y=29)
        self.diam_svai_entry.place(x=220, y=29)

        self.thickness_svai_label.place(x=15, y=52)
        self.thickness_svai_entry.place(x=220, y=52)

        self.deepness_svai_label.place(x=15, y=75)
        self.deepness_svai_entry.place(x=220, y=75)

        self.height_svai_label.place(x=15, y=98)
        self.height_svai_entry.place(x=220, y=98)

        self.is_initial_data_checkbutton.place(x=60, y=119)

        self.typical_ground_label.place(x=15, y=144)
        self.typical_ground_combobox.place(x=220, y=144)

        self.udel_sceplenie_label.place(x=15, y=167)
        self.udel_sceplenie_entry.place(x=220, y=167)

        self.ugol_vntr_trenia_label.place(x=15, y=190)
        self.ugol_vntr_trenia_entry.place(x=220, y=190)

        self.ves_grunta_label.place(x=15, y=213)
        self.ves_grunta_entry.place(x=220, y=213)

        self.deform_module_label.place(x=15, y=235)
        self.deform_module_entry.place(x=220, y=235)

        self.ground_water_lvl_label.place(x=15, y=258)
        self.ground_water_lvl_entry.place(x=220, y=258)

        # self.plus_button.place(x=215, y=288)
        # self.minus_button.place(x=184, y=288)

        self.quantity_of_sloy_label.place(x=72,y=290)
        self.quantity_of_sloy_combobox.place(x=220, y=290)

        self.sloy_nomer1_label.place(x=180, y=318)
        self.sloy_nomer2_label.place(x=290, y=318)
        self.sloy_nomer3_label.place(x=400, y=318)
        self.sloy_nomer4_label.place(x=510, y=318)
        self.sloy_nomer5_label.place(x=620, y=318)

        self.nomer_ige_label.place(x=15, y=341)
        self.nomer_ige1_entry.place(x=164, y=341)
        self.nomer_ige2_entry.place(x=274, y=341)
        self.nomer_ige3_entry.place(x=384, y=341)
        self.nomer_ige4_entry.place(x=494, y=341)
        self.nomer_ige5_entry.place(x=604, y=341)

        self.ground_type_label.place(x=15, y=364)
        self.ground_type1_combobox.place(x=164, y=364)
        self.ground_type2_combobox.place(x=274, y=364)
        self.ground_type3_combobox.place(x=384, y=364)
        self.ground_type4_combobox.place(x=494, y=364)
        self.ground_type5_combobox.place(x=604, y=364)

        self.ground_name_label.place(x=15, y=387)
        self.ground_name1_entry.place(x=164, y=387)
        self.ground_name2_entry.place(x=274, y=387)
        self.ground_name3_entry.place(x=384, y=387)
        self.ground_name4_entry.place(x=494, y=387)
        self.ground_name5_entry.place(x=604, y=387)

        self.verh_sloy_label.place(x=15, y=410)
        self.verh_sloy1_entry.place(x=164, y=410)
        self.verh_sloy2_entry.place(x=274, y=410)
        self.verh_sloy3_entry.place(x=384, y=410)
        self.verh_sloy4_entry.place(x=494, y=410)
        self.verh_sloy5_entry.place(x=604, y=410)

        self.nijn_sloy_label.place(x=15, y=433)
        self.nijn_sloy1_entry.place(x=164, y=433)
        self.nijn_sloy2_entry.place(x=274, y=433)
        self.nijn_sloy3_entry.place(x=384, y=433)
        self.nijn_sloy4_entry.place(x=494, y=433)
        self.nijn_sloy5_entry.place(x=604, y=433)

        self.mochn_sloya_label.place(x=15, y=456)
        self.mochn_sloya1_entry.place(x=164, y=456)
        self.mochn_sloya2_entry.place(x=274, y=456)
        self.mochn_sloya3_entry.place(x=384, y=456)
        self.mochn_sloya4_entry.place(x=494, y=456)
        self.mochn_sloya5_entry.place(x=604, y=456)

    def paste_typical_ground_data(self, event):
        typical_ground_dict = {
            "Песок Крупный (e = 0.45)": {
                "ф": "43",
                "C": "2",
                "y": "2",
                "E": "50000"
            },
            "Песок Крупный (e = 0.65)": {
                "ф": "35",
                "C": "1",
                "y": "1,9",
                "E": "30000"
            },
            "Песок Мелкий (e = 0.45)": {
                "ф": "38",
                "C": "6",
                "y": "1,9",
                "E": "48000"
            },
            "Песок Мелкий (e = 0.75)": {
                "ф": "28",
                "C": "0",
                "y": "1,8",
                "E": "18000"
            },
            "Песок Пылеватый (e = 0.45)": {
                "ф": "36",
                "C": "8",
                "y": "1,8",
                "E": "39000"
            },
            "Песок Пылеватый (e = 0.75)": {
                "ф": "26",
                "C": "2",
                "y": "1,8",
                "E": "11000"
            },
            "Супесь (e = 0.45)": {
                "ф": "30",
                "C": "21",
                "y": "2",
                "E": "32000"
            },
            "Супесь (e = 0.65)": {
                "ф": "27",
                "C": "15",
                "y": "1,9",
                "E": "16000"
            },
            "Супесь (e = 0.85)": {
                "ф": "18",
                "C": "9",
                "y": "1,7",
                "E": "9000"
            },
            "Суглинок (e = 0.45)": {
                "ф": "26",
                "C": "47",
                "y": "2",
                "E": "34000"
            },
            "Суглинок (e = 0.65)": {
                "ф": "24",
                "C": "31",
                "y": "1,9",
                "E": "22000"
            },
            "Суглинок (e = 0.85)": {
                "ф": "19",
                "C": "18",
                "y": "1,8",
                "E": "11000"
            },
            "Суглинок (e = 1.05)": {
                "ф": "12",
                "C": "12",
                "y": "1,8",
                "E": "5000"
            },
            "Глина (e = 0.55)": {
                "ф": "21",
                "C": "81",
                "y": "1,8",
                "E": "28000"
            },
            "Глина (e = 0.75)": {
                "ф": "19",
                "C": "54",
                "y": "1,8",
                "E": "21000"
            },
            "Глина (e = 0.95)": {
                "ф": "14",
                "C": "37",
                "y": "1,7",
                "E": "12000"
            },
            "Глина (e = 1.05)": {
                "ф": "7",
                "C": "29",
                "y": "1,7",
                "E": "7000"
            }
        }
        typical_ground_key = self.typical_ground_combobox.get()
        self.udel_sceplenie_entry.delete(0, tk.END)
        self.udel_sceplenie_entry.insert(0, typical_ground_dict[typical_ground_key]["C"])
        self.ugol_vntr_trenia_entry.delete(0, tk.END)
        self.ugol_vntr_trenia_entry.insert(0, typical_ground_dict[typical_ground_key]["ф"])
        self.ves_grunta_entry.delete(0, tk.END)
        self.ves_grunta_entry.insert(0, typical_ground_dict[typical_ground_key]["y"])
        self.deform_module_entry.delete(0, tk.END)
        self.deform_module_entry.insert(0, typical_ground_dict[typical_ground_key]["E"])
    
    def toggle_state(self):
        if self.is_initial_data_var.get() == 0:
            self.typical_ground_combobox.config(state="normal")
            self.udel_sceplenie_entry.config(state="normal")
            self.ugol_vntr_trenia_entry.config(state="normal")
            self.ves_grunta_entry.config(state="normal")
            self.deform_module_entry.config(state="normal")
            self.ground_water_lvl_entry.config(state="disabled")
        else:
            self.typical_ground_combobox.delete(0, tk.END)
            self.typical_ground_combobox.config(state="disabled")
            self.udel_sceplenie_entry.delete(0, tk.END)
            self.udel_sceplenie_entry.config(state="disabled")
            self.ugol_vntr_trenia_entry.delete(0, tk.END)
            self.ugol_vntr_trenia_entry.config(state="disabled")
            self.ves_grunta_entry.delete(0, tk.END)
            self.ves_grunta_entry.config(state="disabled")
            self.deform_module_entry.delete(0, tk.END)
            self.deform_module_entry.config(state="disabled")
            self.quantity_of_sloy_combobox.delete(0, tk.END)
            self.quantity_of_sloy_combobox.config(state="disabled")
            self.nomer_ige1_entry.delete(0, tk.END)
            self.nomer_ige2_entry.delete(0, tk.END)
            self.nomer_ige3_entry.delete(0, tk.END)
            self.nomer_ige4_entry.delete(0, tk.END)
            self.nomer_ige5_entry.delete(0, tk.END)
            self.nomer_ige1_entry.config(state="disabled")
            self.nomer_ige2_entry.config(state="disabled")
            self.nomer_ige3_entry.config(state="disabled")
            self.nomer_ige4_entry.config(state="disabled")
            self.nomer_ige5_entry.config(state="disabled")

    def add_sloy(self):
        self.layer_number = 1
        self.layer_number += 1
        self.sloy_nomer_

    def del_sloy(self):
        pass

    # def browse_for_pole(self):
    #     self.file_path = make_path_png()
    #     self.pole_entry.delete("0", "end")
    #     self.pole_entry.insert("insert", self.file_path)

    # def browse_for_pole_defl(self):
    #     self.file_path = make_path_png()
    #     self.pole_defl_entry.delete("0", "end")
    #     self.pole_defl_entry.insert("insert", self.file_path)

    # def browse_for_loads(self):
    #     self.file_path = make_multiple_path()
    #     self.loads_entry.delete("0", "end") 
    #     self.loads_entry.insert("insert", self.file_path)

    # def browse_for_mont_schema(self):
    #     self.file_path = make_path_png()
    #     self.is_mont_schema_entry.delete("0", "end") 
    #     self.is_mont_schema_entry.insert("insert", self.file_path)
        
    # def browse_for_txt_1(self):
    #     self.file_path = make_path_txt()
    #     self.path_to_txt_1_entry.delete("0", "end") 
    #     self.path_to_txt_1_entry.insert("insert", self.file_path)

    # def browse_for_txt_2(self):
    #     self.file_path = make_path_txt()
    #     self.path_to_txt_2_entry.delete("0", "end") 
    #     self.path_to_txt_2_entry.insert("insert", self.file_path)

    def save_data(self):
        filename = fd.asksaveasfilename(
        defaultextension=".txt",
        filetypes=[("Text Files", "*.txt")]
        )
        if filename:
            with open(filename, "w") as file:
                file.writelines(
                    [
                    ]
                )
            
    def open_data(self):
        filename = fd.askopenfilename(filetypes=[("Text Files", "*.txt")])
        if filename:
            with open(filename, "r") as file:
                pass 

    # def generate_output(self):
    #     if re.match(r"\w+\s\d+/\d+", self.wire_entry.get()):
    #         self.result = put_data(
    #             project_name=self.project_name_entry.get(),
    #             project_code=self.project_code_entry.get(),
    #             pole_code=self.pole_code_entry.get(),
    #             pole_type=self.pole_type_combobox.get(),
    #             developer=self.developer_combobox.get(),
    #             voltage=self.voltage_combobox.get(),
    #             area=self.area_combobox.get(),
    #             branches=self.branches_combobox.get(),
    #             wind_region=self.wind_region_combobox.get(),
    #             wind_pressure=self.wind_pressure_entry.get(),
    #             ice_region=self.ice_region_combobox.get(),
    #             ice_thickness=self.ice_thickness_entry.get(),
    #             ice_wind_pressure=self.ice_wind_pressure_entry.get(),
    #             year_average_temp=self.year_average_temp_entry.get(),
    #             min_temp=self.min_temp_entry.get(),
    #             max_temp=self.max_temp_entry.get(),
    #             ice_temp=self.ice_temp_entry.get(),
    #             wind_temp=self.wind_temp_entry.get(),
    #             wind_reg_coef=self.wind_reg_coef_entry.get(),
    #             ice_reg_coef=self.ice_reg_coef_entry.get(),
    #             wire_hesitation=self.wire_hesitation_combobox.get(),
    #             wire=self.wire_entry.get(),
    #             wire_tencion=self.wire_tencion_entry.get(),
    #             ground_wire=self.ground_wire_entry.get(),
    #             oksn=self.oksn_entry.get(),
    #             wind_span=self.wind_span_entry.get(),
    #             weight_span=self.weight_span_entry.get(),
    #             is_stand=self.is_stand_combobox.get(),
    #             is_plate=self.is_plate_combobox.get(),
    #             is_ground_wire_davit=self.is_ground_wire_davit_combobox.get(),
    #             deflection=self.deflection_entry.get(),
    #             wire_pos=self.wire_pos_combobox.get(),
    #             ground_wire_attachment=self.ground_wire_attachment_combobox.get(),
    #             quantity_of_ground_wire=self.quantity_of_ground_wire_combobox.get(),
    #             pole=self.pole_entry.get(),
    #             pole_defl_pic=self.pole_defl_entry.get(),
    #             loads_str=self.loads_entry.get(),
    #             mont_schema=self.is_mont_schema_entry.get(),
    #             path_to_txt_1=self.path_to_txt_1_entry.get(),
    #             path_to_txt_2=self.path_to_txt_2_entry.get()
    #         )
    #     else:
    #         print("!!!ОШИБКА!!!: Проверь правильность введенного провода!")

    # def validate_pole_type(self, value):
    #     if value in ["Анкерно-угловая", "Концевая", "Отпаечная", "Промежуточная"]:
    #         return True
    #     return False
    
    # def validate_voltage(self, value):
    #     return value.isdigit()
    
    # def validate_branches(self, value):
    #     if value in ["1", "2"]:
    #         return True
    #     return False

    # def validate_float(self, value):
    #     return re.match(r"^\d*\.?\d*$", value) is not None
    
    # def validate_wire_pos(self, value):
    #     if value in ["Горизонтальное", "Вертикальное", ""]:
    #         return True
    #     return False
    
    # def validate_ground_wire_attach(self, value):
    #     if value in ["Ниже верха опоры", "К верху опоры", ""]:
    #         return True
    #     return False
    
    # def validate_yes_no(self, value):
    #     if value in ["Да", "Нет"]:
    #         return True
    #     return False
    
    # def validate_quantity_of_ground_wire(self, value):
    #     if value in ["1", "2", ""]:
    #         return True
    #     return False
    
    # # def check_entries(self):
    # #     if self.pole_type_combobox.get() and self.voltage_combobox.get()\
    # #     and self.branches_combobox.get() and self.wire_entry.get():
    # #         self.generate_and_save_button.configure(state="normal")
    # #     else:
    # #         self.generate_and_save_button.configure(state='disabled')


if __name__ == "__main__":
    main_page = FoundationCalculation()
    main_page.run()