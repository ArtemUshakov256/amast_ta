import tkinter as tk


from PIL import ImageTk
from core.exceptions import CheckCalculationData
from tkinter.ttk import Combobox
from tkinter import messagebox as mb


from core.utils import (
    tempFile_back,
    tempFile_open,
    tempFile_save,
    tempFile_plus,
    tempFile_minus,
    make_path_txt,
    make_path_png,
    make_multiple_path,
    make_path_xlsx,
    AssemblyAPI
)
from core.constants import typical_ground_dict, coef_nadej_dict
from core.functionality.foundation.utils import *
from core.functionality.foundation.utils import (
    calculate_foundation,
    make_rpzf,
    save_xlsx,
    make_foundation_schema
)


class FoundationCalculation(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self.title("Расчет фундамента")
        self.geometry("781x775+400+5")
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
            width=761,
            height=770,
            borderwidth=2,
            relief="sunken"
        )

        self.back_to_main_window_button = tk.Button(
            self,
            image=self.back_icon,
            command=self.back_to_main_window
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
                "Песок Крупный (e = 0,45)",
                "Песок Крупный (e = 0,65)",
                "Песок Мелкий (e = 0,45)",
                "Песок Мелкий (e = 0,75)",
                "Песок Пылеватый (e = 0,45)",
                "Песок Пылеватый (e = 0,75)",
                "Супесь (e = 0,45)",
                "Супесь (e = 0,65)",
                "Супесь (e = 0,85)",
                "Суглинок (e = 0,45)",
                "Суглинок (e = 0,65)",
                "Суглинок (e = 0,85)",
                "Суглинок (e = 1,05)",
                "Глина (e = 0,55)",
                "Глина (e = 0,75)",
                "Глина (e = 0,95)",
                "Глина (e = 1,05)",
            ),
            textvariable=self.typical_ground_variable,
            validate="key"
        )
        self.typical_ground_combobox.bind("<<ComboboxSelected>>", self.paste_typical_ground_data)
        self.typical_ground_combobox["validatecommand"] = (
            self.typical_ground_combobox.register(self.validate_ground_type),
            "%P"
        )

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
            bd=2,
            state="disabled"
        )

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
            ),
            state="disabled"
        )
        self.quantity_of_sloy_combobox.bind("<<ComboboxSelected>>", self.activate_sloy)

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
            text='Слой №3',
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
            width=11,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.nomer_ige2_entry = tk.Entry(
            self,
            width=11,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.nomer_ige3_entry = tk.Entry(
            self,
            width=11,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.nomer_ige4_entry = tk.Entry(
            self,
            width=11,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.nomer_ige5_entry = tk.Entry(
            self,
            width=11,
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
            width=8,
            values=(
                "Песок",
                "Супесь",
                "Суглинок",
                "Глина"
            ),
            state="disabled",
            validate="key"
        )
        self.ground_type1_combobox["validatecommand"] = (
            self.ground_type1_combobox.register(self.validate_sloy_type),
            "%P"
        )
        self.ground_type2_combobox = Combobox(
            self,
            width=8,
            values=(
                "Песок",
                "Супесь",
                "Суглинок",
                "Глина"
            ),
            state="disabled",
            validate="key"
        )
        self.ground_type2_combobox["validatecommand"] = (
            self.ground_type2_combobox.register(self.validate_sloy_type),
            "%P"
        )
        self.ground_type3_combobox = Combobox(
            self,
            width=8,
            values=(
                "Песок",
                "Супесь",
                "Суглинок",
                "Глина"
            ),
            state="disabled",
            validate="key"
        )
        self.ground_type3_combobox["validatecommand"] = (
            self.ground_type3_combobox.register(self.validate_sloy_type),
            "%P"
        )
        self.ground_type4_combobox = Combobox(
            self,
            width=8,
            values=(
                "Песок",
                "Супесь",
                "Суглинок",
                "Глина"
            ),
            state="disabled",
            validate="key"
        )
        self.ground_type4_combobox["validatecommand"] = (
            self.ground_type4_combobox.register(self.validate_sloy_type),
            "%P"
        )
        self.ground_type5_combobox = Combobox(
            self,
            width=8,
            values=(
                "Песок",
                "Супесь",
                "Суглинок",
                "Глина"
            ),
            state="disabled",
            validate="key"
        )
        self.ground_type5_combobox["validatecommand"] = (
            self.ground_type5_combobox.register(self.validate_sloy_type),
            "%P"
        )

        self.ground_name_label = tk.Label(
            self,
            text='Наименование грунта',
            width=20,
            anchor="e"
        )
        self.ground_name1_entry = tk.Entry(
            self,
            width=11,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.ground_name2_entry = tk.Entry(
            self,
            width=11,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.ground_name3_entry = tk.Entry(
            self,
            width=11,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.ground_name4_entry = tk.Entry(
            self,
            width=11,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.ground_name5_entry = tk.Entry(
            self,
            width=11,
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
            width=11,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.verh_sloy2_entry = tk.Entry(
            self,
            width=11,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.verh_sloy3_entry = tk.Entry(
            self,
            width=11,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.verh_sloy4_entry = tk.Entry(
            self,
            width=11,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.verh_sloy5_entry = tk.Entry(
            self,
            width=11,
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
            width=11,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.nijn_sloy2_entry = tk.Entry(
            self,
            width=11,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.nijn_sloy3_entry = tk.Entry(
            self,
            width=11,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.nijn_sloy4_entry = tk.Entry(
            self,
            width=11,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.nijn_sloy5_entry = tk.Entry(
            self,
            width=11,
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
            width=11,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.mochn_sloya2_entry = tk.Entry(
            self,
            width=11,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.mochn_sloya3_entry = tk.Entry(
            self,
            width=11,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.mochn_sloya4_entry = tk.Entry(
            self,
            width=11,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.mochn_sloya5_entry = tk.Entry(
            self,
            width=11,
            relief="sunken",
            bd=2,
            state="disabled"
        )

        self.tech_param_label = tk.Label(
            self,
            text='Технические характеристики:',
            width=28,
            anchor="e"
        )

        self.coef_poristosti_label = tk.Label(
            self,
            text='Коэффициент пористости, e',
            width=30,
            anchor="e"
        )
        self.coef_poristosti1_entry = tk.Entry(
            self,
            width=11,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.coef_poristosti2_entry = tk.Entry(
            self,
            width=11,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.coef_poristosti3_entry = tk.Entry(
            self,
            width=11,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.coef_poristosti4_entry = tk.Entry(
            self,
            width=11,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.coef_poristosti5_entry = tk.Entry(
            self,
            width=11,
            relief="sunken",
            bd=2,
            state="disabled"
        )

        self.udel_scep_label = tk.Label(
            self,
            text='Удельное сцепление cI, кПа',
            width=30,
            anchor="e"
        )
        self.udel_scep1_entry = tk.Entry(
            self,
            width=11,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.udel_scep2_entry = tk.Entry(
            self,
            width=11,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.udel_scep3_entry = tk.Entry(
            self,
            width=11,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.udel_scep4_entry = tk.Entry(
            self,
            width=11,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.udel_scep5_entry = tk.Entry(
            self,
            width=11,
            relief="sunken",
            bd=2,
            state="disabled"
        )

        self.ugol_vn_tr_label = tk.Label(
            self,
            text='Угол внутреннего трения фI, град',
            width=30,
            anchor="e"
        )
        self.ugol_vn_tr1_entry = tk.Entry(
            self,
            width=11,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.ugol_vn_tr2_entry = tk.Entry(
            self,
            width=11,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.ugol_vn_tr3_entry = tk.Entry(
            self,
            width=11,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.ugol_vn_tr4_entry = tk.Entry(
            self,
            width=11,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.ugol_vn_tr5_entry = tk.Entry(
            self,
            width=11,
            relief="sunken",
            bd=2,
            state="disabled"
        )

        self.ves_gr_prir_label = tk.Label(
            self,
            text='Вес грунта природный, т/м3',
            width=30,
            anchor="e"
        )
        self.ves_gr_prir1_entry = tk.Entry(
            self,
            width=11,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.ves_gr_prir2_entry = tk.Entry(
            self,
            width=11,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.ves_gr_prir3_entry = tk.Entry(
            self,
            width=11,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.ves_gr_prir4_entry = tk.Entry(
            self,
            width=11,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.ves_gr_prir5_entry = tk.Entry(
            self,
            width=11,
            relief="sunken",
            bd=2,
            state="disabled"
        )

        self.ves_gr_ras_label = tk.Label(
            self,
            text='Вес грунта расчетный, т/м3',
            width=30,
            anchor="e"
        )
        self.ves_gr_ras1_entry = tk.Entry(
            self,
            width=11,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.ves_gr_ras2_entry = tk.Entry(
            self,
            width=11,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.ves_gr_ras3_entry = tk.Entry(
            self,
            width=11,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.ves_gr_ras4_entry = tk.Entry(
            self,
            width=11,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.ves_gr_ras5_entry = tk.Entry(
            self,
            width=11,
            relief="sunken",
            bd=2,
            state="disabled"
        )

        self.def_mod_label = tk.Label(
            self,
            text='Модуль деформации E, кПа',
            width=30,
            anchor="e"
        )
        self.def_mod1_entry = tk.Entry(
            self,
            width=11,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.def_mod2_entry = tk.Entry(
            self,
            width=11,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.def_mod3_entry = tk.Entry(
            self,
            width=11,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.def_mod4_entry = tk.Entry(
            self,
            width=11,
            relief="sunken",
            bd=2,
            state="disabled"
        )
        self.def_mod5_entry = tk.Entry(
            self,
            width=11,
            relief="sunken",
            bd=2,
            state="disabled"
        )

        self.pole_type_label = tk.Label(
            self,
            text='Тип опоры',
            width=20,
            anchor="e"
        )
        self.pole_type_combobox = Combobox(
            self,
            width=13,
            values=(
            "Анкерно-угловая",
            "Концевая",
            "Промежуточная"
            )
        )
        self.pole_type_combobox.bind("<<ComboboxSelected>>", self.paste_coef_nadej)

        self.coef_nadej_label = tk.Label(
            self,
            text='7.10 Коэффициент надежности yn',
            width=31,
            anchor="e"
        )
        self.coef_nadej_entry = tk.Entry(
            self,
            width=11,
            relief="sunken",
            bd=2
        )

        self.coef_usl_rab_label = tk.Label(
            self,
            text='7.2 Коэф. усл. работы СП22 yc2',
            width=31,
            anchor="e"
        )
        self.coef_usl_rab_entry = tk.Entry(
            self,
            width=11,
            relief="sunken",
            bd=2
        )

        self.calculate_button = tk.Button(
            self,
            text="Рассчитать",
            command=self.calculate
        )
        self.save_raschet_button = tk.Button(
            self,
            text="Сохранить расчет",
            command=save_xlsx
        )

        self.sr_znach_label = tk.Label(
            self,
            text="Средневзвешенные значения:",
            width=31,
            anchor="e"
        )

        self.sr_udel_scep_label = tk.Label(
            self,
            text="Удельнное сцепление cI, кПа",
            width=31,
            anchor="e"
        )
        self.sr_udel_scep_entry = tk.Entry(
            self,
            width=11,
            relief="sunken",
            bd=2
        )

        self.sr_ugol_vn_tr_label = tk.Label(
            self,
            text="Угол внутреннего трения фI, град",
            width=31,
            anchor="e"
        )
        self.sr_ugol_vn_tr_entry = tk.Entry(
            self,
            width=11,
            relief="sunken",
            bd=2
        )

        self.sr_ves_gr_ras_label = tk.Label(
            self,
            text="Вес грунта расчетный, т/м3",
            width=31,
            anchor="e"
        )
        self.sr_ves_gr_ras_entry = tk.Entry(
            self,
            width=11,
            relief="sunken",
            bd=2
        )

        self.sr_def_mod_label = tk.Label(
            self,
            text="Модуль деформации E, кПа",
            width=31,
            anchor="e"
        )
        self.sr_def_mod_entry = tk.Entry(
            self,
            width=11,
            relief="sunken",
            bd=2
        )

        self.result_label = tk.Label(
            self,
            text="Результаты расчета:",
            width=20,
            anchor="e",
            font=("default", 12, "bold")
        )

        self.ras_svai_pr_label = tk.Label(
            self,
            text="Расчет сваи по прочности",
            width=31,
            anchor="e",
            font=("default", 10, "bold")
        )

        self.coef_isp_s245_label = tk.Label(
            self,
            text="Коэф. использования (C245)",
            width=31,
            anchor="e"
        )

        self.coef_isp_s345_label = tk.Label(
            self,
            text="Коэф. использования (C345)",
            width=31,
            anchor="e"
        )

        self.ras_gor_nagr_label = tk.Label(
            self,
            text="Расчет на горизонтальную нагрузку",
            width=31,
            anchor="e",
            font=("default", 10, "bold")
        )

        self.coef_isp_gor_label = tk.Label(
            self,
            text="Коэффициент использования",
            width=31,
            anchor="e"
        )

        self.ras_def_label = tk.Label(
            self,
            text="Расчет по деформации",
            width=31,
            anchor="e",
            font=("default", 10, "bold")
        )

        self.ugol_pov_label = tk.Label(
            self,
            text="Угол поворота (без ригеля)",
            width=31,
            anchor="e"
        )

        self.ige_name_label = tk.Label(
            self,
            text="Наименование отчета ИГЭ",
            width=22,
            anchor="e"
        )
        self.ige_name_entry = tk.Entry(
            self,
            width=25,
            relief="sunken",
            bd=2
        )

        self.building_adress_label = tk.Label(
            self,
            text="Адрес строительства",
            width=22,
            anchor="e"
        )
        self.building_adress_entry = tk.Entry(
            self,
            width=25,
            relief="sunken",
            bd=2
        )

        self.razrez_skvajin_label = tk.Label(
            self,
            text="Разрез скважины",
            width=22,
            anchor="e"
        )
        self.razrez_skvajin_entry = tk.Entry(
            self,
            width=11,
            relief="sunken",
            bd=2
        )

        self.media_label = tk.Label(
            self,
            text="РПЗФ:",
            width=5,
            anchor="e",
            font=("default", 10, "bold")
        )

        self.picture1_label = tk.Label(
            self,
            text="Параметры выбранной скважины",
            width=28,
            anchor="e"
        )
        self.picture1_entry = tk.Entry(
            self,
            width=25,
            relief="sunken",
            bd=2
        )

        self.picture2_label = tk.Label(
            self,
            text="Рекомедуемые параметры грунтов",
            width=28,
            anchor="e"
        )
        self.picture2_entry = tk.Entry(
            self,
            width=25,
            relief="sunken",
            bd=2
        )

        self.browse_for_pic1_button = tk.Button(
            self,
            text="Обзор",
            command=self.browse_for_pic1
        )
        self.browse_for_pic2_button = tk.Button(
            self,
            text="Обзор",
            command=self.browse_for_pic2
        )

        self.xlsx_svai_label = tk.Label(
            self,
            text="Эксель посчитанной сваи",
            width=28,
            anchor="e"
        )
        self.xlsx_svai_entry = tk.Entry(
            self,
            width=25,
            relief="sunken",
            bd=2
        )
        self.browse_for_xlsx_button = tk.Button(
            self,
            text="Обзор",
            command=self.browse_for_xlsx
        )

        self.rpzf_button = tk.Button(
            self,
            text="Создать РПЗФ",
            command=self.call_make_rpzf
        )

        self.make_schema_button = tk.Button(
            self,
            text="Создать чертеж",
            command=self.make_schema
        )

    def run(self):
        self.draw_widgets()
        self.mainloop()

    def draw_widgets(self):
        self.module_bg.place(x=10, y=0)
        self.back_to_main_window_button.place(x=15, y=2)
        self.diam_svai_label.place(x=15, y=29)
        self.diam_svai_entry.place(x=220, y=29)
        self.thickness_svai_label.place(x=15, y=52)
        self.thickness_svai_entry.place(x=220, y=52)
        self.deepness_svai_label.place(x=15, y=75)
        self.deepness_svai_entry.place(x=220, y=75)
        self.height_svai_label.place(x=15, y=98)
        self.height_svai_entry.place(x=220, y=98)
        self.is_initial_data_checkbutton.place(x=120, y=119)
        self.typical_ground_label.place(x=325, y=29)
        self.typical_ground_combobox.place(x=530, y=29)
        self.udel_sceplenie_label.place(x=325, y=52)
        self.udel_sceplenie_entry.place(x=530, y=52)
        self.ugol_vntr_trenia_label.place(x=325, y=75)
        self.ugol_vntr_trenia_entry.place(x=530, y=75)
        self.ves_grunta_label.place(x=325, y=98)
        self.ves_grunta_entry.place(x=530, y=98)
        self.deform_module_label.place(x=325, y=119)
        self.deform_module_entry.place(x=530, y=119)
        self.ground_water_lvl_label.place(x=325, y=142)
        self.ground_water_lvl_entry.place(x=530, y=142)
        self.quantity_of_sloy_label.place(x=72,y=142)
        self.quantity_of_sloy_combobox.place(x=220, y=142)
        self.sloy_nomer1_label.place(x=255, y=164)
        self.sloy_nomer2_label.place(x=345, y=164)
        self.sloy_nomer3_label.place(x=440, y=164)
        self.sloy_nomer4_label.place(x=535, y=164)
        self.sloy_nomer5_label.place(x=630, y=164)
        self.nomer_ige_label.place(x=97, y=186)
        self.nomer_ige1_entry.place(x=244, y=186)
        self.nomer_ige2_entry.place(x=338, y=186)
        self.nomer_ige3_entry.place(x=432, y=186)
        self.nomer_ige4_entry.place(x=526, y=186)
        self.nomer_ige5_entry.place(x=620, y=186)
        self.ground_type_label.place(x=97, y=209)
        self.ground_type1_combobox.place(x=244, y=209)
        self.ground_type2_combobox.place(x=338, y=209)
        self.ground_type3_combobox.place(x=432, y=209)
        self.ground_type4_combobox.place(x=526, y=209)
        self.ground_type5_combobox.place(x=620, y=209)
        self.ground_name_label.place(x=97, y=232)
        self.ground_name1_entry.place(x=244, y=232)
        self.ground_name2_entry.place(x=338, y=232)
        self.ground_name3_entry.place(x=432, y=232)
        self.ground_name4_entry.place(x=526, y=232)
        self.ground_name5_entry.place(x=620, y=232)
        self.verh_sloy_label.place(x=97, y=255)
        self.verh_sloy1_entry.place(x=244, y=255)
        self.verh_sloy2_entry.place(x=338, y=255)
        self.verh_sloy3_entry.place(x=432, y=255)
        self.verh_sloy4_entry.place(x=526, y=255)
        self.verh_sloy5_entry.place(x=620, y=255)
        self.nijn_sloy_label.place(x=97, y=278)
        self.nijn_sloy1_entry.place(x=244, y=278)
        self.nijn_sloy2_entry.place(x=338, y=278)
        self.nijn_sloy3_entry.place(x=432, y=278)
        self.nijn_sloy4_entry.place(x=526, y=278)
        self.nijn_sloy5_entry.place(x=620, y=278)
        self.mochn_sloya_label.place(x=97, y=301)
        self.mochn_sloya1_entry.place(x=244, y=301)
        self.mochn_sloya2_entry.place(x=338, y=301)
        self.mochn_sloya3_entry.place(x=432, y=301)
        self.mochn_sloya4_entry.place(x=526, y=301)
        self.mochn_sloya5_entry.place(x=620, y=301)
        self.tech_param_label.place(x=40, y=324)
        self.coef_poristosti_label.place(x=25, y=347)
        self.coef_poristosti1_entry.place(x=244, y=347)
        self.coef_poristosti2_entry.place(x=338, y=347)
        self.coef_poristosti3_entry.place(x=432, y=347)
        self.coef_poristosti4_entry.place(x=526, y=347)
        self.coef_poristosti5_entry.place(x=620, y=347)
        self.udel_scep_label.place(x=25, y=370)
        self.udel_scep1_entry.place(x=244, y=370)
        self.udel_scep2_entry.place(x=338, y=370)
        self.udel_scep3_entry.place(x=432, y=370)
        self.udel_scep4_entry.place(x=526, y=370)
        self.udel_scep5_entry.place(x=620, y=370)
        self.ugol_vn_tr_label.place(x=25, y=393)
        self.ugol_vn_tr1_entry.place(x=244, y=393)
        self.ugol_vn_tr2_entry.place(x=338, y=393)
        self.ugol_vn_tr3_entry.place(x=432, y=393)
        self.ugol_vn_tr4_entry.place(x=526, y=393)
        self.ugol_vn_tr5_entry.place(x=620, y=393)
        self.ves_gr_prir_label.place(x=25, y=416)
        self.ves_gr_prir1_entry.place(x=244, y=416)
        self.ves_gr_prir2_entry.place(x=338, y=416)
        self.ves_gr_prir3_entry.place(x=432, y=416)
        self.ves_gr_prir4_entry.place(x=526, y=416)
        self.ves_gr_prir5_entry.place(x=620, y=416)
        # self.ves_gr_ras_label.place(x=25, y=451)
        # self.ves_gr_ras1_entry.place(x=244, y=451)
        # self.ves_gr_ras2_entry.place(x=338, y=451)
        # self.ves_gr_ras3_entry.place(x=432, y=451)
        # self.ves_gr_ras4_entry.place(x=526, y=451)
        # self.ves_gr_ras5_entry.place(x=620, y=451)
        self.def_mod_label.place(x=25, y=439)
        self.def_mod1_entry.place(x=244, y=439)
        self.def_mod2_entry.place(x=338, y=439)
        self.def_mod3_entry.place(x=432, y=439)
        self.def_mod4_entry.place(x=526, y=439)
        self.def_mod5_entry.place(x=620, y=439)
        self.pole_type_label.place(x=95, y=465)
        self.pole_type_combobox.place(x=244, y=465)
        self.coef_nadej_label.place(x=18, y=488)
        self.coef_nadej_entry.place(x=244, y=488)
        self.coef_usl_rab_label.place(x=18, y=511)
        self.coef_usl_rab_entry.place(x=244, y=511)
        self.calculate_button.place(x=30, y=534)
        self.save_raschet_button.place(x=105, y=534)
        self.rpzf_button.place(x=500, y=742)
        self.ige_name_label.place(x=12, y=698)
        self.ige_name_entry.place(x=173, y=698)
        self.building_adress_label.place(x=12, y=721)
        self.building_adress_entry.place(x=173, y=721)
        self.razrez_skvajin_label.place(x=12, y=744)
        self.razrez_skvajin_entry.place(x=173, y=744)
        self.media_label.place(x=360, y=652)
        self.picture1_label.place(x=340, y=675)
        self.picture1_entry.place(x=545, y=675)
        self.browse_for_pic1_button.place(x=702, y=670)
        self.picture2_label.place(x=340, y=698)
        self.picture2_entry.place(x=545, y=698)
        self.browse_for_pic2_button.place(x=702, y=697)
        self.xlsx_svai_label.place(x=340, y=721)
        self.xlsx_svai_entry.place(x=545, y=721)
        self.browse_for_xlsx_button.place(x=702, y=723)
        self.make_schema_button.place(x=218, y=534)

    def paste_typical_ground_data(self, event):
        typical_ground_key = self.typical_ground_combobox.get()
        self.udel_sceplenie_entry.config(state="normal")
        self.udel_sceplenie_entry.delete(0, tk.END)
        self.udel_sceplenie_entry.insert(0, typical_ground_dict[typical_ground_key]["C"])
        self.udel_sceplenie_entry.config(state="readonly")
        self.ugol_vntr_trenia_entry.config(state="normal")
        self.ugol_vntr_trenia_entry.delete(0, tk.END)
        self.ugol_vntr_trenia_entry.insert(0, typical_ground_dict[typical_ground_key]["ф"])
        self.ugol_vntr_trenia_entry.config(state="readonly")
        self.ves_grunta_entry.config(state="normal")
        self.ves_grunta_entry.delete(0, tk.END)
        self.ves_grunta_entry.insert(0, typical_ground_dict[typical_ground_key]["y"])
        self.ves_grunta_entry.config(state="readonly")
        self.deform_module_entry.config(state="normal")
        self.deform_module_entry.delete(0, tk.END)
        self.deform_module_entry.insert(0, typical_ground_dict[typical_ground_key]["E"])
        self.deform_module_entry.config(state="readonly")
        ground_type_key = self.typical_ground_combobox.get().split()[0]
        self.coef_usl_rab_entry.insert(0, coef_usl_rab_dict[ground_type_key])

    def paste_coef_nadej(self, event):
        pole_type_key = self.pole_type_combobox.get()
        self.coef_nadej_entry.config(state="normal")
        self.coef_nadej_entry.delete(0, tk.END)
        self.coef_nadej_entry.insert(0, coef_nadej_dict[pole_type_key])
        self.coef_nadej_entry.config(state="readonly")
    
    def activate_sloy(self, event):
        if self.quantity_of_sloy_combobox.get() == "1":
            self.nomer_ige2_entry.delete(0, tk.END)
            self.nomer_ige3_entry.delete(0, tk.END)
            self.nomer_ige4_entry.delete(0, tk.END)
            self.nomer_ige5_entry.delete(0, tk.END)
            self.nomer_ige2_entry.config(state="disabled")
            self.nomer_ige3_entry.config(state="disabled")
            self.nomer_ige4_entry.config(state="disabled")
            self.nomer_ige5_entry.config(state="disabled")
            self.nomer_ige1_entry.config(state="normal")
            self.ground_type2_combobox.delete(0, tk.END)
            self.ground_type3_combobox.delete(0, tk.END)
            self.ground_type4_combobox.delete(0, tk.END)
            self.ground_type5_combobox.delete(0, tk.END)
            self.ground_type2_combobox.config(state="disabled")
            self.ground_type3_combobox.config(state="disabled")
            self.ground_type4_combobox.config(state="disabled")
            self.ground_type5_combobox.config(state="disabled")
            self.ground_type1_combobox.config(state="normal")
            self.ground_name2_entry.delete(0, tk.END)
            self.ground_name3_entry.delete(0, tk.END)
            self.ground_name4_entry.delete(0, tk.END)
            self.ground_name5_entry.delete(0, tk.END)
            self.ground_name2_entry.config(state="disabled")
            self.ground_name3_entry.config(state="disabled")
            self.ground_name4_entry.config(state="disabled")
            self.ground_name5_entry.config(state="disabled")
            self.ground_name1_entry.config(state="normal")
            self.verh_sloy2_entry.delete(0, tk.END)
            self.verh_sloy3_entry.delete(0, tk.END)
            self.verh_sloy4_entry.delete(0, tk.END)
            self.verh_sloy5_entry.delete(0, tk.END)
            self.verh_sloy2_entry.config(state="disabled")
            self.verh_sloy3_entry.config(state="disabled")
            self.verh_sloy4_entry.config(state="disabled")
            self.verh_sloy5_entry.config(state="disabled")
            self.verh_sloy1_entry.config(state="normal")
            self.nijn_sloy2_entry.delete(0, tk.END)
            self.nijn_sloy3_entry.delete(0, tk.END)
            self.nijn_sloy4_entry.delete(0, tk.END)
            self.nijn_sloy5_entry.delete(0, tk.END)
            self.nijn_sloy2_entry.config(state="disabled")
            self.nijn_sloy3_entry.config(state="disabled")
            self.nijn_sloy4_entry.config(state="disabled")
            self.nijn_sloy5_entry.config(state="disabled")
            self.nijn_sloy1_entry.config(state="normal")
            self.mochn_sloya2_entry.delete(0, tk.END)
            self.mochn_sloya3_entry.delete(0, tk.END)
            self.mochn_sloya4_entry.delete(0, tk.END)
            self.mochn_sloya5_entry.delete(0, tk.END)
            self.mochn_sloya2_entry.config(state="disabled")
            self.mochn_sloya3_entry.config(state="disabled")
            self.mochn_sloya4_entry.config(state="disabled")
            self.mochn_sloya5_entry.config(state="disabled")
            self.mochn_sloya1_entry.config(state="normal")
            self.coef_poristosti2_entry.delete(0, tk.END)
            self.coef_poristosti3_entry.delete(0, tk.END)
            self.coef_poristosti4_entry.delete(0, tk.END)
            self.coef_poristosti5_entry.delete(0, tk.END)
            self.coef_poristosti2_entry.config(state="disabled")
            self.coef_poristosti3_entry.config(state="disabled")
            self.coef_poristosti4_entry.config(state="disabled")
            self.coef_poristosti5_entry.config(state="disabled")
            self.coef_poristosti1_entry.config(state="normal")
            self.udel_scep2_entry.delete(0, tk.END)
            self.udel_scep3_entry.delete(0, tk.END)
            self.udel_scep4_entry.delete(0, tk.END)
            self.udel_scep5_entry.delete(0, tk.END)
            self.udel_scep2_entry.config(state="disabled")
            self.udel_scep3_entry.config(state="disabled")
            self.udel_scep4_entry.config(state="disabled")
            self.udel_scep5_entry.config(state="disabled")
            self.udel_scep1_entry.config(state="normal")
            self.ugol_vn_tr2_entry.delete(0, tk.END)
            self.ugol_vn_tr3_entry.delete(0, tk.END)
            self.ugol_vn_tr4_entry.delete(0, tk.END)
            self.ugol_vn_tr5_entry.delete(0, tk.END)
            self.ugol_vn_tr2_entry.config(state="disabled")
            self.ugol_vn_tr3_entry.config(state="disabled")
            self.ugol_vn_tr4_entry.config(state="disabled")
            self.ugol_vn_tr5_entry.config(state="disabled")
            self.ugol_vn_tr1_entry.config(state="normal")
            self.ves_gr_prir2_entry.delete(0, tk.END)
            self.ves_gr_prir3_entry.delete(0, tk.END)
            self.ves_gr_prir4_entry.delete(0, tk.END)
            self.ves_gr_prir5_entry.delete(0, tk.END)
            self.ves_gr_prir2_entry.config(state="disabled")
            self.ves_gr_prir3_entry.config(state="disabled")
            self.ves_gr_prir4_entry.config(state="disabled")
            self.ves_gr_prir5_entry.config(state="disabled")
            self.ves_gr_prir1_entry.config(state="normal")
            self.ves_gr_ras2_entry.delete(0, tk.END)
            self.ves_gr_ras3_entry.delete(0, tk.END)
            self.ves_gr_ras4_entry.delete(0, tk.END)
            self.ves_gr_ras5_entry.delete(0, tk.END)
            self.ves_gr_ras2_entry.config(state="disabled")
            self.ves_gr_ras3_entry.config(state="disabled")
            self.ves_gr_ras4_entry.config(state="disabled")
            self.ves_gr_ras5_entry.config(state="disabled")
            self.ves_gr_ras1_entry.config(state="normal")
            self.def_mod2_entry.delete(0, tk.END)
            self.def_mod3_entry.delete(0, tk.END)
            self.def_mod4_entry.delete(0, tk.END)
            self.def_mod5_entry.delete(0, tk.END)
            self.def_mod2_entry.config(state="disabled")
            self.def_mod3_entry.config(state="disabled")
            self.def_mod4_entry.config(state="disabled")
            self.def_mod5_entry.config(state="disabled")
            self.def_mod1_entry.config(state="normal")
        elif self.quantity_of_sloy_combobox.get() == "2":
            self.nomer_ige3_entry.delete(0, tk.END)
            self.nomer_ige4_entry.delete(0, tk.END)
            self.nomer_ige5_entry.delete(0, tk.END)
            self.nomer_ige3_entry.config(state="disabled")
            self.nomer_ige4_entry.config(state="disabled")
            self.nomer_ige5_entry.config(state="disabled")
            self.nomer_ige1_entry.config(state="normal")
            self.nomer_ige2_entry.config(state="normal")
            self.ground_type3_combobox.delete(0, tk.END)
            self.ground_type4_combobox.delete(0, tk.END)
            self.ground_type5_combobox.delete(0, tk.END)
            self.ground_type3_combobox.config(state="disabled")
            self.ground_type4_combobox.config(state="disabled")
            self.ground_type5_combobox.config(state="disabled")
            self.ground_type1_combobox.config(state="normal")
            self.ground_type2_combobox.config(state="normal")
            self.ground_name3_entry.delete(0, tk.END)
            self.ground_name4_entry.delete(0, tk.END)
            self.ground_name5_entry.delete(0, tk.END)
            self.ground_name3_entry.config(state="disabled")
            self.ground_name4_entry.config(state="disabled")
            self.ground_name5_entry.config(state="disabled")
            self.ground_name1_entry.config(state="normal")
            self.ground_name2_entry.config(state="normal")
            self.verh_sloy3_entry.delete(0, tk.END)
            self.verh_sloy4_entry.delete(0, tk.END)
            self.verh_sloy5_entry.delete(0, tk.END)
            self.verh_sloy3_entry.config(state="disabled")
            self.verh_sloy4_entry.config(state="disabled")
            self.verh_sloy5_entry.config(state="disabled")
            self.verh_sloy1_entry.config(state="normal")
            self.verh_sloy2_entry.config(state="normal")
            self.nijn_sloy3_entry.delete(0, tk.END)
            self.nijn_sloy4_entry.delete(0, tk.END)
            self.nijn_sloy5_entry.delete(0, tk.END)
            self.nijn_sloy3_entry.config(state="disabled")
            self.nijn_sloy4_entry.config(state="disabled")
            self.nijn_sloy5_entry.config(state="disabled")
            self.nijn_sloy1_entry.config(state="normal")
            self.nijn_sloy2_entry.config(state="normal")
            self.mochn_sloya3_entry.delete(0, tk.END)
            self.mochn_sloya4_entry.delete(0, tk.END)
            self.mochn_sloya5_entry.delete(0, tk.END)
            self.mochn_sloya3_entry.config(state="disabled")
            self.mochn_sloya4_entry.config(state="disabled")
            self.mochn_sloya5_entry.config(state="disabled")
            self.mochn_sloya1_entry.config(state="normal")
            self.mochn_sloya2_entry.config(state="normal")
            self.coef_poristosti3_entry.delete(0, tk.END)
            self.coef_poristosti4_entry.delete(0, tk.END)
            self.coef_poristosti5_entry.delete(0, tk.END)
            self.coef_poristosti3_entry.config(state="disabled")
            self.coef_poristosti4_entry.config(state="disabled")
            self.coef_poristosti5_entry.config(state="disabled")
            self.coef_poristosti1_entry.config(state="normal")
            self.coef_poristosti2_entry.config(state="normal")
            self.udel_scep3_entry.delete(0, tk.END)
            self.udel_scep4_entry.delete(0, tk.END)
            self.udel_scep5_entry.delete(0, tk.END)
            self.udel_scep3_entry.config(state="disabled")
            self.udel_scep4_entry.config(state="disabled")
            self.udel_scep5_entry.config(state="disabled")
            self.udel_scep1_entry.config(state="normal")
            self.udel_scep2_entry.config(state="normal")
            self.ugol_vn_tr3_entry.delete(0, tk.END)
            self.ugol_vn_tr4_entry.delete(0, tk.END)
            self.ugol_vn_tr5_entry.delete(0, tk.END)
            self.ugol_vn_tr3_entry.config(state="disabled")
            self.ugol_vn_tr4_entry.config(state="disabled")
            self.ugol_vn_tr5_entry.config(state="disabled")
            self.ugol_vn_tr1_entry.config(state="normal")
            self.ugol_vn_tr2_entry.config(state="normal")
            self.ves_gr_prir3_entry.delete(0, tk.END)
            self.ves_gr_prir4_entry.delete(0, tk.END)
            self.ves_gr_prir5_entry.delete(0, tk.END)
            self.ves_gr_prir3_entry.config(state="disabled")
            self.ves_gr_prir4_entry.config(state="disabled")
            self.ves_gr_prir5_entry.config(state="disabled")
            self.ves_gr_prir1_entry.config(state="normal")
            self.ves_gr_prir2_entry.config(state="normal")
            self.ves_gr_ras3_entry.delete(0, tk.END)
            self.ves_gr_ras4_entry.delete(0, tk.END)
            self.ves_gr_ras5_entry.delete(0, tk.END)
            self.ves_gr_ras3_entry.config(state="disabled")
            self.ves_gr_ras4_entry.config(state="disabled")
            self.ves_gr_ras5_entry.config(state="disabled")
            self.ves_gr_ras1_entry.config(state="normal")
            self.ves_gr_ras2_entry.config(state="normal")
            self.def_mod3_entry.delete(0, tk.END)
            self.def_mod4_entry.delete(0, tk.END)
            self.def_mod5_entry.delete(0, tk.END)
            self.def_mod3_entry.config(state="disabled")
            self.def_mod4_entry.config(state="disabled")
            self.def_mod5_entry.config(state="disabled")
            self.def_mod1_entry.config(state="normal")
            self.def_mod2_entry.config(state="normal")
        elif self.quantity_of_sloy_combobox.get() == "3":
            self.nomer_ige4_entry.delete(0, tk.END)
            self.nomer_ige5_entry.delete(0, tk.END)
            self.nomer_ige4_entry.config(state="disabled")
            self.nomer_ige5_entry.config(state="disabled")
            self.nomer_ige1_entry.config(state="normal")
            self.nomer_ige2_entry.config(state="normal")
            self.nomer_ige3_entry.config(state="normal")
            self.ground_type4_combobox.delete(0, tk.END)
            self.ground_type5_combobox.delete(0, tk.END)
            self.ground_type4_combobox.config(state="disabled")
            self.ground_type5_combobox.config(state="disabled")
            self.ground_type1_combobox.config(state="normal")
            self.ground_type2_combobox.config(state="normal")
            self.ground_type3_combobox.config(state="normal")
            self.ground_name4_entry.delete(0, tk.END)
            self.ground_name5_entry.delete(0, tk.END)
            self.ground_name4_entry.config(state="disabled")
            self.ground_name5_entry.config(state="disabled")
            self.ground_name1_entry.config(state="normal")
            self.ground_name2_entry.config(state="normal")
            self.ground_name3_entry.config(state="normal")
            self.verh_sloy4_entry.delete(0, tk.END)
            self.verh_sloy5_entry.delete(0, tk.END)
            self.verh_sloy4_entry.config(state="disabled")
            self.verh_sloy5_entry.config(state="disabled")
            self.verh_sloy1_entry.config(state="normal")
            self.verh_sloy2_entry.config(state="normal")
            self.verh_sloy3_entry.config(state="normal")
            self.nijn_sloy4_entry.delete(0, tk.END)
            self.nijn_sloy5_entry.delete(0, tk.END)
            self.nijn_sloy4_entry.config(state="disabled")
            self.nijn_sloy5_entry.config(state="disabled")
            self.nijn_sloy1_entry.config(state="normal")
            self.nijn_sloy2_entry.config(state="normal")
            self.nijn_sloy3_entry.config(state="normal")
            self.mochn_sloya4_entry.delete(0, tk.END)
            self.mochn_sloya5_entry.delete(0, tk.END)
            self.mochn_sloya4_entry.config(state="disabled")
            self.mochn_sloya5_entry.config(state="disabled")
            self.mochn_sloya1_entry.config(state="normal")
            self.mochn_sloya2_entry.config(state="normal")
            self.mochn_sloya3_entry.config(state="normal")
            self.coef_poristosti4_entry.delete(0, tk.END)
            self.coef_poristosti5_entry.delete(0, tk.END)
            self.coef_poristosti4_entry.config(state="disabled")
            self.coef_poristosti5_entry.config(state="disabled")
            self.coef_poristosti1_entry.config(state="normal")
            self.coef_poristosti2_entry.config(state="normal")
            self.coef_poristosti3_entry.config(state="normal")
            self.udel_scep4_entry.delete(0, tk.END)
            self.udel_scep5_entry.delete(0, tk.END)
            self.udel_scep4_entry.config(state="disabled")
            self.udel_scep5_entry.config(state="disabled")
            self.udel_scep1_entry.config(state="normal")
            self.udel_scep2_entry.config(state="normal")
            self.udel_scep3_entry.config(state="normal")
            self.ugol_vn_tr4_entry.delete(0, tk.END)
            self.ugol_vn_tr5_entry.delete(0, tk.END)
            self.ugol_vn_tr4_entry.config(state="disabled")
            self.ugol_vn_tr5_entry.config(state="disabled")
            self.ugol_vn_tr1_entry.config(state="normal")
            self.ugol_vn_tr2_entry.config(state="normal")
            self.ugol_vn_tr3_entry.config(state="normal")
            self.ves_gr_prir4_entry.delete(0, tk.END)
            self.ves_gr_prir5_entry.delete(0, tk.END)
            self.ves_gr_prir4_entry.config(state="disabled")
            self.ves_gr_prir5_entry.config(state="disabled")
            self.ves_gr_prir1_entry.config(state="normal")
            self.ves_gr_prir2_entry.config(state="normal")
            self.ves_gr_prir3_entry.config(state="normal")
            self.ves_gr_ras4_entry.delete(0, tk.END)
            self.ves_gr_ras5_entry.delete(0, tk.END)
            self.ves_gr_ras4_entry.config(state="disabled")
            self.ves_gr_ras5_entry.config(state="disabled")
            self.ves_gr_ras1_entry.config(state="normal")
            self.ves_gr_ras2_entry.config(state="normal")
            self.ves_gr_ras3_entry.config(state="normal")
            self.def_mod4_entry.delete(0, tk.END)
            self.def_mod5_entry.delete(0, tk.END)
            self.def_mod4_entry.config(state="disabled")
            self.def_mod5_entry.config(state="disabled")
            self.def_mod1_entry.config(state="normal")
            self.def_mod2_entry.config(state="normal")
            self.def_mod3_entry.config(state="normal")
        elif self.quantity_of_sloy_combobox.get() == "4":
            self.nomer_ige5_entry.delete(0, tk.END)
            self.nomer_ige5_entry.config(state="disabled")
            self.nomer_ige1_entry.config(state="normal")
            self.nomer_ige2_entry.config(state="normal")
            self.nomer_ige3_entry.config(state="normal")
            self.nomer_ige4_entry.config(state="normal")
            self.ground_type5_combobox.delete(0, tk.END)
            self.ground_type5_combobox.config(state="disabled")
            self.ground_type1_combobox.config(state="normal")
            self.ground_type2_combobox.config(state="normal")
            self.ground_type3_combobox.config(state="normal")
            self.ground_type4_combobox.config(state="normal")
            self.ground_name5_entry.delete(0, tk.END)
            self.ground_name5_entry.config(state="disabled")
            self.ground_name1_entry.config(state="normal")
            self.ground_name2_entry.config(state="normal")
            self.ground_name3_entry.config(state="normal")
            self.ground_name4_entry.config(state="normal")
            self.verh_sloy5_entry.delete(0, tk.END)
            self.verh_sloy5_entry.config(state="disabled")
            self.verh_sloy1_entry.config(state="normal")
            self.verh_sloy2_entry.config(state="normal")
            self.verh_sloy3_entry.config(state="normal")
            self.verh_sloy4_entry.config(state="normal")
            self.nijn_sloy5_entry.delete(0, tk.END)
            self.nijn_sloy5_entry.config(state="disabled")
            self.nijn_sloy1_entry.config(state="normal")
            self.nijn_sloy2_entry.config(state="normal")
            self.nijn_sloy3_entry.config(state="normal")
            self.nijn_sloy4_entry.config(state="normal")
            self.mochn_sloya5_entry.delete(0, tk.END)
            self.mochn_sloya5_entry.config(state="disabled")
            self.mochn_sloya1_entry.config(state="normal")
            self.mochn_sloya2_entry.config(state="normal")
            self.mochn_sloya3_entry.config(state="normal")
            self.mochn_sloya4_entry.config(state="normal")
            self.coef_poristosti5_entry.delete(0, tk.END)
            self.coef_poristosti5_entry.config(state="disabled")
            self.coef_poristosti1_entry.config(state="normal")
            self.coef_poristosti2_entry.config(state="normal")
            self.coef_poristosti3_entry.config(state="normal")
            self.coef_poristosti4_entry.config(state="normal")
            self.udel_scep5_entry.delete(0, tk.END)
            self.udel_scep5_entry.config(state="disabled")
            self.udel_scep1_entry.config(state="normal")
            self.udel_scep2_entry.config(state="normal")
            self.udel_scep3_entry.config(state="normal")
            self.udel_scep4_entry.config(state="normal")
            self.ugol_vn_tr5_entry.delete(0, tk.END)
            self.ugol_vn_tr5_entry.config(state="disabled")
            self.ugol_vn_tr1_entry.config(state="normal")
            self.ugol_vn_tr2_entry.config(state="normal")
            self.ugol_vn_tr3_entry.config(state="normal")
            self.ugol_vn_tr4_entry.config(state="normal")
            self.ves_gr_prir5_entry.delete(0, tk.END)
            self.ves_gr_prir5_entry.config(state="disabled")
            self.ves_gr_prir1_entry.config(state="normal")
            self.ves_gr_prir2_entry.config(state="normal")
            self.ves_gr_prir3_entry.config(state="normal")
            self.ves_gr_prir4_entry.config(state="normal")
            self.ves_gr_ras5_entry.delete(0, tk.END)
            self.ves_gr_ras5_entry.config(state="disabled")
            self.ves_gr_ras1_entry.config(state="normal")
            self.ves_gr_ras2_entry.config(state="normal")
            self.ves_gr_ras3_entry.config(state="normal")
            self.ves_gr_ras4_entry.config(state="normal")
            self.def_mod5_entry.delete(0, tk.END)
            self.def_mod5_entry.config(state="disabled")
            self.def_mod1_entry.config(state="normal")
            self.def_mod2_entry.config(state="normal")
            self.def_mod3_entry.config(state="normal")
            self.def_mod4_entry.config(state="normal")
        else:
            self.nomer_ige1_entry.config(state="normal")
            self.nomer_ige2_entry.config(state="normal")
            self.nomer_ige3_entry.config(state="normal")
            self.nomer_ige4_entry.config(state="normal")
            self.nomer_ige5_entry.config(state="normal")
            self.ground_type1_combobox.config(state="normal")
            self.ground_type2_combobox.config(state="normal")
            self.ground_type3_combobox.config(state="normal")
            self.ground_type4_combobox.config(state="normal")
            self.ground_type5_combobox.config(state="normal")
            self.ground_name1_entry.config(state="normal")
            self.ground_name2_entry.config(state="normal")
            self.ground_name3_entry.config(state="normal")
            self.ground_name4_entry.config(state="normal")
            self.ground_name5_entry.config(state="normal")
            self.verh_sloy1_entry.config(state="normal")
            self.verh_sloy2_entry.config(state="normal")
            self.verh_sloy3_entry.config(state="normal")
            self.verh_sloy4_entry.config(state="normal")
            self.verh_sloy5_entry.config(state="normal")
            self.nijn_sloy1_entry.config(state="normal")
            self.nijn_sloy2_entry.config(state="normal")
            self.nijn_sloy3_entry.config(state="normal")
            self.nijn_sloy4_entry.config(state="normal")
            self.nijn_sloy5_entry.config(state="normal")
            self.mochn_sloya1_entry.config(state="normal")
            self.mochn_sloya2_entry.config(state="normal")
            self.mochn_sloya3_entry.config(state="normal")
            self.mochn_sloya4_entry.config(state="normal")
            self.mochn_sloya5_entry.config(state="normal")
            self.coef_poristosti1_entry.config(state="normal")
            self.coef_poristosti2_entry.config(state="normal")
            self.coef_poristosti3_entry.config(state="normal")
            self.coef_poristosti4_entry.config(state="normal")
            self.coef_poristosti5_entry.config(state="normal")
            self.udel_scep1_entry.config(state="normal")
            self.udel_scep2_entry.config(state="normal")
            self.udel_scep3_entry.config(state="normal")
            self.udel_scep4_entry.config(state="normal")
            self.udel_scep5_entry.config(state="normal")
            self.ugol_vn_tr1_entry.config(state="normal")
            self.ugol_vn_tr2_entry.config(state="normal")
            self.ugol_vn_tr3_entry.config(state="normal")
            self.ugol_vn_tr4_entry.config(state="normal")
            self.ugol_vn_tr5_entry.config(state="normal")
            self.ves_gr_prir1_entry.config(state="normal")
            self.ves_gr_prir2_entry.config(state="normal")
            self.ves_gr_prir3_entry.config(state="normal")
            self.ves_gr_prir4_entry.config(state="normal")
            self.ves_gr_prir5_entry.config(state="normal")
            self.ves_gr_ras1_entry.config(state="normal")
            self.ves_gr_ras2_entry.config(state="normal")
            self.ves_gr_ras3_entry.config(state="normal")
            self.ves_gr_ras4_entry.config(state="normal")
            self.ves_gr_ras5_entry.config(state="normal")
            self.def_mod1_entry.config(state="normal")
            self.def_mod2_entry.config(state="normal")
            self.def_mod3_entry.config(state="normal")
            self.def_mod4_entry.config(state="normal")
            self.def_mod5_entry.config(state="normal")

    def toggle_state(self):
        if self.is_initial_data_var.get() == 0:
            self.typical_ground_combobox.config(state="normal")
            self.udel_sceplenie_entry.config(state="normal")
            self.ugol_vntr_trenia_entry.config(state="normal")
            self.ves_grunta_entry.config(state="normal")
            self.deform_module_entry.config(state="normal")
            self.ground_water_lvl_entry.config(state="disabled")
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
            self.ground_type1_combobox.delete(0, tk.END)
            self.ground_type2_combobox.delete(0, tk.END)
            self.ground_type3_combobox.delete(0, tk.END)
            self.ground_type4_combobox.delete(0, tk.END)
            self.ground_type5_combobox.delete(0, tk.END)
            self.ground_type1_combobox.config(state="disabled")
            self.ground_type2_combobox.config(state="disabled")
            self.ground_type3_combobox.config(state="disabled")
            self.ground_type4_combobox.config(state="disabled")
            self.ground_type5_combobox.config(state="disabled")
            self.ground_name1_entry.delete(0, tk.END)
            self.ground_name2_entry.delete(0, tk.END)
            self.ground_name3_entry.delete(0, tk.END)
            self.ground_name4_entry.delete(0, tk.END)
            self.ground_name5_entry.delete(0, tk.END)
            self.ground_name1_entry.config(state="disabled")
            self.ground_name2_entry.config(state="disabled")
            self.ground_name3_entry.config(state="disabled")
            self.ground_name4_entry.config(state="disabled")
            self.ground_name5_entry.config(state="disabled")
            self.verh_sloy1_entry.delete(0, tk.END)
            self.verh_sloy2_entry.delete(0, tk.END)
            self.verh_sloy3_entry.delete(0, tk.END)
            self.verh_sloy4_entry.delete(0, tk.END)
            self.verh_sloy5_entry.delete(0, tk.END)
            self.verh_sloy1_entry.config(state="disabled")
            self.verh_sloy2_entry.config(state="disabled")
            self.verh_sloy3_entry.config(state="disabled")
            self.verh_sloy4_entry.config(state="disabled")
            self.verh_sloy5_entry.config(state="disabled")
            self.nijn_sloy1_entry.delete(0, tk.END)
            self.nijn_sloy2_entry.delete(0, tk.END)
            self.nijn_sloy3_entry.delete(0, tk.END)
            self.nijn_sloy4_entry.delete(0, tk.END)
            self.nijn_sloy5_entry.delete(0, tk.END)
            self.nijn_sloy1_entry.config(state="disabled")
            self.nijn_sloy2_entry.config(state="disabled")
            self.nijn_sloy3_entry.config(state="disabled")
            self.nijn_sloy4_entry.config(state="disabled")
            self.nijn_sloy5_entry.config(state="disabled")
            self.mochn_sloya1_entry.delete(0, tk.END)
            self.mochn_sloya2_entry.delete(0, tk.END)
            self.mochn_sloya3_entry.delete(0, tk.END)
            self.mochn_sloya4_entry.delete(0, tk.END)
            self.mochn_sloya5_entry.delete(0, tk.END)
            self.mochn_sloya1_entry.config(state="disabled")
            self.mochn_sloya2_entry.config(state="disabled")
            self.mochn_sloya3_entry.config(state="disabled")
            self.mochn_sloya4_entry.config(state="disabled")
            self.mochn_sloya5_entry.config(state="disabled")
            self.coef_poristosti1_entry.delete(0, tk.END)
            self.coef_poristosti2_entry.delete(0, tk.END)
            self.coef_poristosti3_entry.delete(0, tk.END)
            self.coef_poristosti4_entry.delete(0, tk.END)
            self.coef_poristosti5_entry.delete(0, tk.END)
            self.coef_poristosti1_entry.config(state="disabled")
            self.coef_poristosti2_entry.config(state="disabled")
            self.coef_poristosti3_entry.config(state="disabled")
            self.coef_poristosti4_entry.config(state="disabled")
            self.coef_poristosti5_entry.config(state="disabled")
            self.udel_scep1_entry.delete(0, tk.END)
            self.udel_scep2_entry.delete(0, tk.END)
            self.udel_scep3_entry.delete(0, tk.END)
            self.udel_scep4_entry.delete(0, tk.END)
            self.udel_scep5_entry.delete(0, tk.END)
            self.udel_scep1_entry.config(state="disabled")
            self.udel_scep2_entry.config(state="disabled")
            self.udel_scep3_entry.config(state="disabled")
            self.udel_scep4_entry.config(state="disabled")
            self.udel_scep5_entry.config(state="disabled")
            self.ugol_vn_tr1_entry.delete(0, tk.END)
            self.ugol_vn_tr2_entry.delete(0, tk.END)
            self.ugol_vn_tr3_entry.delete(0, tk.END)
            self.ugol_vn_tr4_entry.delete(0, tk.END)
            self.ugol_vn_tr5_entry.delete(0, tk.END)
            self.ugol_vn_tr1_entry.config(state="disabled")
            self.ugol_vn_tr2_entry.config(state="disabled")
            self.ugol_vn_tr3_entry.config(state="disabled")
            self.ugol_vn_tr4_entry.config(state="disabled")
            self.ugol_vn_tr5_entry.config(state="disabled")
            self.ves_gr_prir1_entry.delete(0, tk.END)
            self.ves_gr_prir2_entry.delete(0, tk.END)
            self.ves_gr_prir3_entry.delete(0, tk.END)
            self.ves_gr_prir4_entry.delete(0, tk.END)
            self.ves_gr_prir5_entry.delete(0, tk.END)
            self.ves_gr_prir1_entry.config(state="disabled")
            self.ves_gr_prir2_entry.config(state="disabled")
            self.ves_gr_prir3_entry.config(state="disabled")
            self.ves_gr_prir4_entry.config(state="disabled")
            self.ves_gr_prir5_entry.config(state="disabled")
            self.ves_gr_ras1_entry.delete(0, tk.END)
            self.ves_gr_ras2_entry.delete(0, tk.END)
            self.ves_gr_ras3_entry.delete(0, tk.END)
            self.ves_gr_ras4_entry.delete(0, tk.END)
            self.ves_gr_ras5_entry.delete(0, tk.END)
            self.ves_gr_ras1_entry.config(state="disabled")
            self.ves_gr_ras2_entry.config(state="disabled")
            self.ves_gr_ras3_entry.config(state="disabled")
            self.ves_gr_ras4_entry.config(state="disabled")
            self.ves_gr_ras5_entry.config(state="disabled")
            self.def_mod1_entry.delete(0, tk.END)
            self.def_mod2_entry.delete(0, tk.END)
            self.def_mod3_entry.delete(0, tk.END)
            self.def_mod4_entry.delete(0, tk.END)
            self.def_mod5_entry.delete(0, tk.END)
            self.def_mod1_entry.config(state="disabled")
            self.def_mod2_entry.config(state="disabled")
            self.def_mod3_entry.config(state="disabled")
            self.def_mod4_entry.config(state="disabled")
            self.def_mod5_entry.config(state="disabled")
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
            self.quantity_of_sloy_combobox.config(state="normal")
            self.ground_water_lvl_entry.config(state="normal")

    def insert_result(self):
        if self.is_initial_data_var.get():
            self.sr_znach_label.place(x=50, y=561)
            self.sr_udel_scep_label.place(x=19, y=584)
            self.sr_udel_scep_entry.place(x=244, y=584)
            self.sr_ugol_vn_tr_label.place(x=19, y=607)
            self.sr_ugol_vn_tr_entry.place(x=244, y=607)
            self.sr_ves_gr_ras_label.place(x=19, y=630)
            self.sr_ves_gr_ras_entry.place(x=244, y=630)
            self.sr_def_mod_label.place(x=19, y=653)
            self.sr_def_mod_entry.place(x=244, y=653)
            self.sr_udel_scep_entry.delete(0, tk.END)
            self.sr_udel_scep_entry.insert(0, self.result["sr_udel_scep"])
            self.sr_ugol_vn_tr_entry.delete(0, tk.END)
            self.sr_ugol_vn_tr_entry.insert(0, self.result["sr_ugol_vn_tr"])
            self.sr_ves_gr_ras_entry.delete(0, tk.END)
            self.sr_ves_gr_ras_entry.insert(0, self.result["sr_ves_gr_ras"])
            self.sr_def_mod_entry.delete(0, tk.END)
            self.sr_def_mod_entry.insert(0, self.result["sr_def_mod"])
        self.coef_isp_s245_entry.delete(0, tk.END)
        self.coef_isp_s245_entry.insert(0, self.result["coef_isp_s245"])
        self.coef_isp_s345_entry.delete(0, tk.END)
        self.coef_isp_s345_entry.insert(0, self.result["coef_isp_s345"])
        self.coef_isp_gor_entry.delete(0, tk.END)
        self.coef_isp_gor_entry.insert(0, self.result["coef_isp_gor"])
        self.ugol_pov_entry.delete(0, tk.END)
        self.ugol_pov_entry.insert(0, self.result["ugol_pov"])
    
    def calculate(self):
        try:
            self.result = calculate_foundation(
                moment=self.parent.pls_pole_data["moment"],
                vert_force=self.parent.pls_pole_data["vert_force"],
                shear_force=self.parent.pls_pole_data["shear_force"],
                flanec_diam=self.diam_svai_entry.get(),
                thickness_svai=self.thickness_svai_entry.get(),
                deepness_svai=self.deepness_svai_entry.get(),
                height_svai=self.height_svai_entry.get(),
                typical_ground=self.typical_ground_combobox.get(),
                is_init_data=self.is_initial_data_var.get(),
                ground_water_lvl=self.ground_water_lvl_entry.get(),
                quantity_of_ige=self.quantity_of_sloy_combobox.get(), 
                nomer_ige1=self.nomer_ige1_entry.get(),
                nomer_ige2=self.nomer_ige2_entry.get(),
                nomer_ige3=self.nomer_ige3_entry.get(),
                nomer_ige4=self.nomer_ige4_entry.get(),
                nomer_ige5=self.nomer_ige5_entry.get(),
                ground_type1=self.ground_type1_combobox.get(),
                ground_type2=self.ground_type2_combobox.get(),
                ground_type3=self.ground_type3_combobox.get(),
                ground_type4=self.ground_type4_combobox.get(),
                ground_type5=self.ground_type5_combobox.get(),
                ground_name1=self.ground_name1_entry.get(),
                ground_name2=self.ground_name2_entry.get(),
                ground_name3=self.ground_name3_entry.get(),
                ground_name4=self.ground_name4_entry.get(),
                ground_name5=self.ground_name5_entry.get(),
                verh_sloy1=self.verh_sloy1_entry.get(),
                verh_sloy2=self.verh_sloy2_entry.get(),
                verh_sloy3=self.verh_sloy3_entry.get(),
                verh_sloy4=self.verh_sloy4_entry.get(),
                verh_sloy5=self.verh_sloy5_entry.get(),
                nijn_sloy1=self.nijn_sloy1_entry.get(),
                nijn_sloy2=self.nijn_sloy2_entry.get(),
                nijn_sloy3=self.nijn_sloy3_entry.get(),
                nijn_sloy4=self.nijn_sloy4_entry.get(),
                nijn_sloy5=self.nijn_sloy5_entry.get(),
                coef_poristosti1=self.coef_poristosti1_entry.get(),
                coef_poristosti2=self.coef_poristosti2_entry.get(),
                coef_poristosti3=self.coef_poristosti3_entry.get(),
                coef_poristosti4=self.coef_poristosti4_entry.get(),
                coef_poristosti5=self.coef_poristosti5_entry.get(),
                udel_scep1=self.udel_scep1_entry.get(),
                udel_scep2=self.udel_scep2_entry.get(),
                udel_scep3=self.udel_scep3_entry.get(),
                udel_scep4=self.udel_scep4_entry.get(),
                udel_scep5=self.udel_scep5_entry.get(),
                ugol_vn_tr1=self.ugol_vn_tr1_entry.get(),
                ugol_vn_tr2=self.ugol_vn_tr2_entry.get(),
                ugol_vn_tr3=self.ugol_vn_tr3_entry.get(),
                ugol_vn_tr4=self.ugol_vn_tr4_entry.get(),
                ugol_vn_tr5=self.ugol_vn_tr5_entry.get(),
                ves_gr_prir1=self.ves_gr_prir1_entry.get(),
                ves_gr_prir2=self.ves_gr_prir2_entry.get(),
                ves_gr_prir3=self.ves_gr_prir3_entry.get(),
                ves_gr_prir4=self.ves_gr_prir4_entry.get(),
                ves_gr_prir5=self.ves_gr_prir5_entry.get(),
                def_mod1=self.def_mod1_entry.get(),
                def_mod2=self.def_mod2_entry.get(),
                def_mod3=self.def_mod3_entry.get(),
                def_mod4=self.def_mod4_entry.get(),
                def_mod5=self.def_mod5_entry.get(),
                pole_type=self.pole_type_combobox.get(),
                ugol_vntr_tr=self.ugol_vntr_trenia_entry.get(),
                udel_sceplenie=self.udel_sceplenie_entry.get(),
                ves_grunta=self.ves_grunta_entry.get(),
                deform_module=self.deform_module_entry.get(),
                coef_nadej=self.coef_nadej_entry.get(),
                coef_usl_rab=self.coef_usl_rab_entry.get()
            )
        except CheckCalculationData as e:
            mb.showinfo("Ошибка", e)
        self.coef_isp_s245_entry = tk.Entry(
            self,
            width=11,
            relief="sunken",
            bd=2,
            bg=self.result["coef_isp_s245_bg"]
        )
        self.coef_isp_s345_entry = tk.Entry(
            self,
            width=11,
            relief="sunken",
            bd=2,
            bg=self.result["coef_isp_s345_bg"]
        )
        self.coef_isp_gor_entry = tk.Entry(
            self,
            width=11,
            relief="sunken",
            bd=2,
            bg=self.result["coef_isp_gor_bg"]
        )
        self.ugol_pov_entry = tk.Entry(
            self,
            width=11,
            relief="sunken",
            bd=2,
            bg=self.result["ugol_pov_bg"]
        )
        self.result_label.place(x=410, y=462)
        self.ras_svai_pr_label.place(x=360, y=483)
        self.coef_isp_s245_label.place(x=350, y=506)
        self.coef_isp_s245_entry.place(x=572, y=506)
        self.coef_isp_s345_label.place(x=350, y=529)
        self.coef_isp_s345_entry.place(x=572, y=529)
        self.ras_gor_nagr_label.place(x=390, y=552)
        self.coef_isp_gor_label.place(x=350, y=575)
        self.coef_isp_gor_entry.place(x=572, y=575)
        self.ras_def_label.place(x=360, y=598)
        self.ugol_pov_label.place(x=350, y=621)
        self.ugol_pov_entry.place(x=572, y=621)

        self.insert_result()

    def call_make_rpzf(self):
        make_rpzf(
            project_name=self.parent.project_name,
            project_code=self.parent.project_code,
            pole_code=self.parent.pole_code,
            developer=self.parent.developer,
            diam_svai=self.diam_svai_entry.get(),
            deepness_svai=self.deepness_svai_entry.get(),
            height_svai=self.height_svai_entry.get(),
            moment1=self.parent.pls_pole_data["moment"],
            vert_force1=self.parent.pls_pole_data["vert_force"],
            shear_force1=self.parent.pls_pole_data["shear_force"],
            ige_name=self.ige_name_entry.get(),
            building_adress=self.building_adress_entry.get(),
            razrez_skvajin=self.razrez_skvajin_entry.get(),
            picture1=self.picture1_entry.get(),
            picture2=self.picture2_entry.get(),
            xlsx_svai=self.xlsx_svai_entry.get()
        )

    def make_schema(self):
        make_foundation_schema(
            diam_svai=self.deepness_svai_entry.get(),
            
        )

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
    
    def browse_for_pic1(self):
        self.file_path = make_path_png()
        self.picture1_entry.delete("0", "end") 
        self.picture1_entry.insert("insert", self.file_path)

    def browse_for_pic2(self):
        self.file_path = make_path_png()
        self.picture2_entry.delete("0", "end") 
        self.picture2_entry.insert("insert", self.file_path)

    def browse_for_xlsx(self):
        self.file_path = make_path_xlsx()
        self.xlsx_svai_entry.delete("0", "end") 
        self.xlsx_svai_entry.insert("insert", self.file_path)

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
    
    def validate_ground_type(self, value):
        if value in [
                "Песок Крупный (e = 0,45)",
                "Песок Крупный (e = 0,65)",
                "Песок Мелкий (e = 0,45)",
                "Песок Мелкий (e = 0,75)",
                "Песок Пылеватый (e = 0,45)",
                "Песок Пылеватый (e = 0,75)",
                "Супесь (e = 0,45)",
                "Супесь (e = 0,65)",
                "Супесь (e = 0,85)",
                "Суглинок (e = 0,45)",
                "Суглинок (e = 0,65)",
                "Суглинок (e = 0,85)",
                "Суглинок (e = 1,05)",
                "Глина (e = 0,55)",
                "Глина (e = 0,75)",
                "Глина (e = 0,95)",
                "Глина (e = 1,05)",
        ]:
            return True
        return False


if __name__ == "__main__":
    main_page = FoundationCalculation()
    main_page.run()