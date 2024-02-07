import tkinter as tk


from PIL import ImageTk
from core.db.db_connector import Database
from core.exceptions import CheckCalculationData
from tkinter.ttk import Combobox
from tkinter import messagebox as mb


from core.utils import (
    # tempFile_back,
    make_path_png,
    make_path_xlsx,
    AssemblyAPI
)
from core.constants import typical_ground_dict, coef_nadej_dict
from core.functionality.foundation.utils import *
from core.functionality.foundation.utils import (
    calculate_foundation,
    make_rpzf,
    save_xlsx,
)


class FoundationCalculation(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self.title("Расчет фундамента")
        self.geometry("781x775+400+5")
        self.resizable(False, False)
        self.config(bg="#FFFFFF")
        self.db = Database()

        # self.back_icon = ImageTk.PhotoImage(
        #     file=tempFile_back
        # )

        self.module_bg = tk.Frame(
            self,
            width=761,
            height=770,
            borderwidth=2,
            relief="sunken"
        )

        self.back_to_main_window_button = tk.Button(
            self,
            # image=self.back_icon,
            text="Назад",
            command=self.back_to_main_window
        )

        self.bending_moment_label = tk.Label(
            self,
            text='Момент у основания, кНм',
            width=21,
            anchor="e"
        )
        self.bending_moment_entry = tk.Entry(
            self,
            width=10,
            relief="sunken",
            bd=2
        )

        self.vert_force_label = tk.Label(
            self,
            text='Вертикальная сила, кН',
            width=20,
            anchor="e"
        )
        self.vert_force_entry = tk.Entry(
            self,
            width=10,
            relief="sunken",
            bd=2
        )

        self.shear_force_label = tk.Label(
            self,
            text='Горизонтальная сила, кН',
            width=20,
            anchor="e"
        )
        self.shear_force_entry = tk.Entry(
            self,
            width=10,
            relief="sunken",
            bd=2
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
            text='Высота головы сваи, мм',
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
            text='Уровень грунтовых вод, м',
            width=28,
            anchor="e"
        )
        self.ground_water_lvl_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            state="disable"
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
            state="disable"
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
            state="disable"
        )
        self.nomer_ige2_entry = tk.Entry(
            self,
            width=11,
            relief="sunken",
            bd=2,
            state="disable"
        )
        self.nomer_ige3_entry = tk.Entry(
            self,
            width=11,
            relief="sunken",
            bd=2,
            state="disable"
        )
        self.nomer_ige4_entry = tk.Entry(
            self,
            width=11,
            relief="sunken",
            bd=2,
            state="disable"
        )
        self.nomer_ige5_entry = tk.Entry(
            self,
            width=11,
            relief="sunken",
            bd=2,
            state="disable"
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
            state="disable",
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
            state="disable",
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
            state="disable",
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
            state="disable",
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
            state="disable",
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
            state="disable"
        )
        self.ground_name2_entry = tk.Entry(
            self,
            width=11,
            relief="sunken",
            bd=2,
            state="disable"
        )
        self.ground_name3_entry = tk.Entry(
            self,
            width=11,
            relief="sunken",
            bd=2,
            state="disable"
        )
        self.ground_name4_entry = tk.Entry(
            self,
            width=11,
            relief="sunken",
            bd=2,
            state="disable"
        )
        self.ground_name5_entry = tk.Entry(
            self,
            width=11,
            relief="sunken",
            bd=2,
            state="disable"
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
            state="disable"
        )
        self.verh_sloy2_entry = tk.Entry(
            self,
            width=11,
            relief="sunken",
            bd=2,
            state="disable"
        )
        self.verh_sloy3_entry = tk.Entry(
            self,
            width=11,
            relief="sunken",
            bd=2,
            state="disable"
        )
        self.verh_sloy4_entry = tk.Entry(
            self,
            width=11,
            relief="sunken",
            bd=2,
            state="disable"
        )
        self.verh_sloy5_entry = tk.Entry(
            self,
            width=11,
            relief="sunken",
            bd=2,
            state="disable"
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
            state="disable"
        )
        self.nijn_sloy2_entry = tk.Entry(
            self,
            width=11,
            relief="sunken",
            bd=2,
            state="disable"
        )
        self.nijn_sloy3_entry = tk.Entry(
            self,
            width=11,
            relief="sunken",
            bd=2,
            state="disable"
        )
        self.nijn_sloy4_entry = tk.Entry(
            self,
            width=11,
            relief="sunken",
            bd=2,
            state="disable"
        )
        self.nijn_sloy5_entry = tk.Entry(
            self,
            width=11,
            relief="sunken",
            bd=2,
            state="disable"
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
            state="disable"
        )
        self.coef_poristosti2_entry = tk.Entry(
            self,
            width=11,
            relief="sunken",
            bd=2,
            state="disable"
        )
        self.coef_poristosti3_entry = tk.Entry(
            self,
            width=11,
            relief="sunken",
            bd=2,
            state="disable"
        )
        self.coef_poristosti4_entry = tk.Entry(
            self,
            width=11,
            relief="sunken",
            bd=2,
            state="disable"
        )
        self.coef_poristosti5_entry = tk.Entry(
            self,
            width=11,
            relief="sunken",
            bd=2,
            state="disable"
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
            state="disable"
        )
        self.udel_scep2_entry = tk.Entry(
            self,
            width=11,
            relief="sunken",
            bd=2,
            state="disable"
        )
        self.udel_scep3_entry = tk.Entry(
            self,
            width=11,
            relief="sunken",
            bd=2,
            state="disable"
        )
        self.udel_scep4_entry = tk.Entry(
            self,
            width=11,
            relief="sunken",
            bd=2,
            state="disable"
        )
        self.udel_scep5_entry = tk.Entry(
            self,
            width=11,
            relief="sunken",
            bd=2,
            state="disable"
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
            state="disable"
        )
        self.ugol_vn_tr2_entry = tk.Entry(
            self,
            width=11,
            relief="sunken",
            bd=2,
            state="disable"
        )
        self.ugol_vn_tr3_entry = tk.Entry(
            self,
            width=11,
            relief="sunken",
            bd=2,
            state="disable"
        )
        self.ugol_vn_tr4_entry = tk.Entry(
            self,
            width=11,
            relief="sunken",
            bd=2,
            state="disable"
        )
        self.ugol_vn_tr5_entry = tk.Entry(
            self,
            width=11,
            relief="sunken",
            bd=2,
            state="disable"
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
            state="disable"
        )
        self.ves_gr_prir2_entry = tk.Entry(
            self,
            width=11,
            relief="sunken",
            bd=2,
            state="disable"
        )
        self.ves_gr_prir3_entry = tk.Entry(
            self,
            width=11,
            relief="sunken",
            bd=2,
            state="disable"
        )
        self.ves_gr_prir4_entry = tk.Entry(
            self,
            width=11,
            relief="sunken",
            bd=2,
            state="disable"
        )
        self.ves_gr_prir5_entry = tk.Entry(
            self,
            width=11,
            relief="sunken",
            bd=2,
            state="disable"
        )

        self.def_mod_label = tk.Label(
            self,
            text='Модуль деформации E, МПа',
            width=30,
            anchor="e"
        )
        self.def_mod1_entry = tk.Entry(
            self,
            width=11,
            relief="sunken",
            bd=2,
            state="disable"
        )
        self.def_mod2_entry = tk.Entry(
            self,
            width=11,
            relief="sunken",
            bd=2,
            state="disable"
        )
        self.def_mod3_entry = tk.Entry(
            self,
            width=11,
            relief="sunken",
            bd=2,
            state="disable"
        )
        self.def_mod4_entry = tk.Entry(
            self,
            width=11,
            relief="sunken",
            bd=2,
            state="disable"
        )
        self.def_mod5_entry = tk.Entry(
            self,
            width=11,
            relief="sunken",
            bd=2,
            state="disable"
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

        self.calculate_button = tk.Button(
            self,
            text="Рассчитать",
            command=self.calculate
        )
        self.save_raschet_button = tk.Button(
            self,
            text="Сохранить расчет",
            command=self.call_save_xlsx
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

    def run(self):
        self.draw_widgets()
        self.paste_init_data()
        self.mainloop()

    def paste_init_data(self):
        self.bending_moment_entry.delete(0, tk.END)
        self.bending_moment_entry.insert(0, self.parent.pls_pole_data["moment"])
        self.vert_force_entry.delete(0, tk.END)
        self.vert_force_entry.insert(0, self.parent.pls_pole_data["vert_force"])
        self.shear_force_entry.delete(0, tk.END)
        self.shear_force_entry.insert(0, self.parent.pls_pole_data["shear_force"])
        self.diam_svai_entry.delete(0, tk.END)
        self.diam_svai_entry.insert(0, self.parent.pls_pole_data["bot_diam"])
        try:
            self.thickness_svai_entry.delete(0, tk.END)
            self.thickness_svai_entry.insert(0, self.parent.thickness_svai)
            self.deepness_svai_entry.delete(0, tk.END)
            self.deepness_svai_entry.insert(0, self.parent.deepness_svai)
            self.height_svai_entry.delete(0, tk.END)
            self.height_svai_entry.insert(0, self.parent.height_svai)
            # self.is_initial_data_var_value=int(self.parent.is_initial_data)
            self.pole_type_combobox.set(self.parent.pole_type_foundation)
            self.coef_nadej_entry.delete(0, tk.END)
            self.coef_nadej_entry.insert(0, self.parent.coef_nadej)
            if self.parent.is_initial_data:
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
                self.is_initial_data_checkbutton.select()
                self.quantity_of_sloy_combobox.config(state="normal")
                self.quantity_of_sloy_combobox.set(self.parent.quantity_of_sloy)
                self.activate_sloy(None)
                self.ground_water_lvl_entry.delete(0, tk.END)
                self.ground_water_lvl_entry.insert(0, self.parent.ground_water_lvl)
                self.nomer_ige1_entry.delete(0, tk.END)
                self.nomer_ige1_entry.insert(0, self.parent.nomer_ige1)
                self.nomer_ige2_entry.delete(0, tk.END)
                self.nomer_ige2_entry.insert(0, self.parent.nomer_ige2)
                self.nomer_ige3_entry.delete(0, tk.END)
                self.nomer_ige3_entry.insert(0, self.parent.nomer_ige3)
                self.nomer_ige4_entry.delete(0, tk.END)
                self.nomer_ige4_entry.insert(0, self.parent.nomer_ige4)
                self.nomer_ige5_entry.delete(0, tk.END)
                self.nomer_ige5_entry.insert(0, self.parent.nomer_ige5)
                self.ground_type1_combobox.set(self.parent.ground_type1)
                self.ground_type2_combobox.set(self.parent.ground_type2)
                self.ground_type3_combobox.set(self.parent.ground_type3)
                self.ground_type4_combobox.set(self.parent.ground_type4)
                self.ground_type5_combobox.set(self.parent.ground_type5)
                self.ground_name1_entry.delete(0, tk.END)
                self.ground_name1_entry.insert(0, self.parent.ground_name1)
                self.ground_name2_entry.delete(0, tk.END)
                self.ground_name2_entry.insert(0, self.parent.ground_name2)
                self.ground_name3_entry.delete(0, tk.END)
                self.ground_name3_entry.insert(0, self.parent.ground_name3)
                self.ground_name4_entry.delete(0, tk.END)
                self.ground_name4_entry.insert(0, self.parent.ground_name4)
                self.ground_name5_entry.delete(0, tk.END)
                self.ground_name5_entry.insert(0, self.parent.ground_name5)
                self.verh_sloy1_entry.delete(0, tk.END)
                self.verh_sloy1_entry.insert(0, self.parent.verh_sloy1)
                self.verh_sloy2_entry.delete(0, tk.END)
                self.verh_sloy2_entry.insert(0, self.parent.verh_sloy2)
                self.verh_sloy3_entry.delete(0, tk.END)
                self.verh_sloy3_entry.insert(0, self.parent.verh_sloy3)
                self.verh_sloy4_entry.delete(0, tk.END)
                self.verh_sloy4_entry.insert(0, self.parent.verh_sloy4)
                self.verh_sloy5_entry.delete(0, tk.END)
                self.verh_sloy5_entry.insert(0, self.parent.verh_sloy5)
                self.nijn_sloy1_entry.delete(0, tk.END)
                self.nijn_sloy1_entry.insert(0, self.parent.nijn_sloy1)
                self.nijn_sloy2_entry.delete(0, tk.END)
                self.nijn_sloy2_entry.insert(0, self.parent.nijn_sloy2)
                self.nijn_sloy3_entry.delete(0, tk.END)
                self.nijn_sloy3_entry.insert(0, self.parent.nijn_sloy3)
                self.nijn_sloy4_entry.delete(0, tk.END)
                self.nijn_sloy4_entry.insert(0, self.parent.nijn_sloy4)
                self.nijn_sloy5_entry.delete(0, tk.END)
                self.nijn_sloy5_entry.insert(0, self.parent.nijn_sloy5)
                self.coef_poristosti1_entry.delete(0, tk.END)
                self.coef_poristosti1_entry.insert(0, self.parent.coef_poristosti1)
                self.coef_poristosti2_entry.delete(0, tk.END)
                self.coef_poristosti2_entry.insert(0, self.parent.coef_poristosti2)
                self.coef_poristosti3_entry.delete(0, tk.END)
                self.coef_poristosti3_entry.insert(0, self.parent.coef_poristosti3)
                self.coef_poristosti4_entry.delete(0, tk.END)
                self.coef_poristosti4_entry.insert(0, self.parent.coef_poristosti4)
                self.coef_poristosti5_entry.delete(0, tk.END)
                self.coef_poristosti5_entry.insert(0, self.parent.coef_poristosti5)
                self.udel_scep1_entry.delete(0, tk.END)
                self.udel_scep1_entry.insert(0, self.parent.udel_scep1)
                self.udel_scep2_entry.delete(0, tk.END)
                self.udel_scep2_entry.insert(0, self.parent.udel_scep2)
                self.udel_scep3_entry.delete(0, tk.END)
                self.udel_scep3_entry.insert(0, self.parent.udel_scep3)
                self.udel_scep4_entry.delete(0, tk.END)
                self.udel_scep4_entry.insert(0, self.parent.udel_scep4)
                self.udel_scep5_entry.delete(0, tk.END)
                self.udel_scep5_entry.insert(0, self.parent.udel_scep5)
                self.ugol_vn_tr1_entry.delete(0, tk.END)
                self.ugol_vn_tr1_entry.insert(0, self.parent.ugol_vn_tr1)
                self.ugol_vn_tr2_entry.delete(0, tk.END)
                self.ugol_vn_tr2_entry.insert(0, self.parent.ugol_vn_tr2)
                self.ugol_vn_tr3_entry.delete(0, tk.END)
                self.ugol_vn_tr3_entry.insert(0, self.parent.ugol_vn_tr3)
                self.ugol_vn_tr4_entry.delete(0, tk.END)
                self.ugol_vn_tr4_entry.insert(0, self.parent.ugol_vn_tr4)
                self.ugol_vn_tr5_entry.delete(0, tk.END)
                self.ugol_vn_tr5_entry.insert(0, self.parent.ugol_vn_tr5)
                self.ves_gr_prir1_entry.delete(0, tk.END)
                self.ves_gr_prir1_entry.insert(0, self.parent.ves_gr_prir1)
                self.ves_gr_prir2_entry.delete(0, tk.END)
                self.ves_gr_prir2_entry.insert(0, self.parent.ves_gr_prir2)
                self.ves_gr_prir3_entry.delete(0, tk.END)
                self.ves_gr_prir3_entry.insert(0, self.parent.ves_gr_prir3)
                self.ves_gr_prir4_entry.delete(0, tk.END)
                self.ves_gr_prir4_entry.insert(0, self.parent.ves_gr_prir4)
                self.ves_gr_prir5_entry.delete(0, tk.END)
                self.ves_gr_prir5_entry.insert(0, self.parent.ves_gr_prir5)
                self.def_mod1_entry.delete(0, tk.END)
                self.def_mod1_entry.insert(0, self.parent.def_mod1)
                self.def_mod2_entry.delete(0, tk.END)
                self.def_mod2_entry.insert(0, self.parent.def_mod2)
                self.def_mod3_entry.delete(0, tk.END)
                self.def_mod3_entry.insert(0, self.parent.def_mod3)
                self.def_mod4_entry.delete(0, tk.END)
                self.def_mod4_entry.insert(0, self.parent.def_mod4)
                self.def_mod5_entry.delete(0, tk.END)
                self.def_mod5_entry.insert(0, self.parent.def_mod5)
            else:
                self.typical_ground_combobox.set(self.parent.typical_ground)
                self.udel_sceplenie_entry.delete(0, tk.END)
                self.udel_sceplenie_entry.insert(0, self.parent.udel_sceplenie)
                self.ugol_vntr_trenia_entry.delete(0, tk.END)
                self.ugol_vntr_trenia_entry.insert(0, self.parent.ugol_vntr_trenia)
                self.ves_grunta_entry.delete(0, tk.END)
                self.ves_grunta_entry.insert(0, self.parent.ves_grunta)
                self.deform_module_entry.delete(0, tk.END)
                self.deform_module_entry.insert(0, self.parent.deform_module)
        except Exception as e:
            print("!INFO!: Сохраненные данные отсутствуют.")

    def draw_widgets(self):
        self.module_bg.place(x=10, y=0)
        self.back_to_main_window_button.place(x=15, y=3)
        self.bending_moment_label.place(x=70, y=4)
        self.bending_moment_entry.place(x=225, y=4)
        self.vert_force_label.place(x=310, y=4)
        self.vert_force_entry.place(x=457, y=4)
        self.shear_force_label.place(x=540, y=4)
        self.shear_force_entry.place(x=687, y=4)
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
        self.tech_param_label.place(x=40, y=301)
        self.coef_poristosti_label.place(x=25, y=324)
        self.coef_poristosti1_entry.place(x=244, y=324)
        self.coef_poristosti2_entry.place(x=338, y=324)
        self.coef_poristosti3_entry.place(x=432, y=324)
        self.coef_poristosti4_entry.place(x=526, y=324)
        self.coef_poristosti5_entry.place(x=620, y=324)
        self.udel_scep_label.place(x=25, y=347)
        self.udel_scep1_entry.place(x=244, y=347)
        self.udel_scep2_entry.place(x=338, y=347)
        self.udel_scep3_entry.place(x=432, y=347)
        self.udel_scep4_entry.place(x=526, y=347)
        self.udel_scep5_entry.place(x=620, y=347)
        self.ugol_vn_tr_label.place(x=25, y=370)
        self.ugol_vn_tr1_entry.place(x=244, y=370)
        self.ugol_vn_tr2_entry.place(x=338, y=370)
        self.ugol_vn_tr3_entry.place(x=432, y=370)
        self.ugol_vn_tr4_entry.place(x=526, y=370)
        self.ugol_vn_tr5_entry.place(x=620, y=370)
        self.ves_gr_prir_label.place(x=25, y=393)
        self.ves_gr_prir1_entry.place(x=244, y=393)
        self.ves_gr_prir2_entry.place(x=338, y=393)
        self.ves_gr_prir3_entry.place(x=432, y=393)
        self.ves_gr_prir4_entry.place(x=526, y=393)
        self.ves_gr_prir5_entry.place(x=620, y=393)
        self.def_mod_label.place(x=25, y=416)
        self.def_mod1_entry.place(x=244, y=416)
        self.def_mod2_entry.place(x=338, y=416)
        self.def_mod3_entry.place(x=432, y=416)
        self.def_mod4_entry.place(x=526, y=416)
        self.def_mod5_entry.place(x=620, y=416)
        self.pole_type_label.place(x=95, y=442)
        self.pole_type_combobox.place(x=244, y=442)
        self.coef_nadej_label.place(x=18, y=465)
        self.coef_nadej_entry.place(x=244, y=465)
        self.calculate_button.place(x=30, y=488)
        self.save_raschet_button.place(x=105, y=488)

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
        # ground_type_key = self.typical_ground_combobox.get().split()[0]
        # self.coef_usl_rab_entry.insert(0, coef_usl_rab_dict[ground_type_key])

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
            self.def_mod2_entry.delete(0, tk.END)
            self.def_mod3_entry.delete(0, tk.END)
            self.def_mod4_entry.delete(0, tk.END)
            self.def_mod5_entry.delete(0, tk.END)
            self.def_mod2_entry.config(state="disabled")
            self.def_mod3_entry.config(state="disabled")
            self.def_mod4_entry.config(state="disabled")
            self.def_mod5_entry.config(state="disabled")
            self.def_mod1_entry.config(state="normal")
            self.ground_water_lvl_entry.config(state="normal")
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
            self.def_mod3_entry.delete(0, tk.END)
            self.def_mod4_entry.delete(0, tk.END)
            self.def_mod5_entry.delete(0, tk.END)
            self.def_mod3_entry.config(state="disabled")
            self.def_mod4_entry.config(state="disabled")
            self.def_mod5_entry.config(state="disabled")
            self.def_mod1_entry.config(state="normal")
            self.def_mod2_entry.config(state="normal")
            self.ground_water_lvl_entry.config(state="normal")
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
            self.def_mod4_entry.delete(0, tk.END)
            self.def_mod5_entry.delete(0, tk.END)
            self.def_mod4_entry.config(state="disabled")
            self.def_mod5_entry.config(state="disabled")
            self.def_mod1_entry.config(state="normal")
            self.def_mod2_entry.config(state="normal")
            self.def_mod3_entry.config(state="normal")
            self.ground_water_lvl_entry.config(state="normal")
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
            self.def_mod5_entry.delete(0, tk.END)
            self.def_mod5_entry.config(state="disabled")
            self.def_mod1_entry.config(state="normal")
            self.def_mod2_entry.config(state="normal")
            self.def_mod3_entry.config(state="normal")
            self.def_mod4_entry.config(state="normal")
            self.ground_water_lvl_entry.config(state="normal")
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
            self.def_mod1_entry.config(state="normal")
            self.def_mod2_entry.config(state="normal")
            self.def_mod3_entry.config(state="normal")
            self.def_mod4_entry.config(state="normal")
            self.def_mod5_entry.config(state="normal")
            self.ground_water_lvl_entry.config(state="normal")

    def toggle_state(self):
        if self.is_initial_data_var.get():
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
        else:
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

    def insert_result(self):
        if self.is_initial_data_var.get():
            self.sr_znach_label.place(x=50, y=515)
            self.sr_udel_scep_label.place(x=19, y=538)
            self.sr_udel_scep_entry.place(x=244, y=538)
            self.sr_ugol_vn_tr_label.place(x=19, y=561)
            self.sr_ugol_vn_tr_entry.place(x=244, y=561)
            self.sr_ves_gr_ras_label.place(x=19, y=584)
            self.sr_ves_gr_ras_entry.place(x=244, y=584)
            self.sr_def_mod_label.place(x=19, y=607)
            self.sr_def_mod_entry.place(x=244, y=607)
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
                coef_nadej=self.coef_nadej_entry.get()
            )
        except Exception as e:
            mb.showinfo("Ошибка", f"{e}")
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
        self.result_label.place(x=410, y=439)
        self.ras_svai_pr_label.place(x=360, y=460)
        self.coef_isp_s245_label.place(x=350, y=483)
        self.coef_isp_s245_entry.place(x=572, y=483)
        self.coef_isp_s345_label.place(x=350, y=506)
        self.coef_isp_s345_entry.place(x=572, y=506)
        self.ras_gor_nagr_label.place(x=390, y=529)
        self.coef_isp_gor_label.place(x=350, y=552)
        self.coef_isp_gor_entry.place(x=572, y=552)
        self.ras_def_label.place(x=360, y=575)
        self.ugol_pov_label.place(x=350, y=598)
        self.ugol_pov_entry.place(x=572, y=598)

        self.insert_result()

    def call_save_xlsx(self):
        self.db.add_foundation_data(
            initial_data_id=self.parent.initial_data_id,
            thickness_svai=self.thickness_svai_entry.get(),
            deepness_svai=self.deepness_svai_entry.get(),
            height_svai=self.height_svai_entry.get(),
            is_initial_data=self.is_initial_data_var.get(),
            typical_ground=self.typical_ground_combobox.get(),
            udel_sceplenie=self.udel_sceplenie_entry.get(),
            ugol_vntr_trenia=self.ugol_vntr_trenia_entry.get(),
            ves_grunta=self.ves_grunta_entry.get(),
            deform_module=self.deform_module_entry.get(),
            ground_water_lvl=self.ground_water_lvl_entry.get(),
            quantity_of_sloy=self.quantity_of_sloy_combobox.get(),
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
            coef_nadej=self.coef_nadej_entry.get()
        )
        save_xlsx()

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