import os
import tkinter as tk
from tkinter.ttk import Combobox
from tkinter import filedialog as fd
from tkinter import messagebox as mb

from PIL import Image, ImageTk

from core.constants import (
    wind_table,
    ice_thickness_table,
    ice_wind_table
)
from core.db.db_connector import Database
from core.functionality.lattice_tower import lattice_tower
from core.functionality.multifaceted_tower import multifaceted_tower
from core.functionality.foundation import foundation
from core.functionality.ankernie_zakladnie import ankernie_zakladnie
from core.utils import (
    tempFile_back,
    tempFile_lupa,
    tempFile_save,
    make_path_txt,
    make_path_png,
    make_multiple_path,
    extract_foundation_loads_and_diam
)


class MainWindow(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("AmastA")
        self.geometry("730x520+500+150")
        self.resizable(False, False)
        self.config(bg="#FFFFFF")
        self.db = Database()

        self.back_icon = ImageTk.PhotoImage(
            file=tempFile_back
        )
        self.lupa_icon = ImageTk.PhotoImage(
            file=tempFile_lupa
        )
        self.save_icon = ImageTk.PhotoImage(
            file=tempFile_save
        )

        self.project_info_bg = tk.Frame(
            self,
            width=710,
            height=100,
            borderwidth=2,
            relief="sunken"
        )

        self.initial_data1_bg = tk.Frame(
            self,
            width=350,
            height=283,
            borderwidth=2,
            relief="sunken"
        )

        self.initial_data2_bg = tk.Frame(
            self,
            width=350,
            height=283,
            borderwidth=2,
            relief="sunken"
        )

        self.open_button = tk.Button(
            self,
            image=self.lupa_icon,
            command=self.find_initial_data
        )
        
        self.save_button = tk.Button(
            self,
            image=self.save_icon,
            command=self.save_initial_data
        )

        self.db_label = tk.Label(self,
            text="Введите шифр проекта и шифр опоры для поиска исходных данных проекта",
            anchor="w",
            font=("default", 10, "bold")
        )

        self.project_name = tk.Label(self, text="Название проекта", anchor="w")
        self.project_name_entry = tk.Entry(
            self,
            width=97,
            justify="left",
            relief="sunken",
            bd=2
        )

        self.project_code = tk.Label(self, text="Шифр проекта", anchor="w")
        self.project_code_entry = tk.Entry(
            self,
            width=35,
            justify="left",
            relief="sunken",
            bd=2
        )

        self.pole_code = tk.Label(self, text="Шифр опоры", anchor="w")
        self.pole_code_entry = tk.Entry(
            self,
            width=35,
            justify="left",
            relief="sunken",
            bd=2
        )

        self.pole_type = tk.Label(self, text="Тип опоры", anchor="e")
        self.pole_type_combobox = Combobox(
            self,
            values=("Анкерно-угловая", "Концевая", "Отпаечная", "Промежуточная"),
            width=20,
            validate="key"
        )
        self.pole_type_combobox["validatecommand"] = (
            self.pole_type_combobox.register(self.validate_pole_type), 
            "%P"
        )

        self.developer = tk.Label(self, text="Разработал", anchor="e")
        self.developer_combobox = Combobox(self, values=("Мельситов", "Ушаков"))

        self.initial_data = tk.Label(self, text="Исходные данные:", width=46, bg="#ffffff")
        
        self.voltage = tk.Label(self, text="Класс напряжения, кВ", anchor="e", width=33)
        self.voltage_combobox = Combobox(
            self,
            values=("6", "10", "35", "110", "220", "330", "500"),
            width=12,
            validate="key"
        )
        self.voltage_combobox["validatecommand"] = (
            self.voltage_combobox.register(self.validate_voltage),
            "%P"
        )

        self.area = tk.Label(self, text="Тип местности", anchor="e", width=33)
        self.area_combobox = Combobox(self, values=("А", "B", "C"), width=12)

        self.branches = tk.Label(self, text="Количество цепей, шт", anchor="e", width=33)
        self.branches_combobox = Combobox(
            self,
            values=("1", "2"),
            width=12,
            validate="key"
        )
        self.branches_combobox["validatecommand"] = (
            self.branches_combobox.register(self.validate_branches),
            "%P"
        )

        self.wind_region = tk.Label(self, text="Район по ветру", anchor="e", width=33)
        self.wind_region_variable = tk.StringVar()
        self.wind_region_combobox = Combobox(
            self,
            values=("I", "II", "III", "IV", "V", "VI", "VII"),
            width=12,
            textvariable=self.wind_region_variable,
            validate="key"
        )
        self.wind_region_combobox.bind("<<ComboboxSelected>>", self.paste_wind_pressure)

        self.wind_pressure = tk.Label(self, text="Ветровое давление, Па", anchor="e", width=33)
        self.wind_pressure_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            # textvariable=self.wind_pressure_value
        )

        self.ice_region = tk.Label(self, text="Район по гололеду", anchor="e", width=33)
        self.ice_reg_variable = tk.StringVar()
        self.ice_region_combobox = Combobox(
            self,
            values=("I", "II", "III", "IV", "V", "VI", "VII"),
            width=12,
            textvariable=self.ice_reg_variable,
            validate="key"
        )
        self.ice_region_combobox.bind("<<ComboboxSelected>>", self.paste_ice_thickness)

        self.ice_thickness = tk.Label(self, text="Толщина стенки гололеда, мм", anchor="e", width=33)
        self.ice_thickness_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2
        )

        self.ice_wind_pressure = tk.Label(self, text="Ветровое давление при гололёде, Па", anchor="e", width=33)
        self.ice_wind_pressure_variable = tk.StringVar()
        self.ice_wind_pressure_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2
        )

        self.wind_reg_coef = tk.Label(self, text="Региональный коэффициент по ветру", anchor="e", width=33)
        self.wind_reg_coef_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2
        )

        self.ice_reg_coef = tk.Label(self, text="Региональный коэффициент по гололеду", anchor="e", width=33)
        self.ice_reg_coef_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2
        )

        self.wire_hesitation = tk.Label(self, text="Пляска проводов", anchor="e", width=33)
        self.wire_hesitation_combobox = Combobox(
            self,
            width=12,
            values=("Умеренная", "Частая и интенсивная"),
        )

        self.year_average_temp = tk.Label(self, text="Среднегодовая температура, °С", anchor="e", width=33)
        self.year_average_temp_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2
        )

        self.min_temp = tk.Label(self, text="Минимальная температура, °С", anchor="e", width=33)
        self.min_temp_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2
        )

        self.max_temp = tk.Label(self, text="Максимальная температура, °С", anchor="e", width=33)
        self.max_temp_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2
        )

        self.ice_temp = tk.Label(self, text="Температура при гололеде, °С", anchor="e", width=33)
        self.ice_temp_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2
        )

        self.wind_temp = tk.Label(self, text="Температура при ветре, °С", anchor="e", width=33)
        self.wind_temp_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2
        )

        self.wire = tk.Label(self, text="Марка провода (формат ввода: АС 000/00)", anchor="e", width=33)
        self.wire_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2
        )

        self.wire_tencion = tk.Label(self, text="Макс. напряжение в проводе, кгс/мм²", anchor="e", width=33)
        self.wire_tencion_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2
        )

        self.ground_wire = tk.Label(self, text="Марка троса", anchor="e", width=33)
        self.ground_wire_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2
        )

        self.oksn = tk.Label(self, text="Марка ОКСН", anchor="e", width=33)
        self.oksn_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2
        )

        self.wind_span = tk.Label(self, text="Длина ветрового пролета, м", anchor="e", width=33)
        self.wind_span_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2
        )

        self.weight_span = tk.Label(self, text="Длина весового пролета, м", anchor="e", width=33)
        self.weight_span_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2
        )

        self.path_to_txt_1_label = tk.Label(
            self,
            text="1 отчет PLS POLE",
            anchor="e",
            width=13,
            bg="#FFFFFF"
        )
        self.path_to_txt_1_entry = tk.Entry(
            self,
            width=30,
            justify="left",
            relief="sunken",
            bd=2
        )
        self.browse_txt_1_button = tk.Button(
            self,
            text="Обзор",
            command=self.browse_for_txt_1
        )

        self.path_to_txt_2_label = tk.Label(
            self,
            text="2 отчет PLS POLE",
            anchor="e",
            width=13,
            bg="#FFFFFF"
        )
        self.path_to_txt_2_entry = tk.Entry(
            self,
            width=30,
            justify="left",
            relief="sunken",
            bd=2
        )
        self.browse_txt_2_button = tk.Button(
            self,
            text="Обзор",
            command=self.browse_for_txt_2
        )

        self.lattice_button = tk.Button(
            self,
            text="РПЗО решетчатая",
            command=self.go_to_lattice_calculation
        )
        self.multifaceted_button = tk.Button(
            self,
            text="РПЗО многогранная",
            command=self.go_to_multifaceted_calculation
        )
        self.foundation_calculation_button = tk.Button(
            self,
            text="Расчет фундамента",
            command=self.go_to_foundation_calculation
        )
        self.raschet_ankera_button = tk.Button(
            self,
            text="Расчет анкерных закладных",
            command=self.go_to_raschet_ankera
        )
        
        self.stand_flag = "Нет"
        self.is_stand_var = tk.IntVar()
        self.is_stand_checkbutton = tk.Checkbutton(
            self,
            text="Есть подставка",
            variable=self.is_stand_var,
            command=self.toggle_stand_state
        )
        self.plate_flag = "Нет"
        self.is_plate_var = tk.IntVar()
        self.is_plate_checkbutton = tk.Checkbutton(
            self,
            text="Есть посчитанный фланец",
            variable=self.is_plate_var,
            command=self.toggle_stand_state
        )

    def run(self):
        self.draw_widgets()
        self.mainloop()

    def draw_widgets(self):
        self.project_info_bg.place(x=10, y=0)
        self.initial_data1_bg.place(x=10, y=133)
        self.initial_data2_bg.place(x=370, y=133)
        self.db_label.place(x=120, y=2)
        self.open_button.place(x=660, y=21)
        self.save_button.place(x=686, y=21)
        self.project_name.place(x=15, y=52)
        self.project_name_entry.place(x=125, y=52)
        self.project_code.place(x=48, y=25)
        self.project_code_entry.place(x=138, y=25)
        self.pole_code.place(x=358, y=25)
        self.pole_code_entry.place(x=440, y=25)
        self.pole_type.place(x=380, y=75)
        self.pole_type_combobox.place(x=450, y=75)
        self.developer.place(x=49, y=75)
        self.developer_combobox.place(x=125, y=75)
        self.initial_data.place(x=200,y=112)
        self.voltage.place(x=15,y=136)
        self.voltage_combobox.place(x=255,y=136)
        self.area.place(x=15,y=159)
        self.area_combobox.place(x=255,y=159)
        self.branches.place(x=15,y=182)
        self.branches_combobox.place(x=255,y=182)
        self.wind_region.place(x=15,y=205)
        self.wind_region_combobox.place(x=255,y=205)
        self.wind_pressure.place(x=15,y=228)
        self.wind_pressure_entry.place(x=255,y=228)
        self.ice_region.place(x=15,y=251)
        self.ice_region_combobox.place(x=255,y=251)
        self.ice_thickness.place(x=15,y=274)
        self.ice_thickness_entry.place(x=255,y=274)
        self.ice_wind_pressure.place(x=15,y=297)
        self.ice_wind_pressure_entry.place(x=255,y=297)
        self.wind_reg_coef.place(x=15,y=320)
        self.wind_reg_coef_entry.place(x=255,y=320)
        self.ice_reg_coef.place(x=15,y=343)
        self.ice_reg_coef_entry.place(x=255,y=343)
        self.wire_hesitation.place(x=15,y=366)
        self.wire_hesitation_combobox.place(x=255,y=366)
        self.year_average_temp.place(x=375,y=136)
        self.year_average_temp_entry.place(x=615,y=136)
        self.min_temp.place(x=375,y=159)
        self.min_temp_entry.place(x=615,y=159)
        self.max_temp.place(x=375,y=182)
        self.max_temp_entry.place(x=615,y=182)
        self.ice_temp.place(x=375,y=205)
        self.ice_temp_entry.place(x=615,y=205)
        self.wind_temp.place(x=375,y=228)
        self.wind_temp_entry.place(x=615,y=228)
        self.wire.place(x=375,y=251)
        self.wire_entry.place(x=615,y=251)
        self.wire_tencion.place(x=375,y=274)
        self.wire_tencion_entry.place(x=615,y=274)
        self.ground_wire.place(x=375,y=297)
        self.ground_wire_entry.place(x=615,y=297)
        self.oksn.place(x=375,y=320)
        self.oksn_entry.place(x=615,y=320)
        self.wind_span.place(x=375,y=343)
        self.wind_span_entry.place(x=615,y=343)
        self.weight_span.place(x=375,y=366)
        self.weight_span_entry.place(x=615,y=366)
        self.is_stand_checkbutton.place(x=240, y=387)
        self.is_plate_checkbutton.place(x=535, y=387)
        self.path_to_txt_1_label.place(x=15, y=421)
        self.path_to_txt_1_entry.place(x=112, y=421)
        self.browse_txt_1_button.place(x=299, y=418)
        self.path_to_txt_2_label.place(x=373, y=421)
        self.path_to_txt_2_entry.place(x=470, y=421)
        self.browse_txt_2_button.place(x=657, y=418)
        self.lattice_button.place(x=35, y=449)
        self.multifaceted_button.place(x=153, y=449)
        self.foundation_calculation_button.place(x=285, y=449)
        self.raschet_ankera_button.place(x=410, y=449)

    def save_project_data(self):
        self.project_name=self.project_name_entry.get()
        self.project_code=self.project_code_entry.get()
        self.pole_code=self.pole_code_entry.get()
        self.pole_type=self.pole_type_combobox.get()
        self.developer=self.developer_combobox.get()
        self.voltage=self.voltage_combobox.get()
        self.area=self.area_combobox.get()
        self.branches=self.branches_combobox.get()
        self.wind_region=self.wind_region_combobox.get()
        self.wind_pressure=self.wind_pressure_entry.get()
        self.ice_region=self.ice_region_combobox.get()
        self.ice_thickness=self.ice_thickness_entry.get()
        self.ice_wind_pressure=self.ice_wind_pressure_entry.get()
        self.year_average_temp=self.year_average_temp_entry.get()
        self.min_temp=self.min_temp_entry.get()
        self.max_temp=self.max_temp_entry.get()
        self.ice_temp=self.ice_temp_entry.get()
        self.wind_temp=self.wind_temp_entry.get()
        self.wind_reg_coef=self.wind_reg_coef_entry.get()
        self.ice_reg_coef=self.ice_reg_coef_entry.get()
        self.wire_hesitation=self.wire_hesitation_combobox.get()
        self.wire=self.wire_entry.get()
        self.wire_tencion=self.wire_tencion_entry.get()
        self.ground_wire=self.ground_wire_entry.get()
        self.oksn=self.oksn_entry.get()
        self.wind_span=self.wind_span_entry.get()
        self.weight_span=self.weight_span_entry.get()
        self.is_stand=self.is_stand_var.get()
        self.is_plate=self.is_plate_var.get()
        self.pls_pole_data = extract_foundation_loads_and_diam(
            path_to_txt_1=self.path_to_txt_1_entry.get(),
            path_to_txt_2=self.path_to_txt_2_entry.get(),
            is_stand=self.stand_flag,
            is_plate=self.plate_flag
        )
    
    def toggle_stand_state(self):
        if self.is_stand_var.get() == 1:
            self.stand_flag = "Да"
        else:
            self.stand_flag = "Нет"
        if self.is_plate_var.get() ==  1:
            self.plate_flag = "Да"
        else:
            self.plate_flag = "Нет"
    
    def go_to_lattice_calculation(self):
        self.save_project_data()
        lattice_window = lattice_tower.LatticeTower()
        lattice_window.run()

    def go_to_multifaceted_calculation(self):
        self.save_project_data()
        lattice_window = multifaceted_tower.MultifacetedTower(self)
        self.withdraw()
        lattice_window.run()

    def go_to_foundation_calculation(self):
        self.save_project_data()
        foundation_calculation_window = foundation.FoundationCalculation(self)
        self.withdraw()
        foundation_calculation_window.run()

    def go_to_raschet_ankera(self):
        self.save_project_data()
        ankernie_zakladnie_window = ankernie_zakladnie.AnkernieZakladnie(self)
        self.withdraw()
        ankernie_zakladnie_window.run()

    def paste_wind_pressure(self, event):
        wind_pressure_key = self.wind_region_combobox.get()
        self.wind_pressure_entry.delete(0, tk.END)
        self.wind_pressure_entry.insert(0, wind_table[wind_pressure_key])
        self.ice_wind_pressure_entry.delete(0, tk.END)
        self.ice_wind_pressure_entry.insert(0, ice_wind_table[wind_pressure_key]) 

    def paste_ice_thickness(self, event):
        ice_thickness_key = self.ice_region_combobox.get()
        self.ice_thickness_entry.delete(0, tk.END)
        self.ice_thickness_entry.insert(0, ice_thickness_table[ice_thickness_key])
        
    def save_initial_data(self):
        txt_1_list = self.path_to_txt_1_entry.get().split("Удаленка")
        txt_2_list = self.path_to_txt_2_entry.get().split("Удаленка")
        self.db.add_initial_data(
            project_name=self.project_name_entry.get(),
            project_code=self.project_code_entry.get(),
            pole_code=self.pole_code_entry.get(),
            pole_type=self.pole_type_combobox.get(),
            developer=self.developer_combobox.get(),
            voltage=self.voltage_combobox.get(),
            area=self.area_combobox.get(),
            branches=self.branches_combobox.get(),
            wind_region=self.wind_region_combobox.get(),
            wind_pressure=self.wind_pressure_entry.get(),
            ice_region=self.ice_region_combobox.get(),
            ice_thickness=self.ice_thickness_entry.get(),
            ice_wind_pressure=self.ice_wind_pressure_entry.get(),
            year_average_temp=self.year_average_temp_entry.get(),
            min_temp=self.min_temp_entry.get(),
            max_temp=self.max_temp_entry.get(),
            ice_temp=self.ice_temp_entry.get(),
            wind_temp=self.wind_temp_entry.get(),
            wind_reg_coef=self.wind_reg_coef_entry.get(),
            ice_reg_coef=self.ice_reg_coef_entry.get(),
            wire_hesitation=self.wire_hesitation_combobox.get(),
            wire=self.wire_entry.get(),
            wire_tencion=self.wire_tencion_entry.get(),
            ground_wire=self.ground_wire_entry.get(),
            oksn=self.oksn_entry.get(),
            wind_span=self.wind_span_entry.get(),
            weight_span=self.weight_span_entry.get(),
            is_stand=self.is_stand_var.get(),
            is_plate=self.is_plate_var.get(),
            txt_1=txt_1_list[1],
            txt_2=txt_2_list[1]
        )
    
    def find_initial_data(self):
        initial_data = self.db.get_initial_data(
            project_code=self.project_code_entry.get().strip(),
            pole_code=self.pole_code_entry.get().strip()
        )
        if not initial_data:
            mb.showinfo("INFO", "Данные по проекту не найдены.")
        else:
            self.project_name_entry.delete(0, "end"),
            self.project_name_entry.insert(0, initial_data["project_name"]),
            self.project_code_entry.delete(0, "end"),
            self.project_code_entry.insert(0, initial_data["project_code"]),
            self.pole_code_entry.delete(0, "end"),
            self.pole_code_entry.insert(0, initial_data["pole_code"]),
            self.pole_type_combobox.set(initial_data["pole_type"]),
            self.developer_combobox.set(initial_data["developer"]),
            self.voltage_combobox.set(initial_data["voltage"]),
            self.area_combobox.set(initial_data["area_type"]),
            self.branches_combobox.set(initial_data["branch"]),
            self.wind_region_combobox.set(initial_data["wind_area"]),
            self.wind_pressure_entry.delete(0, "end"),
            self.wind_pressure_entry.insert(0, initial_data["wind_pressure"]),
            self.ice_region_combobox.set(initial_data["ice_area"]),
            self.ice_thickness_entry.delete(0, "end"),
            self.ice_thickness_entry.insert(0, initial_data["ice_thickness"]),
            self.ice_wind_pressure_entry.delete(0, "end"),
            self.ice_wind_pressure_entry.insert(0, initial_data["ice_wind_pressure"]),
            self.year_average_temp_entry.delete(0, "end"),
            self.year_average_temp_entry.insert(0, initial_data["avg_temp"]),
            self.min_temp_entry.delete(0, "end"),
            self.min_temp_entry.insert(0, initial_data["min_temp"]),
            self.max_temp_entry.delete(0, "end"),
            self.max_temp_entry.insert(0, initial_data["max_temp"]),
            self.ice_temp_entry.delete(0, "end"),
            self.ice_temp_entry.insert(0, initial_data["ice_temp"]),
            self.wind_temp_entry.delete(0, "end"),
            self.wind_temp_entry.insert(0, initial_data["wind_temp"]),
            self.wind_reg_coef_entry.delete(0, "end"),
            self.wind_reg_coef_entry.insert(0, initial_data["wind_reg_coef"]),
            self.ice_reg_coef_entry.delete(0, "end"),
            self.ice_reg_coef_entry.insert(0, initial_data["ice_reg_coef"]),
            self.wire_hesitation_combobox.set(initial_data["wire_hesitation"]),
            self.wire_entry.delete(0, "end"),
            self.wire_entry.insert(0, initial_data["wire"]),
            self.wire_tencion_entry.delete(0, "end"),
            self.wire_tencion_entry.insert(0, initial_data["wire_tension"]),
            self.ground_wire_entry.delete(0, "end"),
            self.ground_wire_entry.insert(0, initial_data["ground_wire"]),
            self.oksn_entry.delete(0, "end"),
            self.oksn_entry.insert(0, initial_data["oksn"]),
            self.wind_span_entry.delete(0, "end"),
            self.wind_span_entry.insert(0, initial_data["wind_span"]),
            self.weight_span_entry.delete(0, "end"),
            self.weight_span_entry.insert(0, initial_data["weight_span"]),
            self.is_stand_var=tk.IntVar(value=int(initial_data["is_stand"]))
            if self.is_stand_var.get(): self.is_stand_checkbutton.select()
            self.is_plate_var=tk.IntVar(value=int(initial_data["is_stand"]))
            if self.is_plate_var.get():\
                self.is_plate_checkbutton.select()
            self.path_to_txt_1_entry.delete(0, "end")
            first_part_of_path = os.path.abspath("lupa.png").split("\Удаленка")[0]
            first_part_of_path = "/".join(first_part_of_path.split("\\"))
            self.path_to_txt_1_entry.insert(
                0,
                first_part_of_path + "/Удаленка" + initial_data["txt_1"]
            )
            self.path_to_txt_2_entry.delete(0, "end")
            self.path_to_txt_2_entry.insert(
                0,
                first_part_of_path + "/Удаленка" + initial_data["txt_2"]
            )

    def browse_for_txt_1(self):
        self.file_path = make_path_txt()
        self.path_to_txt_1_entry.delete("0", "end") 
        self.path_to_txt_1_entry.insert("insert", self.file_path)

    def browse_for_txt_2(self):
        self.file_path = make_path_txt()
        self.path_to_txt_2_entry.delete("0", "end") 
        self.path_to_txt_2_entry.insert("insert", self.file_path)
    
    def validate_pole_type(self, value):
        if value in ["Анкерно-угловая", "Концевая", "Отпаечная", "Промежуточная"]:
            return True
        return False
    
    def validate_voltage(self, value):
        return value.isdigit()
    
    def validate_branches(self, value):
        if value in ["1", "2"]:
            return True
        return False

if __name__ == "__main__":
    main_page = MainWindow()
    main_page.run()