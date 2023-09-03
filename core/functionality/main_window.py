import tkinter as tk
from tkinter.ttk import Combobox
from tkinter import filedialog as fd

from PIL import Image, ImageTk

from core.functionality.lattice_tower import lattice_tower
from core.functionality.multifaceted_tower import multifaceted_tower
from core.functionality.foundation import foundation
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


class MainWindow(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("AmastA")
        self.geometry("730x500+500+150")
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

        self.back_to_main_window_button = tk.Button(
            self,
            image=self.back_icon,
            command=self.destroy
        )

        self.open_button = tk.Button(
            self,
            image=self.open_icon,
            command=self.open_initial_data
        )
        
        self.save_button = tk.Button(
            self,
            image=self.save_icon,
            command=self.save_initial_data
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
            height=260,
            borderwidth=2,
            relief="sunken"
        )

        self.initial_data2_bg = tk.Frame(
            self,
            width=350,
            height=260,
            borderwidth=2,
            relief="sunken"
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

        self.pole_type = tk.Label(self, text="Тип опоры", anchor="w")
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

        self.developer = tk.Label(self, text="Разработал", anchor="w")
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
        self.wind_region_combobox = Combobox(
            self,
            values=("I", "II", "III", "IV", "V", "VI", "VII"),
            width=12
        )

        self.wind_pressure = tk.Label(self, text="Ветровое давление, Па", anchor="e", width=33)
        self.wind_pressure_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2,
            # textvariable=self.wind_pressure_value
        )

        self.ice_region = tk.Label(self, text="Район по гололеду", anchor="e", width=33)
        self.ice_region_combobox = Combobox(
            self,
            values=("I", "II", "III", "IV", "V", "VI", "VII"),
            width=12
        )

        self.ice_thickness = tk.Label(self, text="Толщина стенки гололеда, мм", anchor="e", width=33)
        self.ice_thickness_entry = tk.Entry(
            self,
            width=15,
            relief="sunken",
            bd=2
        )

        self.ice_wind_pressure = tk.Label(self, text="Ветровое давление при гололёде, Па", anchor="e", width=33)
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
            text="Путь к 1 отчету",
            anchor="e",
            width=13
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
            text="Путь ко 2 отчету",
            anchor="e",
            width=13
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
            text="Решетчатая",
            command=self.go_to_lattice_calculation,
            width=12
        )
        self.multifaceted_button = tk.Button(
            self,
            text="Многогранная",
            command=self.go_to_multifaceted_calculation,
            width=12
        )
        self.foundation_calculation_button = tk.Button(
            self,
            text="Расчет фундамента",
            command=self.go_to_foundation_calculation
        )


    def run(self):
        self.draw_widgets()
        self.mainloop()

    def draw_widgets(self):
        self.project_info_bg.place(x=10, y=0)
        self.initial_data1_bg.place(x=10, y=133)
        self.initial_data2_bg.place(x=370, y=133)

        self.back_to_main_window_button.place(x=15, y=2)
        self.open_button.place(x=41, y=2)
        self.save_button.place(x=67, y=2)

        self.project_name.place(x=15, y=29)
        self.project_name_entry.place(x=125, y=29)

        self.project_code.place(x=15, y=52)
        self.project_code_entry.place(x=125, y=52)

        self.pole_code.place(x=365, y=52)
        self.pole_code_entry.place(x=450, y=52)

        self.pole_type.place(x=365, y=75)
        self.pole_type_combobox.place(x=450, y=75)

        self.developer.place(x=15, y=75)
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

        self.lattice_button.place(x=35, y=430)
        self.multifaceted_button.place(x=135, y=430)
        self.foundation_calculation_button.place(x=235, y=430)

    def go_to_lattice_calculation(self):
        lattice_window = lattice_tower.LatticeTower()
        lattice_window.run()

    def go_to_multifaceted_calculation(self):
        lattice_window = multifaceted_tower.MultifacetedTower()
        lattice_window.run()

    def go_to_foundation_calculation(self):
        foundation_calculation_window = foundation.FoundationCalculation(self)
        foundation_calculation_window.run()

    def save_initial_data(self):
        filename = fd.asksaveasfilename(
        defaultextension=".txt",
        filetypes=[("Text Files", "*.txt")]
        )
        if filename:
            with open(filename, "w") as file:
                file.writelines(
                    [self.project_name_entry.get() + "\n",
                    self.project_code_entry.get() + "\n",
                    self.pole_code_entry.get() + "\n",
                    self.pole_type_combobox.get() + "\n",
                    self.developer_combobox.get() + "\n",
                    self.voltage_combobox.get() + "\n",
                    self.area_combobox.get() + "\n",
                    self.branches_combobox.get() + "\n",
                    self.wind_region_combobox.get() + "\n",
                    self.wind_pressure_entry.get() + "\n",
                    self.ice_region_combobox.get() + "\n",
                    self.ice_thickness_entry.get() + "\n",
                    self.ice_wind_pressure_entry.get() + "\n",
                    self.year_average_temp_entry.get() + "\n",
                    self.min_temp_entry.get() + "\n",
                    self.max_temp_entry.get() + "\n",
                    self.ice_temp_entry.get() + "\n",
                    self.wind_temp_entry.get() + "\n",
                    self.wind_reg_coef_entry.get() + "\n",
                    self.ice_reg_coef_entry.get() + "\n",
                    self.wire_hesitation_combobox.get() + "\n",
                    self.wire_entry.get() + "\n",
                    self.wire_tencion_entry.get() + "\n",
                    self.ground_wire_entry.get() + "\n",
                    self.oksn_entry.get() + "\n",
                    self.wind_span_entry.get() + "\n",
                    self.weight_span_entry.get() + "\n"]
                )
    
    def open_initial_data(self):
        filename = fd.askopenfilename(filetypes=[("Text Files", "*.txt")])
        if filename:
            with open(filename, "r") as file:
                self.project_name_entry.delete(0, "end"),
                self.project_name_entry.insert(0, file.readline().rstrip("\n")),
                self.project_code_entry.delete(0, "end"),
                self.project_code_entry.insert(0, file.readline().rstrip("\n")),
                self.pole_code_entry.delete(0, "end"),
                self.pole_code_entry.insert(0, file.readline().rstrip("\n")),
                self.pole_type_combobox.set(file.readline().rstrip("\n")),
                self.developer_combobox.set(file.readline().rstrip("\n")),
                self.voltage_combobox.set(file.readline().rstrip("\n")),
                self.area_combobox.set(file.readline().rstrip("\n")),
                self.branches_combobox.set(file.readline().rstrip("\n")),
                self.wind_region_combobox.set(file.readline().rstrip("\n")),
                self.wind_pressure_entry.delete(0, "end"),
                self.wind_pressure_entry.insert(0, file.readline().rstrip("\n")),
                self.ice_region_combobox.set(file.readline().rstrip("\n")),
                self.ice_thickness_entry.delete(0, "end"),
                self.ice_thickness_entry.insert(0, file.readline().rstrip("\n")),
                self.ice_wind_pressure_entry.delete(0, "end"),
                self.ice_wind_pressure_entry.insert(0, file.readline().rstrip("\n")),
                self.year_average_temp_entry.delete(0, "end"),
                self.year_average_temp_entry.insert(0, file.readline().rstrip("\n")),
                self.min_temp_entry.delete(0, "end"),
                self.min_temp_entry.insert(0, file.readline().rstrip("\n")),
                self.max_temp_entry.delete(0, "end"),
                self.max_temp_entry.insert(0, file.readline().rstrip("\n")),
                self.ice_temp_entry.delete(0, "end"),
                self.ice_temp_entry.insert(0, file.readline().rstrip("\n")),
                self.wind_temp_entry.delete(0, "end"),
                self.wind_temp_entry.insert(0, file.readline().rstrip("\n")),
                self.wind_reg_coef_entry.delete(0, "end"),
                self.wind_reg_coef_entry.insert(0, file.readline().rstrip("\n")),
                self.ice_reg_coef_entry.delete(0, "end"),
                self.ice_reg_coef_entry.insert(0, file.readline().rstrip("\n")),
                self.wire_hesitation_combobox.set(file.readline().rstrip("\n")),
                self.wire_entry.delete(0, "end"),
                self.wire_entry.insert(0, file.readline().rstrip("\n")),
                self.wire_tencion_entry.delete(0, "end"),
                self.wire_tencion_entry.insert(0, file.readline().rstrip("\n")),
                self.ground_wire_entry.delete(0, "end"),
                self.ground_wire_entry.insert(0, file.readline().rstrip("\n")),
                self.oksn_entry.delete(0, "end"),
                self.oksn_entry.insert(0, file.readline().rstrip("\n")),
                self.wind_span_entry.delete(0, "end"),
                self.wind_span_entry.insert(0, file.readline().rstrip("\n")),
                self.weight_span_entry.delete(0, "end"),
                self.weight_span_entry.insert(0, file.readline().rstrip("\n")),

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