import re
import tkinter as tk


from tkinter.ttk import Combobox


from logic.utils import *
from logic.functionality.description_window import DescriptionWindow


class MultifacetedTower():
    def __init__(
        self,
        width=730,
        height=690,
        title="РПЗО многогранных опор",
        resizable=(False, False),
        bg="#FFFFFF",
        bg_pic=None,
        icon=None
    ):
        self.ice_thickness_table = {
            "I": "10",
            "II": "15",
            "III": "20",
            "IV": "25",
            "V": "30",
            "VI": "35",
            "VII": "40",
            "Особый": "Выше 40"
        }

        self.wind_table = {
            "I": "400",
            "II": "500",
            "III": "650",
            "IV": "800",
            "V": "1000",
            "VI": "1250",
            "VII": "1500",
            "Особый": "Выше 1500"
        }

        self.multifaceted = tk.Tk()
        self.multifaceted.title(title)
        self.multifaceted.geometry(f"{width}x{height}+500+150")
        self.multifaceted.resizable(resizable[0], resizable[1])
        self.multifaceted.config(bg=bg)
        if icon:
            self.multifaceted.iconbitmap(icon)

        self.project_info_bg = tk.Frame(
            self.multifaceted,
            width=710,
            height=100,
            borderwidth=2,
            relief="sunken"
        )

        self.initial_data_bg = tk.Frame(
            self.multifaceted,
            width=350,
            height=490,
            borderwidth=2,
            relief="sunken"
        )

        self.construction_data_bg = tk.Frame(
            self.multifaceted,
            width=350,
            height=490,
            borderwidth=2,
            relief="sunken"
        )

        self.media_bg = tk.Frame(
            self.multifaceted,
            width=350,
            height=75,
            borderwidth=2,
            relief="sunken"
        )

        self.generation_bg = tk.Frame(
            self.multifaceted,
            width=350,
            height=75,
            borderwidth=2,
            relief="sunken"
        )

        self.back_to_main_window_button = tk.Button(
            self.multifaceted,
            text="Назад в меню",
            command=self.multifaceted.destroy
        )

        self.project_info = tk.Label(self.multifaceted, text="Информация о проекте:", anchor="w")

        self.project_name = tk.Label(self.multifaceted, text="Название проекта", anchor="w")
        self.project_name_entry = tk.Entry(
            self.multifaceted,
            width=97,
            justify="left",
            relief="sunken",
            bd=2
        )

        self.project_code = tk.Label(self.multifaceted, text="Шифр проекта", anchor="w")
        self.project_code_entry = tk.Entry(
            self.multifaceted,
            width=35,
            justify="left",
            relief="sunken",
            bd=2
        )

        self.pole_code = tk.Label(self.multifaceted, text="Шифр опоры", anchor="w")
        self.pole_code_entry = tk.Entry(
            self.multifaceted,
            width=35,
            justify="left",
            relief="sunken",
            bd=2
        )

        self.pole_type = tk.Label(self.multifaceted, text="Тип опоры", anchor="w")
        self.pole_type_combobox = Combobox(
            self.multifaceted,
            values=("Анкерно-угловая", "Концевая", "Отпаечная", "Промежуточная"),
            width=20,
            validate="key"
        )
        self.pole_type_combobox["validatecommand"] = (
            self.pole_type_combobox.register(self.validate_pole_type), 
            "%P"
        )
        self.pole_type_combobox.bind("<KeyRelease>", lambda x: self.check_entries())

        self.developer = tk.Label(self.multifaceted, text="Разработал", anchor="w")
        self.developer_combobox = Combobox(self.multifaceted, values=("Мельситов", "Ушаков"))

        self.initial_data = tk.Label(self.multifaceted, text="Исходные данные:", width=46)
        
        self.voltage = tk.Label(self.multifaceted, text="Класс напряжения, кВ", anchor="e", width=33)
        self.voltage_combobox = Combobox(
            self.multifaceted,
            values=("6", "10", "35", "110", "220", "330", "500"),
            width=12,
            validate="key"
        )
        self.voltage_combobox["validatecommand"] = (
            self.voltage_combobox.register(self.validate_voltage),
            "%P"
        )
        self.voltage_combobox.bind("<KeyRelease>", lambda x: self.check_entries())

        self.area = tk.Label(self.multifaceted, text="Тип местности", anchor="e", width=33)
        self.area_combobox = Combobox(self.multifaceted, values=("А", "B", "C"), width=12)

        self.branches = tk.Label(self.multifaceted, text="Количество цепей, шт", anchor="e", width=33)
        self.branches_combobox = Combobox(
            self.multifaceted,
            values=("1", "2"),
            width=12,
            validate="key"
        )
        self.branches_combobox["validatecommand"] = (
            self.branches_combobox.register(self.validate_branches),
            "%P"
        )
        self.branches_combobox.bind("<KeyRelease>", lambda x: self.check_entries())

        self.wind_region = tk.Label(self.multifaceted, text="Район по ветру", anchor="e", width=33)
        self.wind_region_combobox = Combobox(
            self.multifaceted,
            values=("I", "II", "III", "IV", "V", "VI", "VII"),
            width=12
        )

        # Try to make autofill
        # self.wind_pressure_value = tk.StringVar()
        # self.wind_pressure_value.trace_variable("w", lambda *a: self.wind_pressure_value.set(self.wind_table[self.wind_region_combobox.get()]))

        self.wind_pressure = tk.Label(self.multifaceted, text="Ветровое давление, Па", anchor="e", width=33)
        self.wind_pressure_entry = tk.Entry(
            self.multifaceted,
            width=15,
            relief="sunken",
            bd=2,
            # textvariable=self.wind_pressure_value
        )

        self.ice_region = tk.Label(self.multifaceted, text="Район по гололеду", anchor="e", width=33)
        self.ice_region_combobox = Combobox(
            self.multifaceted,
            values=("I", "II", "III", "IV", "V", "VI", "VII"),
            width=12
        )

        self.ice_thickness = tk.Label(self.multifaceted, text="Толщина стенки гололеда, мм", anchor="e", width=33)
        self.ice_thickness_entry = tk.Entry(
            self.multifaceted,
            width=15,
            relief="sunken",
            bd=2
        )

        self.ice_wind_pressure = tk.Label(self.multifaceted, text="Ветровое давление при гололёде, Па", anchor="e", width=33)
        self.ice_wind_pressure_entry = tk.Entry(
            self.multifaceted,
            width=15,
            relief="sunken",
            bd=2
        )

        self.year_average_temp = tk.Label(self.multifaceted, text="Среднегодовая температура, °С", anchor="e", width=33)
        self.year_average_temp_entry = tk.Entry(
            self.multifaceted,
            width=15,
            relief="sunken",
            bd=2
        )

        self.min_temp = tk.Label(self.multifaceted, text="Минимальная температура, °С", anchor="e", width=33)
        self.min_temp_entry = tk.Entry(
            self.multifaceted,
            width=15,
            relief="sunken",
            bd=2
        )

        self.max_temp = tk.Label(self.multifaceted, text="Максимальная температура, °С", anchor="e", width=33)
        self.max_temp_entry = tk.Entry(
            self.multifaceted,
            width=15,
            relief="sunken",
            bd=2
        )

        self.ice_temp = tk.Label(self.multifaceted, text="Температура при гололеде, °С", anchor="e", width=33)
        self.ice_temp_entry = tk.Entry(
            self.multifaceted,
            width=15,
            relief="sunken",
            bd=2
        )

        self.wind_temp = tk.Label(self.multifaceted, text="Температура при ветре, °С", anchor="e", width=33)
        self.wind_temp_entry = tk.Entry(
            self.multifaceted,
            width=15,
            relief="sunken",
            bd=2
        )

        self.wind_reg_coef = tk.Label(self.multifaceted, text="Региональный коэффициент по ветру", anchor="e", width=33)
        self.wind_reg_coef_entry = tk.Entry(
            self.multifaceted,
            width=15,
            relief="sunken",
            bd=2
        )

        self.ice_reg_coef = tk.Label(self.multifaceted, text="Региональный коэффициент по гололеду", anchor="e", width=33)
        self.ice_reg_coef_entry = tk.Entry(
            self.multifaceted,
            width=15,
            relief="sunken",
            bd=2
        )

        self.wire_hesitation = tk.Label(self.multifaceted, text="Пляска проводов", anchor="e", width=33)
        self.wire_hesitation_combobox = Combobox(
            self.multifaceted,
            width=12,
            values=("Умеренная", "Частая и интенсивная"),
        )

        # self.seismicity = tk.Label(self.multifaceted, text="Сейсмичность площадки строительства", anchor="e", width=33)
        # self.seismicity_entry = tk.Entry(
        #     self.multifaceted,
        #     width=15,
        #     relief="sunken",
        #     bd=2
        # )

        # self.wire_pattern = re.compile(r"^[A-Za-zА-Яа-я]{0,10}\s\d{3}/\d{2}$")
        self.wire = tk.Label(self.multifaceted, text="Марка провода (формат ввода: АС 000/00)", anchor="e", width=33)
        self.wire_entry = tk.Entry(
            self.multifaceted,
            width=15,
            relief="sunken",
            bd=2
            # validate="key",
            # validatecommand=(self.multifaceted.register(self.validate_wire), "%P"),
            # invalidcommand=lambda: print("Провод неверного формата")
        )
        # self.wire_entry['validatecommand'] = (self.wire_entry.register(self.validate_wire),
        #                                       '%d', '%i', '%P', '%s', '%S', '%v', '%W')
        self.wire_entry.bind("<KeyRelease>", lambda x: self.check_entries())

        self.wire_tencion = tk.Label(self.multifaceted, text="Макс. напряжение в проводе, кгс/мм²", anchor="e", width=33)
        self.wire_tencion_entry = tk.Entry(
            self.multifaceted,
            width=15,
            relief="sunken",
            bd=2
        )

        self.ground_wire = tk.Label(self.multifaceted, text="Марка троса", anchor="e", width=33)
        self.ground_wire_entry = tk.Entry(
            self.multifaceted,
            width=15,
            relief="sunken",
            bd=2
        )

        self.media = tk.Label(
            self.multifaceted,
            text="Ссылки на медиа-файлы:",
            anchor="e",
            width=20
        )

        self.pole = tk.Label(
            self.multifaceted,
            text="Общий вид опоры",
            anchor="e",
            width=15
        )
        self.pole_entry = tk.Entry(
            self.multifaceted,
            width=30,
            relief="sunken",
            bd=2
        )
        self.pole_entry.bind("<KeyRelease>", lambda x: self.check_entries())
        self.browse_for_pole_button = tk.Button(
            self.multifaceted,
            text="Обзор",
            command=self.browse_for_pole
        )

        self.loads = tk.Label(
            self.multifaceted,
            text="Нагрузки",
            anchor="e",
            width=15
        )
        self.loads_entry = tk.Entry(
            self.multifaceted,
            width=30,
            relief="sunken",
            bd=2
        )
        self.loads_entry.bind("<KeyRelease>", lambda x: self.check_entries())
        self.browse_for_loads_button = tk.Button(
            self.multifaceted,
            text="Обзор",
            command=self.browse_for_loads
        )

        self.generation = tk.Label(
            self.multifaceted,
            text="Сгенерировать и сохранить отчет:",
            anchor="e",
            width=32
        )

        self.path_to_txt_label = tk.Label(
            self.multifaceted,
            text="Путь к txt файлу",
            anchor="e",
            width=13
        )
        self.path_to_txt_entry = tk.Entry(
            self.multifaceted,
            width=30,
            justify="left",
            relief="sunken",
            bd=2
        )
        self.path_to_txt_entry.bind("<FocusOut>", lambda x: self.check_entries())
        self.browse_txt_button = tk.Button(
            self.multifaceted,
            text="Обзор",
            command=self.browse_for_txt
        )
        self.browse_txt_button.bind("<Button-1>", lambda x: self.check_entries())
        
        self.generate_and_save_button = tk.Button(
            self.multifaceted, text="Сгенерировать отчет",
            command=self.generate_output,
            state="disabled"
        )

        self.generate_and_save_appendix_button = tk.Button(
            self.multifaceted, text="Сгенерировать приложение 1",
            command=self.generate_appendix_1,
            state="disabled"
        )

    def run(self):
        self.draw_widgets()
        self.multifaceted.mainloop()

    def draw_widgets(self):
        self.project_info_bg.place(x=10, y=0)
        self.initial_data_bg.place(x=10, y=110)
        self.construction_data_bg.place(x=370, y=110)
        self.media_bg.place(x=10, y=610)
        self.generation_bg.place(x=370, y=610)

        self.back_to_main_window_button.place(x=15, y=2)

        # self.project_info.place(x=300, y=3)

        self.project_name.place(x=15, y=27)
        self.project_name_entry.place(x=125, y=27)

        self.project_code.place(x=15, y=50)
        self.project_code_entry.place(x=125, y=50)

        self.pole_code.place(x=365, y=50)
        self.pole_code_entry.place(x=450, y=50)

        self.pole_type.place(x=365, y=73)
        self.pole_type_combobox.place(x=450, y=73)

        self.developer.place(x=15, y=73)
        self.developer_combobox.place(x=125, y=73)

        self.initial_data.place(x=25,y=112)

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

        self.year_average_temp.place(x=15,y=320)
        self.year_average_temp_entry.place(x=255,y=320)

        self.min_temp.place(x=15,y=343)
        self.min_temp_entry.place(x=255,y=343)

        self.max_temp.place(x=15,y=366)
        self.max_temp_entry.place(x=255,y=366)

        self.ice_temp.place(x=15,y=389)
        self.ice_temp_entry.place(x=255,y=389)

        self.wind_temp.place(x=15,y=411)
        self.wind_temp_entry.place(x=255,y=411)

        self.wind_reg_coef.place(x=15,y=433)
        self.wind_reg_coef_entry.place(x=255,y=433)

        self.ice_reg_coef.place(x=15,y=456)
        self.ice_reg_coef_entry.place(x=255,y=456)

        self.wire_hesitation.place(x=15,y=479)
        self.wire_hesitation_combobox.place(x=255,y=479)

        # self.seismicity.place(x=15,y=502)
        # self.seismicity_entry.place(x=255,y=502)

        self.wire.place(x=15,y=525)
        self.wire_entry.place(x=255,y=525)
        
        self.wire_tencion.place(x=15,y=548)
        self.wire_tencion_entry.place(x=255,y=548)

        self.ground_wire.place(x=15,y=571)
        self.ground_wire_entry.place(x=255,y=571)

        self.media.place(x=115,y=612)

        self.pole.place(x=13,y=634)
        self.pole_entry.place(x=125,y=634)
        self.browse_for_pole_button.place(x=312,y=631)

        self.loads.place(x=13,y=659)
        self.loads_entry.place(x=125,y=659)
        self.browse_for_loads_button.place(x=312,y=656)

        self.generation.place(x=412,y=612)

        self.path_to_txt_label.place(x=373,y=634)
        self.path_to_txt_entry.place(x=470,y=634)
        self.browse_txt_button.place(x=657,y=631)

        self.generate_and_save_button.place(x=400,y=656)
        self.generate_and_save_appendix_button.place(x=527,y=656)

    def browse_for_pole(self):
        self.file_path = make_path_png() 
        self.pole_entry.insert("insert", self.file_path)

    def browse_for_loads(self):
        self.file_path = make_multiple_path() 
        self.loads_entry.insert("insert", self.file_path)
        
    def browse_for_txt(self):
        self.file_path = make_path_txt() 
        self.path_to_txt_entry.insert("insert", self.file_path)
        if self.file_path:
            self.generate_and_save_appendix_button.config(state="active")

    def generate_output(self):
        self.result = put_data(
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
            seismicity=self.seismicity_entry.get(),
            wire=self.wire_entry.get(),
            wire_tencion=self.wire_tencion_entry.get(),
            ground_wire=self.ground_wire_entry.get(),
            sections=self.sections_entry.get(),
            pole_height=self.pole_height_entry.get(),
            pole_base=self.pole_base_entry.get(),
            pole_top=self.pole_top_entry.get(),
            fracture_1=self.fracture_1_entry.get(),
            fracture_2=self.fracture_2_entry.get(),
            height_davit_low=self.height_davit_low_entry.get(),
            height_davit_mid=self.height_davit_mid_entry.get(),
            height_davit_up=self.height_davit_up_entry.get(),
            length_davit_low_r=self.lenght_davit_low_r_entry.get(),
            length_davit_low_l=self.lenght_davit_low_l_entry.get(),
            length_davit_mid_r=self.lenght_davit_mid_r_entry.get(),
            length_davit_mid_l=self.lenght_davit_mid_l_entry.get(),
            length_davit_up_r=self.lenght_davit_up_r_entry.get(),
            length_davit_up_l=self.lenght_davit_up_l_entry.get(),
            length_davit_ground_r=self.length_davit_ground_r_entry.get(),
            length_davit_ground_l=self.length_davit_ground_l_entry.get(),
            wind_span=self.wind_span_entry.get(),
            weight_span=self.weight_span_entry.get(),
            pole=self.pole_entry.get(),
            loads_str=self.loads_entry.get(),
            path_to_txt=self.path_to_txt_entry.get()
        )

    def generate_appendix_1(self):
        self.appendix_1 = generate_appendix(
            path_to_txt=self.path_to_txt_entry.get()
        )

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

    def validate_float(self, value):
        return re.match(r"^\d*\.?\d*$", value) is not None
    
    def check_entries(self):
        if self.pole_type_combobox.get() and self.voltage_combobox.get()\
        and self.branches_combobox.get() and self.wire_entry.get() and\
        self.pole_height_entry.get():
            self.generate_and_save_button.configure(state="normal")
        else:
            self.generate_and_save_button.configure(state='disabled')

    # def open_description_window(self,
    #     width=400,
    #     height=200,
    #     title="Описание программы",
    #     resizable=(False, False),
    #     bg="#DFD8FF",
    #     icon=None
    #     ):
    #     DescriptionWindow(self.multifaceted,
    #         width,
    #         height,
    #         title,
    #         resizable,
    #         bg,
    #         icon)


if __name__ == "__main__":
    main_page = multifacetedTower(icon="logo.ico")
    main_page.open_description_window(icon="logo.ico")
    main_page.run()