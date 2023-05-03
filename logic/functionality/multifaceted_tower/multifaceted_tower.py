import re
import tkinter as tk


from tkinter.ttk import Combobox


from logic.functionality.multifaceted_tower.utils import *


class MultifacetedTower():
    def __init__(
        self,
        width=730,
        height=666,
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

        self.initial_data1_bg = tk.Frame(
            self.multifaceted,
            width=350,
            height=260,
            borderwidth=2,
            relief="sunken"
        )

        self.initial_data2_bg = tk.Frame(
            self.multifaceted,
            width=350,
            height=260,
            borderwidth=2,
            relief="sunken"
        )

        self.calculation_clarification_bg1 = tk.Frame(
            self.multifaceted,
            width=350,
            height=96,
            borderwidth=2,
            relief="sunken"
        )

        self.calculation_clarification_bg2 = tk.Frame(
            self.multifaceted,
            width=350,
            height=96,
            borderwidth=2,
            relief="sunken"
        )

        self.media_bg = tk.Frame(
            self.multifaceted,
            width=350,
            height=109,
            borderwidth=2,
            relief="sunken"
        )

        self.generation_bg = tk.Frame(
            self.multifaceted,
            width=350,
            height=109,
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

        self.initial_data = tk.Label(self.multifaceted, text="Исходные данные:", width=46, bg="#ffffff")
        
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

        self.oksn = tk.Label(self.multifaceted, text="Марка ОКСН", anchor="e", width=33)
        self.oksn_entry = tk.Entry(
            self.multifaceted,
            width=15,
            relief="sunken",
            bd=2
        )

        self.wind_span = tk.Label(self.multifaceted, text="Длина ветрового пролета, м", anchor="e", width=33)
        self.wind_span_entry = tk.Entry(
            self.multifaceted,
            width=15,
            relief="sunken",
            bd=2
        )

        self.weight_span = tk.Label(self.multifaceted, text="Длина весового пролета, м", anchor="e", width=33)
        self.weight_span_entry = tk.Entry(
            self.multifaceted,
            width=15,
            relief="sunken",
            bd=2
        )

        self.calculation_clarification = tk.Label(
            self.multifaceted, 
            text="Уточнения по расчету:", 
            width=46, 
            bg="#ffffff"
        )

        # self.is_stand_var = tk.IntVar()
        # self.is_stand = tk.Checkbutton(
        #     self.multifaceted,
        #     text="Опора с подставкой",
        #     variable=self.is_stand_var,
        #     onvalue=1,
        #     offvalue=0
        # )

        # self.is_plate_var = tk.IntVar()
        # self.is_plate = tk.Checkbutton(
        #     self.multifaceted,
        #     text="Почитан опорный фланец в PLS POLE",
        #     variable=self.is_plate_var
        # )

        # self.is_mont_schema_var = tk.IntVar()
        # self.is_mont_schema = tk.Checkbutton(
        #     self.multifaceted,
        #     text="Есть КМД",
        #     variable=self.is_mont_schema_var
        # )

        # self.is_ground_wire_davit_var = tk.IntVar()
        # self.is_ground_wire_davit = tk.Checkbutton(
        #     self.multifaceted,
        #     text="Есть тросовая(ые) траверса",
        #     variable=self.is_ground_wire_davit_var
        # )
        
        self.is_stand = tk.Label(
            self.multifaceted,
            text="Подставка",
            anchor="e", 
            width=30
        )
        self.is_stand_combobox = Combobox(
            self.multifaceted,
            values=("Да", "Нет"),
            width=16
        )

        self.is_plate = tk.Label(
            self.multifaceted,
            text="Почитан опорный фланец в PLS POLE",
            anchor="e", 
            width=30
        )
        self.is_plate_combobox = Combobox(
            self.multifaceted,
            values=("Да", "Нет"),
            width=16
        )

        self.is_ground_wire_davit = tk.Label(
            self.multifaceted,
            text="Есть тросовая(ые) траверса(ы)",
            anchor="e", 
            width=30
        )
        self.is_ground_wire_davit_combobox = Combobox(
            self.multifaceted,
            values=("Да", "Нет"),
            width=16
        )

        self.deflection = tk.Label(
            self.multifaceted,
            text="Отклонение (если выше нормы), мм",
            anchor="e", 
            width=30
        )
        self.deflection_entry = tk.Entry(
            self.multifaceted,
            width=19,
            relief="sunken",
            bd=2
        )

        self.wire_pos = tk.Label(self.multifaceted, text="Расположение проводов", anchor="e", width=30)
        self.wire_pos_combobox = Combobox(
            self.multifaceted,
            values=("Горизонтальное", "Вертикальное"),
            width=16
        )

        self.ground_wire_attachment = tk.Label(self.multifaceted, text="Крепление троса", anchor="e", width=30)
        self.ground_wire_attachment_combobox = Combobox(
            self.multifaceted,
            values=("Ниже верха опоры", "К верху опоры"),
            width=16
        )

        self.quantity_of_ground_wire = tk.Label(self.multifaceted, text="Количество тросов", anchor="e", width=30)
        self.quantity_of_ground_wire_combobox = Combobox(
            self.multifaceted,
            values=("1", "2"),
            width=16
        )

        self.media = tk.Label(
            self.multifaceted,
            text="Ссылки на медиа-файлы:",
            anchor="e",
            width=20,
            bg="#FFFFFF"
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

        self.pole_defl = tk.Label(
            self.multifaceted,
            text="Отклонение опоры",
            anchor="e",
            width=15
        )
        self.pole_defl_entry = tk.Entry(
            self.multifaceted,
            width=30,
            relief="sunken",
            bd=2
        )
        self.pole_defl_entry.bind("<KeyRelease>", lambda x: self.check_entries())
        self.browse_for_pole_defl_button = tk.Button(
            self.multifaceted,
            text="Обзор",
            command=self.browse_for_pole_defl
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

        self.is_mont_schema = tk.Label(
            self.multifaceted,
            text="Чертеж",
            anchor="e", 
            width=15
        )
        self.is_mont_schema_entry = tk.Entry(
            self.multifaceted,
            width=30,
            relief="sunken",
            bd=2
        )
        self.browse_for_mont_schema_button = tk.Button(
            self.multifaceted,
            text="Обзор",
            command=self.browse_for_mont_schema
        )

        self.generation = tk.Label(
            self.multifaceted,
            text="Сгенерировать и сохранить отчет:",
            anchor="e",
            width=32,
            bg="#FFFFFF"
        )

        self.path_to_txt_1_label = tk.Label(
            self.multifaceted,
            text="Путь к 1 отчету",
            anchor="e",
            width=13
        )
        self.path_to_txt_1_entry = tk.Entry(
            self.multifaceted,
            width=30,
            justify="left",
            relief="sunken",
            bd=2
        )
        self.path_to_txt_1_entry.bind("<FocusOut>", lambda x: self.check_entries())
        self.browse_txt_1_button = tk.Button(
            self.multifaceted,
            text="Обзор",
            command=self.browse_for_txt_1
        )
        self.browse_txt_1_button.bind("<Button-1>", lambda x: self.check_entries())

        self.path_to_txt_2_label = tk.Label(
            self.multifaceted,
            text="Путь ко 2 отчету",
            anchor="e",
            width=13
        )
        self.path_to_txt_2_entry = tk.Entry(
            self.multifaceted,
            width=30,
            justify="left",
            relief="sunken",
            bd=2
        )
        self.path_to_txt_2_entry.bind("<FocusOut>", lambda x: self.check_entries())
        self.browse_txt_2_button = tk.Button(
            self.multifaceted,
            text="Обзор",
            command=self.browse_for_txt_2
        )
        self.browse_txt_2_button.bind("<Button-1>", lambda x: self.check_entries())
        
        self.generate_and_save_button = tk.Button(
            self.multifaceted, text="Сгенерировать отчет",
            command=self.generate_output,
            state="disabled"
        )

        # self.generate_and_save_appendix_button = tk.Button(
        #     self.multifaceted, text="Сгенерировать приложение 1",
        #     command=self.generate_appendix_1,
        #     state="disabled"
        # )

    def run(self):
        self.draw_widgets()
        self.multifaceted.mainloop()

    def draw_widgets(self):
        self.project_info_bg.place(x=10, y=0)
        self.initial_data1_bg.place(x=10, y=133)
        self.initial_data2_bg.place(x=370, y=133)
        self.calculation_clarification_bg1.place(x=10, y=426),
        self.calculation_clarification_bg2.place(x=370, y=426)
        self.media_bg.place(x=10, y=552)
        self.generation_bg.place(x=370, y=552)

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

        # self.seismicity.place(x=15,y=502)
        # self.seismicity_entry.place(x=255,y=502)

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

        self.calculation_clarification.place(x=203, y=403)

        self.wire_pos.place(x=14, y=429)
        self.wire_pos_combobox.place(x=231, y=429)

        self.ground_wire_attachment.place(x=14, y=452)
        self.ground_wire_attachment_combobox.place(x=231, y=452)

        self.quantity_of_ground_wire.place(x=14, y=475)
        self.quantity_of_ground_wire_combobox.place(x=231, y=475)
        
        self.is_stand.place(x=375, y=429)
        self.is_stand_combobox.place(x=595, y=429)

        self.is_plate.place(x=375, y=452)
        self.is_plate_combobox.place(x=595, y=452)

        self.is_ground_wire_davit.place(x=14, y=498)
        self.is_ground_wire_davit_combobox.place(x=231, y=498)

        self.deflection.place(x=375, y=475)
        self.deflection_entry.place(x=595, y=475)

        self.media.place(x=115,y=531)

        self.pole.place(x=15,y=558)
        self.pole_entry.place(x=126,y=558)
        self.browse_for_pole_button.place(x=313,y=555)

        self.pole_defl.place(x=15,y=584)
        self.pole_defl_entry.place(x=126,y=584)
        self.browse_for_pole_defl_button.place(x=313,y=581)

        self.loads.place(x=15,y=610)
        self.loads_entry.place(x=126,y=610)
        self.browse_for_loads_button.place(x=313,y=607)

        self.is_mont_schema.place(x=15, y=633)
        self.is_mont_schema_entry.place(x=126, y=633)
        self.browse_for_mont_schema_button.place(x=313,y=633)

        self.generation.place(x=115,y=531)

        self.path_to_txt_1_label.place(x=373,y=558)
        self.path_to_txt_1_entry.place(x=470,y=558)
        self.browse_txt_1_button.place(x=657,y=555)

        self.path_to_txt_2_label.place(x=373,y=584)
        self.path_to_txt_2_entry.place(x=470,y=584)
        self.browse_txt_2_button.place(x=657,y=581)

        self.generate_and_save_button.place(x=400,y=608)
        # self.generate_and_save_appendix_button.place(x=527,y=631)

    def browse_for_pole(self):
        self.file_path = make_path_png()
        self.pole_entry.delete("0", "end")
        self.pole_entry.insert("insert", self.file_path)

    def browse_for_pole_defl(self):
        self.file_path = make_path_png()
        self.pole_defl_entry.delete("0", "end")
        self.pole_defl_entry.insert("insert", self.file_path)

    def browse_for_loads(self):
        self.file_path = make_multiple_path()
        self.loads_entry.delete("0", "end") 
        self.loads_entry.insert("insert", self.file_path)

    def browse_for_mont_schema(self):
        self.file_path = make_multiple_path()
        self.is_mont_schema_entry.delete("0", "end") 
        self.is_mont_schema_entry.insert("insert", self.file_path)
        
    def browse_for_txt_1(self):
        self.file_path = make_path_txt()
        self.path_to_txt_1_entry.delete("0", "end") 
        self.path_to_txt_1_entry.insert("insert", self.file_path)

    def browse_for_txt_2(self):
        self.file_path = make_path_txt()
        self.path_to_txt_2_entry.delete("0", "end") 
        self.path_to_txt_2_entry.insert("insert", self.file_path)

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
            wire=self.wire_entry.get(),
            wire_tencion=self.wire_tencion_entry.get(),
            ground_wire=self.ground_wire_entry.get(),
            oksn=self.oksn_entry.get(),
            wind_span=self.wind_span_entry.get(),
            weight_span=self.weight_span_entry.get(),
            is_stand=self.is_stand_combobox.get(),
            is_plate=self.is_plate_combobox.get(),
            is_ground_wire_davit=self.is_ground_wire_davit_combobox.get(),
            deflection=self.deflection_entry.get(),
            wire_pos=self.wire_pos_combobox.get(),
            ground_wire_attachment=self.ground_wire_attachment_combobox.get(),
            quantity_of_ground_wire=self.quantity_of_ground_wire_combobox.get(),
            pole=self.pole_entry.get(),
            pole_defl_pic=self.pole_defl_entry.get(),
            loads_str=self.loads_entry.get(),
            mont_schema=self.is_mont_schema_entry.get(),
            path_to_txt_1=self.path_to_txt_1_entry.get(),
            path_to_txt_2=self.path_to_txt_2_entry.get()
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
        and self.branches_combobox.get() and self.wire_entry.get():
            self.generate_and_save_button.configure(state="normal")
        else:
            self.generate_and_save_button.configure(state='disabled')


if __name__ == "__main__":
    main_page = MultifacetedTower(icon="logo.ico")
    main_page.run()