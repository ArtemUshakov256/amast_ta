import tkinter as tk


from tkinter.ttk import Combobox


from ..utils import *
from .description_window import DescriptionWindow


class MainWindow:
    def __init__(
        self,
        width=730,
        height=695,
        title="Amast TA",
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

        self.root = tk.Tk()
        self.root.title(title)
        self.root.geometry(f"{width}x{height}+500+150")
        self.root.resizable(resizable[0], resizable[1])
        self.root.config(bg=bg)
        if icon:
            self.root.iconbitmap(icon)

        self.project_info_bg = tk.Frame(
            self.root,
            width=710,
            height=100,
            borderwidth=2,
            relief="sunken"
        )

        self.initial_data_bg = tk.Frame(
            self.root,
            width=350,
            height=490,
            borderwidth=2,
            relief="sunken"
        )

        self.construction_data_bg = tk.Frame(
            self.root,
            width=350,
            height=490,
            borderwidth=2,
            relief="sunken"
        )

        self.media_bg = tk.Frame(
            self.root,
            width=350,
            height=75,
            borderwidth=2,
            relief="sunken"
        )

        self.convertation_bg = tk.Frame(
            self.root,
            width=350,
            height=75,
            borderwidth=2,
            relief="sunken"
        )

        self.project_info = tk.Label(self.root, text="Информация о проекте:", anchor="w")

        self.project_name = tk.Label(self.root, text="Название проекта", anchor="w")
        self.project_name_entry = tk.Entry(
            self.root,
            width=97,
            justify="left",
            relief="sunken",
            bd=2
        )

        self.project_code = tk.Label(self.root, text="Шифр проекта", anchor="w")
        self.project_code_entry = tk.Entry(
            self.root,
            width=40,
            justify="left",
            relief="sunken",
            bd=2
        )

        self.developer = tk.Label(self.root, text="Разработал", anchor="w")
        self.developer_combobox = Combobox(self.root, values=("Мельситов", "Ушаков"))

        self.initial_data = tk.Label(self.root, text="Исходные данные:", width=46)
        
        self.voltage = tk.Label(self.root, text="Класс напряжения, кВ", anchor="e", width=33)
        self.voltage_combobox = Combobox(
            self.root,
            values=("6", "10", "35", "110", "220", "330", "500"),
            width=12
        )

        self.area = tk.Label(self.root, text="Тип местности", anchor="e", width=33)
        self.area_combobox = Combobox(self.root, values=("А", "B", "C"), width=12)

        self.branches = tk.Label(self.root, text="Количество цепей, шт", anchor="e", width=33)
        self.branches_combobox = Combobox(
            self.root,
            values=("1", "2"),
            width=12
        )

        self.wind_region = tk.Label(self.root, text="Район по ветру", anchor="e", width=33)
        self.wind_region_combobox = Combobox(
            self.root,
            values=("I", "II", "II", "III", "IV", "V", "VI", "VII"),
            width=12
        )

        # Try to make autofill
        # self.wind_pressure_value = tk.StringVar()
        # self.wind_pressure_value.trace_variable("w", lambda *a: self.wind_pressure_value.set(self.wind_table[self.wind_region_combobox.get()]))

        self.wind_pressure = tk.Label(self.root, text="Ветровое давление, Па", anchor="e", width=33)
        self.wind_pressure_entry = tk.Entry(
            self.root,
            width=15,
            relief="sunken",
            bd=2,
            # textvariable=self.wind_pressure_value
        )

        self.ice_region = tk.Label(self.root, text="Район по гололеду", anchor="e", width=33)
        self.ice_region_combobox = Combobox(
            self.root,
            values=("I", "II", "II", "III", "IV", "V", "VI", "VII"),
            width=12
        )

        self.ice_thickness = tk.Label(self.root, text="Толщина стенки гололеда, мм", anchor="e", width=33)
        self.ice_thickness_entry = tk.Entry(
            self.root,
            width=15,
            relief="sunken",
            bd=2
        )

        self.ice_wind_pressure = tk.Label(self.root, text="Ветровое давление при гололёде, Па", anchor="e", width=33)
        self.ice_wind_pressure_entry = tk.Entry(
            self.root,
            width=15,
            relief="sunken",
            bd=2
        )

        self.year_average_temp = tk.Label(self.root, text="Среднегодовая температура, °С", anchor="e", width=33)
        self.year_average_temp_entry = tk.Entry(
            self.root,
            width=15,
            relief="sunken",
            bd=2
        )

        self.min_temp = tk.Label(self.root, text="Минимальная температура, °С", anchor="e", width=33)
        self.min_temp_entry = tk.Entry(
            self.root,
            width=15,
            relief="sunken",
            bd=2
        )

        self.max_temp = tk.Label(self.root, text="Максимальная температура, °С", anchor="e", width=33)
        self.max_temp_entry = tk.Entry(
            self.root,
            width=15,
            relief="sunken",
            bd=2
        )

        self.ice_temp = tk.Label(self.root, text="Температура при гололеде, °С", anchor="e", width=33)
        self.ice_temp_entry = tk.Entry(
            self.root,
            width=15,
            relief="sunken",
            bd=2
        )

        self.wind_temp = tk.Label(self.root, text="Температура при ветре, °С", anchor="e", width=33)
        self.wind_temp_entry = tk.Entry(
            self.root,
            width=15,
            relief="sunken",
            bd=2
        )

        self.wind_reg_coef = tk.Label(self.root, text="Региональный коэффициент по ветру", anchor="e", width=33)
        self.wind_reg_coef_entry = tk.Entry(
            self.root,
            width=15,
            relief="sunken",
            bd=2
        )

        self.ice_reg_coef = tk.Label(self.root, text="Региональный коэффициент по гололеду", anchor="e", width=33)
        self.ice_reg_coef_entry = tk.Entry(
            self.root,
            width=15,
            relief="sunken",
            bd=2
        )

        self.wire_hesitation = tk.Label(self.root, text="Пляска проводов", anchor="e", width=33)
        self.wire_hesitation_entry = tk.Entry(
            self.root,
            width=15,
            relief="sunken",
            bd=2
        )

        self.seismicity = tk.Label(self.root, text="Сейсмичность площадки строительства", anchor="e", width=33)
        self.seismicity_entry = tk.Entry(
            self.root,
            width=15,
            relief="sunken",
            bd=2
        )

        self.wire = tk.Label(self.root, text="Марка провода", anchor="e", width=33)
        self.wire_entry = tk.Entry(
            self.root,
            width=15,
            relief="sunken",
            bd=2
        )

        self.wire_tencion = tk.Label(self.root, text="Макс. напряжение в проводе, кгс/мм²", anchor="e", width=33)
        self.wire_tencion_entry = tk.Entry(
            self.root,
            width=15,
            relief="sunken",
            bd=2
        )

        self.ground_wire = tk.Label(self.root, text="Марка троса", anchor="e", width=33)
        self.ground_wire_entry = tk.Entry(
            self.root,
            width=15,
            relief="sunken",
            bd=2
        )

        self.construction_data = tk.Label(self.root, text="Конструктивные решения:", anchor="e", width=35)

        self.sections = tk.Label(self.root, text="Количество секций, шт", anchor="e", width=33)
        self.sections_entry = tk.Entry(
            self.root,
            width=15,
            relief="sunken",
            bd=2
        )

        self.pole_height = tk.Label(self.root, text="Высота опоры, м", anchor="e", width=33)
        self.pole_height_entry = tk.Entry(
            self.root,
            width=15,
            relief="sunken",
            bd=2
        )

        self.pole_base = tk.Label(self.root, text="Длина стороны основания (база), м", anchor="e", width=33)
        self.pole_base_entry = tk.Entry(
            self.root,
            width=15,
            relief="sunken",
            bd=2
        )

        self.pole_top = tk.Label(self.root, text="Длина стороны верха, м", anchor="e", width=33)
        self.pole_top_entry = tk.Entry(
            self.root,
            width=15,
            relief="sunken",
            bd=2
        )

        self.pole_base = tk.Label(self.root, text="Длина стороны основания (база), м", anchor="e", width=33)
        self.pole_base_entry = tk.Entry(
            self.root,
            width=15,
            relief="sunken",
            bd=2
        )

        self.fracture_1 = tk.Label(self.root, text="Перелом поясов 1, м *", anchor="e", width=33)
        self.fracture_1_entry = tk.Entry(
            self.root,
            width=15,
            relief="sunken",
            bd=2
        )

        self.fracture_2 = tk.Label(self.root, text="Перелом поясов 2, м *", anchor="e", width=33)
        self.fracture_2_entry = tk.Entry(
            self.root,
            width=15,
            relief="sunken",
            bd=2
        )

        self.height_davit_low = tk.Label(self.root, text="Высота крепления нижн. траверс, м", anchor="e", width=33)
        self.height_davit_low_entry = tk.Entry(
            self.root,
            width=15,
            relief="sunken",
            bd=2
        )

        self.height_davit_mid = tk.Label(self.root, text="Высота крепления сред. траверс, м **", anchor="e", width=33)
        self.height_davit_mid_entry = tk.Entry(
            self.root,
            width=15,
            relief="sunken",
            bd=2
        )

        self.height_davit_up = tk.Label(self.root, text="Высота крепления верх. траверс, м", anchor="e", width=33)
        self.height_davit_up_entry = tk.Entry(
            self.root,
            width=15,
            relief="sunken",
            bd=2
        )

        self.length_davit_low_r = tk.Label(self.root, text="Длина нижн. траверсы правой, м", anchor="e", width=33)
        self.lenght_davit_low_r_entry = tk.Entry(
            self.root,
            width=15,
            relief="sunken",
            bd=2
        )

        self.length_davit_low_l = tk.Label(self.root, text="Длина нижн. траверсы левой, м", anchor="e", width=33)
        self.lenght_davit_low_l_entry = tk.Entry(
            self.root,
            width=15,
            relief="sunken",
            bd=2
        )

        self.length_davit_mid_r = tk.Label(self.root, text="Длина сред. траверсы правой, м **", anchor="e", width=33)
        self.lenght_davit_mid_r_entry = tk.Entry(
            self.root,
            width=15,
            relief="sunken",
            bd=2
        )

        self.length_davit_mid_l = tk.Label(self.root, text="Длина сред. траверсы левой, м **", anchor="e", width=33)
        self.lenght_davit_mid_l_entry = tk.Entry(
            self.root,
            width=15,
            relief="sunken",
            bd=2
        )

        self.length_davit_up_r = tk.Label(self.root, text="Длина верх. траверсы правой, м", anchor="e", width=33)
        self.lenght_davit_up_r_entry = tk.Entry(
            self.root,
            width=15,
            relief="sunken",
            bd=2
        )

        self.length_davit_up_l = tk.Label(self.root, text="Длина верх. траверсы левой, м", anchor="e", width=33)
        self.lenght_davit_up_l_entry = tk.Entry(
            self.root,
            width=15,
            relief="sunken",
            bd=2
        )

        self.span_data = tk.Label(self.root, text="Расчетные пролеты:", anchor="e", width=28)

        self.wind_span = tk.Label(self.root, text="Длина ветрового пролета, м", anchor="e", width=33)
        self.wind_span_entry = tk.Entry(
            self.root,
            width=15,
            relief="sunken",
            bd=2
        )

        self.weight_span = tk.Label(self.root, text="Длина весового пролета, м", anchor="e", width=33)
        self.weight_span_entry = tk.Entry(
            self.root,
            width=15,
            relief="sunken",
            bd=2
        )

        self.notes = tk.Label(
            self.root,
            text='* - уровни изменения "конусности"/наклона поясов\n'\
                '** - если опора одноцепная, оставить поля пустыми',
            anchor="w",
            width=45
        )

        self.media = tk.Label(
            self.root,
            text="Ссылки на медиа-файлы:",
            anchor="e",
            width=20
        )

        self.pole = tk.Label(
            self.root,
            text="Общий вид опоры",
            anchor="e",
            width=15
        )

        self.loads = tk.Label(
            self.root,
            text="Нагрузки",
            anchor="e",
            width=15
        )

        # self.entry = tk.Entry(
        #     self.root,
        #     width=30,
        #     justify="left",
        #     relief="sunken",
        #     bd=2
        # )
        # self.path_label = tk.Label(self.root, text="Путь к файлу", bg="#FFFFFF")
        # self.browse_button = tk.Button(self.root, text="Обзор", command=self.callback)
        # self.convert_and_save_button = tk.Button(
        #     self.root, text="Конвертировать и сохранить",
        #     command=self.convert_txt_to_xlsx
        # )

        # self.convertation_result = tk.Label(self.root, text=f"{self.result}", bg="#FFFFFF")

    def run(self):
        self.draw_widgets()
        self.root.mainloop()

    def draw_widgets(self):
        self.project_info_bg.place(x=10, y=0)
        self.initial_data_bg.place(x=10, y=110)
        self.construction_data_bg.place(x=370, y=110)
        self.media_bg.place(x=10, y=610)
        self.convertation_bg.place(x=370, y=610)

        self.project_info.place(x=15, y=3)

        self.project_name.place(x=15, y=26)
        self.project_name_entry.place(x=125, y=26)

        self.project_code.place(x=15, y=49)
        self.project_code_entry.place(x=125, y=49)

        self.developer.place(x=15, y=72)
        self.developer_combobox.place(x=125, y=72)

        self.initial_data.place(x=25,y=113)

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
        self.wire_hesitation_entry.place(x=255,y=479)

        self.seismicity.place(x=15,y=502)
        self.seismicity_entry.place(x=255,y=502)

        self.wire.place(x=15,y=525)
        self.wire_entry.place(x=255,y=525)
        
        self.wire_tencion.place(x=15,y=548)
        self.wire_tencion_entry.place(x=255,y=548)

        self.ground_wire.place(x=15,y=571)
        self.ground_wire_entry.place(x=255,y=571)

        self.construction_data.place(x=375,y=113)

        self.sections.place(x=375,y=136)
        self.sections_entry.place(x=615,y=136)

        self.pole_height.place(x=375,y=159)
        self.pole_height_entry.place(x=615,y=159)

        self.pole_base.place(x=375,y=182)
        self.pole_base_entry.place(x=615,y=182)

        self.pole_top.place(x=375,y=205)
        self.pole_top_entry.place(x=615,y=205)

        self.fracture_1.place(x=375,y=228)
        self.fracture_1_entry.place(x=615,y=228)

        self.fracture_2.place(x=375,y=251)
        self.fracture_2_entry.place(x=615,y=251)

        self.height_davit_low.place(x=375,y=274)
        self.height_davit_low_entry.place(x=615,y=274)

        self.height_davit_mid.place(x=375,y=297)
        self.height_davit_mid_entry.place(x=615,y=297)

        self.height_davit_up.place(x=375,y=320)
        self.height_davit_up_entry.place(x=615,y=320)

        self.length_davit_low_r.place(x=375,y=343)
        self.lenght_davit_low_r_entry.place(x=615,y=343)

        self.length_davit_low_l.place(x=375,y=366)
        self.lenght_davit_low_l_entry.place(x=615,y=366)

        self.length_davit_mid_r.place(x=375,y=389)
        self.lenght_davit_mid_r_entry.place(x=615,y=389)

        self.length_davit_mid_l.place(x=375,y=412)
        self.lenght_davit_mid_l_entry.place(x=615,y=412)

        self.length_davit_up_r.place(x=375,y=435)
        self.lenght_davit_up_r_entry.place(x=615,y=435)

        self.length_davit_up_l.place(x=375,y=458)
        self.lenght_davit_up_l_entry.place(x=615,y=458)

        self.span_data.place(x=400,y=481)

        self.wind_span.place(x=375,y=504)
        self.wind_span_entry.place(x=615,y=504)

        self.weight_span.place(x=375,y=527)
        self.weight_span_entry.place(x=615,y=527)

        self.notes.place(x=375,y=560)

        self.media.place(x=115,y=613)

        self.pole.place(x=13,y=636)

        self.loads.place(x=13,y=659)

        # self.path_label.place(x=0, y=160)
        # self.entry.place(x=80, y=160)
        # self.browse_button.place(x=270, y=160)
        # self.convert_and_save_button.place(x=87, y=190)

        # self.convertation_result.place(x=120, y=230)

    def callback(self):
        self.file_path = make_path() 
        self.entry.insert("insert", self.file_path)

    def get_input_data(self):
        self.project_name_data = self.project_name_entry.get()
        self.project_code_data = self.project_code_entry.get()
        self.developer_data = self.developer_combobox.get()
        self.voltage_data = self.voltage_combobox.get()
        self.area_data = self.area_combobox.get()
        self.branches_data = self.branches_combobox.get()
        self.wind_region_data = self.wind_region_combobox.get()
        self.wind_pressure_data = self.wind_pressure_entry.get()
        self.ice_region_data = self.ice_region_combobox.get()
        self.ice_thickness_data = self.ice_thickness_entry.get()
        self.ice_wind_pressure_data = self.ice_wind_pressure_entry.get()
        self.year_average_temp_data = self.year_average_temp_entry.get()
        self.min_temp_data = self.min_temp_entry.get()
        self.max_temp_data = self.max_temp_entry.get()
        self.ice_temp_data = self.ice_temp_entry.get()
        self.wind_temp_data = self.wind_temp_entry.get()
        self.wind_reg_coef_data = self.wind_reg_coef_entry.get()
        self.ice_reg_coef_data = self.ice_reg_coef_entry.get()
        self.wire_hesitation_data = self.wire_hesitation_entry.get()
        self.seismicity_data = self.seismicity_entry.get()
        self.wire_data = self.wire_entry.get()
        self.wire_tencion_data = self.wire_tencion_entry.get()
        self.ground_wire_data = self.ground_wire_entry.get()
        self.sections_data = self.sections_entry.get()
        self.pole_height_data = self.pole_height_entry.get()
        self.pole_base_data = self.pole_base_entry.get()
        self.pole_top_data = self.pole_top_entry.get()
        self.fracture_1_data = self.fracture_1_entry.get()
        self.fracture_2_data = self.fracture_2_entry.get()
        self.height_davit_low_data = self.height_davit_low_entry.get()
        self.height_davit_mid_data = self.height_davit_mid_entry.get()
        self.height_davit_up_data = self.height_davit_up_entry.get()
        self.length_davit_low_r_data = self.lenght_davit_low_r_entry.get()
        self.length_davit_low_l_data = self.lenght_davit_low_l_entry.get()
        self.length_davit_mid_r_data = self.lenght_davit_mid_r_entry.get()
        self.length_davit_mid_l_data = self.lenght_davit_mid_l_entry.get()
        self.length_davit_up_r_data = self.lenght_davit_up_r_entry.get()
        self.length_davit_up_l_data = self.lenght_davit_up_l_entry.get()
        self.wind_span_data = self.wind_span_entry.get()
        self.weight_span_data = self.weight_span_entry.get()


    def convert_txt_to_xlsx(self):
        self.result = extract_txt_data(path=self.file_path)
        return self.result

    # def open_description_window(self,
    #     width=400,
    #     height=200,
    #     title="Описание программы",
    #     resizable=(False, False),
    #     bg="#DFD8FF",
    #     icon=None
    #     ):
    #     DescriptionWindow(self.root,
    #         width,
    #         height,
    #         title,
    #         resizable,
    #         bg,
    #         icon)


if __name__ == "__main__":
    main_page = MainWindow(icon="logo.ico")
    main_page.open_description_window(icon="logo.ico")
    main_page.run()