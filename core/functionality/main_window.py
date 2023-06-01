import tkinter as tk

from core.functionality.lattice_tower import lattice_tower
from core.functionality.multifaceted_tower import multifaceted_tower
from core.functionality.foundation import foundation
from core.utils import *


class MainWindow(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("AmastA")
        self.geometry("500x500+500+150")
        self.resizable(False, False)
        self.config(bg="#FFFFFF")

        self.sapr_txt_label = tk.Label(
            self,
            text="Путь до отчета САПРа",
            anchor="w",
            bg="#FFFFFF"
        )
        self.sapr_txt_entry = tk.Entry(
            self,
            justify="left",
            relief="sunken",
            bd=2,
            width=50
        )
        self.sapr_txt_button = tk.Button(
            self,
            text="Обзор",
            command=make_path_txt
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
        self.sapr_txt_label.place(x=5, y=4)
        self.sapr_txt_entry.place(x=133, y=4)
        self.sapr_txt_button.place(x=440, y=1)

        self.lattice_button.place(x=35, y=26)
        self.multifaceted_button.place(x=135, y=26)
        self.foundation_calculation_button.place(x=235, y=26)

    def go_to_lattice_calculation(self):
        # lattice_window = lattice_tower.LatticeTower()
        lattice_window = lattice_tower.LatticeTower()
        lattice_window.run()

    def go_to_multifaceted_calculation(self):
        # lattice_window = multifaceted_tower.MultifacetedTower()
        lattice_window = multifaceted_tower.MultifacetedTower()
        lattice_window.run()

    def go_to_foundation_calculation(self):
        foundation_calculation_window = foundation.FoundationCalculation(self)
        foundation_calculation_window.run()

if __name__ == "__main__":
    main_page = MainWindow()
    main_page.run()