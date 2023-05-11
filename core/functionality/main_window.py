import tkinter as tk

from core.functionality.lattice_tower import lattice_tower
from core.functionality.multifaceted_tower import multifaceted_tower
from core.utils import *


class MainWindow():
    def __init__(
        self,
        width=270,
        height=55,
        title="Tower Automation",
        resizable=(False, False),
        bg="#FFFFFF",
        bg_pic=None,
        icon=None
    ):
        self.root = tk.Tk()
        self.root.title(title)
        self.root.geometry(f"{width}x{height}+500+150")
        self.root.resizable(resizable[0], resizable[1])
        self.root.config(bg=bg)
        if icon:
            self.root.iconbitmap(icon)

        self.information_label = tk.Label(
            self.root,
            text="Выберите тип опоры для создания отчета:",
            anchor="center",
            bg="#FFFFFF"
        )
        self.lattice_button = tk.Button(
            self.root,
            text="Решетчатая",
            command=self.go_to_lattice_calculation,
            width=12
        )
        self.multifaceted_button = tk.Button(
            self.root,
            text="Многогранная",
            command=self.go_to_multifaceted_calculation,
            width=12
        )

    def run(self):
        self.draw_widgets()
        self.root.mainloop()

    def draw_widgets(self):
        self.information_label.place(x=20, y=3)

        self.lattice_button.place(x=35, y=26)
        self.multifaceted_button.place(x=135, y=26)

    def go_to_lattice_calculation(self):
        lattice_window = lattice_tower.LatticeTower()
        # lattice_window = lattice_tower.LatticeTower(icon="logo.ico")
        lattice_window.run()

    def go_to_multifaceted_calculation(self):
        lattice_window = lattice_tower.LatticeTower()
        # lattice_window = multifaceted_tower.MultifacetedTower(icon="logo.ico")
        lattice_window.run()

if __name__ == "__main__":
    main_page = MainWindow()
    main_page.run()