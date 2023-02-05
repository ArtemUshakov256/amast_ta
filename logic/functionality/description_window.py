import tkinter as tk


class DescriptionWindow():
    def __init__(
        self,
        parent,
        width=400,
        height=200,
        title="Основное окно",
        resizable=(False, False),
        bg="#DFD8FF",
        icon=None
    ):
        self.root = tk.Toplevel(parent)
        self.root.title(title)
        self.root.geometry(f"{width}x{height}+200+200")
        self.root.resizable(resizable[0], resizable[1])
        self.root.config(bg=bg)
        if icon:
            self.root.iconbitmap(icon)