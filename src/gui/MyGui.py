import os
import sys
from tkinter import Tk, Toplevel, Frame
from tkinter.ttk import Notebook

import xlwings

from src.gui.SpExGui import SpExGui
from src.gui.ToggleGui import ToggleGui

# TODO 실행중, 실행끝 표시

inKind = ['성풍 출고장', '토글(네이버)']
outKind = ['위하고', '이카운트', '얼마에요', 'CJ택배송장']


class MyGui:
    def __init__(self):
        self.root = Tk()
        # self.toggleRoot = Toplevel(self.root)

        notebook = Notebook(self.root, width=800, height=500)
        notebook.pack()

        tab1 = Frame(self.root)
        tab2 = Frame(self.root)
        tab3 = Frame(self.root)
        notebook.add(tab1, text="출고장 -> ")
        notebook.add(tab2, text="출고장 -> ")
        notebook.add(tab3, text="토글 -> ")

        SpExGui(tab1, tab2)
        ToggleGui(tab3)

        self.root.mainloop()


if __name__ == '__main__':
    myGui = MyGui()
    os.system('pause')
