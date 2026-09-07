import logging
import sys
from tkinter import Tk, Frame, messagebox
from tkinter.ttk import Notebook

from src.gui.SpExGui import SpExGui
from src.gui.ToggleGui import ToggleGui

# TODO 실행중, 실행끝 표시

inKind = ['성풍 출고장', '토글(네이버)']
outKind = ['위하고', '이카운트', '얼마에요', 'CJ택배송장']


class MyGui:
    def __init__(self):
        self.root = Tk()
        if sys.platform == 'darwin':
            # A Finder-launched app has no console for Tk's default error report.
            self.root.report_callback_exception = self.report_callback_exception
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

    def run(self):
        self.root.mainloop()

    def report_callback_exception(self, exc_type, value, traceback):
        logging.getLogger(__name__).error('Conversion failed', exc_info=(exc_type, value, traceback))
        messagebox.showerror('작업을 완료하지 못했습니다', str(value), parent=self.root)


if __name__ == '__main__':
    MyGui().run()
