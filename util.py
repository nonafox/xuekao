import tkinter
from tkinter import *
from tkinter import ttk
import conf

def window_center(window: Tk):
    width = window.winfo_reqwidth()
    height = window.winfo_reqheight()
    screenwidth = window.winfo_screenwidth()
    screenheight = window.winfo_screenheight()
    window.geometry('+%d+%d' % ((screenwidth - width) / 2, (screenheight - height) / 2))
def split_array(arr: list, chunk_size: int):
    return [arr[i:i + chunk_size] for i in range(0, len(arr), chunk_size)]
def str_select(str1: str, str2: str):
    return str1 if str1 != '' else str2
def count_diff_new_classes():
    return (len(conf.students) / (conf.gen_rows * conf.gen_cols)).__ceil__() - len(conf.new_classes)
