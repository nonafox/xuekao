import tkinter
from tkinter import *
from tkinter import ttk
from tkinter.scrolledtext import ScrolledText

def window_center(window: Tk):
    width = window.winfo_width()
    height = window.winfo_height()
    screenwidth = window.winfo_screenwidth()
    screenheight = window.winfo_screenheight()
    window.geometry('+%d+%d' % ((screenwidth - width) / 2, (screenheight - height) / 2))
def ask_input(prompt = '', callback = lambda val: None, wrap = 0):
    window = Tk()
    window.title('输入')
    window.resizable(False, False)
    label = Label(window, text = prompt, height = 3)
    label.grid(row = 1, column = 0, columnspan = 3)
    if wrap:
        entry = ScrolledText(window, width = 30, height = 5)
    else:
        entry = Entry(window, width = 30)
    entry.grid(row = 2, column = 1, columnspan = 2)
    def confirm():
        if wrap:
            get = entry.get('1.0', 'end').rstrip()
        else:
            get = entry.get()
        if get != '':
            callback(get)
        window.destroy()
    def cancel():
        window.destroy()
    button_confirm = Button(window, text = '确定', command = confirm, width = 20)
    button_confirm.grid(row = 3, column = 1)
    button_cancel = Button(window, text = '取消', command = cancel, width = 20)
    button_cancel.grid(row = 3, column = 2)
    window_center(window)
    entry.focus_force()
    window.mainloop()
