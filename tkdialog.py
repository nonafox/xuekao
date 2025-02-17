import tkinter
from tkinter import *
from tkinter import ttk
import sv_ttk
import util

def ask_input(prompt = '', callback = lambda val: None, wrap = 0, default = ''):
    window = Tk()
    window.title('输入')
    window.resizable(False, False)
    label = ttk.Label(window, text=prompt)
    label.grid(row=1, column=0, columnspan=3)
    if wrap:
        entry = Text(window, width=30, height=5)
        entry.insert(END, default)
    else:
        entry = ttk.Entry(window, width=30)
        entry.insert(END, default)
    entry.grid(row=2, column=1, columnspan=2)
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
    ttk.Label(window).grid(row=3, column=0, columnspan=3)
    button_confirm = ttk.Button(window, text='确定', command=confirm, width=20)
    button_confirm.grid(row=4, column=1)
    button_cancel = ttk.Button(window, text='取消', command=cancel, width=20)
    button_cancel.grid(row=4, column=2)
    util.window_center(window)
    entry.focus_force()
    window.mainloop()
