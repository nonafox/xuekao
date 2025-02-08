import tkinter
from tkinter import *
from tkinter import ttk
from tkinter import messagebox
from tkinter import filedialog
import sv_ttk
import csv
import os
import conf
import util
import tkdialog
import render
import webbrowser

def save_table():
    with open(conf.path_students, 'w', encoding = 'utf-8') as file:
        file.truncate(0)
        writer = csv.DictWriter(file, conf.columns, lineterminator='\n')
        writer.writeheader()
        writer.writerows(conf.students)

def window_main():
    window = Tk()

    def window_title():
        window.title(f'{conf.title} - {os.path.join(os.getcwd(), conf.path_students)}')
    window_title()
    window.resizable(0, 0)

    table_frame = Frame(window)
    table_frame.grid(row=1, column=2, rowspan=100)
    scrollbar_y = ttk.Scrollbar(table_frame, orient=VERTICAL)
    table = ttk.Treeview(
        master=table_frame,
        height=20,
        columns=conf.columns,
        show='headings',
        yscrollcommand=scrollbar_y.set,
    )
    for i, col in enumerate(conf.columns):
        table.heading(col, text=col)
        table.column(col, width=conf.columns_width[i], anchor=CENTER)
    scrollbar_y.config(command=table.yview)
    scrollbar_y.pack(side=RIGHT, fill=Y)
    table.pack(fill=BOTH, expand=1)
    cbox = ttk.Combobox(window, state='readonly')
    def update_cbox():
        cbox['value'] = (conf.txt_all_classes, ) + tuple(conf.classes)
    update_cbox()
    cbox.current(0)
    def clear_table():
        obj = table.get_children()
        for o in obj:
            table.delete(o)
    def update_table(_ = None):
        selected = cbox.get()
        clear_table()
        for student in conf.students:
            if selected == '' or selected == conf.txt_all_classes or student[conf.key_class] == selected:
                table.insert('', END, values=list(student.values()))
    update_table()

    cbox.bind('<<ComboboxSelected>>', update_table)
    cbox.grid(row=1, column=1)
    def add_class():
        def callback(get: str):
            conf.classes.append(get)
            update_cbox()
            cbox.current(len(cbox['value']) - 1)
            update_table()
        tkdialog.ask_input('请输入班级名称', callback)
    def add_student():
        curr_class = cbox.get()
        class_spec = curr_class != conf.txt_all_classes
        def callback(get: str):
            for line in get.split('\n'):
                line = line.strip()
                if line != '':
                    line = line.split(' ')
                    if class_spec:
                        line.insert(0, curr_class)
                    line = dict(zip(conf.columns, line))
                    conf.students.append(line)
                    if line[conf.key_class] not in conf.classes:
                        conf.classes.append(line[conf.key_class])
                    update_cbox()
            save_table()
            update_cbox()
            update_table()
        tkdialog.ask_input(f'请输入学生信息（每行一个学生，格式为“%s{conf.key_sid} {conf.key_name} {conf.key_eid}”）' % ('' if class_spec else f'{conf.key_class} '), callback, 1)
    def del_student():
        if messagebox.askyesno('确认', '确定要删除选中学生吗？'):
            del conf.students[table.index(table.selection())]
            save_table()
            update_table()
    def del_class():
        curr_class = cbox.get()
        if messagebox.askyesno('确认', '确定要删除或清空当前班级吗？'):
            k = len(conf.students) - 1
            while k >= 0:
                v = conf.students[k]
                if curr_class == conf.txt_all_classes or v[conf.key_class] == curr_class:
                    del conf.students[k]
                k -= 1
            save_table()
            if curr_class != conf.txt_all_classes:
                del conf.classes[conf.classes.index(curr_class)]
            update_cbox()
            cbox.current(0)
            update_table()
    def open_table():
        os.startfile(os.path.join(os.getcwd(), conf.path_students))
    def new_table():
        try:
            file = filedialog.asksaveasfilename(filetypes=[('表格文件', '*.csv')], defaultextension='.csv')
            if file != '':
                with open(file, 'w'):
                    pass
                reload_data(file)
        except Exception as e:
            messagebox.showerror('错误', f'新建文件失败，文件名可能被占用：{str(e)}')
    def reload_data(path = conf.path_students_test):
        try:
            conf.load_data(path)
            update_cbox()
            cbox.current(0)
            update_table()
            window_title()
            messagebox.showinfo('提示', '已打开！')
        except Exception as e:
            messagebox.showerror('错误', f'打开文件失败，可能格式有误：{str(e)}')
    def reload_data_default():
        reload_data(conf.path_students_default)
    def reload_data_select():
        messagebox.showinfo('提示', f'请选择具有“{conf.key_class} {conf.key_sid} {conf.key_name} {conf.key_eid}”字段的 .csv 格式的表格文件！')
        file = filedialog.askopenfilename(filetypes=[('表格文件', '*.csv'), ('所有文件', '*.*')])
        if file != '':
            reload_data(file)
    def gen():
        try:
            render.render(conf.students)
            update_table()
            os.startfile(os.path.join(os.getcwd(), conf.path_out_1))
            os.startfile(os.path.join(os.getcwd(), conf.path_out_2))
        except PermissionError:
            messagebox.showerror('错误', '文件写入失败，请检查其是否被占用！')
        except Exception as e:
            messagebox.showerror('错误', f'文件写入失败，原因未知：\n{str(e)}')
    def about():
        webbrowser.open(conf.about_url)
    
    ttk.Button(window, text='添加班级', command=add_class).grid(row=2, column=1)
    ttk.Button(window, text='添加学生', command=add_student).grid(row=3, column=1)
    ttk.Button(window, text='删除所选学生', command=del_student).grid(row=4, column=1)
    ttk.Button(window, text='删除当前班级', command=del_class).grid(row=5, column=1)
    ttk.Button(window, text='生成考试座位表', command=gen).grid(row=6, column=1)

    menu = Menu(type='menubar', tearoff=False)
    menu_file = Menu(menu, tearoff=False)
    menu.add_cascade(label='文件', menu=menu_file)
    menu_file.add_command(label='新建花名册', command=new_table)
    menu_file.add_command(label='打开默认花名册', command=reload_data_default)
    menu_file.add_command(label='打开测试花名册', command=reload_data)
    menu_file.add_command(label='打开其他花名册', command=reload_data_select)
    menu_edit = Menu(menu, tearoff=False)
    menu.add_cascade(label='编辑', menu=menu_edit)
    menu_edit.add_command(label='用其他软件编辑', command=open_table)
    menu.add_command(label='关于', accelerator='A', command=about)
    window.config(menu=menu)
    sv_ttk.set_theme('light')
    window.mainloop()

window_main()
