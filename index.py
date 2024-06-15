import tkinter
from tkinter import *
from tkinter import ttk
from tkinter import messagebox
from tkinter import filedialog
from conf import *
import tkdialog
import csv
import os
import shutil

def save_table():
    with open(path_students, 'w', encoding = 'utf-8') as file:
        file.truncate(0)
        writer = csv.DictWriter(file, columns, lineterminator = '\n')
        writer.writeheader()
        writer.writerows(students)

def window_main():
    window = Tk()
    window.title(window_title)
    window.resizable(0, 0)

    table_frame = Frame(window)
    table_frame.grid(row = 1, column = 1, rowspan = 100)
    scrollbar_y = Scrollbar(table_frame, orient = VERTICAL)
    table = ttk.Treeview(
        master = table_frame,
        height = 20,
        columns = columns,
        show = 'headings',
        yscrollcommand = scrollbar_y.set,
    )
    for i, col in enumerate(columns):
        table.heading(col, text = col)
        table.column(col, width = columns_width[i], anchor = CENTER)
    scrollbar_y.config(command = table.yview)
    scrollbar_y.pack(side = RIGHT, fill = Y)
    table.pack(fill = BOTH, expand = 1)
    cbox = ttk.Combobox(window, state = 'readonly')
    def update_cbox():
        cbox['value'] = (txt_all_classes, ) + tuple(classes)
    update_cbox()
    cbox.current(0)
    def clear_table():
        obj = table.get_children()
        for o in obj:
            table.delete(o)
    def update_table(_ = None):
        selected = cbox.get()
        clear_table()
        for student in students:
            if selected == '' or selected == txt_all_classes or student[key_class] == selected:
                table.insert('', END, values = list(student.values()))
    update_table()

    cbox.bind('<<ComboboxSelected>>', update_table)
    cbox.grid(row = 1, column = 3, sticky = 'NW')
    def add_class():
        def callback(get: str):
            classes.append(get)
            update_cbox()
            cbox.current(len(cbox['value']) - 1)
            update_table()
        tkdialog.ask_input('请输入班级名称', callback)
    def add_student():
        curr_class = cbox.get()
        class_spec = curr_class != txt_all_classes
        def callback(get: str):
            for line in get.split('\n'):
                line = line.strip()
                if line != '':
                    line = line.split(' ')
                    if class_spec:
                        line.insert(0, curr_class)
                    line = dict(zip(columns, line))
                    students.append(line)
                    if line[key_class] not in classes:
                        classes.append(line[key_class])
                    update_cbox()
            save_table()
            update_cbox()
            update_table()
        tkdialog.ask_input(f'请输入学生信息（每行一个学生，格式为“%s{key_sid} {key_name} {key_eid}' % ('' if class_spec else f'{key_class} '), callback, 1)
    def del_student():
        if messagebox.askyesno('确认', '确定要删除选中学生吗？'):
            del students[table.index(table.selection())]
            save_table()
            update_table()
    def del_class():
        curr_class = cbox.get()
        if messagebox.askyesno('确认', '确定要删除或清空当前班级吗？'):
            k = len(students) - 1
            while k >= 0:
                v = students[k]
                if curr_class == txt_all_classes or v[key_class] == curr_class:
                    del students[k]
                k -= 1
            save_table()
            if curr_class != txt_all_classes:
                del classes[classes.index(curr_class)]
            update_cbox()
            cbox.current(0)
            update_table()
    def open_table():
        os.startfile(os.path.join(os.getcwd(), path_students))
    def reload_data(path = path_students_default):
        try:
            shutil.copy(path, path_students)
            load_data()
            update_cbox()
            cbox.current(0)
            update_table()
        except:
            messagebox.showerror('错误', '导入文件可能格式有误！')
    def reload_data_select():
        messagebox.showinfo('提示', f'请选择具有“{key_class} {key_sid} {key_name} {key_eid}”字段的 .csv 格式的表格文件！')
        file = filedialog.askopenfilename()
        if file != '':
            reload_data(file)
    button_add_class = Button(window, text='添加班级', command=add_class)
    button_add_class.grid(row = 2, column = 3, sticky = 'NW')
    button_add_student = Button(window, text='添加学生', command=add_student)
    button_add_student.grid(row = 3, column = 3, sticky = 'NW')
    button_del_student = Button(window, text='删除所选学生', command=del_student)
    button_del_student.grid(row = 4, column = 3, sticky = 'NW')
    button_del_class = Button(window, text='删除当前班级', command=del_class)
    button_del_class.grid(row = 5, column = 3, sticky = 'NW')
    button_del_class = Button(window, text='用其他软件编辑表格', command=open_table)
    button_del_class.grid(row = 6, column = 3, sticky = 'NW')
    button_del_class = Button(window, text='导入数据', command=reload_data_select)
    button_del_class.grid(row = 7, column = 3, sticky = 'NW')
    button_del_class = Button(window, text='导入测试数据', command=reload_data)
    button_del_class.grid(row = 8, column = 3, sticky = 'NW')

    window.mainloop()

window_main()
