import tkinter
from tkinter import *
from tkinter import ttk
from tkinter import messagebox
from tkinter import filedialog
import sv_ttk
import csv
import os
import re
import shutil
import conf
import sv_ttk.theme
import util
import tkdialog
import render
import webbrowser

def save_table():
    with open(conf.path_students, 'w', encoding = 'utf-8') as file:
        file.truncate(0)
        file.write(f'{conf.txt_new_classes_list}\n')
        for new_class in conf.new_classes:
            file.write(f'{new_class}\n')
        file.write('\n')
        writer = csv.DictWriter(file, conf.columns_classes, lineterminator='\n')
        writer.writeheader()
        writer.writerows(conf.students)

def window_main():
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
                    line = re.sub('(\s)+', r'\1', line)
                    line = line.split(' ')
                    if class_spec:
                        line.insert(0, curr_class)
                    line = dict(zip(conf.columns_classes_inputable, line))
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
            del conf.students[curr_table.index(curr_table.selection())]
            save_table()
            update_table()
    def del_class():
        curr_class = cbox.get()
        if messagebox.askyesno('确认', '确定要清空并删除当前班级吗？' if curr_class != conf.txt_all_classes else '确定要清空所有学生和班级吗？'):
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
            cbox.current(1)
            update_table()
    def add_new_class():
        def callback(get: str):
            conf.new_classes.append(get)
            save_table()
            update_table()
        tkdialog.ask_input('请输入考场名称', callback)
    def del_new_class():
        if messagebox.askyesno('确认', '确定要删除选中考场吗？'):
            del conf.new_classes[curr_table.index(curr_table.selection())]
            save_table()
            update_table()
    def edit_new_class():
        k = curr_table.index(curr_table.selection())
        def callback(get: str):
            conf.new_classes[k] = get
            save_table()
            update_table()
        tkdialog.ask_input('请输入新的考场名称', callback, default=conf.new_classes[k])
    def edit_table_outside():
        ori_dir = os.path.join(os.getcwd(), conf.path_students)
        temp_dir = os.path.join(os.getcwd(), conf.path_students_edit_temp)
        shutil.copyfile(ori_dir, temp_dir)
        os.startfile(temp_dir)
        if messagebox.askyesno('编辑花名册', '请在弹出的文件中编辑，编辑完成后将其关闭，点击“是”'):
            shutil.copyfile(temp_dir, ori_dir)
            os.remove(temp_dir)
            reload_data(conf.path_students)
        else:
            os.remove(temp_dir)
    def new_table():
        try:
            file = filedialog.asksaveasfilename(filetypes=[('花名册文件', f'*{conf.file_ext}')], defaultextension=conf.file_ext)
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
            cbox.current(1)
            update_table()
            window_title()
        except Exception as e:
            messagebox.showerror('错误', f'打开文件失败，可能格式有误：{str(e)}')
    def reload_data_default():
        reload_data(conf.path_students_default)
    def reload_data_select():
        file = filedialog.askopenfilename(filetypes=[('表格文件', f'*{conf.file_ext}'), ('所有文件', '*.*')])
        if file != '':
            reload_data(file)
    def gen():
        try:
            dn = util.count_diff_new_classes()
            if dn > 0:
                messagebox.showerror('错误', f'当前考场数量不足，请再添加{dn}个考场！')
                return
            render.render(conf.students, conf.new_classes)
            update_table()
            os.startfile(os.path.join(os.getcwd(), conf.path_out_1))
            os.startfile(os.path.join(os.getcwd(), conf.path_out_2))
        except PermissionError:
            messagebox.showerror('错误', '文件写入失败，请检查其是否被占用！')
        except Exception as e:
            messagebox.showerror('错误', f'文件写入失败，原因未知：\n{str(e)}')
    def about():
        webbrowser.open(conf.about_url)

    window = Tk()

    def window_title():
        window.title(f'{conf.title} - {os.path.join(os.getcwd(), conf.path_students)}')
    window_title()
    window.resizable(0, 0)
    
    cbox = ttk.Combobox(window, state='readonly')
    def update_cbox():
        cbox['value'] = (conf.txt_new_classes, conf.txt_all_classes) + tuple(conf.classes)
    update_cbox()
    cbox.current(1)
    cbox.grid(row=1, column=1)
    
    btns_1 = []
    btns_1.append(ttk.Button(window, text='添加班级', command=add_class))
    btns_1.append(ttk.Button(window, text='添加学生', command=add_student))
    btns_1.append(ttk.Button(window, text='删除所选学生', command=del_student))
    btn_del_class = ttk.Button(window, text='删除当前班级', command=del_class)
    btns_1.append(btn_del_class)
    btns_2 = []
    btns_2.append(ttk.Button(window, text='添加考场', command=add_new_class))
    btns_2.append(ttk.Button(window, text='修改所选考场', command=edit_new_class))
    btns_2.append(ttk.Button(window, text='删除所选考场', command=del_new_class))
    last_row = 2
    btns = btns_1
    def grid_btns():
        nonlocal last_row, btns
        last_row = 2
        selected = cbox.get()
        for btn in btns:
            btn.grid_forget()
        if selected == conf.txt_new_classes:
            btns = btns_2
        else:
            btns = btns_1
        for btn in btns:
            btn.grid(row=last_row, column=1)
            last_row += 1
    grid_btns()
    ttk.Button(window, text='生成考试座位表', command=gen).grid(row=last_row, column=1)
    last_row += 1
    number_tip_label = ttk.Label(window, text='')
    number_tip_label.grid(row=last_row, column=1, rowspan=10)
    last_row += 1

    table_frame = Frame(window)
    table_frame.grid(row=1, column=2, rowspan=100)
    scrollbar_y = ttk.Scrollbar(table_frame, orient=VERTICAL)
    table_1 = ttk.Treeview(
        master=table_frame,
        height=20,
        columns=conf.columns_classes,
        show='headings',
        yscrollcommand=scrollbar_y.set
    )
    table_2 = ttk.Treeview(
        master=table_frame,
        height=20,
        columns=conf.columns_new_classes,
        show='headings',
        yscrollcommand=scrollbar_y.set
    )
    curr_table = table_1
    for i, col in enumerate(conf.columns_classes):
        table_1.heading(col, text=col)
        table_1.column(col, width=conf.columns_classes_width[i], anchor=CENTER)
    for i, col in enumerate(conf.columns_new_classes):
        table_2.heading(col, text=col)
        table_2.column(col, width=conf.columns_new_classes_width[i], anchor=CENTER)
    scrollbar_y.config(command=table_1.yview)
    scrollbar_y.pack(side=RIGHT, fill=Y)
    curr_table.pack(fill=BOTH, expand=1)
    def clear_table():
        obj = curr_table.get_children()
        for o in obj:
            curr_table.delete(o)
    def update_table(_ = None):
        nonlocal curr_table
        selected = cbox.get()
        curr_table.pack_forget()
        if selected == conf.txt_new_classes:
            curr_table = table_2
        else:
            curr_table = table_1
        scrollbar_y.config(command=curr_table.yview)
        curr_table.pack(fill=BOTH, expand=1)
        clear_table()
        grid_btns()
        if selected == conf.txt_new_classes:
            for i, cl in enumerate(conf.new_classes):
                curr_table.insert('', END, values=[i + 1, cl])
        else:
            if selected == conf.txt_all_classes:
                btn_del_class.config(text='清空所有学生和班级')
            else:
                btn_del_class.config(text='删除当前班级')
            for student in conf.students:
                if selected == '' or selected == conf.txt_all_classes or student[conf.key_class] == selected:
                    curr_table.insert('', END, values=list(student.values()))
        dn = util.count_diff_new_classes()
        if dn > 0:
            number_tip_label.config(text=f'★当前考场数量不足，请再添加{dn}个考场！')
        else:
            number_tip_label.config(text='')
    update_table()
    cbox.bind('<<ComboboxSelected>>', update_table)

    menu = Menu(type='menubar', tearoff=False)
    menu_file = Menu(menu, tearoff=False)
    menu.add_cascade(label='文件', menu=menu_file)
    menu_file.add_command(label='新建花名册', command=new_table)
    menu_file.add_command(label='打开默认花名册', command=reload_data_default)
    menu_file.add_command(label='打开测试花名册', command=reload_data)
    menu_file.add_command(label='打开其他花名册', command=reload_data_select)
    menu_edit = Menu(menu, tearoff=False)
    menu.add_cascade(label='编辑', menu=menu_edit)
    menu_edit.add_command(label='用其他软件编辑', command=edit_table_outside)
    menu.add_command(label='关于', accelerator='A', command=about)
    window.config(menu=menu)
    sv_ttk.set_theme('light')
    window.mainloop()

window_main()
