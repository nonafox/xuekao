import csv
import os
from docx.shared import *

title = '广州市西关外国语学校 高二1班 谭镇洋 - 考试座位表生成工具'

key_class = '班级'
key_sid = '学号'
key_name = '姓名'
key_eid = '考号'
key_new_class = '考场'
key_new_class_sid = '考场座位号'

key_new_class_id = '考场号'
key_new_class_name = '考场名称'

txt_new_classes = '【所有考场】'
txt_all_classes = '【所有班级】'
txt_new_classes_list = '考场列表'

columns_classes = [key_class, key_sid, key_name, key_eid, key_new_class, key_new_class_sid]
columns_classes_inputable = [key_class, key_sid, key_name, key_eid]
columns_classes_width = [100, 100, 200, 300, 100, 100]
columns_new_classes = [key_new_class_id, key_new_class_name]
columns_new_classes_width = [100, 200]

gen_rows = 6
gen_cols = 5
gen_i_dig = len(str(gen_rows * gen_cols))
page_margin_top = page_margin_left = page_margin_bottom = page_margin_right = Cm(1.27)

paths_to_check = ['./data/out']
for path in paths_to_check:
    os.makedirs(path, exist_ok=1)
file_ext = '.xk'
file_edit_outside_ext = '.csv'
file_out_ext = '.docx'
path_students_default = f'./data/students{file_ext}'
path_students = path_students_default
path_students_test = f'./data/students_default{file_ext}'
path_students_edit_temp = f'./data/students_edit_temp{file_edit_outside_ext}'
path_out_1 = f'./data/out/tables{file_out_ext}'
path_out_2 = f'./data/out/classified_tables{file_out_ext}'

about_url = 'https://github.com/nonafox/xuekao/blob/master/README.md'

students = []
classes = []
new_classes = []
def load_data(path = path_students):
    global students, classes, path_students, new_classes
    students = []
    classes = []
    new_classes = []
    path_students = path

    with open(path_students, 'r', encoding='utf-8') as file:
        file.readline()
        while True:
            new_class = file.readline().strip()
            if new_class != '':
                new_classes.append(new_class)
            else:
                break
        
        reader = csv.DictReader(file)
        for student in reader:
            students.append(student.copy())
            if student[key_class] not in classes:
                classes.append(student[key_class])
load_data()
