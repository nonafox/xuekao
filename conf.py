import csv
import os

title = '广州市西关外国语学校 高二1班 谭镇洋 - 考试座位表生成工具'

key_class = '班级'
key_sid = '学号'
key_name = '姓名'
key_eid = '考号'
txt_all_classes = '所有班级'

columns = [key_class, key_sid, key_name, key_eid]
columns_width = [100, 100, 200, 300]

gen_rows = 6
gen_cols = 5
gen_i_dig = len(str(gen_rows * gen_cols))

paths_to_check = ['./data/out']
for path in paths_to_check:
    os.makedirs(path, exist_ok=1)
path_students_default = './data/students.csv'
path_students = path_students_default
path_students_test = './data/students_default.csv'
path_out_template = './data/out_template.docx'
path_out = './data/out/tables.docx'

about_url = 'https://github.com/nonafox/xuekao/blob/master/README.md'

students = []
classes = []
def load_data(path = path_students):
    global students, classes, path_students
    students = []
    classes = []
    path_students = path
    with open(path_students, 'r', encoding='utf-8') as file:
        reader = csv.DictReader(file)
        for student in reader:
            students.append(student.copy())
            if student[key_class] not in classes:
                classes.append(student[key_class])
load_data()
