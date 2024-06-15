import csv

window_title = '高一4班谭镇洋 学考项目 - 考号'

key_class = '班级'
key_sid = '学号'
key_name = '姓名'
key_eid = '考号'
txt_all_classes = '所有班级'

columns = [key_class, key_sid, key_name, key_eid]
columns_width = [100, 100, 200, 300]

path_students = './data/students.csv'
path_students_default = './data/students_default.csv'

students = []
classes = []
def load_data(path = path_students):
    global path_students, students, classes
    students = []
    classes = []
    path_students = path
    with open(path_students, 'r', encoding = 'utf-8') as file:
        reader = csv.DictReader(file)
        for student in reader:
            students.append(student.copy())
            if student[key_class] not in classes:
                classes.append(student[key_class])
load_data()
