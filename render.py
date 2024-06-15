import docx
from docx import document
from docx.enum.text import *
from tkinter import messagebox
import conf

def render(data: list):
    doc = docx.Document(conf.path_out_template)
    doc.styles['Normal'].font.name = '宋体'
    assert isinstance(doc, document.Document)
    virgin = 1
    for room in data:
        if virgin:
            virgin = 1
        else:
            doc.add_page_break()
        table = doc.add_table(conf.gen_rows, conf.gen_cols)
        i = 1
        orient = 1
        row, col = 0, conf.gen_cols - 1
        for student in room:
            cell = table.cell(row, col)
            p = cell.add_paragraph()
            l = p.add_run('座位号 %s\n' % str(i).rjust(conf.gen_i_dig, '0'))
            l.font.size = 9
            p = cell.add_paragraph()
            l = p.add_run('%s %s号\n' % (student[conf.key_class], student[conf.key_sid]))
            l.font.size = 9
            p.alignment = 
            p = cell.add_paragraph()
            l = p.add_run('%s\n' % student[conf.key_name])
            l.font.size = 12
            l.font.bold = 1
            p = cell.add_paragraph()
            l = p.add_run('考号 %s\n' % student[conf.key_eid])
            l.font.size = 9
            row += orient
            if row < 0:
                row = 0
                col -= 1
                orient *= - 1
            elif row > conf.gen_rows - 1:
                row = conf.gen_rows - 1
                col -= 1
                orient *= -1
            i += 1
    doc.save(conf.path_out)
