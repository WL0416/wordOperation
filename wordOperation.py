import sys
import xlrd
import xdrlib
from docxtpl import DocxTemplate

NAME = ''
ID = ''
PROGRAM = ''
PHONE = 'Phone: '
EMAIL = 'Email: '
ADDRESS = 'Address: '
PASSPORT = ''

def open_excel(file = "student.xlsx"):
    try:
        data = xlrd.open_workbook(file)
        return data
    except Exception as e:
        print(str(e))

def parse_data():
    data = open_excel() #打开excel文件
    table = data.sheet_by_name('Checklist') #根据sheet名字来获取excel中的sheet
    nrows = table.nrows #行数
    list = []
    for rownum in range(1, nrows):
        row = table.row_values(rownum)
        if row:          
            list.append(row)
    return list

def write_to_word():
    data = parse_data()
    count_num = 1
    for student in data:
        if student[1] is not '':
            if student[5] is '':
                NAME = student[4] + ' ' + str(student[6]).upper()
            else:
                NAME = student[4] + ' ' + student[5] + ' ' + str(student[6]).upper()
            print(NAME)
            ID = student[2]
            print(ID)
            PROGRAM = student[17]
            print(PROGRAM)
            PASSPORT = student[10]
            print(PASSPORT)
            if student[11] is not '':
                ADDRESS = 'Address: ' + str(student[11]).rstrip() + ' ' + \
                        str(student[12]).rstrip() + ' ' + \
                        str(student[13]).rstrip() + ' ' + \
                        str(student[15]).rstrip()
            print(ADDRESS)

            tpl = DocxTemplate('check.docx')

            context = {'name': NAME,
                       'id': ID,
                       'program': PROGRAM,
                       'passport': PASSPORT,
                       'phone': PHONE,
                       'email': EMAIL,
                       'address': ADDRESS}
            
            tpl.render(context)
            tpl.save('output/' + str(count_num) + '.docx')
            count_num += 1

if __name__ == "__main__":
    write_to_word()