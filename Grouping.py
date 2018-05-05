from xlrd import open_workbook
from xlwt import Workbook
from tkinter import *
from tkinter import filedialog, messagebox
from random import shuffle
from math import ceil
from os.path import dirname


class Student(object):
    def __init__(self, id, name, gender, want_leader):
        self.id = id
        self.name = name
        self.gender = gender
        self.want_leader = want_leader
        self.is_leader = ''


class Group(object):
    def __init__(self, size):
        self.size = size
        self.members = []

    def has_leader(self):
        for student in self.members:
            if student.is_leader == '是':
                return True
        return False

    def number_of_girls(self):
        return len(
            [student for student in self.members if student.gender == "女"])


class App(Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.pack()
        self.create_widgets()
        self.filename = ''

    def create_widgets(self):
        self.choose_file = Button(
            self, text='Select file', command=self.select_file)
        self.choose_file.pack()
        self.submit = Button(self, text='Start Grouping', command=self.begin)
        self.submit.pack(side=BOTTOM)
        self.message = Message(self)
        self.message.pack()
        self.var = StringVar()
        self.filename_label = Label(self, textvariable=self.var)

    def begin(self):
        if self.filename == '':
            messagebox.showerror('Error', 'File not selected')
            return
        # Read data from excel
        wb = open_workbook(self.filename)
        sheet = wb.sheet_by_index(0)
        number_of_students = sheet.nrows
        students = []
        for row in range(number_of_students):
            id = sheet.cell(row, 0).value
            name = sheet.cell(row, 1).value
            gender = sheet.cell(row, 2).value
            want_leader = sheet.cell(row, 3).value
            students.append(Student(id, name, gender, want_leader))
        shuffle(students)

        # Create groups
        groups = []
        number_of_girls = len(
            [student for student in students if student.gender == '女'])
        number_of_groups = ceil(number_of_students / 6)
        number_of_groups_with_five_students = 6 * number_of_groups - number_of_students
        for i in range(number_of_groups_with_five_students):
            groups.append(Group(5))
        for i in range(number_of_groups - number_of_groups_with_five_students):
            groups.append(Group(6))

        # Get leaders
        leaders = []
        candidates = []
        for student in students:
            if student.want_leader == 1:
                leaders.append(student)
            elif student.want_leader != 0:
                candidates.append(student)
        shuffle(candidates)
        shuffle(leaders)
        if (len(leaders) < number_of_groups):
            while len(leaders) < number_of_groups:
                leaders.append(candidates.pop())
        else:
            while len(leaders) > number_of_groups:
                leaders.pop()
        for student in leaders:
            student.is_leader = '是'
        students = [student for student in students if student not in leaders]

        # Assign a leader to each group
        for group in groups:
            group.members.append(leaders.pop())

        # Spread girls to groups
        girls = [student for student in students if student.gender == "女"]
        students = [student for student in students if student not in girls]
        max_number_of_girls_each_group = ceil(
            number_of_girls / number_of_groups)
        for i in range(max_number_of_girls_each_group):
            for group in groups:
                if len(girls) != 0 and group.number_of_girls() < i + 1:
                    group.members.append(girls.pop())

        # Assign other students
        for group in groups:
            while len(group.members) < group.size:
                group.members.append(students.pop())

        # Write into file
        out_wb = Workbook()
        out_filename = dirname(self.filename) + '/result.xls'
        sheet = out_wb.add_sheet('result')
        row = 1
        col_name = ['组号', '学号', '姓名', '性别', '组长']
        for k, v in enumerate(col_name):
            sheet.write(0, k, v)
        for group_id, group in enumerate(groups):
            for student in group.members:
                sheet.write(row, 0, group_id + 1)
                sheet.write(row, 1, student.id)
                sheet.write(row, 2, student.name)
                sheet.write(row, 3, student.gender)
                sheet.write(row, 4, student.is_leader)
                row += 1
        out_wb.save(out_filename)
        messagebox.showinfo('Success', 'result is stored in ' + out_filename)

    def select_file(self):
        self.filename = filedialog.askopenfilename(
            title='Select file',
            filetypes=[('xlsx', '*.xlsx'), ('xls', 'xls')])
        print(self.filename)
        self.var.set(self.filename)


root = Tk()
root.title('Automatic Group Generator by CSH')
root.geometry('200x100')

# Start program
grouping = App(master=root)
grouping.mainloop()
