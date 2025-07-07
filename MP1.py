from tkinter import *
from tkinter import filedialog, messagebox
import pandas as pd
from tkinter import ttk

class Student:
    def __init__(self, name, id):
        self.name = name
        self.id = id

class StudentList:
    def __init__(self):
        self.students = []

    def add_student(self, student):
        self.students.append(student)

    def remove_student(self, student):
        self.students.remove(student)

class GUI:
    def __init__(self, master):
        self.master = master
        self.added_students = set()
        self.file = None

        self.create_widgets()

    def create_widgets(self):
        # GUI setup
        self.master.geometry("600x260+450+250")
        self.master.title("AttendanceKeeper v1.0")
        self.master.config(background="silver")

        self.label = Label(self.master, text="AttendanceKeeper v1.0", font=('', 19), fg='black', bg='silver')
        self.label.grid(row=0, column=0, columnspan=4)

        self.label1 = Label(self.master, text="Select student list Excel file:", font=('', 12), fg='black', bg='silver',)
        self.label1.grid(row=1, column=0)

        self.label2 = Label(self.master, text="Select a student:", font=('', 12), fg='black', bg='silver',)
        self.label2.grid(row=2, column=0)

        self.label3 = Label(self.master, text="Section: ", font=('', 13), fg='black', bg='silver')
        self.label3.grid(row=2, column=1)

        self.label4 = Label(self.master, text="Please select file type: ", font=('', 10), fg='black', bg='silver')
        self.label4.grid(row=4, column=0, sticky='W')

        self.label5 = Label(self.master, text="Please enter week: ", font=('', 10), fg='black', bg='silver')
        self.label5.grid(row=4, column=1)

        self.label6 = Label(self.master, text="Attended students: ", font=('', 12), fg='black', bg='silver',)
        self.label6.grid(row=2, column=2)

        self.button = Button(self.master, text="Import List", width=20, command=self.Import)
        self.button.grid(row=1, column=1)

        self.button1 = Button(self.master, text="Add = >", width=20, command=self.Add)
        self.button1.grid(row=3, column=1)

        self.button2 = Button(self.master, text="< = Remove", width=20, command=self.Remove)
        self.button2.grid(row=3, column=1, sticky='S')

        self.button3 = Button(self.master, text="Export as File", width=10, command=self.Export)
        self.button3.grid(row=4, column=2, sticky='E')

        self.listbox_frame = Frame(self.master)
        self.listbox_frame.grid(row=3, column=0, sticky='W')
        self.listbox = Listbox(self.listbox_frame, width=35, height=5, selectmode='multiple')
        self.listbox.pack(side=LEFT, fill=BOTH)

        self.listbox_scroll = Scrollbar(self.listbox_frame, orient=VERTICAL)
        self.listbox_scroll.pack(side=RIGHT, fill=Y)

        self.listbox.config(yscrollcommand=self.listbox_scroll.set)
        self.listbox_scroll.config(command=self.listbox.yview)

        self.listbox1_frame = Frame(self.master)
        self.listbox1_frame.grid(row=3, column=2, sticky='E')
        self.listbox1 = Listbox(self.listbox1_frame, width=32, height=5, selectmode='multiple')
        self.listbox1.pack(side=LEFT, fill=BOTH)

        self.listbox1_scroll = Scrollbar(self.listbox1_frame, orient=VERTICAL)
        self.listbox1_scroll.pack(side=RIGHT, fill=Y)

        self.listbox1.config(yscrollcommand=self.listbox1_scroll.set)
        self.listbox1_scroll.config(command=self.listbox1.yview)

        self.drop = ttk.Combobox(self.master)
        self.drop.grid(row=3, column=1, sticky='N')
        self.drop.bind("<<ComboboxSelected>>", self.select_section)

        self.drop1 = ttk.Combobox(self.master, width=5, height=5, values=['xls', 'csv', 'txt'])
        self.drop1.grid(row=4, column=0, sticky='E')
        self.drop1.bind("<<ComboboxSelected>>", self.select_type)

        self.entry = Entry(self.master)
        self.entry.grid(row=4, column=2, sticky='W')

    def Import(self):
        filepath = filedialog.askopenfilename()
        if filepath:
            self.file = pd.read_excel(filepath)
            self.listbox.delete(0, END)
            sections = [str(section).strip() for section in self.file['Section'].unique()]
            self.drop['values'] = sections
            if "AP 01" in sections:
                self.drop.set("AP 01")
        else:
            messagebox.showerror("Error", "No file selected.")

    def select_section(self, event):
        section = self.drop.get()
        if self.file is not None:
            selected_students = self.file[self.file['Section'] == section]
            selected_students['Last Name'] = selected_students['Name'].apply(lambda x: x.split()[1])
            selected_students_sorted = selected_students.sort_values(by='Last Name')
            self.listbox.delete(0, END)
            for index, row in selected_students_sorted.iterrows():
                formatted_row = f"{row['Name'].split()[1]}, {row['Name'].split()[0]}, {row['Id']}"
                self.listbox.insert(END, formatted_row)

            for student in self.added_students.copy():
                if student not in self.listbox.get(0, END):
                    self.added_students.remove(student)
                    self.listbox1.delete(self.listbox1.get(0, END).index(student))

    def Add(self):
        selected_indices = self.listbox.curselection()
        selected_students = [self.listbox.get(index) for index in selected_indices]
        for student in selected_students:
            if student not in self.added_students:
                self.listbox1.insert(END, student)
                self.added_students.add(student)

    def Remove(self):
        selected_indices = self.listbox1.curselection()
        for index in selected_indices[::-1]:
            removed_student = self.listbox1.get(index)
            self.listbox1.delete(index)
            self.added_students.remove(removed_student)

    def select_type(self, event):
            file_type = self.drop1.get()
            if file_type not in ['xls', 'csv']:
                messagebox.showinfo("Info", f"File type selected: {file_type}")
            else:
                messagebox.showerror("Error", "File type is not supported")

    def Export(self):
        if self.file is None:
            messagebox.showerror("Error", "No student list imported")
            return

        section = self.drop.get()
        week = self.entry.get()

        file_type = self.drop1.get()
        if file_type == 'xls':
            filename = f"{section} Week {week}.xls"
            writer = pd.ExcelWriter(filename)
            df = pd.DataFrame(columns=['ID', 'Name', 'Department'])
            for student in self.added_students:
                data = student.split(',')
                df = df.append({'ID': data[2].strip(), 'Name': data[1].strip(), 'Department': section},
                               ignore_index=True)
            df.to_excel(writer, index=False)
            writer.save()
            messagebox.showinfo("Success", f"Attendance saved as {filename}")

        elif file_type == 'csv':
            messagebox.showerror("Error", "File type is not supported")

        elif file_type == 'txt':
            filename = f"{section} Week {week}.txt"
            with open(filename, 'w', encoding='utf-8') as f:
                for student in self.added_students:
                    data = student.split(',')
                    name = data[1].strip() + ' ' + data[0].strip()
                    f.write(f"{data[2].strip()}\t{name}\t{section}\n")
            messagebox.showinfo("Success", f"Attendance saved as {filename}")


def main():
        root = Tk()
        app = GUI(root)
        root.mainloop()

if __name__ == "__main__":
        main()


