import pandas as pd
import random
import tkinter as tk 
from tkinter import ttk
from fpdf import FPDF
from datetime import datetime
import os

classes = { 
    "35_capacity": {"columns": 5, "rows": 7, "benches" : 35, "classrooms_list" : ['A301', 'A302', 'A303', 'A304', 'A305', 'A308', 'A401', 'A402', 'A403', 'A404', 'A405', 'A408', 'C104', 'C203', 'C205', 'C302', 'C303', 'C304', 'C305', 'C306', 'C402', 'C403', 'C404', 'C405', 'C406', 'C408']},
    "40_capacity": {"columns": 5, "rows": 8, "benches" : 40, "classrooms_list" : ['A306', 'A406', 'E203A', 'E203B', 'A108', 'A107',  'B104', 'B106', 'B108', 'C201', 'C206', 'C208', 'C211', 'C212', 'C301', 'C307', 'C401', 'C407']},
    "36_capacity": {"columns": 6, "rows": 6, "benches" : 36, "classrooms_list" : ['A106', 'E205', 'E206', 'E207', 'E208', 'E209', 'E210']},
}
columns = 5
benches = 7
total_benches = columns * benches
sections = ['CSE', 'AIE', 'AID', 'ECE', 'EEE', 'EAC', 'ELC', 'MEE', 'RAE']

df = pd.read_csv("student list master upto 2023 (1).csv")

first_year_students = []
second_year_students = []
third_year_students = []

student_year_lists = {
    "CSE_year_1": [], "CSE_year_2": [], "CSE_year_3": [],
    "AIE_year_1": [], "AIE_year_2": [], "AIE_year_3": [],
    "AID_year_1": [], "AID_year_2": [], "AID_year_3": [],
    "ECE_year_1": [], "ECE_year_2": [], "ECE_year_3": [],
    "EAC_year_1": [], "EAC_year_2": [], "EAC_year_3": [],
    "ELC_year_1": [], "ELC_year_2": [], "ELC_year_3": [],
    "EEE_year_1": [], "EEE_year_2": [], "EEE_year_3": [],
    "MEE_year_1": [], "MEE_year_2": [], "MEE_year_3": [],
    "RAE_year_1": [], "RAE_year_2": [], "RAE_year_3": []
}

# This input is used to get the current year from the user to determine the first, second and third year students
current_year = input("Enter the current year: ")

for reg_no in df["Registration numbers"]:
    if reg_no[11:13] == current_year[-2:]:
        first_year_students.append(reg_no)
    elif reg_no[11:13] == str(int(current_year[-2:]) - 1).zfill(2):
        second_year_students.append(reg_no)
    elif reg_no[11:13] == str(int(current_year[-2:]) - 2).zfill(2):
        third_year_students.append(reg_no)
    else:
        print("Error")

def student_classification(first_year_students, second_year_students, third_year_students, student_year_lists):
    for student_roll_no1 in first_year_students:
        if student_roll_no1[8:11] == 'CSE':
            student_year_lists["CSE_year_1"].append(student_roll_no1)
        elif student_roll_no1[8:11] == 'AIE':
            student_year_lists["AIE_year_1"].append(student_roll_no1)
        elif student_roll_no1[8:11] == 'AID':
            student_year_lists["AID_year_1"].append(student_roll_no1)
        elif student_roll_no1[8:11] == 'ECE':
            student_year_lists["ECE_year_1"].append(student_roll_no1)
        elif student_roll_no1[8:11] == 'ELC':
            student_year_lists["ELC_year_1"].append(student_roll_no1)
        elif student_roll_no1[8:11] == 'EAC':
            student_year_lists["EAC_year_1"].append(student_roll_no1)
        elif student_roll_no1[8:11] == 'EEE':
            student_year_lists["EEE_year_1"].append(student_roll_no1)
        elif student_roll_no1[8:11] == 'MEE':
            student_year_lists["MEE_year_1"].append(student_roll_no1)
        elif student_roll_no1[8:11] == 'RAE':
            student_year_lists["RAE_year_1"].append(student_roll_no1)
        else:
            print("Error")

    for student_roll_no2 in second_year_students:
        if student_roll_no2[8:11] == 'CSE':
            student_year_lists["CSE_year_2"].append(student_roll_no2)
        elif student_roll_no2[8:11] == 'AIE':
            student_year_lists["AIE_year_2"].append(student_roll_no2)
        elif student_roll_no2[8:11] == 'AID':
            student_year_lists["AID_year_2"].append(student_roll_no2)
        elif student_roll_no2[8:11] == 'ECE':
            student_year_lists["ECE_year_2"].append(student_roll_no2)
        elif student_roll_no2[8:11] == 'ELC':
            student_year_lists["ELC_year_2"].append(student_roll_no2)
        elif student_roll_no2[8:11] == 'EAC':
            student_year_lists["EAC_year_2"].append(student_roll_no2)
        elif student_roll_no2[8:11] == 'EEE':
            student_year_lists["EEE_year_2"].append(student_roll_no2)
        elif student_roll_no2[8:11] == 'MEE':
            student_year_lists["MEE_year_2"].append(student_roll_no2)
        elif student_roll_no2[8:11] == 'RAE':
            student_year_lists["RAE_year_2"].append(student_roll_no2)
        else:
            print("Error")

    for student_roll_no3 in third_year_students:
        if student_roll_no3[8:11] == 'CSE':
            student_year_lists["CSE_year_3"].append(student_roll_no3)
        elif student_roll_no3[8:11] == 'AIE':
            student_year_lists["AIE_year_3"].append(student_roll_no3)
        elif student_roll_no3[8:11] == 'AID':
            student_year_lists["AID_year_3"].append(student_roll_no3)
        elif student_roll_no3[8:11] == 'ECE':
            student_year_lists["ECE_year_3"].append(student_roll_no3)
        elif student_roll_no3[8:11] == 'ELC':
            student_year_lists["ELC_year_3"].append(student_roll_no3)
        elif student_roll_no3[8:11] == 'EAC':
            student_year_lists["EAC_year_3"].append(student_roll_no3)
        elif student_roll_no3[8:11] == 'EEE':
            student_year_lists["EEE_year_3"].append(student_roll_no3)
        elif student_roll_no3[8:11] == 'MEE':
            student_year_lists["MEE_year_3"].append(student_roll_no3)
        elif student_roll_no3[8:11] == 'RAE':
            student_year_lists["RAE_year_3"].append(student_roll_no3)
        else:
            print("Error")

    return student_year_lists

students_data = student_classification(first_year_students, second_year_students, third_year_students, student_year_lists)

def seating_arrangement(classes, students_data):
    active_courses = {k: v for k, v in students_data.items() if v and len(v) > 0}
    active_course_names = list(active_courses.keys())
    random.shuffle(active_course_names)

    all_students = []
    for name in active_course_names:
        all_students.extend(active_courses[name])

    bench_count = 0
    student_index = 0
    empty_seats = []
    classrooms_content = {}

    arrangement = {}
    for capacity, details in classes.items():
        arrangement[capacity] = [
            [["" for _ in range(details['rows'])] for _ in range(details['columns'])]
            for _ in details['classrooms_list']
        ]

    for capacity, details in classes.items():
        for classroom_index, classroom_name in enumerate(details['classrooms_list']):
            classrooms_content[f"classroom_{classroom_name}"] = []
            current_column = 0
            while current_column < details['columns']:
                current_bench = 0
                while current_bench < details['rows']:
                    if capacity in ["40_capacity", "36_capacity"]:
                        if (current_column + current_bench) % 2 == 0 and student_index < len(all_students):
                            student = all_students[student_index]
                            arrangement[capacity][classroom_index][current_column][current_bench] = student
                            classrooms_content[f"classroom_{classroom_name}"].append(student)
                            student_index += 1
                        else:
                            arrangement[capacity][classroom_index][current_column][current_bench] = "EMPTY"
                            empty_seats.append((capacity, classroom_index, current_column, current_bench))
                    else:
                        if bench_count % 2 == 0 and student_index < len(all_students):
                            student = all_students[student_index]
                            arrangement[capacity][classroom_index][current_column][current_bench] = student
                            classrooms_content[f"classroom_{classroom_name}"].append(student)
                            student_index += 1
                        else:
                            arrangement[capacity][classroom_index][current_column][current_bench] = "EMPTY"
                            empty_seats.append((capacity, classroom_index, current_column, current_bench))

                    bench_count += 1
                    current_bench += 1

                current_column += 1

    for seat in empty_seats:
        if student_index < len(all_students):
            capacity, classroom_index, col, bch = seat
            student = all_students[student_index]
            arrangement[capacity][classroom_index][col][bch] = student
            classroom_name = classes[capacity]['classrooms_list'][classroom_index]
            classrooms_content[f"classroom_{classroom_name}"].append(student)
            student_index += 1

    empty_seats = [
        seat for seat in empty_seats
        if arrangement[seat[0]][seat[1]][seat[2]][seat[3]] == "EMPTY"
    ]

    total_students = len(all_students)
    seated_students = total_students - len(all_students[student_index:])
    non_seated_students = total_students - seated_students

    if non_seated_students > 0:
        print(f"Not all students could be seated. Seated: {seated_students}, Non-seated: {non_seated_students}")

    return arrangement, classrooms_content

def classroom_data(classrooms_content):
    print("\nClassroom Data:")
    for classroom, students in classrooms_content.items():
        num_students = len(students)
        class_counts = {}
        for student in students:
            class_name = student[8:11]
            year = student[11:13]
            key = f"{class_name}_year_{year}"
            if key in class_counts:
                class_counts[key] += 1
            else:
                class_counts[key] = 1

arrangement, classrooms_content = seating_arrangement(classes, students_data)
classroom_data(classrooms_content)

def save_as_pdf(arrangement, classes):
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.set_font("Arial", size=10, style='B')

    today_date = datetime.today().strftime("%d-%m-%Y")
    college_name = "Amrita Vishwa Vidyapeetham, Bengaluru Campus"

    pdf.add_page()
    pdf.set_font("Arial", size=14, style='B')
    pdf.cell(200, 10, txt=college_name, ln=True, align='C')
    pdf.cell(200, 10, txt="Seating Arrangement", ln=True, align='C')
    pdf.cell(200, 10, txt=f"Date: {today_date}", ln=True, align='C')
    pdf.set_font("Arial", size=12, style='B')

    for capacity, details in classes.items():
        for classroom_index, classroom in enumerate(details['classrooms_list']):
            if classroom_index % 2 == 0:
                pdf.add_page()

            pdf.set_font("Arial", size=16, style='B')
            pdf.cell(200, 10, txt=f"Classroom {classroom}", ln=True, align='C')
            pdf.set_font("Arial", size=12, style='B')

            if capacity == "36_capacity":
                table_width = details['columns'] * 30  
                cell_width = 30
                font_size = 8
            else:
                table_width = details['columns'] * 35
                cell_width = 35
                font_size = 10
            start_x = (210 - table_width) / 2  

            pdf.set_x(start_x)
            for col in range(details['columns']):
                if col == 0:
                    pdf.cell(cell_width, 6, txt=f"Date: {today_date}", border=0, align='C')
                elif col == details['columns'] - 1:
                    pdf.cell(cell_width, 6, txt="Door Side", border=0, align='C')
                else:
                    pdf.cell(cell_width, 6, txt="", border=0)
            pdf.ln()

            pdf.set_x(start_x)
            for col in range(details['columns']):
                pdf.cell(cell_width, 6, txt=f"Column {col + 1}", border=1, align='C')
            pdf.ln()

            pdf.set_font("Arial", size=font_size, style='B')
            for row in range(details['rows']): 
                pdf.set_x(start_x)
                for col in range(details['columns']):
                    seat = arrangement[capacity][classroom_index][col][row]
                    pdf.cell(cell_width, 12, txt=seat if seat else "EMPTY", border=1, align='C')
                pdf.ln()
            pdf.ln()

    downloads_path = os.path.join(os.path.expanduser("~"), "Downloads", "seating_arrangement.pdf")
    pdf.output(downloads_path)
    print(f"PDF saved to {downloads_path}")

def seating_gui(arrangement, classrooms_content, classes):
    root = tk.Tk()
    root.title("Seating Arrangement")
    root.state("zoomed")

    canvas = tk.Canvas(root)
    scrollbar = ttk.Scrollbar(root, orient="vertical", command=canvas.yview)
    scrollable_frame = ttk.Frame(canvas)

    scrollable_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
    )

    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)

    save_button = ttk.Button(root, text="Save as PDF", command=lambda: save_as_pdf(arrangement, classes))
    save_button.pack(side="top", anchor="ne", padx=10, pady=10)

    title_font = ("Arial", 28, "bold")
    subtitle_font = ("Arial", 18, "bold")
    classroom_font = ("Arial", 30, "bold") 
    header_font = ("Arial", 16, "bold")
    seat_font = ("Arial", 14, "bold")  

    college_name = "Amrita Vishwa Vidyapeetham, Bengaluru Campus"
    today_date = datetime.today().strftime("%d-%m-%Y")

    title_label = tk.Label(scrollable_frame, text=college_name, font=title_font, anchor="center")
    title_label.grid(row=0, column=0, columnspan=1, pady=10)
    subtitle_label = tk.Label(scrollable_frame, text=f"Seating Arrangement\nDate: {today_date}", font=subtitle_font, anchor="center")
    subtitle_label.grid(row=1, column=0, columnspan=1, pady=20)

    classroom_row = 2
    total_students = 0
    for capacity, details in classes.items():
        for classroom_index, classroom in enumerate(details['classrooms_list']):
            frame = tk.Frame(scrollable_frame, pady=20)
            frame.grid(row=classroom_row, column=0, padx=50, pady=10, sticky="nsew")
            classroom_row += 1

            room_label = tk.Label(frame, text=f"Room No.: {classroom}", font=classroom_font, anchor="center")
            room_label.pack(pady=10)

            seating_frame = tk.Frame(frame)
            seating_frame.pack(fill="both", expand=True)

            for col in range(details['columns']):
                if col == 0:
                    date_label = tk.Label(seating_frame, text=f"Date: {today_date}", font=header_font, borderwidth=1, relief="solid", anchor="center")
                    date_label.grid(row=0, column=col, padx=5, pady=5, sticky="nsew")
                elif col == details['columns'] - 1:
                    door_side_label = tk.Label(seating_frame, text="Door Side", font=header_font, borderwidth=1, relief="solid", anchor="center")
                    door_side_label.grid(row=0, column=col, padx=5, pady=5, sticky="nsew")
                elif capacity == "36_capacity" and col == 5:
                    door_side_label = tk.Label(seating_frame, text="Door Side", font=header_font, borderwidth=1, relief="solid", anchor="center")
                    door_side_label.grid(row=0, column=col, padx=5, pady=5, sticky="nsew")
                else:
                    empty_label = tk.Label(seating_frame, text="", font=header_font, anchor="center")
                    empty_label.grid(row=0, column=col, padx=5, pady=5, sticky="nsew")

            for col in range(details['columns']):
                col_label = tk.Label(seating_frame, text=f"Column {col + 1}", font=header_font, borderwidth=1, relief="solid", anchor="center")
                col_label.grid(row=1, column=col, padx=5, pady=5, sticky="nsew")

            for row in range(details['rows']):
                for col in range(details['columns']):
                    seat = arrangement[capacity][classroom_index][col][row]
                    seat_label = tk.Label(
                        seating_frame,
                        text=seat if seat else "EMPTY",
                        font=seat_font,
                        borderwidth=1,
                        relief="solid",
                        anchor="center"
                    )
                    seat_label.grid(row=row + 2, column=col, padx=5, pady=5, sticky="nsew")
                    if seat and seat != "EMPTY":
                        total_students += 1

            for col in range(details['columns']):
                seating_frame.columnconfigure(col, weight=1)
            for row in range(details['rows'] + 2):
                seating_frame.rowconfigure(row, weight=1)

    total_students_label = tk.Label(scrollable_frame, text=f"Total Students Seated: {total_students}", font=subtitle_font, anchor="center")
    total_students_label.grid(row=classroom_row, column=0, columnspan=1, pady=20)

    scrollable_frame.columnconfigure(0, weight=1)

    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    root.mainloop()

seating_gui(arrangement, classrooms_content, classes)

def attendance_sheet(classrooms_content, df):
    attendance_data = {}
    for index, row in df.iterrows():
        reg_no = row.get("Registration numbers")
        name = row.get("Name")
        if reg_no and name:
            attendance_data[reg_no] = name

    skip_words = {"Column 1", "Column 2", "Column 3", "Column 4", "Column 5", "Column 6", "Door", "side", "Door Side", "Date"}
    attendance_sheet_data = {}
    print("\nAttendance Sheet:")
    for classroom, students in classrooms_content.items():
        attendance_sheet_data[classroom] = []
        for student in students:
            if student not in skip_words:
                name = attendance_data.get(student, "Unknown")
                attendance_sheet_data[classroom].append((name, student))
    return attendance_sheet_data

attendance_data = attendance_sheet(classrooms_content, df)

def pdf_attendance_sheet(attendance_data):
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)

    for classroom, students in attendance_data.items():
        pdf.add_page()
        pdf.set_font("Arial", size=22, style='B')
        pdf.cell(200, 10, txt=classroom.encode('latin-1','replace').decode('latin-1'), ln=True, align='C')
        pdf.ln(10)

        pdf.set_font("Arial", size=14, style='B')
        column_widths = [15, 45, 70, 30, 40]
        table_width = sum(column_widths)
        start_x = (210 - table_width) / 2

        pdf.set_x(start_x)
        headers = ["S.No", "Register No.", "Name", "Booklet No.", "Signature"]
        for i, header in enumerate(headers):
            pdf.cell(column_widths[i], 10, txt=header.encode('latin-1','replace').decode('latin-1'), border=1, align='C')
        pdf.ln()

        pdf.set_font("Arial", size=12)
        for idx, (name, reg_no) in enumerate(students, start=1):
            pdf.set_x(start_x)
            pdf.cell(column_widths[0], 10, txt=str(idx).encode('latin-1','replace').decode('latin-1'), border=1, align='C')
            pdf.cell(column_widths[1], 10, txt=reg_no.encode('latin-1','replace').decode('latin-1'), border=1, align='C')
            if len(name) > 30:
                pdf.set_font("Arial", size=8)
                pdf.cell(column_widths[2], 10, txt=name.encode('latin-1','replace').decode('latin-1'), border=1, align='C')
                pdf.set_font("Arial", size=12)
            else:
                pdf.cell(column_widths[2], 10, txt=name.encode('latin-1','replace').decode('latin-1'), border=1, align='C')
            pdf.cell(column_widths[3], 10, txt="", border=1, align='C')
            pdf.cell(column_widths[4], 10, txt="", border=1, ln=True, align='C')

        pdf.ln(10)
        pdf.set_font("Arial", size=10, style='B')

        box_width = 50
        box_height = 15

        pdf.set_x(start_x + 30) 
        pdf.cell(box_width, box_height, txt="Total No of\nStudents Present:", border=1, align='C')
        pdf.cell(box_width, box_height, txt="Register Nos.\n(Malpractice):", border=1, align='C')
        pdf.cell(box_width, box_height, txt="Room\nSuperintendent:", border=1, align='C')
        pdf.ln(box_height)

        pdf.set_x(start_x + 30)
        pdf.cell(box_width, box_height, txt="", border=1, align='C')
        pdf.cell(box_width, box_height, txt="", border=1, align='C')
        pdf.cell(box_width, box_height, txt="", border=1, align='C')
        pdf.ln(box_height)

        pdf.set_x(start_x + 30)
        pdf.cell(box_width, box_height, txt="Total No of\nStudents Absent:", border=1, align='C')
        pdf.cell(box_width, box_height, txt="Register Nos.\n(absentees):", border=1, align='C')
        pdf.cell(box_width, box_height, txt="Deputy\nController of Exams:", border=1, align='C')
        pdf.ln(box_height)

        pdf.set_x(start_x + 30)
        pdf.cell(box_width, box_height, txt="", border=1, align='C')
        pdf.cell(box_width, box_height, txt="", border=1, align='C')
        pdf.cell(box_width, box_height, txt="", border=1, align='C')
        pdf.ln(box_height)

    downloads_path = os.path.join(os.path.expanduser("~"), "Downloads", "attendance_sheet.pdf")
    pdf.output(downloads_path)
    print(f"PDF saved to {downloads_path}")

def attendance_sheet_gui(attendance_data):
    root = tk.Tk()
    root.title("Attendance Sheet")
    root.state("zoomed")

    canvas = tk.Canvas(root)
    scrollbar = ttk.Scrollbar(root, orient="vertical", command=canvas.yview)
    canvas.configure(yscrollcommand=scrollbar.set)
    scrollbar.pack(side="right", fill="y")
    canvas.pack(side="left", fill="both", expand=True)

    scrollable_frame = ttk.Frame(canvas)
    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")

    save_pdf_button = ttk.Button(root, text="Save as PDF", command=lambda: pdf_attendance_sheet(attendance_data))
    save_pdf_button.pack(side="top", anchor="ne", padx=10, pady=10)

    title_font = ("Arial", 32, "bold")
    subtitle_font = ("Arial", 22, "bold")
    header_font = ("Arial", 20, "bold")
    cell_font = ("Arial", 12, "bold")

    for classroom, students in attendance_data.items():
        classroom_label = ttk.Label(scrollable_frame, text=classroom, font=title_font, padding=10)
        classroom_label.pack(anchor="center", padx=10, pady=5)

        columns = ("S.No", "Register No.", "Name", "Booklet No.", "Signature")
        tree = ttk.Treeview(scrollable_frame, columns=columns, show="headings")

        tree.heading("S.No", text="Sl No.", anchor="center")
        tree.heading("Register No.", text="Register No.", anchor="center")
        tree.heading("Name", text="Name", anchor="center")
        tree.heading("Booklet No.", text="Booklet No.", anchor="center")
        tree.heading("Signature", text="Signature", anchor="center")

        tree.column("S.No", width=40, anchor="center")
        tree.column("Register No.", width=120, anchor="center")
        tree.column("Name", width=250, anchor="w")
        tree.column("Booklet No.", width=100, anchor="center")
        tree.column("Signature", width=120, anchor="center")

        for idx, (name, reg_no) in enumerate(students, start=1):
            if len(name) > 30:
                name = '\n'.join([name[i:i+30] for i in range(0, len(name), 30)])
            tree.insert("", "end", values=(idx, reg_no, name, "", ""))

        tree.pack(fill="x", padx=10, pady=5)

        layout_frame = ttk.Frame(scrollable_frame)
        layout_frame.pack(fill="x", padx=10, pady=5)

        box_frame_1 = ttk.Frame(layout_frame, borderwidth=1, relief="solid")
        box_frame_1.grid(row=0, column=0, padx=5, pady=5)
        ttk.Label(box_frame_1, text="Total No of\nStudents Present:", font=cell_font).pack(anchor="w")
        ttk.Entry(box_frame_1, width=20).pack(fill="x", padx=5)
        ttk.Label(box_frame_1, text="Total No of\nStudents Absent:", font=cell_font).pack(anchor="w")
        ttk.Entry(box_frame_1, width=20).pack(fill="x", padx=5)

        box_frame_2 = ttk.Frame(layout_frame, borderwidth=1, relief="solid")
        box_frame_2.grid(row=0, column=1, padx=5, pady=5)
        ttk.Label(box_frame_2, text="Register Nos.\n(Malpractice):", font=cell_font).pack(anchor="w")
        ttk.Entry(box_frame_2, width=20).pack(fill="x", padx=5)
        ttk.Label(box_frame_2, text="Register Nos.\n(absentees):", font=cell_font).pack(anchor="w")
        ttk.Entry(box_frame_2, width=20).pack(fill="x", padx=5)

        box_frame_3 = ttk.Frame(layout_frame, borderwidth=1, relief="solid")
        box_frame_3.grid(row=0, column=2, padx=5, pady=5)
        ttk.Label(box_frame_3, text="Room Superintendent", font=cell_font).pack(anchor="w")
        ttk.Entry(box_frame_3, width=20).pack(fill="x", padx=5)
        ttk.Label(box_frame_3, text="Deputy Controller of Exams", font=cell_font).pack(anchor="w")
        ttk.Entry(box_frame_3, width=20).pack(fill="x", padx=5)

    scrollable_frame.update_idletasks()
    canvas.configure(scrollregion=canvas.bbox("all"))

    root.mainloop()

attendance_sheet_gui(attendance_data)

def print_classroom_details(classrooms_content, first_year_students, second_year_students, third_year_students):
    for classroom, students in classrooms_content.items():
        print(f"\nClassroom {classroom}:")
        class_counts = {}
        for student in students:
            class_name = student[8:11]
            if class_name in class_counts:
                class_counts[class_name].append(student)
            else:
                class_counts[class_name] = [student]
        
        for class_name, student_list in class_counts.items():
            first_year_list = [s for s in student_list if s in first_year_students]
            second_year_list = [s for s in student_list if s in second_year_students]
            third_year_list = [s for s in student_list if s in third_year_students]

            print(f"  {class_name}: {len(student_list)} students")
            if first_year_list:
                print(f"    1st Year: {len(first_year_list)} students")
                print(f"      Roll numbers: {first_year_list[0]} to {first_year_list[-1]}")
            if second_year_list:
                print(f"    2nd Year: {len(second_year_list)} students")
                print(f"      Roll numbers: {second_year_list[0]} to {second_year_list[-1]}")
            if third_year_list:
                print(f"    3rd Year: {len(third_year_list)} students")
                print(f"      Roll numbers: {third_year_list[0]} to {third_year_list[-1]}")

print_classroom_details(classrooms_content, first_year_students, second_year_students, third_year_students)