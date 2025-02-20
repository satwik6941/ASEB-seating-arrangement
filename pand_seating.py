import pandas as pd
import random
import tkinter as tk 
from tkinter import ttk
from fpdf import FPDF
from datetime import datetime
from PIL import Image
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

df = pd.read_excel("Subjects (2).xlsx", sheet_name="names")

first_year_students = []
second_year_students = []
third_year_students = []
fourth_year_students = []

student_year_lists = {
    "CSE_year_1": [], "CSE_year_2": [], "CSE_year_3": [],"CSE_year_4": [],
    "AIE_year_1": [], "AIE_year_2": [], "AIE_year_3": [],"AIE_year_4": [],
    "AID_year_1": [], "AID_year_2": [], "AID_year_3": [],"AID_year_4": [],
    "ECE_year_1": [], "ECE_year_2": [], "ECE_year_3": [],"ECE_year_4": [],
    "EAC_year_1": [], "EAC_year_2": [], "EAC_year_3": [],"EAC_year_4": [],
    "ELC_year_1": [], "ELC_year_2": [], "ELC_year_3": [],"ELC_year_4": [],
    "EEE_year_1": [], "EEE_year_2": [], "EEE_year_3": [],"EEE_year_4": [],
    "MEE_year_1": [], "MEE_year_2": [], "MEE_year_3": [],"MEE_year_4": [],
    "RAE_year_1": [], "RAE_year_2": [], "RAE_year_3": [],"RAE_year_4": [],
}

def exam_details():
    type_input = input("Are these Mid semester Exams or End Semester Exams (1 - Mid Semester / 2 - End semester): ")
    if type_input == "1":
        exam_type = "Mid Sem"
    elif type_input == "2":
        exam_type = "End Sem"
    else:
        print("ERROR")
        exit(1)

    academic_year_from = input("Enter the academic year: ")
    academic_year_to = str(int(academic_year_from) + 1)
    academic_year = f"{academic_year_from} - {academic_year_to}"

    exam_month_start = input("Which month the exams start: ")
    exam_month_end = input("Which month the exams end: ")
    if exam_month_start == exam_month_end:
        exam_month = exam_month_start
    else:
        exam_month = f"{exam_month_start} - {exam_month_end}"

    sem_input = input("Which kind of semesters are these (1 - Odd / 2 - Even): ")
    if sem_input == "1":
        semester_level = "I, III, V"
        sem_type = "Odd"
    elif sem_input == "2":
        semester_level = "II, IV, VI"
        sem_type = "Even"
    else:
        print("ERROR")
        exit(1)

    time_slot = input("Enter the time slot for the exams (1 - Morning , 2 - Afternoon): ")
    if exam_type == "Mid Sem":
        if time_slot == "1":
            time_slot = "09:30 AM to 11:30 AM"
        elif time_slot == "2":
            time_slot = "01:30 PM to 03:30 PM"
        else:
            print("ERROR")
    elif exam_type == "End Sem":
        if time_slot == "1":
            time_slot = "09:30 AM to 12:30 PM"
        elif time_slot == "2":
            time_slot = "01:30 PM to 04:30 PM"
        else:
            print("ERROR")
    else:
        print("ERROR")

    return {
        "college_name": "Amrita Vishwa Vidyapeetham, Bengaluru Campus",
        "report_title": "ATTENDANCE & ROOM SUPERINTENDENT'S REPORT",
        "sub_title": f"B-Tech   {semester_level} {exam_type} Exams",
        "exam_details": f"{semester_level} - {exam_type} Exam",
        'sem_type': sem_type,
        'academic_year': academic_year,
        'exam_type': exam_type,
        "time_slot": time_slot,
        'month_details': exam_month
    }

exam_info = exam_details()

# This input is used to get the current year from the user to determine the first, second and third year students
current_year = input("Enter the current year: ")

for reg_no in df["Roll No"]:
    if reg_no[11:13] == current_year[-2:]:
        first_year_students.append(reg_no)
    elif reg_no[11:13] == str(int(current_year[-2:]) - 1).zfill(2):
        second_year_students.append(reg_no)
    elif reg_no[11:13] == str(int(current_year[-2:]) - 2).zfill(2):
        third_year_students.append(reg_no)
    elif reg_no[11:13] == str(int(current_year[-2:]) - 3).zfill(2):
        fourth_year_students.append(reg_no)
    else:
        print("Error...")

def student_classification(first_year_students, second_year_students, third_year_students, fourth_year_students, student_year_lists):
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

    for student_roll_no4 in fourth_year_students:
        if student_roll_no4[8:11] == 'CSE':
            student_year_lists["CSE_year_4"].append(student_roll_no4)
        elif student_roll_no4[8:11] == 'AIE':
            student_year_lists["AIE_year_4"].append(student_roll_no4)
        elif student_roll_no4[8:11] == 'AID':
            student_year_lists["AID_year_4"].append(student_roll_no4)
        elif student_roll_no4[8:11] == 'ECE':
            student_year_lists["ECE_year_4"].append(student_roll_no4)
        elif student_roll_no4[8:11] == 'ELC':
            student_year_lists["ELC_year_4"].append(student_roll_no4)
        elif student_roll_no4[8:11] == 'EAC':
            student_year_lists["EAC_year_4"].append(student_roll_no4)
        elif student_roll_no4[8:11] == 'EEE':
            student_year_lists["EEE_year_4"].append(student_roll_no4)
        elif student_roll_no4[8:11] == 'MEE':
            student_year_lists["MEE_year_4"].append(student_roll_no4)
        elif student_roll_no4[8:11] == 'RAE':
            student_year_lists["RAE_year_4"].append(student_roll_no4)
        else:
            print("Error")

    return student_year_lists

students_data = student_classification(first_year_students, second_year_students, third_year_students, fourth_year_students, student_year_lists)


def seating_arrangement(classes, students_data):
    random.seed(42)
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
    print()
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

def save_as_pdf(arrangement, classes, exam_info):
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)

    today_date = datetime.today().strftime("%d-%m-%Y")
    college_name = "Amrita Vishwa Vidyapeetham, Bengaluru Campus"
    current_year = datetime.today().strftime("%Y")

    for capacity, details in classes.items():
        for classroom_index, classroom in enumerate(details['classrooms_list']):
            if capacity == "40_capacity" or classroom_index % 2 == 0:
                pdf.add_page()
            pdf.set_font("Arial", size=16, style='B')  
            pdf.cell(200, 10, txt=college_name, ln=True, align='C')
            pdf.set_font("Arial", size=12, style='B')  
            pdf.cell(200, 10, txt=f"Seating Arrangement for {exam_info['sub_title']} - {exam_info['month_details']} {current_year}", ln=True, align='C')
            pdf.set_font("Arial", size=12, style='B')

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

def seating_gui(arrangement, classrooms_content, classes, exam_info):
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

    save_button = ttk.Button(root, text="Save as PDF", command=lambda: save_as_pdf(arrangement, classes, exam_info))
    save_button.pack(side="top", anchor="ne", padx=10, pady=10)

    title_font = ("Arial", 28, "bold")
    subtitle_font = ("Arial", 18, "bold")
    classroom_font = ("Arial", 30, "bold")
    header_font = ("Arial", 16, "bold")
    seat_font = ("Arial", 14, "bold")

    college_name = "Amrita Vishwa Vidyapeetham, Bengaluru Campus"
    today_date = datetime.today().strftime("%d-%m-%Y")
    current_year = datetime.today().strftime("%Y")

    total_students = 0
    row_counter = 0
    capacities_order = ["35_capacity", "40_capacity", "36_capacity"]
    for capacity in capacities_order:
        details = classes[capacity]
        for classroom_index, classroom in enumerate(details['classrooms_list']):
            frame = tk.Frame(scrollable_frame, pady=20)
            frame.grid(row=row_counter, column=0, padx=50, pady=10, sticky="nsew")
            row_counter += 1

            tk.Label(frame, text=college_name, font=title_font, anchor="center").pack()
            tk.Label(frame, text=f"Seating Arrangement for {exam_info['sub_title']} - {exam_info['month_details']} {current_year}", font=subtitle_font, anchor="center").pack(pady=5)

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
    total_students_label.grid(row=row_counter, column=0, pady=20)

    scrollable_frame.columnconfigure(0, weight=1)

    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    root.mainloop()

seating_gui(arrangement, classrooms_content, classes, exam_info)

def attendance_sheet(classrooms_content, df):
    attendance_data = {}
    for index, row in df.iterrows():
        reg_no = row.get("Registration numbers")
        name = row.get("Name")
        if reg_no and name:
            attendance_data[reg_no] = name

    skip_words = {"Column 1", "Column 2", "Column 3", "Column 4", "Column 5", "Column 6", "Door", "side", "Door Side", "Date"}
    attendance_sheet_data = {}
    for classroom, students in classrooms_content.items():
        attendance_sheet_data[classroom] = []
        for student in students:
            if student not in skip_words:
                name = attendance_data.get(student, "Unknown")
                attendance_sheet_data[classroom].append((name, student))
    return attendance_sheet_data

attendance_data = attendance_sheet(classrooms_content, df)

def pdf_attendance_sheet(attendance_data, exam_info):
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)

    def safe_text(s):
        return s.encode('latin-1', 'replace').decode('latin-1')

    basic = exam_info
    college_name = basic["college_name"]
    report_title = basic["report_title"]
    sub_title = basic["sub_title"]
    exam_details_text = basic["exam_details"]
    time_slot = basic["time_slot"]

    date_str = datetime.today().strftime("%d-%m-%Y")

    bw_image_path = "download_bw.png"
    Image.open("download_bw.png").convert('L').save(bw_image_path)

    for classroom, students in attendance_data.items():
        pdf.add_page()

        pdf.image(bw_image_path, 10, 5, 20)

        pdf.set_font("Times", style='B', size=12)
        pdf.set_xy(10, 30)  # Set the position for the college name
        pdf.cell(190, 6, safe_text(college_name), ln=True, align='C')

        pdf.set_font("Arial", style='BU', size=9)
        pdf.set_xy(10, 40)  # Set the position for the report title
        pdf.cell(190, 5, safe_text(report_title), ln=True, align='C')

        pdf.set_font("Arial", style='B', size=9)
        pdf.set_xy(10, 50)  # Set the position for the sub title
        pdf.cell(190, 5, safe_text(sub_title), ln=True, align='C')

        pdf.set_xy(10, 60)  # Set the position for the semester and exam details
        pdf.cell(190, 5, safe_text(f"{basic['sem_type']} Semester - {basic['exam_type']} Exam - {basic['month_details']} {datetime.today().year}"), ln=True, align='C')

        actual_classroom_name = classroom.replace("classroom_", "")
        pdf.ln(2)
        pdf.set_font("Arial", style='B', size=9)
        pdf.cell(0, 7, safe_text(f"Room No.: {actual_classroom_name}"), ln=False, align='L')

        pdf.set_xy(100, pdf.get_y())
        pdf.set_font("Arial", size=8, style='B')
        pdf.cell(55, 5, f"Date : {date_str}", border=0, align='C')  
        pdf.cell(55, 5, f"Time : {time_slot}", border=0, align='C')  
        pdf.ln(7)

        pdf.set_font("Arial", size=8, style='B') 
        column_widths = [15, 50, 70, 30, 35]
        row_height = 6  
        start_x = (210 - sum(column_widths)) / 2
        pdf.set_x(start_x)
        headers = ["S.No", "Register No.", "Name", "Booklet No.", "Signature"]
        for i, header in enumerate(headers):
            pdf.cell(column_widths[i], row_height, safe_text(header), border=1, align='C')
        pdf.ln(row_height)

        pdf.set_font("Arial", size=8)  
        for idx, (name, reg_no) in enumerate(students, start=1):
            pdf.set_x(start_x)
            pdf.cell(column_widths[0], row_height, str(idx), border=1, align='C')
            pdf.cell(column_widths[1], row_height, reg_no, border=1, align='C')
            if len(name) > 30:
                pdf.set_font("Arial", size=7)  
                pdf.cell(column_widths[2], row_height, safe_text(name), border=1, align='C')
                pdf.set_font("Arial", size=8)  
            else:
                pdf.cell(column_widths[2], row_height, safe_text(name), border=1, align='C')
            pdf.cell(column_widths[3], row_height, "", border=1, align='C')
            pdf.cell(column_widths[4], row_height, "", border=1, ln=True, align='C')

        pdf.ln(7)
        summary_width = 60
        summary_height = 18  
        total_width = summary_width * 3
        start_x = (210 - total_width) / 2
        current_y = pdf.get_y()

        pdf.set_xy(start_x, current_y)
        pdf.set_font("Arial", size=7)  
        pdf.multi_cell(summary_width, summary_height / 3, "Total No of Students Present: __________\n\nTotal No of Students Absent: __________\n\n", border=1, align='L')

        pdf.set_xy(start_x + summary_width, current_y)
        pdf.set_font("Arial", size=7)  
        pdf.multi_cell(summary_width, summary_height / 3, "Register Nos. (Malpractice): ___________\n\nRegister Nos. (absentees): ___________\n\n", border=1, align='L')

        pdf.set_xy(start_x + 2 * summary_width, current_y)
        pdf.set_font("Arial", size=7)  
        pdf.multi_cell(summary_width, summary_height / 3, "Room Superintendent\n\nDeputy Controller of Exams", border=1, align='L')

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
    
    save_pdf_button = ttk.Button(root, text="Save as PDF", command=lambda: pdf_attendance_sheet(attendance_data, exam_info))
    save_pdf_button.pack(side="top", anchor="ne", padx=10, pady=10)
    
    basic = exam_info
    sem_type = basic['sem_type']
    exam_type = basic['exam_type']
    college_name = basic["college_name"]
    report_title = basic["report_title"]
    sub_title = basic["sub_title"]
    exam_details_text = basic["exam_details"]
    exam_month = basic["month_details"]

    current_year = datetime.today().strftime("%Y")
    date_str = datetime.today().strftime("%d-%m-%Y")
    time_slot = basic["time_slot"]

    for classroom, students in attendance_data.items():
        ttk.Label(scrollable_frame, text=college_name, font=("Times", 18, "bold")).pack(anchor="center")
        ttk.Label(scrollable_frame, text=report_title, font=("Arial", 16, "bold")).pack(anchor="center", pady=5)
        ttk.Label(scrollable_frame, text=sub_title, font=("Arial", 14)).pack(anchor="center")
        ttk.Label(scrollable_frame, text=f"{sem_type} - {exam_type}. Exam - {exam_month}. {current_year}", font=("Arial", 12)).pack(anchor="center")
        
        date_time_frame = ttk.Frame(scrollable_frame)
        date_time_frame.pack(fill="x", padx=10, pady=5)
        
        actual_classroom_name = classroom.replace("classroom_", "")
        ttk.Label(date_time_frame, text=f"Room No. {actual_classroom_name}", font=("Arial", 32, "bold"), padding=10).pack(side="left", padx=10, pady=5)
        
        ttk.Label(date_time_frame, text=f"Date: {date_str}", font=("Arial", 12)).pack(side="left", padx=10)
        ttk.Label(date_time_frame, text=f"Time: {time_slot}", font=("Arial", 12)).pack(side="left", padx=10)
    
        columns = ("S.No", "Register No.", "Name", "Booklet No.", "Signature")
        tree = ttk.Treeview(scrollable_frame, columns=columns, show="headings")
    
        for col in columns:
            tree.heading(col, text=col, anchor="center")
            tree.column(col, anchor="center")
    
        for idx, (name, reg_no) in enumerate(students, start=1):
            tree.insert("", "end", values=(idx, reg_no, name, "", ""))
    
        tree.pack(fill="x", padx=10, pady=5)

        summary_frame = ttk.Frame(scrollable_frame)
        summary_frame.pack(fill="x", padx=10, pady=10)

        ttk.Label(
            summary_frame,
            text="Total No of Students Present: __________\n\nTotal No of Students Absent: __________",
            font=("Arial", 10),
            borderwidth=1,
            relief="solid",
            padding=5
        ).grid(row=0, column=0, sticky="nsew", padx=5)

        ttk.Label(
            summary_frame,
            text="Register Nos. (Malpractice): ___________\n\nRegister Nos. (absentees): ___________",
            font=("Arial", 10),
            borderwidth=1,
            relief="solid",
            padding=5
        ).grid(row=0, column=1, sticky="nsew", padx=5)

        ttk.Label(
            summary_frame,
            text="Room Superintendent\n\nDeputy Controller of Exams",
            font=("Arial", 10),
            borderwidth=1,
            relief="solid",
            padding=5
        ).grid(row=0, column=2, sticky="nsew", padx=5)

    scrollable_frame.update_idletasks()
    canvas.configure(scrollregion=canvas.bbox("all"))
    root.mainloop()

attendance_sheet_gui(attendance_data)

def print_classroom_details(classrooms_content, student_year_lists):
    course_year_to_classrooms = {}
    for course_year, students_list in student_year_lists.items():
        for classroom, studs in classrooms_content.items():
            intersection = sorted(set(studs) & set(students_list))
            if intersection:
                if course_year not in course_year_to_classrooms:
                    course_year_to_classrooms[course_year] = {}
                course_year_to_classrooms[course_year][classroom] = intersection

def pdf_classroom_details(course_year_to_classrooms, exam_info): 
    def safe_text(s):
        return s.encode('latin-1', 'replace').decode('latin-1')

    current_year = datetime.now().year

    basic = exam_info
    academic_year = basic['academic_year']

    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.set_font("Arial", size=12, style='B')
    pdf.add_page()

    bw_image_path = "download1_bw.png"
    pdf.image(bw_image_path, x=(210 - 200) / 2, y=10, w=200) 

    pdf.ln(30)
    academic_year_text = f"ACADEMIC YEAR {academic_year} SEATING ARRANGEMENT FOR {exam_info['exam_type'].upper()}/SUPPLEMENTARY EXAMS - {exam_info['month_details'].upper()} {current_year}"
    pdf.set_font("Arial", size=10,style='B')
    pdf.cell(200, 10, txt=safe_text(academic_year_text), ln=True, align='C')

    pdf.cell(200, 10, txt=safe_text("Classroom Details"), ln=True, align='C')

    column_headers = ["Room No", "University Registration Number", "Branch", "Total"]
    column_widths = [50, 80, 40, 20]
    start_x = (210 - sum(column_widths)) / 2

    current_course = None
    for course_year, classroom_dict in course_year_to_classrooms.items():
        course_display = f"{course_year.split('_')[0]} - {course_year.split('_')[2]}"
        if course_display != current_course:
            pdf.ln(5)
            pdf.set_font("Arial", size=10, style='B')
            pdf.cell(200, 10, txt=safe_text(course_display), ln=True, align='C')
            pdf.set_x(start_x)
            for i, header in enumerate(column_headers):
                pdf.cell(column_widths[i], 8, txt=safe_text(header), border=1, align='C')
            pdf.ln()
            current_course = course_display

        for classroom_name, studs in classroom_dict.items():
            pdf.set_x(start_x)
            roll_range = f"{studs[0]} to {studs[-1]}" if len(studs) > 1 else studs[0]
            branch = studs[0][8:11] if studs else ""
            total_count = len(studs)
            row_values = [classroom_name, roll_range, branch, str(total_count)]
            for i, val in enumerate(row_values):
                pdf.cell(column_widths[i], 8, txt=safe_text(val), border=1, align='C')
            pdf.ln()

    downloads_path = os.path.join(os.path.expanduser("~"), "Downloads", "classroom_details.pdf")
    pdf.output(downloads_path)
    print(f"PDF saved to {downloads_path}")

def display_classroom_details_gui(classrooms_content, student_year_lists, exam_info):
    course_year_to_classrooms = {}
    for course_year, students_list in student_year_lists.items():
        for classroom, studs in classrooms_content.items():
            intersection = sorted(set(studs) & set(students_list))
            if intersection:
                if course_year not in course_year_to_classrooms:
                    course_year_to_classrooms[course_year] = {}
                course_year_to_classrooms[course_year][classroom] = intersection

    root = tk.Tk()
    root.title("Classroom Details")
    root.state("zoomed")

    style = ttk.Style()
    style.theme_use("default")
    style.configure("Treeview", rowheight=25, highlightthickness=1, bd=1, relief="solid", font=("Arial", 10))
    style.configure("Treeview.Heading", font=("Arial", 12, "bold"))
    style.configure("CourseYear.Treeview", font=("Arial", 12, "bold"), background="lightblue")

    tree = ttk.Treeview(root, columns=("room_no", "reg_numbers", "branch", "total"), show="headings")
    tree.heading("room_no", text="Room No")
    tree.heading("reg_numbers", text="University Registration Number")
    tree.heading("branch", text="Branch")
    tree.heading("total", text="Total")
    tree.column("room_no", width=100, anchor="center")
    tree.column("reg_numbers", width=300, anchor="center")
    tree.column("branch", width=100, anchor="center")
    tree.column("total", width=50, anchor="center")
    tree.pack(fill="both", expand=True)

    save_pdf_button = ttk.Button(root, text="Save as PDF",
                                    command=lambda: pdf_classroom_details(course_year_to_classrooms,exam_info))
    save_pdf_button.pack(side="top", anchor="ne", padx=10, pady=10)

    current_course = None
    for course_year, classroom_dict in course_year_to_classrooms.items():
        course_name = course_year.split('_')[0]
        year = course_year.split('_')[2]
        course_display = f"{course_name} - {year}"
        if course_display != current_course:
            tree.insert("", "end", values=(course_display, "", "", ""), tags=("course_year",))
            current_course = course_display

        for classroom_name, studs in classroom_dict.items():
            branch = studs[0][8:11] if studs else ""
            roll_range = f"{studs[0]} to {studs[-1]}" if len(studs) > 1 else studs[0]
            tree.insert("", "end", values=(classroom_name, roll_range, branch, str(len(studs))))

    tree.tag_configure("course_year", font=("Arial", 12, "bold"), anchor="center", background="lightblue")

    root.mainloop()

print_classroom_details(classrooms_content, student_year_lists)
display_classroom_details_gui(classrooms_content, student_year_lists, exam_info)