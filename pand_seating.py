import pandas as pd
import random
import tkinter as tk
from tkinter import ttk
from fpdf import FPDF
from datetime import datetime
from PIL import Image
import os

classes = {
    "35_capacity": {"columns": 5, "rows": 7, "benches": 35, "classrooms_list": ['A301', 'A302', 'A303', 'A304', 'A305', 'A308', 'A401', 'A402', 'A403', 'A404', 'A405', 'A408', 'C104', 'C203', 'C205', 'C302', 'C303', 'C304', 'C305', 'C306', 'C402', 'C403', 'C404', 'C405', 'C406', 'C408']},
    "40_capacity": {"columns": 5, "rows": 8, "benches": 40, "classrooms_list": ['A306', 'A406', 'E203A', 'E203B', 'A108', 'A107', 'B104', 'B106', 'B108', 'C201', 'C206', 'C208', 'C211', 'C212', 'C301', 'C307', 'C401', 'C407']},
    "36_capacity": {"columns": 6, "rows": 6, "benches": 36, "classrooms_list": ['A106', 'E205', 'E206', 'E207', 'E208', 'E209', 'E210']},
}

columns = 5
benches = 7
total_benches = columns * benches
sections = ['CSE', 'AIE', 'AID', 'ECE', 'EEE', 'EAC', 'ELC', 'MEE', 'RAE']

df = pd.read_excel("Subjects (2).xlsx", sheet_name="names")
df_morning = pd.read_excel("Mid Sem.xlsx", sheet_name="Morning", parse_dates=["Date"])
df_afternoon = pd.read_excel("Mid Sem.xlsx", sheet_name="Afternoon", parse_dates=["Date"])

df_morning["Date"] = pd.to_datetime(df_morning["Date"], dayfirst=True)
df_afternoon["Date"] = pd.to_datetime(df_afternoon["Date"], dayfirst=True)

morning_data = df_morning.to_dict("records")
afternoon_data = df_afternoon.to_dict("records")
for record in morning_data:
    record["Date"] = datetime.strftime(record["Date"], "%d/%m/%Y")
for record in afternoon_data:
    record["Date"] = datetime.strftime(record["Date"], "%d/%m/%Y")

first_year_students = []
second_year_students = []
third_year_students = []
fourth_year_students = []

student_year_lists = {
    "CSE_year_1": [], "CSE_year_2": [], "CSE_year_3": [], "CSE_year_4": [],
    "AIE_year_1": [], "AIE_year_2": [], "AIE_year_3": [], "AIE_year_4": [],
    "AID_year_1": [], "AID_year_2": [], "AID_year_3": [], "AID_year_4": [],
    "ECE_year_1": [], "ECE_year_2": [], "ECE_year_3": [], "ECE_year_4": [],
    "EAC_year_1": [], "EAC_year_2": [], "EAC_year_3": [], "EAC_year_4": [],
    "ELC_year_1": [], "ELC_year_2": [], "ELC_year_3": [], "ELC_year_4": [],
    "EEE_year_1": [], "EEE_year_2": [], "EEE_year_3": [], "EEE_year_4": [],
    "MEE_year_1": [], "MEE_year_2": [], "MEE_year_3": [], "MEE_year_4": [],
    "RAE_year_1": [], "RAE_year_2": [], "RAE_year_3": [], "RAE_year_4": [],
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

    return {
        "college_name": "Amrita Vishwa Vidyapeetham, Bengaluru Campus",
        "report_title": "ATTENDANCE & ROOM SUPERINTENDENT'S REPORT",
        "sub_title": f"B-Tech   {semester_level} {exam_type} Exams",
        "exam_details": f"{semester_level} - {exam_type} Exam",
        'sem_type': sem_type,
        'academic_year': academic_year,
        'exam_type': exam_type,
        'month_details': exam_month
    }

def get_exam_info(session, exam_details):
    exam_info = exam_details.copy()
    exam_info['time_slot'] = "9:00 AM - 12:00 PM" if session == "Morning" else "1:00 PM - 4:00 PM"
    return exam_info

exam_details_data = exam_details()

morning_exam_info = get_exam_info("Morning", exam_details_data)
afternoon_exam_info = get_exam_info("Afternoon", exam_details_data)

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
    for student_roll_no in first_year_students:
        student_year_lists[f"{student_roll_no[8:11]}_year_1"].append(student_roll_no)
    for student_roll_no in second_year_students:
        student_year_lists[f"{student_roll_no[8:11]}_year_2"].append(student_roll_no)
    for student_roll_no in third_year_students:
        student_year_lists[f"{student_roll_no[8:11]}_year_3"].append(student_roll_no)
    for student_roll_no in fourth_year_students:
        student_year_lists[f"{student_roll_no[8:11]}_year_4"].append(student_roll_no)
    return student_year_lists

students_data = student_classification(first_year_students, second_year_students, third_year_students, fourth_year_students, student_year_lists)

def normalize_date(date_str):
    for sep in ['-', '/']:
        if sep in date_str:
            parts = date_str.split(sep)
            break
    day = str(int(parts[0]))
    month = str(int(parts[1]))
    year = parts[2].strip()
    return f"{day}-{month}-{year}"

def format_target_date(date_str):
    if "-" in date_str:
        fmt = "%d-%m-%Y"
    elif "/" in date_str:
        fmt = "%d/%m/%Y"
    else:
        raise ValueError("Date format not recognized")
    dt = datetime.strptime(date_str, fmt)
    return dt.strftime("%d/%m/%Y")

def get_exam_schedule_until_date(data_records, target_date_str):
    if not target_date_str:
        return {}
    target_norm = format_target_date(target_date_str)
    exam_schedule = {}
    for idx, record in enumerate(data_records):
        if record["Date"] == target_norm:
            exam_schedule[idx] = {
                'Subject Name': record['Subject Name'],
                'Subject Code': record['Subject Code'],
                'Date': record["Date"]
            }
    print(f"Exam schedule for {target_date_str}: {exam_schedule}")  # Debug print
    return exam_schedule

def save_seating_arrangement_pdf(arrangement, classes, exam_info, session, date):
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    today_date = datetime.today().strftime("%d/%m/%Y")
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

    downloads_path = os.path.join(os.path.expanduser("~"), "Downloads", "Seating Arrangement", session, date)
    os.makedirs(downloads_path, exist_ok=True)
    pdf.output(os.path.join(downloads_path, "seating_arrangement.pdf"))
    print(f"PDF saved to {downloads_path}")

def count_students_for_courses(students_data, courses_list):
    """Count the total number of students for the selected courses"""
    student_count = 0
    for key, student_list in students_data.items():
        if any(course in key for course in courses_list):
            student_count += len(student_list)
    return student_count

def select_classrooms(classes, student_count):
    selected_classrooms = {}
    total_capacity = 0
    
    required_capacity = student_count
    
    custom_order = ["E203A", "E203B"]
    
    for prefix in ["E", "A", "B", "C"]:
        for capacity_key, capacity_details in classes.items():
            for classroom in capacity_details["classrooms_list"]:
                if classroom not in custom_order and classroom.startswith(prefix):
                    custom_order.append(classroom)
    
    print(f"\nSelecting classrooms for {student_count} students (need capacity for {required_capacity})")  # Debug print
    
    for classroom in custom_order:
        for capacity_key, capacity_details in classes.items():
            if classroom in capacity_details["classrooms_list"]:
                classroom_capacity = capacity_details["benches"]
                
                if total_capacity >= required_capacity:
                    break
                
                if capacity_key not in selected_classrooms:
                    selected_classrooms[capacity_key] = {
                        "columns": capacity_details["columns"],
                        "rows": capacity_details["rows"],
                        "benches": capacity_details["benches"],
                        "classrooms_list": []
                    }
                
                selected_classrooms[capacity_key]["classrooms_list"].append(classroom)
                total_capacity += classroom_capacity
                print(f"  Selected classroom {classroom} (capacity {classroom_capacity}) - Total capacity now: {total_capacity}")  # Debug print
    
    if total_capacity < required_capacity:
        print(f"WARNING: Not enough classroom capacity for all students! Need {required_capacity}, have {total_capacity}")
    
    total_rooms = sum(len(details['classrooms_list']) for details in selected_classrooms.values())
    print(f"Selected {total_rooms} classrooms with total capacity {total_capacity} for {student_count} students")  # Debug print
    
    return selected_classrooms

def seating_arrangement_for_course(classes, students_data, course):
    random.seed(42)
    active_courses = {
        k: v for k, v in students_data.items() 
        if any(k.startswith(c) for c in course) and v and len(v) > 0
    }
    active_course_names = list(active_courses.keys())
    random.shuffle(active_course_names)
    
    all_students = []
    for name in active_course_names:
        all_students.extend(active_courses[name])

    total_students = len(all_students)
    print(f"Total students for course {course}: {total_students}")  # Debug print
    student_index = 0
    bench_count = 0
    empty_seats = []
    classrooms_content = {}

    arrangement = {}
    for capacity, details in classes.items():
        arrangement[capacity] = [
            [["" for _ in range(details['rows'])] for _ in range(details['columns'])]
            for _ in details['classrooms_list']
        ]

    for capacity, details in classes.items():
        for classroom_index, classroom in enumerate(details['classrooms_list']):
            current_column = 0
            classrooms_content[f"classroom_{classroom}"] = []

            while current_column < details['columns']:
                current_bench = 0

                while current_bench < details['rows']:

                    if capacity in ["40_capacity", "36_capacity"]:
                        if (current_column + current_bench) % 2 == 0 and student_index < len(all_students):
                            student = all_students[student_index]
                            arrangement[capacity][classroom_index][current_column][current_bench] = student
                            classrooms_content[f"classroom_{classroom}"].append(student)
                            student_index += 1
                        else:
                            arrangement[capacity][classroom_index][current_column][current_bench] = "EMPTY"
                            empty_seats.append((capacity, classroom_index, current_column, current_bench))
                    else:
                        if bench_count % 2 == 0 and student_index < len(all_students):
                            student = all_students[student_index]
                            arrangement[capacity][classroom_index][current_column][current_bench] = student
                            classrooms_content[f"classroom_{classroom}"].append(student)
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

    print(f"Classrooms content for course {course}: {classrooms_content}")  # Debug print
    return arrangement, classrooms_content

def generate_seating_arrangement(exam_details, session, exam_info, courses_list):
    if exam_details:
        print(f"\nCourses selected: {courses_list}")
        print(f"Generating seating arrangement for {session} session...")
        
        # Count students for selected courses
        student_count = count_students_for_courses(students_data, courses_list)
        print(f"Total students for {session} session: {student_count}")
        
        if student_count == 0:
            print("No students found for the selected courses. Please check course names.")
            return {}, ""
        
        # Select appropriate classrooms based on student count
        selected_classrooms = select_classrooms(classes, student_count)
        
        # Use selected classrooms for seating arrangement
        arrangement, classrooms_content = seating_arrangement_for_course(selected_classrooms, students_data, courses_list)
        
        base_folder = os.path.join(os.path.expanduser("~"), "Downloads", "Seating Arrangement", session, exam_details['Date'])
        os.makedirs(base_folder, exist_ok=True)
        save_seating_arrangement_pdf(arrangement, selected_classrooms, exam_info, session, exam_details['Date'])
        
        # Print a summary of the classrooms used
        print("\nClassrooms used in this arrangement:")
        for capacity, details in selected_classrooms.items():
            print(f"  {capacity.split('_')[0]} capacity rooms: {', '.join(details['classrooms_list'])}")
        
        return classrooms_content, base_folder
    return {}, ""

def save_as_pdf(arrangement, classes):
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.set_font("Arial", size=10, style='B')
    today_date = datetime.today().strftime("%d/%m/%Y")
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
    today_date = datetime.today().strftime("%d/%m/%Y")
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

def attendance_sheet(classrooms_content, df):
    attendance_data = {}
    for index, row in df.iterrows():
        reg_no = row.get("Roll No")
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

def pdf_attendance_sheet(attendance_data, exam_info, base_folder):
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

    date_str = datetime.today().strftime("%d/%m/%Y")

    bw_image_path = "download_bw.png"
    Image.open("download_bw.png").convert('L').save(bw_image_path)

    # Get all subject codes and names
    all_subject_codes = basic.get("Subject Code", "N/A").split(" / ")
    all_subject_names = basic.get("Subject Name", "N/A").split(" / ")
    
    # Create a mapping of subject code to subject name for easy lookup
    subject_map = {code: name for code, name in zip(all_subject_codes, all_subject_names)}

    for classroom, students in attendance_data.items():
        pdf.add_page()

        pdf.image(bw_image_path, 10, 5, 20)

        pdf.set_y(5)

        pdf.set_font("Times", style='B', size=16)
        pdf.cell(210, 10, safe_text(college_name), ln=True, align='C')

        pdf.set_font("Arial", style='BU', size=12)
        pdf.cell(210, 6, safe_text(report_title), ln=True, align='C')

        pdf.set_font("Arial", style='B', size=12)
        pdf.cell(210, 6, safe_text(sub_title), ln=True, align='C')

        pdf.cell(210, 6, safe_text(f"{basic['sem_type']} Semester - {basic['exam_type']} Exam - {basic['month_details']} {datetime.today().year}"), ln=True, align='C')

        # Move room number, date, and time info up before subject details
        actual_classroom_name = classroom.replace("classroom_", "")
        pdf.set_font("Arial", style='B', size=12)
        pdf.cell(0, 8, safe_text(f"Room No.: {actual_classroom_name}"), ln=False, align='L')

        pdf.set_font("Arial", size=10, style='B')
        pdf.set_xy(160, pdf.get_y())
        pdf.cell(40, 8, f"Date: {date_str}    Time: {time_slot}", ln=True, align='R')

        # Add Subject Code and Subject Name on the same line, with code on left and name on right
        if "Subject Code" in basic and "Subject Name" in basic:
            # We'll show all subjects for the session (simplified approach)
            # In a more complex implementation, we could determine which subjects 
            # are actually being written in this classroom
            pdf.ln(1)  # Small space before subject details
            pdf.set_font("Arial", style='B', size=7)  # Bold font for both
            
            for i, (code, name) in enumerate(zip(all_subject_codes, all_subject_names)):
                # Left side: Subject Code
                pdf.cell(90, 5, safe_text(f"Subject Code: {code}"), align='L')
                
                # Right side: Subject Name
                pdf.cell(90, 5, safe_text(f"Subject Name: {name}"), align='R', ln=True)
                
                if i < len(all_subject_codes) - 1:  # Add space between multiple subjects
                    pdf.ln(1)
            
            pdf.ln(2)  # Space after subject details

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
        pdf.multi_cell(summary_width, summary_height / 3, "Room Superintendent\n\nDeputy Controller of Exams\n\n", border=1, align='L')

    downloads_path = os.path.join(base_folder, "attendance_sheet.pdf")
    pdf.output(downloads_path)
    print(f"PDF saved to {downloads_path}")

def attendance_sheet_gui(attendance_data, exam_info):
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

def print_classroom_details(classrooms_content, student_year_lists):
    course_year_to_classrooms = {}
    for course_year, students_list in student_year_lists.items():
        for classroom, studs in classrooms_content.items():
            intersection = sorted(set(studs) & set(students_list))
            if intersection:
                if course_year not in course_year_to_classrooms:
                    course_year_to_classrooms[course_year] = {}
                course_year_to_classrooms[course_year][classroom] = intersection
    return course_year_to_classrooms

def pdf_classroom_details(course_year_to_classrooms, exam_info, base_folder):
    """
    Creates a PDF that shows classroom-wise student details for each subject.
    """
    from fpdf import FPDF
    from datetime import datetime
    
    def safe_text(s):
        return s.encode('latin-1', 'replace').decode('latin-1')

    subject_codes = exam_info.get("Subject Code", "").split(" / ")
    subject_names = exam_info.get("Subject Name", "").split(" / ")
    session = "Morning" if "9:00" in exam_info.get("time_slot", "") else "Afternoon"
    
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    
    total_students_overall = 0
    found_any_students = False

    for idx, (code, name) in enumerate(zip(subject_codes, subject_names)):
        pdf.add_page()
        
        pdf.set_font("Arial", size=14, style='B')
        pdf.cell(0, 10, safe_text(f"Subject: {code} - {name}"), ln=True, align='C')
        pdf.ln(5)
        
        headers = ["Room No", "Register No. Range", "Branch", "Year", "Count"]
        col_widths = [35, 80, 25, 15, 20]
        
        pdf.set_font("Arial", size=10, style='B')
        for h, w in zip(headers, col_widths):
            pdf.cell(w, 8, safe_text(h), border=1, align='C')
        pdf.ln()

        subject_students_count = 0
        pdf.set_font("Arial", size=10)
        
        for course_year_key in sorted(course_year_to_classrooms.keys()):
            # Parse the branch and year from the key (format: "CSE_year_1")
            parts = course_year_key.split('_year_')
            if len(parts) != 2:
                continue
            
            branch = parts[0]
            year = parts[1]
            
            classrooms_dict = course_year_to_classrooms[course_year_key]
            
            for classroom_name, students in sorted(classrooms_dict.items()):
                if students:
                    actual_room = classroom_name.replace("classroom_", "")
                    students_sorted = sorted(students)
                    roll_info = f"{students_sorted[0]} to {students_sorted[-1]}" if len(students_sorted) > 1 else students_sorted[0]
                    
                    pdf.cell(col_widths[0], 8, safe_text(actual_room), border=1, align='C')
                    pdf.cell(col_widths[1], 8, safe_text(roll_info), border=1, align='C')
                    pdf.cell(col_widths[2], 8, safe_text(branch), border=1, align='C')
                    pdf.cell(col_widths[3], 8, safe_text(year), border=1, align='C')
                    pdf.cell(col_widths[4], 8, str(len(students_sorted)), border=1, align='C')
                    pdf.ln()
                    subject_students_count += len(students_sorted)
                    found_any_students = True
        
        pdf.ln(5)
        pdf.set_font("Arial", style='B', size=10)
        pdf.cell(0, 8, safe_text(f"Total for {code}: {subject_students_count}"), ln=True, align='C')
        total_students_overall += subject_students_count
    
    if not found_any_students:
        pdf.add_page()
        pdf.set_font("Arial", size=14, style='B')
        pdf.cell(0, 10, "No classroom details found.", ln=True, align='C')
    else:
        pdf.ln(5)
        pdf.set_font("Arial", style='B', size=10)
        pdf.cell(0, 8, safe_text(f"Overall total: {total_students_overall}"), ln=True, align='C')

    pdf.output(os.path.join(base_folder, f"{session.lower()}_classroom_details.pdf"))

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

    # Create tabs for each subject
    notebook = ttk.Notebook(root)
    notebook.pack(fill="both", expand=True)
    
    subject_codes = exam_info.get("Subject Code", "N/A").split(" / ")
    subject_names = exam_info.get("Subject Name", "N/A").split(" / ")
    
    # Create a mapping of branch codes to subject prefixes
    branch_subject_mapping = {
        "CSE": ["CS", "IS"],
        "AIE": ["AI"],
        "AID": ["AD", "DS"],
        "ECE": ["EC"],
        "EEE": ["EE"],
        "EAC": ["AC"],
        "ELC": ["LC"],
        "MEE": ["ME"],
        "RAE": ["RA"]
    }
    
    for subject_code, subject_name in zip(subject_codes, subject_names):
        # Create a tab for this subject
        subject_frame = ttk.Frame(notebook)
        notebook.add(subject_frame, text=f"{subject_code}")
        
        subject_header = f"Subject Code: {subject_code} | Subject Name: {subject_name}"
        header_label = ttk.Label(subject_frame, text=subject_header, font=("Arial", 16, "bold"))
        header_label.pack(pady=10)
        
        style = ttk.Style()
        style.theme_use("default")
        style.configure("Treeview", rowheight=25, highlightthickness=1, bd=1, relief="solid", font=("Arial", 10))
        style.configure("Treeview.Heading", font=("Arial", 12, "bold"))
        
        tree = ttk.Treeview(subject_frame, columns=("room_no", "reg_numbers", "branch", "year", "total"), show="headings")
        tree.heading("room_no", text="Room No")
        tree.heading("reg_numbers", text="University Registration Number")
        tree.heading("branch", text="Branch")
        tree.heading("year", text="Year")
        tree.heading("total", text="Total")
        tree.column("room_no", width=100, anchor="center")
        tree.column("reg_numbers", width=300, anchor="center")
        tree.column("branch", width=100, anchor="center")
        tree.column("year", width=80, anchor="center")
        tree.column("total", width=50, anchor="center")
        tree.pack(fill="both", expand=True)
        
        # Extract subject prefix for filtering
        subject_prefix = subject_code[:2].upper()
        if subject_prefix in ["19", "20", "21"]:  # Handle curriculum code prefixes
            subject_prefix = subject_code[5:7].upper()
        
        common_subject_prefixes = ["MA", "PH", "CY", "HS", "ID", "GS", "CS"]
        is_common_subject = subject_prefix in common_subject_prefixes
        
        total_students_subject = 0
        
        for branch, years_dict in sorted(course_year_to_classrooms.items()):
            branch_code = branch.split('_')[0]
            
            should_show_branch = False
            # For common subjects, show all branches
            if is_common_subject:
                should_show_branch = True
            # For branch-specific subjects
            elif any(code in subject_code.upper() for code in branch_subject_mapping.get(branch_code, [])):
                should_show_branch = True
            # Check if branch code is explicitly in the subject code
            elif branch_code.upper() in subject_code.upper():
                should_show_branch = True
                
            if not should_show_branch:
                continue
                
            for year, classrooms in sorted(years_dict.items()):
                for classroom_name, students in sorted(classrooms.items()):
                    if not students:
                        continue
                        
                    display_classroom = classroom_name.replace("classroom_", "")
                    year_num = year.split('_')[1].replace('year_', '')
                    
                    students = sorted(students)
                    roll_range = f"{students[0]} to {students[-1]}" if len(students) > 1 else students[0]
                    
                    tree.insert("", "end", values=(
                        display_classroom,
                        roll_range,
                        branch_code,
                        f"Year {year_num}",
                        str(len(students))
                    ))
                    
                    total_students_subject += len(students)
        
        total_label = ttk.Label(subject_frame, text=f"Total Students for {subject_code}: {total_students_subject}", 
                              font=("Arial", 12, "bold"))
        total_label.pack(pady=10)

    save_pdf_button = ttk.Button(
        root, 
        text="Save as PDF",
        command=lambda: pdf_classroom_details(
            course_year_to_classrooms, 
            exam_info, 
            os.path.join(os.path.expanduser("~"), "Downloads", "Seating Arrangement")
        )
    )
    save_pdf_button.pack(side="top", anchor="ne", padx=10, pady=10)

    root.mainloop()

while True:
    date_input = input("Enter the target date (DD/MM/YYYY) or type 'exit' to finish: ")
    if date_input.lower() == "exit":
        print("Exiting...")
        break

    morning_schedule = get_exam_schedule_until_date(morning_data, date_input)
    afternoon_schedule = get_exam_schedule_until_date(afternoon_data, date_input)

    print(f"Found {len(morning_schedule)} morning exams and {len(afternoon_schedule)} afternoon exams on {date_input}")
    if not morning_schedule and not afternoon_schedule:
        print(f"No exams found on {date_input}. Please check the Excel file.")
        continue

    show_gui = input("Do you want to see the GUI after processing all exams? (yes/no): ").strip().lower()

    morning_exams_data = []
    print("\n--- Getting inputs for Morning Exams ---")
    for idx, exam_details_item in morning_schedule.items():
        current_exam = {**exam_details_item, "Date": format_target_date(date_input)}
        print(f"\nMorning Exam: {current_exam['Subject Code']} - {current_exam['Subject Name']}")
        exam_info = morning_exam_info.copy()
        exam_info["Date"] = current_exam["Date"]
        exam_info["Subject Code"] = current_exam["Subject Code"]
        exam_info["Subject Name"] = current_exam["Subject Name"]
        
        prompt_str = (f"Are there multiple courses writing {current_exam['Subject Code']} - "
                    f"{current_exam['Subject Name']} on {current_exam['Date']} (Morning)? (yes/no): ")
        multiple_courses = input(prompt_str).strip().lower()
        courses_list = []
        
        if multiple_courses == "yes":
            print(f"Enter each course one by one. Type 'done' when finished.")
            while True:
                course = input("Course: ").strip()
                if course.lower() == "done":
                    break
                if course: 
                    courses_list.append(course)
        else:
            course = input(f"Which course is writing this exam? ").strip()
            if course:
                courses_list.append(course)
                
        current_exam["courses"] = courses_list
        morning_exams_data.append({
            "exam_details": current_exam,
            "exam_info": exam_info
        })

    afternoon_exams_data = []
    print("\n--- Getting inputs for Afternoon Exams ---")
    for idx, exam_details_item in afternoon_schedule.items():
        current_exam = {**exam_details_item, "Date": format_target_date(date_input)}
        print(f"\nAfternoon Exam: {current_exam['Subject Code']} - {current_exam['Subject Name']}")
        exam_info = afternoon_exam_info.copy()
        exam_info["Date"] = current_exam["Date"]
        exam_info["Subject Code"] = current_exam["Subject Code"]
        exam_info["Subject Name"] = current_exam["Subject Name"]
        
        prompt_str = (f"Are there multiple courses writing {current_exam['Subject Code']} - "
                    f"{current_exam['Subject Name']} on {current_exam['Date']} (Afternoon)? (yes/no): ")
        multiple_courses = input(prompt_str).strip().lower()
        courses_list = []
        
        if multiple_courses == "yes":
            print(f"Enter each course one by one. Type 'done' when finished.")
            while True:
                course = input("Course: ").strip()
                if course.lower() == "done":
                    break
                if course: 
                    courses_list.append(course)
        else:
            course = input(f"Which course is writing this exam? ").strip()
            if course:
                courses_list.append(course)
                
        current_exam["courses"] = courses_list
        afternoon_exams_data.append({
            "exam_details": current_exam,
            "exam_info": exam_info
        })

    if morning_exams_data:
        combined_morning_courses = []
        morning_subject_codes = []
        morning_subject_names = []
        for exam_data in morning_exams_data:
            combined_morning_courses.extend(exam_data["exam_details"]["courses"])
            morning_subject_codes.append(exam_data["exam_details"]["Subject Code"])
            morning_subject_names.append(exam_data["exam_details"]["Subject Name"])
        combined_morning_courses = list(set(combined_morning_courses))
        combined_morning_info = morning_exam_info.copy()
        combined_morning_info["Date"] = morning_exams_data[0]["exam_info"]["Date"]
        combined_morning_info["Subject Code"] = " / ".join(morning_subject_codes)
        combined_morning_info["Subject Name"] = " / ".join(morning_subject_names)

        print("\n--- Generating combined PDFs for Morning Exams ---")
        classrooms_content, base_folder = generate_seating_arrangement(
            combined_morning_info, 
            "Morning", 
            combined_morning_info, 
            combined_morning_courses
        )
        if classrooms_content:
            attendance_data = attendance_sheet(classrooms_content, df)
            pdf_attendance_sheet(attendance_data, combined_morning_info, base_folder)
            course_year_to_classrooms = print_classroom_details(classrooms_content, student_year_lists)
            pdf_classroom_details(course_year_to_classrooms, combined_morning_info, base_folder)

    if afternoon_exams_data:
        combined_afternoon_courses = []
        afternoon_subject_codes = []
        afternoon_subject_names = []
        for exam_data in afternoon_exams_data:
            combined_afternoon_courses.extend(exam_data["exam_details"]["courses"])
            afternoon_subject_codes.append(exam_data["exam_details"]["Subject Code"])
            afternoon_subject_names.append(exam_data["exam_details"]["Subject Name"])
        combined_afternoon_courses = list(set(combined_afternoon_courses))
        combined_afternoon_info = afternoon_exam_info.copy()
        combined_afternoon_info["Date"] = afternoon_exams_data[0]["exam_info"]["Date"]
        combined_afternoon_info["Subject Code"] = " / ".join(afternoon_subject_codes)
        combined_afternoon_info["Subject Name"] = " / ".join(afternoon_subject_names)

        print("\n--- Generating combined PDFs for Afternoon Exams ---")
        classrooms_content, base_folder = generate_seating_arrangement(
            combined_afternoon_info, 
            "Afternoon", 
            combined_afternoon_info, 
            combined_afternoon_courses
        )
        if classrooms_content:
            attendance_data = attendance_sheet(classrooms_content, df)
            pdf_attendance_sheet(attendance_data, combined_afternoon_info, base_folder)
            course_year_to_classrooms = print_classroom_details(classrooms_content, student_year_lists)
            pdf_classroom_details(course_year_to_classrooms, combined_afternoon_info, base_folder)

    if show_gui == "yes":
        print("\nShowing GUI for combined Morning and Afternoon PDFs...")
        if morning_exams_data:
            attendance_sheet_gui(attendance_data, combined_morning_info)
            display_classroom_details_gui(classrooms_content, student_year_lists, combined_morning_info)
        if afternoon_exams_data:
            attendance_sheet_gui(attendance_data, combined_afternoon_info)
            display_classroom_details_gui(classrooms_content, student_year_lists, combined_afternoon_info)

    print(f"\nAll combined PDFs have been generated successfully for the exams on {date_input}")