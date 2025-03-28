import pandas as pd
import random
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
    print(f"Exam schedule for {target_date_str}: {exam_schedule}")  
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
    
    print(f"\nSelecting classrooms for {student_count} students (need capacity for {required_capacity})")  
    
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
                print(f"  Selected classroom {classroom} (capacity {classroom_capacity}) - Total capacity now: {total_capacity}")  
    
    if total_capacity < required_capacity:
        print(f"WARNING: Not enough classroom capacity for all students! Need {required_capacity}, have {total_capacity}")
    
    total_rooms = sum(len(details['classrooms_list']) for details in selected_classrooms.values())
    print(f"Selected {total_rooms} classrooms with total capacity {total_capacity} for {student_count} students")  
    
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
    print(f"Total students for course {course}: {total_students}")  
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

    print(f"Classrooms content for course {course}: {classrooms_content}")  
    return arrangement, classrooms_content

def generate_seating_arrangement(exam_details, session, exam_info, courses_list):
    if exam_details:
        print(f"\nCourses selected: {courses_list}")
        print(f"Generating seating arrangement for {session} session...")
        student_count = count_students_for_courses(students_data, courses_list)
        print(f"Total students for {session} session: {student_count}")
        
        if student_count == 0:
            print("No students found for the selected courses. Please check course names.")
            return {}, ""
        
        selected_classrooms = select_classrooms(classes, student_count)
        
        arrangement, classrooms_content = seating_arrangement_for_course(selected_classrooms, students_data, courses_list)
        
        base_folder = os.path.join(os.path.expanduser("~"), "Downloads", "Seating Arrangement", session, exam_details['Date'])
        os.makedirs(base_folder, exist_ok=True)
        save_seating_arrangement_pdf(arrangement, selected_classrooms, exam_info, session, exam_details['Date'])
        
        print("\nClassrooms used in this arrangement:")
        for capacity, details in selected_classrooms.items():
            print(f"  {capacity.split('_')[0]} capacity rooms: {', '.join(details['classrooms_list'])}")
        
        return classrooms_content, base_folder
    return {}, ""

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

    all_subject_codes = basic.get("Subject Code", "N/A").split(" / ")
    all_subject_names = basic.get("Subject Name", "N/A").split(" / ")
    
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

        actual_classroom_name = classroom.replace("classroom_", "")
        pdf.set_font("Arial", style='B', size=12)
        pdf.cell(0, 8, safe_text(f"Room No.: {actual_classroom_name}"), ln=False, align='L')

        pdf.set_font("Arial", size=10, style='B')
        pdf.set_xy(160, pdf.get_y())
        pdf.cell(40, 8, f"Date: {date_str}    Time: {time_slot}", ln=True, align='R')

        if "Subject Code" in basic and "Subject Name" in basic:
            pdf.set_font("Arial", style='B', size=7)  
            for i, (code, name) in enumerate(zip(all_subject_codes, all_subject_names)):
                pdf.cell(90, 5, safe_text(f"Subject Code: {code}"), align='L')
                pdf.cell(90, 5, safe_text(f"Subject Name: {name}"), align='R', ln=True)
                
                if i < len(all_subject_codes) - 1:  
                    pdf.ln(1)
            
            pdf.ln(2)  

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

def print_classroom_details(classrooms_content, student_year_lists):
    course_year_to_classrooms = {}
    print(f"Processing {len(classrooms_content)} classrooms")
    print(f"Classroom keys: {list(classrooms_content.keys())[:5]}...")
    
    for course_year, students_list in student_year_lists.items():
        for classroom, studs in classrooms_content.items():
            intersection = sorted(set(studs) & set(students_list))
            if intersection:
                if course_year not in course_year_to_classrooms:
                    course_year_to_classrooms[course_year] = {}
                course_year_to_classrooms[course_year][classroom] = intersection
                
    print(f"Created course_year_to_classrooms with {len(course_year_to_classrooms)} branch-year entries")
    if course_year_to_classrooms:
        sample_key = next(iter(course_year_to_classrooms))
        print(f"Sample data for {sample_key}: {len(course_year_to_classrooms[sample_key])} classrooms")
    
    return course_year_to_classrooms

def pdf_classroom_details(course_year_to_classrooms, exam_info, base_folder):
    def safe_text(s):
        return s.encode('latin-1', 'replace').decode('latin-1')

    subject_codes = exam_info.get("Subject Code", "").split(" / ")
    subject_names = exam_info.get("Subject Name", "").split(" / ")
    session = "Morning" if "9:00" in exam_info.get("time_slot", "") else "Afternoon"
    exam_courses = exam_info.get("courses", [])  
    
    print(f"PDF Classroom Details - Filtering for courses: {exam_courses}")
    print(f"Total branch-year keys before filtering: {len(course_year_to_classrooms)}")
    
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    
    all_classroom_data = {}
    branch_year_students = {}
    total_students = 0
    
    for branch_year_key, classrooms_dict in course_year_to_classrooms.items():
        parts = branch_year_key.split('_year_')
        if len(parts) != 2:
            continue
            
        branch = parts[0]
        year = parts[1]
        
        should_include = True
        if exam_courses:
            should_include = False
            for course in exam_courses:
                if branch.startswith(course) or course.startswith(branch):
                    should_include = True
                    break
        
        if not should_include:
            print(f"Skipping {branch_year_key} - not in exam courses {exam_courses}")
            continue
        else:
            print(f"Including {branch_year_key} for courses {exam_courses}")
        
        for classroom_key, students in classrooms_dict.items():
            room_name = classroom_key.replace("classroom_", "")
            
            if room_name not in all_classroom_data:
                all_classroom_data[room_name] = []
                
            for student_id in students:
                if student_id == "EMPTY":
                    continue
                student_info = {
                    "id": student_id,
                    "branch": branch,
                    "year": year
                }
                all_classroom_data[room_name].append(student_info)
                
                if branch_year_key not in branch_year_students:
                    branch_year_students[branch_year_key] = []
                branch_year_students[branch_year_key].append((room_name, student_id))
                
                total_students += 1
    
    print(f"Total students after filtering: {total_students}")
    print(f"Rooms with data: {list(all_classroom_data.keys())}")
    
    if total_students == 0:
        print(f"Warning: No students found for courses {exam_courses}. PDF will be empty.")
        print("Check if course codes match with branch prefixes in student data.")
        pdf.add_page()
        pdf.set_font("Arial", size=14, style='B')
        pdf.cell(0, 10, "No student data found", ln=True, align='C')
        pdf.set_font("Arial", size=12)
        pdf.cell(0, 10, f"No students matched for courses: {', '.join(exam_courses)}", ln=True, align='C')
        pdf.cell(0, 10, "Please check if the course codes match branch prefixes in student data", ln=True, align='C')
        
        filename = f"{session.lower()}_classroom_details.pdf"
        filepath = os.path.join(base_folder, filename)
        pdf.output(filepath)
        print(f"Created empty PDF with warning message: {filepath}")
        return
    
    for idx, (subject_code, subject_name) in enumerate(zip(subject_codes, subject_names)):
        pdf.add_page()
        
        pdf.set_font("Arial", size=14, style='B')
        pdf.cell(0, 10, safe_text(f"Subject: {subject_code} - {subject_name}"), ln=True, align='C')
        pdf.set_font("Arial", size=12, style='B')
        pdf.cell(0, 10, safe_text(f"Date: {exam_info['Date']} - {exam_info['time_slot']}"), ln=True, align='C')
        
        if exam_courses:
            pdf.set_font("Arial", size=10)
            pdf.cell(0, 8, safe_text(f"Courses: {', '.join(exam_courses)}"), ln=True, align='C')
        
        pdf.ln(5)
        
        pdf.set_font("Arial", size=10, style='B')
        headers = ["Room No", "Branch", "Year", "Roll No. Range", "Count"]
        col_widths = [30, 25, 20, 85, 20]
        
        for h, w in zip(headers, col_widths):
            pdf.cell(w, 8, safe_text(h), border=1, align='C')
        pdf.ln()
        
        room_branch_year_data = {}
        
        for room_name, students in sorted(all_classroom_data.items()):
            by_branch_year = {}
            for student in students:
                key = f"{student['branch']}_{student['year']}"
                if key not in by_branch_year:
                    by_branch_year[key] = []
                by_branch_year[key].append(student['id'])
            
            room_branch_year_data[room_name] = by_branch_year
        
        pdf.set_font("Arial", size=9)
        total_in_subject = 0
        
        for room_name in sorted(room_branch_year_data.keys()):
            for branch_year_key, students in sorted(room_branch_year_data[room_name].items()):
                if not students:
                    continue
                
                branch, year = branch_year_key.split('_')
                students_sorted = sorted(students)
                roll_range = f"{students_sorted[0]} to {students_sorted[-1]}" if len(students_sorted) > 1 else students_sorted[0]
                
                pdf.cell(col_widths[0], 8, safe_text(room_name), border=1, align='C')
                pdf.cell(col_widths[1], 8, safe_text(branch), border=1, align='C')
                pdf.cell(col_widths[2], 8, safe_text(year), border=1, align='C')
                pdf.cell(col_widths[3], 8, safe_text(roll_range), border=1, align='C')
                pdf.cell(col_widths[4], 8, str(len(students_sorted)), border=1, align='C')
                pdf.ln()
                
                total_in_subject += len(students_sorted)
        
        pdf.ln(5)
        pdf.set_font("Arial", style='B', size=10)
        pdf.cell(0, 8, safe_text(f"Total students for {subject_code}: {total_in_subject}"), ln=True, align='C')
    
    pdf.add_page()
    pdf.set_font("Arial", size=14, style='B')
    pdf.cell(0, 10, safe_text("Summary by Branch and Year"), ln=True, align='C')
    if exam_courses:
        pdf.set_font("Arial", size=11)
        pdf.cell(0, 8, safe_text(f"Courses: {', '.join(exam_courses)}"), ln=True, align='C')
    pdf.ln(5)
    
    headers = ["Branch", "Year", "Rooms", "Student Count"]
    col_widths = [30, 20, 90, 30]
    
    pdf.set_font("Arial", size=10, style='B')
    for h, w in zip(headers, col_widths):
        pdf.cell(w, 8, safe_text(h), border=1, align='C')
    pdf.ln()
    
    pdf.set_font("Arial", size=9)
    for branch_year_key, room_students in sorted(branch_year_students.items()):
        if not room_students:
            continue
        
        parts = branch_year_key.split('_year_')
        if len(parts) != 2:
            continue
        
        branch, year = parts
        
        rooms_dict = {}
        for room, student_id in room_students:
            if room not in rooms_dict:
                rooms_dict[room] = 0
            rooms_dict[room] += 1
        
        rooms_formatted = ", ".join([f"{r}({c})" for r, c in sorted(rooms_dict.items())])
        
        pdf.cell(col_widths[0], 8, safe_text(branch), border=1, align='C')
        pdf.cell(col_widths[1], 8, safe_text(year), border=1, align='C')
        pdf.cell(col_widths[2], 8, safe_text(rooms_formatted), border=1, align='L')
        pdf.cell(col_widths[3], 8, str(len(room_students)), border=1, align='C')
        pdf.ln()
    
    pdf.ln(5)
    pdf.set_font("Arial", style='B', size=11)
    pdf.cell(0, 8, safe_text(f"Overall total students: {total_students}"), ln=True, align='C')
    
    filename = f"{session.lower()}_classroom_details.pdf"
    filepath = os.path.join(base_folder, filename)
    pdf.output(filepath)
    print(f"PDF saved to {filepath}")

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
        combined_morning_info["courses"] = combined_morning_courses  # Add courses to exam_info
        
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
        combined_afternoon_info["courses"] = combined_afternoon_courses  
        
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

    print(f"\nAll combined PDFs have been generated successfully for the exams on {date_input}")