import pandas as pd
import random
import tkinter as tk 
from tkinter import ttk
from fpdf import FPDF
from datetime import datetime
import os

classes = { 
    "35_capacity": {"columns": 5, "rows": 7, "benches" : 35, "classrooms_list" : list(range(1,19))},
    "40_capacity": {"columns": 5, "rows": 8, "benches" : 40, "classrooms_list" : list(range(19,38))},
    "36_capacity": {"columns": 6, "rows": 6, "benches" : 36, "classrooms_list" : list(range(38,57))},
}
columns = 5
benches = 7
total_benches = columns * benches
sections = ['CSE', 'AIE', 'AID', 'ECE', 'EEE', 'MEE', 'RAE']

df = pd.read_csv("seatingarrangement_data.csv")

first_year_students = []
second_year_students = []
third_year_students = []

student_year_lists = {
    "CSE_year_1": [], "CSE_year_2": [], "CSE_year_3": [],
    "AIE_year_1": [], "AIE_year_2": [], "AIE_year_3": [],
    "AID_year_1": [], "AID_year_2": [], "AID_year_3": [],
    "ECE_year_1": [], "ECE_year_2": [], "ECE_year_3": [],
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
        arrangement[capacity] = [[[''] * details['rows'] for _ in range(details['columns'])] for _ in details['classrooms_list']]

    for capacity, details in classes.items():
        for classroom in details['classrooms_list']:
            current_column = 0
            classrooms_content[f"classroom_{classroom}"] = []
            while current_column < details['columns']:
                current_bench = 0
                while current_bench < details['rows']:
                    if bench_count % 2 == 0 and student_index < len(all_students):
                        student = all_students[student_index]
                        arrangement[capacity][classroom - details['classrooms_list'][0]][current_column][current_bench] = student
                        classrooms_content[f"classroom_{classroom}"].append(student) 
                        student_index += 1
                    else:
                        arrangement[capacity][classroom - details['classrooms_list'][0]][current_column][current_bench] = "EMPTY"
                        empty_seats.append((capacity, classroom, current_column, current_bench))
                    bench_count += 1
                    current_bench += 1
                current_column += 1

    for seat in empty_seats:
        if student_index < len(all_students):
            capacity, classroom, col, bch = seat
            student = all_students[student_index]
            arrangement[capacity][classroom - classes[capacity]['classrooms_list'][0]][col][bch] = student
            classrooms_content[f"classroom_{classroom}"].append(student)
            student_index += 1

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
        print(f"{classroom}:")
        print(f"  Number of students: {num_students}")
        for key, count in class_counts.items():
            print(f"  {key}: {count} students")
        print(f"  Roll numbers: {', '.join(students)}")

arrangement, classrooms_content = seating_arrangement(classes, students_data)
classroom_data(classrooms_content)

def save_as_pdf(arrangement, classes):
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.set_font("Arial", size=10, style='B')

    today_date = datetime.today().strftime("%d-%m-%Y")
    college_name = "Amrita Vishwa Vidyapeetham, Bengaluru Campus"

    # Title page
    pdf.add_page()
    pdf.set_font("Arial", size=14, style='B')
    pdf.cell(200, 10, txt=college_name, ln=True, align='C')
    pdf.cell(200, 10, txt="Seating Arrangement", ln=True, align='C')
    pdf.cell(200, 10, txt=f"Date: {today_date}", ln=True, align='C')
    pdf.set_font("Arial", size=12, style='B')

    for capacity, details in classes.items():
        for classroom_index, classroom in enumerate(details['classrooms_list']):
            if classroom_index % 2 == 0 and classroom_index != 0:
                pdf.add_page()

            # Add classroom title
            pdf.set_font("Arial", size=16, style='B')
            pdf.cell(200, 10, txt=f"Classroom {classroom}", ln=True, align='C')
            pdf.set_font("Arial", size=12, style='B')

            # Calculate table alignment
            table_width = details['columns'] * 40
            start_x = (210 - table_width) / 2  

            # Add "Date" above Column 1 and "Door Side" above Column 5
            pdf.set_x(start_x)
            for col in range(details['columns']):
                if col == 0:
                    pdf.cell(40, 6, txt=f"Date: {today_date}", border=0, align='C')
                elif col == 4:
                    pdf.cell(40, 6, txt="Door Side", border=0, align='C')
                else:
                    pdf.cell(40, 6, txt="", border=0)
            pdf.ln()

            # Add column headers
            pdf.set_x(start_x)
            for col in range(details['columns']):
                pdf.cell(40, 6, txt=f"Column {col + 1}", border=1, align='C')
            pdf.ln()

            # Add seating arrangement
            pdf.set_font("Arial", size=10, style='B')
            for row in range(details['rows']):
                pdf.set_x(start_x)
                for col in range(details['columns']):
                    seat = arrangement[capacity][classroom_index][col][row]
                    pdf.cell(40, 12, txt=seat if seat else "EMPTY", border=1, align='C')
                pdf.ln()
            pdf.ln()

    # Save the PDF
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
    classroom_font = ("Arial", 30, "bold")  # Larger font for classrooms
    header_font = ("Arial", 16, "bold")
    seat_font = ("Arial", 14, "bold")  # Slightly larger font for seats

    college_name = "Amrita Vishwa Vidyapeetham, Bengaluru Campus"
    today_date = datetime.today().strftime("%d-%m-%Y")

    # Title and subtitle (centered)
    title_label = tk.Label(scrollable_frame, text=college_name, font=title_font, anchor="center")
    title_label.grid(row=0, column=0, columnspan=1, pady=10)
    subtitle_label = tk.Label(scrollable_frame, text=f"Seating Arrangement\nDate: {today_date}", font=subtitle_font, anchor="center")
    subtitle_label.grid(row=1, column=0, columnspan=1, pady=20)

    classroom_row = 2
    for capacity, details in classes.items():
        for classroom_index, classroom in enumerate(details['classrooms_list']):
            frame = tk.Frame(scrollable_frame, pady=20)
            frame.grid(row=classroom_row, column=0, padx=50, pady=10, sticky="nsew")
            classroom_row += 1

            # Classroom title (centered)
            room_label = tk.Label(frame, text=f"Room No.: {classroom}", font=classroom_font, anchor="center")
            room_label.pack(pady=10)

            # Seating frame
            seating_frame = tk.Frame(frame)
            seating_frame.pack(fill="both", expand=True)

            # Add "Date" above Column 1 and "Door Side" above Column 5
            for col in range(details['columns']):
                if col == 0:
                    date_label = tk.Label(seating_frame, text=f"Date: {today_date}", font=header_font, borderwidth=1, relief="solid", anchor="center")
                    date_label.grid(row=0, column=col, padx=5, pady=5, sticky="nsew")
                elif col == 4:
                    door_side_label = tk.Label(seating_frame, text="Door Side", font=header_font, borderwidth=1, relief="solid", anchor="center")
                    door_side_label.grid(row=0, column=col, padx=5, pady=5, sticky="nsew")
                else:
                    empty_label = tk.Label(seating_frame, text="", font=header_font, anchor="center")
                    empty_label.grid(row=0, column=col, padx=5, pady=5, sticky="nsew")

            # Add column headers (centered)
            for col in range(details['columns']):
                col_label = tk.Label(seating_frame, text=f"Column {col + 1}", font=header_font, borderwidth=1, relief="solid", anchor="center")
                col_label.grid(row=1, column=col, padx=5, pady=5, sticky="nsew")

            # Add seating arrangement (centered)
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

            # Configure grid weights for centering
            for col in range(details['columns']):
                seating_frame.columnconfigure(col, weight=1)
            for row in range(details['rows'] + 2):
                seating_frame.rowconfigure(row, weight=1)

    # Ensure the scrollable frame is properly centered
    scrollable_frame.columnconfigure(0, weight=1)

    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    root.mainloop()

arrangement, classrooms_content = seating_arrangement(classes, students_data)

seating_gui(arrangement, classrooms_content, classes)