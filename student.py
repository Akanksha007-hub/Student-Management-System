import tkinter as tk
from tkinter import ttk, messagebox
import mysql.connector
from openpyxl import Workbook, load_workbook
import os



# MYSQL CONNECTION


def connect_mysql():
    try:
        conn = mysql.connector.connect(
            host="localhost",
            user="root",
            password="ramaji007",
            database="student_result"
        )
        return conn

    except mysql.connector.Error as e:
        messagebox.showerror("Database Error", str(e))
        return None



# CALCULATE RESULT


def calculate_result():
    try:
        maths = int(math_entry.get())
        science = int(science_entry.get())
        english = int(english_entry.get())

        # Validate marks
        if not (
            0 <= maths <= 100
            and 0 <= science <= 100
            and 0 <= english <= 100
        ):
            messagebox.showerror(
                "Error",
                "Marks must be between 0 and 100"
            )
            return None

        total = maths + science + english
        percentage = total / 3

        # Calculate grade and status
        if maths < 35 or science < 35 or english < 35:
            grade = "F"
            status = "FAIL"

        elif percentage >= 90:
            grade = "A+"
            status = "PASS"

        elif percentage >= 80:
            grade = "A"
            status = "PASS"

        elif percentage >= 70:
            grade = "B"
            status = "PASS"

        elif percentage >= 60:
            grade = "C"
            status = "PASS"

        elif percentage >= 50:
            grade = "D"
            status = "PASS"

        else:
            grade = "E"
            status = "PASS"

        return total, round(percentage, 2), grade, status

    except ValueError:
        messagebox.showerror(
            "Error",
            "Please enter valid marks"
        )
        return None



# SAVE TO MYSQL AND EXCEL


def save_student():
    roll_no = roll_entry.get().strip()
    name = name_entry.get().strip()
    class_name = class_entry.get().strip()

    # Validate basic details
    if not roll_no or not name or not class_name:
        messagebox.showerror(
            "Error",
            "Please enter Roll No, Name and Class"
        )
        return

    result = calculate_result()

    if result is None:
        return

    total, percentage, grade, status = result

    # Get marks once for both MySQL and Excel
    try:
        maths = int(math_entry.get())
        science = int(science_entry.get())
        english = int(english_entry.get())
    except ValueError:
        return

    conn = connect_mysql()

    if conn is None:
        return

    try:
        cursor = conn.cursor()

        sql = """
        INSERT INTO students
        (roll_no, name, class_name,
         maths, science, english,
         total, percentage, grade, status)
        VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
        """

        values = (
            roll_no,
            name,
            class_name,
            maths,
            science,
            english,
            total,
            percentage,
            grade,
            status
        )

        cursor.execute(sql, values)
        conn.commit()

        cursor.close()
        conn.close()

        # Also save to Excel
        save_to_excel(
            roll_no,
            name,
            class_name,
            maths,
            science,
            english,
            total,
            percentage,
            grade,
            status
        )

        messagebox.showinfo(
            "Success",
            "Student result saved to MySQL and Excel!"
        )

        clear_form()
        load_data()

    except mysql.connector.Error as e:
        messagebox.showerror(
            "MySQL Error",
            str(e)
        )

        if conn.is_connected():
            conn.close()


# SAVE DATA TO EXCEL
def save_to_excel(
    roll_no,
    name,
    class_name,
    maths,
    science,
    english,
    total,
    percentage,
    grade,
    status
):
    filename = "student_results.xlsx"

    try:
        # Create Excel file if it doesn't exist
        if not os.path.exists(filename):
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "Student Results"

            sheet.append([
                "Roll No",
                "Name",
                "Class",
                "Maths",
                "Science",
                "English",
                "Total",
                "Percentage",
                "Grade",
                "Status"
            ])

            workbook.save(filename)

        # Open existing Excel file
        workbook = load_workbook(filename)

        if "Student Results" not in workbook.sheetnames:
            sheet = workbook.create_sheet("Student Results")

            sheet.append([
                "Roll No",
                "Name",
                "Class",
                "Maths",
                "Science",
                "English",
                "Total",
                "Percentage",
                "Grade",
                "Status"
            ])
        else:
            sheet = workbook["Student Results"]

        sheet.append([
            roll_no,
            name,
            class_name,
            maths,
            science,
            english,
            total,
            percentage,
            grade,
            status
        ])

        workbook.save(filename)
        workbook.close()

    except Exception as e:
        messagebox.showerror(
            "Excel Error",
            str(e)
        )


# DISPLAY MYSQL DATA

def load_data():
    # Clear old table data
    for item in tree.get_children():
        tree.delete(item)

    conn = connect_mysql()

    if conn is None:
        return

    try:
        cursor = conn.cursor()

        cursor.execute("""
            SELECT roll_no, name, class_name,
                   maths, science, english,
                   total, percentage, grade, status
            FROM students
            ORDER BY id DESC
        """)

        rows = cursor.fetchall()

        for row in rows:
            tree.insert("", tk.END, values=row)

        cursor.close()
        conn.close()

    except mysql.connector.Error as e:
        messagebox.showerror(
            "Database Error",
            str(e)
        )

        if conn.is_connected():
            conn.close()


# CLEAR FORM

def clear_form():
    roll_entry.delete(0, tk.END)
    name_entry.delete(0, tk.END)
    class_entry.delete(0, tk.END)

    math_entry.delete(0, tk.END)
    science_entry.delete(0, tk.END)
    english_entry.delete(0, tk.END)


# TKINTER WINDOW


root = tk.Tk()

root.title("Student Result Management System")
root.geometry("1100x650")
root.configure(bg="#eef2f7")
# TITLE

title = tk.Label(
    root,
    text="STUDENT RESULT MANAGEMENT SYSTEM",
    font=("Arial", 22, "bold"),
    bg="#1f4e78",
    fg="white",
    pady=15
)

title.pack(fill=tk.X)



# FORM FRAME

form_frame = tk.Frame(
    root,
    bg="white",
    padx=20,
    pady=20
)

form_frame.pack(
    fill=tk.X,
    padx=20,
    pady=20
)


# ROLL NO

tk.Label(
    form_frame,
    text="Roll No",
    bg="white",
    font=("Arial", 11)
).grid(row=0, column=0, padx=10, pady=8)

roll_entry = tk.Entry(
    form_frame,
    width=25
)

roll_entry.grid(
    row=0,
    column=1,
    padx=10,
    pady=8
)



# NAME

tk.Label(
    form_frame,
    text="Student Name",
    bg="white",
    font=("Arial", 11)
).grid(row=0, column=2, padx=10, pady=8)

name_entry = tk.Entry(
    form_frame,
    width=25
)

name_entry.grid(
    row=0,
    column=3,
    padx=10,
    pady=8
)


# CLASS

tk.Label(
    form_frame,
    text="Class",
    bg="white",
    font=("Arial", 11)
).grid(row=1, column=0, padx=10, pady=8)

class_entry = tk.Entry(
    form_frame,
    width=25
)

class_entry.grid(
    row=1,
    column=1,
    padx=10,
    pady=8
)

# MATHS

tk.Label(
    form_frame,
    text="Maths",
    bg="white",
    font=("Arial", 11)
).grid(row=1, column=2, padx=10, pady=8)

math_entry = tk.Entry(
    form_frame,
    width=25
)

math_entry.grid(
    row=1,
    column=3,
    padx=10,
    pady=8
    )
# SCIENCE


tk.Label(
    form_frame,
    text="Science",
    bg="white",
    font=("Arial", 11)
).grid(row=2, column=0, padx=10, pady=8)

science_entry = tk.Entry(
    form_frame,
    width=25
)

science_entry.grid(
    row=2,
    column=1,
    padx=10,
    pady=8
)
# ENGLISH

tk.Label(
    form_frame,
    text="English",
    bg="white",
    font=("Arial", 11)
).grid(row=2, column=2, padx=10, pady=8)

english_entry = tk.Entry(
    form_frame,
    width=25
)

english_entry.grid(
    row=2,
    column=3,
    padx=10,
    pady=8
)

# BUTTONS

button_frame = tk.Frame(
    root,
    bg="#eef2f7"
)

button_frame.pack(pady=5)


tk.Button(
    button_frame,
    text="SAVE RESULT",
    command=save_student,
    bg="#198754",
    fg="white",
    font=("Arial", 11, "bold"),
    width=18
).grid(row=0, column=0, padx=10)


tk.Button(
    button_frame,
    text="CLEAR",
    command=clear_form,
    bg="#dc3545",
    fg="white",
    font=("Arial", 11, "bold"),
    width=18
).grid(row=0, column=1, padx=10)


tk.Button(
    button_frame,
    text="REFRESH",
    command=load_data,
    bg="#0d6efd",
    fg="white",
    font=("Arial", 11, "bold"),
    width=18
).grid(row=0, column=2, padx=10)

# TABLE

table_frame = tk.Frame(root)

table_frame.pack(
    fill=tk.BOTH,
    expand=True,
    padx=20,
    pady=15
)

columns = (
    "Roll",
    "Name",
    "Class",
    "Maths",
    "Science",
    "English",
    "Total",
    "Percentage",
    "Grade",
    "Status"
)

tree = ttk.Treeview(
    table_frame,
    columns=columns,
    show="headings"
)

for column in columns:
    tree.heading(
        column,
        text=column
    )

    tree.column(
        column,
        width=100,
        anchor=tk.CENTER
    )

tree.pack(
    fill=tk.BOTH,
    expand=True
)

# START APPLICATION


load_data()

root.mainloop()
