from flask import Flask, render_template,request,redirect,flash,session,send_file
from functools import wraps
import pyodbc
import os

from openpyxl import Workbook
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import Paragraph
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

app = Flask(__name__)
app.secret_key = "employee_management_secret "
def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):

        if "username" not in session:
            flash("Please login first.", "warning")
            return redirect("/login")

        return f(*args, **kwargs)

    return decorated_function

@app.route("/")
@login_required
def home():
    if "username" not in session:
        flash("Please login first!", "warning")
        return redirect("/login") 
    conn= pyodbc.connect(
            "DRIVER={ODBC Driver 17 for SQL Server};"
            "SERVER=STUNNER;"
            "DATABASE=StudentDB;"
            "Trusted_connection=yes;"
        )
    cursor = conn.cursor()
    search = request.args.get("search","")
    if search:
        cursor.execute(
            """
            SELECT * FROM Employee 
            WHERE name LIKE ?
            OR department LIKE ?
            """,
            ('%' + search + '%', '%' + search + '%',)
            )
    else:
         cursor.execute("SELECT * FROM Employee")
    rows = cursor.fetchall()
    employee_count = len(rows)
    cursor.execute("""
    SELECT
        MAX(salary),
        MIN(salary),
        AVG(salary)
    FROM Employee
    """)

    stats = cursor.fetchone()

    highest_salary = stats[0] if stats[0] is not None else 0
    lowest_salary = stats[1] if stats[1] is not None else 0
    average_salary = round(stats[2], 2) if stats[2] is not None else 0
    cursor.execute("""
    SELECT COUNT(DISTINCT department)
    FROM Employee
    """)

    department_count = cursor.fetchone()[0]
    conn.close()

    return render_template("index.html", employees=rows, search=search, 
                           employee_count=employee_count, 
                           highest_salary=highest_salary, 
                           lowest_salary=lowest_salary, 
                           average_salary=average_salary,
                           department_count=department_count)
@app.route("/add",methods=["GET","POST"])
@login_required
def add_employee():
    if "username" not in session:
        flash("Please login first!", "warning")
        return redirect("/login") 
    if request.method == "POST":
        name = request.form["name"]
        department = request.form["department"]
        salary = request.form["salary"]
        photo = request.files["photo"]
        filename = ""
        if photo.filename != "":
            filename = photo.filename

            photo.save(
                os.path.join(
                    "static",
                    "uploads",
                    filename
                )
            )
        if not name.strip():
            flash("Name cannot be empty", "danger")
            return redirect("/add")
        if not department.strip():
            flash("Department cannot be empty", "danger")
            return redirect("/add")
        if int(salary) <= 0:
            flash("Sallary must be greater than 0", "danger")
            return redirect("/add")
        conn= pyodbc.connect(
            "DRIVER={ODBC Driver 17 for SQL Server};"
            "SERVER=STUNNER;"
            "DATABASE=StudentDB;"
            "Trusted_connection=yes;"
        )
        cursor = conn.cursor()
        cursor.execute(
            """
        INSERT INTO Employee(name, department, salary, photo)
        VALUES (?, ?, ?, ?)
        """,
        (name, department, salary, filename)
        )
        conn.commit()
        conn.close()

        flash("Employee Added Successfully!", "success")
        return redirect("/")
       
    return render_template("add_employee.html")

@app.route("/edit/<int:id>", methods=["GET", "POST"])
@login_required
def edit_employee(id):
    if "username" not in session:
        flash("Please login first!", "warning")
        return redirect("/login") 
    conn = pyodbc.connect(
        "DRIVER={ODBC Driver 17 for SQL Server};"
        "SERVER=STUNNER;"
        "DATABASE=StudentDB;"
        "Trusted_Connection=yes;"
    )
    cursor = conn.cursor()
    if request.method == "POST":
        name = request.form["name"]
        department = request.form["department"]
        salary = request.form["salary"]

        cursor.execute(
        """
            UPDATE Employee
            SET name=?, department=?, salary=?
            WHERE id=?
            """,
            (name, department, salary, id)
    )
        conn.commit()
        conn.close()
        flash("Employee Updated Successfully", "info")

        return redirect("/")
    
    cursor.execute(
        "SELECT * FROM Employee WHERE id=?",
        (id,)
    )

    employee = cursor.fetchone()
    conn.close()

    return render_template("edit_employee.html", employee=employee)
@app.route("/delete/<int:id>")
@login_required
def delete_employee(id):
    if "username" not in session:
        flash("Please login first!", "warning")
        return redirect("/login") 
    conn = pyodbc.connect(
        "DRIVER={ODBC Driver 17 for SQL Server};"
        "SERVER=STUNNER;"
        "DATABASE=StudentDB;"
        "Trusted_Connection=yes;"
    )
    cursor = conn.cursor()
    cursor.execute(
        "DELETE FROM Employee WHERE id=?",
        (id,)
    )
    conn.commit()
    conn.close()
    flash("Employee Deleted Successfully","warning")
    
    return redirect("/")
@app.route("/login",methods=["GET", "POST"])
def login():
    if "username" in session:
        return redirect("/")
    if request.method =="POST":
        username = request.form.get("username")
        password = request.form.get("password")
        conn = pyodbc.connect(
            "DRIVER={ODBC Driver 17 for SQL Server};"
            "SERVER=STUNNER;"
            "DATABASE=StudentDB;"
            "Trusted_connection=yes;"
        )

        cursor = conn.cursor()
        cursor.execute(
            "SELECT * FROM Users WHERE username=? and password=?",
            (username,password)
        )
        user = cursor.fetchone()
        if user:
            session["username"] = username
            flash("Login Successful!","success")
            return redirect("/")
        else:
            flash("Invalid Username or Password","danger")

    return render_template("login.html")

@app.route("/logout")
def logout():
    session.pop("username", None)
    flash("Logged out successfully!","info")
    return redirect("/login")

@app.route("/export/excel")
@login_required
def export_excel():

    conn = pyodbc.connect(
        "DRIVER={ODBC Driver 17 for SQL Server};"
        "SERVER=STUNNER;"
        "DATABASE=StudentDB;"
        "Trusted_Connection=yes;"
    )

    cursor = conn.cursor()

    cursor.execute("""
        SELECT id, name, department, salary
        FROM Employee
    """)

    employees = cursor.fetchall()

    conn.close()

    wb = Workbook()

    ws = wb.active

    ws.title = "Employees"

    # Header Row
    ws.append([
        "Employee ID",
        "Name",
        "Department",
        "Salary"
    ])

    # Employee Data
    for emp in employees:

        ws.append([
            emp.id,
            emp.name,
            emp.department,
            emp.salary
        ])

    filename = "employees.xlsx"

    wb.save(filename)

    return send_file(
        filename,
        as_attachment=True
    )

@app.route("/export/pdf")
@login_required
def export_pdf():

    conn = pyodbc.connect(
        "DRIVER={ODBC Driver 17 for SQL Server};"
        "SERVER=STUNNER;"
        "DATABASE=StudentDB;"
        "Trusted_connection=yes;"
    )

    cursor = conn.cursor()
    cursor.execute("SELECT * FROM Employee")
    employees = cursor.fetchall()
    conn.close()

    filename = "employees.pdf"

    c = canvas.Canvas(filename, pagesize=letter)

    c.setFont("Helvetica-Bold", 16)
    c.drawString(200, 770, "Employee Report")

    c.setFont("Helvetica-Bold", 11)
    c.drawString(40, 740, "ID")
    c.drawString(80, 740, "Name")
    c.drawString(220, 740, "Department")
    c.drawString(380, 740, "Salary")

    y = 720

    c.setFont("Helvetica", 11)

    for emp in employees:

        c.drawString(40, y, str(emp.id))
        c.drawString(80, y, emp.name)
        c.drawString(220, y, emp.department)
        c.drawString(380, y, str(emp.salary))

        y -= 20

        # Start a new page if the current one is full
        if y < 50:
            c.showPage()
            c.setFont("Helvetica", 11)
            y = 770

    c.save()

    return send_file(
        filename,
        as_attachment=True
    )

@app.route("/employee/<int:id>")
def employee_details(id):

    conn = pyodbc.connect(
        "DRIVER={ODBC Driver 17 for SQL Server};"
        "SERVER=STUNNER;"
        "DATABASE=StudentDB;"
        "Trusted_Connection=yes;"
    )

    cursor = conn.cursor()

    cursor.execute(
        "SELECT * FROM Employee WHERE id=?",
        (id,)
    )

    employee = cursor.fetchone()

    conn.close()

    return render_template(
        "employee_details.html",
        employee=employee
    )

@app.errorhandler(404)
def page_not_found(e):
    return render_template("404.html"), 404

@app.errorhandler(500)
def internal_server_error(e):
    return render_template("500.html"), 500

if __name__ == "__main__":
    app.run(debug=True)
    