from flask import Flask, render_template, request, send_file
import pandas as pd
from io import BytesIO
from datetime import datetime, timedelta

app = Flask(__name__)

# ---------------- Home Page ----------------
@app.route('/')
def home():
    today = datetime.now().strftime("%d-%m-%Y")
    return render_template('index.html', date=today)

# ---------------- Sitting Plan ----------------
@app.route('/upload', methods=['POST'])
def upload_files():
    student = pd.read_csv(request.files['student'])
    teacher = pd.read_csv(request.files['teacher'])
    room = pd.read_csv(request.files['room'])

    # Random allocation
    sitting = pd.DataFrame({
        'Date': [datetime.now().strftime("%d-%m-%Y")] * len(student),
        'Roll_No': student['Roll_No'],

        'Department': student['Department'],
        'Semester': student['Semester'],
        'Room_No': room['Room_No'].sample(n=len(student), replace=True).values,
        'Row_No': [(i % 10) + 1 for i in range(len(student))],
        'Desk_No': [(i % 5) + 1 for i in range(len(student))],
        'Invigilator': teacher['Teacher_Name'].sample(n=len(student), replace=True).values
    })

    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        sitting.to_excel(writer, index=False, sheet_name='Sitting_Plan')

    output.seek(0)
    return send_file(output, download_name='Sitting_Plan.xlsx', as_attachment=True)

# ---------------- Teacher Duty ----------------
@app.route('/duty', methods=['POST'])
def generate_duty():
    teacher = pd.read_csv(request.files['teacher'])
    days = ['Day 1', 'Day 2', 'Day 3']
    shifts = ['Morning', 'Evening']

    duty_data = []
    for t in teacher['Teacher_Name']:
        for day in days:
            for shift in shifts:
                duty_data.append({'Teacher_Name': t, 'Day': day, 'Shift': shift})

    duty = pd.DataFrame(duty_data)

    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        duty.to_excel(writer, index=False, sheet_name='Teacher_Duty')

    output.seek(0)
    return send_file(output, download_name='Teacher_Duty.xlsx', as_attachment=True)

# ---------------- DateSheet ----------------
@app.route('/datesheet', methods=['POST'])
def generate_datesheet():
    subjects = pd.read_csv(request.files['subjects'])

    start_date = datetime.now()
    dates = [(start_date + timedelta(days=i)).strftime("%d-%m-%Y") for i in range(len(subjects))]

    # Morning & Evening shifts for each day
    sessions = ['Morning', 'Evening'] * (len(subjects) // 2 + 1)
    sessions = sessions[:len(subjects)]

    datesheet = pd.DataFrame({
        'Date': dates,
        'Shift': sessions,
        'Branch': subjects['Branch'],
        'Semester': subjects['Semester'],
        'Subject_Code': subjects['Subject_Code'],
        'Subject_Name': subjects['Subject_Name']
    })

    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        datesheet.to_excel(writer, index=False, sheet_name='DateSheet')

    output.seek(0)
    return send_file(output, download_name='DateSheet.xlsx', as_attachment=True)

# ---------------- Run App ----------------
if __name__ == '__main__':
    app.run(debug=True)
