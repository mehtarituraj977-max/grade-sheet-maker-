import os
import datetime
import pandas as pd
from flask import Flask, render_template, request, redirect, url_for

app = Flask(__name__)

# Determine the Excel filename for this session
DATA_DIR = 'data'
if not os.path.exists(DATA_DIR):
    os.makedirs(DATA_DIR)

session_time = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
EXCEL_FILE = os.path.join(DATA_DIR, f'marksheet_{session_time}.xlsx')

# Columns for our dataframe
COLUMNS = ['ID', 'Name', 'Math', 'Science', 'English', 'History', 'Art', 'Total', 'Percentage', 'Rank']

def init_excel():
    """Initializes a new Excel file with the required columns."""
    df = pd.DataFrame(columns=COLUMNS)
    df.to_excel(EXCEL_FILE, index=False)
    print(f"Started new session. Created {EXCEL_FILE}")

def load_data():
    """Loads data from the current Excel file."""
    if os.path.exists(EXCEL_FILE):
         return pd.read_excel(EXCEL_FILE)
    return pd.DataFrame(columns=COLUMNS)

def save_data(df):
    """Saves the dataframe to the current Excel file and updates ranks."""
    # Recalculate ranks before saving
    if not df.empty:
        df['Rank'] = df['Total'].rank(method='min', ascending=False).astype(int)
        df = df.sort_values(by='Rank')
    df.to_excel(EXCEL_FILE, index=False)

# Initialize the file on startup
init_excel()


@app.route('/')
def dashboard():
    df = load_data()
    students = df.to_dict('records')
    
    # Get top 3
    top_3 = []
    if not df.empty:
        # Since it's sorted by rank in save_data, we can just take the top ones loosely
        # or exactly match Rank 1, 2, 3
        top_3 = df[df['Rank'] <= 3].to_dict('records')
        
    return render_template('index.html', students=students, top_3=top_3)

@app.route('/add', methods=['GET', 'POST'])
def add_student():
    if request.method == 'POST':
        name = request.form['name']
        math = float(request.form['math'])
        science = float(request.form['science'])
        english = float(request.form['english'])
        history = float(request.form['history'])
        art = float(request.form['art'])

        total = math + science + english + history + art
        percentage = (total / 500) * 100

        df = load_data()
        
        # Determine new ID
        new_id = 1 if df.empty else df['ID'].max() + 1
        
        new_row = {
            'ID': new_id,
            'Name': name,
            'Math': math,
            'Science': science,
            'English': english,
            'History': history,
            'Art': art,
            'Total': total,
            'Percentage': percentage,
            'Rank': 0 # Will be recalculated on save
        }
        
        df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
        save_data(df)
        
        return redirect(url_for('report', student_id=new_id))

    return render_template('add_student.html')

@app.route('/report/<int:student_id>')
def report(student_id):
    df = load_data()
    student_data = df[df['ID'] == student_id]
    
    if student_data.empty:
        return "Student not found", 404
        
    student = student_data.iloc[0].to_dict()
    return render_template('report_card.html', student=student)

if __name__ == '__main__':
    app.run(debug=True)
