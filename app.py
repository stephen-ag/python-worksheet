from flask import Flask, render_template, request

import sqlite3

app = Flask(__name__)
db_locale = 'students.db'

@app.route('/')
@app.route('/home')
def home_page():
    student_data= query_contact_details()
    return render_template('index.html',student_data=student_data)



@app.route('/update')
def update_student():
    return render_template('update_student.html')

@app.route('/add', methods = ['GET','POST'])
def add_student():
    if request.method == 'GET':
        return render_template('add_student.html')
    else:
        students_details = (
            request.form['first_name'],
            request.form['last_name'],
            request.form['street_address'],
            request.form['suburb']
        )
        insert_students(students_details)
        return render_template('add_sucess.html')



def insert_students(student_details):
    connie = sqlite3.connect(db_locale)
    c = connie.cursor()
    sql_execute_string='INSERT INTO contact_details (first_name, last_name, street_address, suburb) VALUES(?,?,?,?)'

    c.execute(sql_execute_string, student_details)
    connie.commit()
    connie.close()


def query_contact_details():
    connie= sqlite3.connect(db_locale)
    c = connie.cursor()
    c.execute("""
    SELECT * FROM contact_details
    
    """)
    student_data = c.fetchall()
    return(student_data)


if __name__ == '__main__':
    app.run()