import sqlite3

db_locale = 'students.db'

connie = sqlite3.connect(db_locale)
c=connie.cursor()

c.execute("""
SELECT * FROM contact_details
""")

student_info = c.fetchall()
print(student_info)
connie.commit()
connie.close()