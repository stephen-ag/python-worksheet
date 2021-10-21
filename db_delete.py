import sqlite3

db_locale = 'students.db'

connie = sqlite3.connect(db_locale)
c=connie.cursor()


c.execute("""
DROP TABLE contact_details
""")

connie.commit()
connie.close()