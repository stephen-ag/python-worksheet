import sqlite3

db_locale = 'students.db'

connie = sqlite3.connect(db_locale)
c=connie.cursor()


c.execute("""
INSERT INTO contact_details (first_name, last_name, street_address, suburb) VALUES
('David', 'rosario', '11 pascal street', 'duston'),
('Gerald', 'stephen', '39 harrison ave', 'illinois'),
('sharron', 'quen', '12 silver lane', 'rockford'),
('sheryn', 'kingson', '6 golden lane', 'illinois')
""")

connie.commit()
connie.close()