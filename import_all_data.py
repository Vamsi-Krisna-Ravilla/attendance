import pandas as pd
import sqlite3
from datetime import datetime

def init_db():
    conn = sqlite3.connect('attendance.db')
    c = conn.cursor()
    
    # Create Students table
    c.execute('''CREATE TABLE IF NOT EXISTS students
                 (ht_number TEXT PRIMARY KEY, name TEXT, original_section TEXT, merged_section TEXT)''')
    
    # Create Faculty table
    c.execute('''CREATE TABLE IF NOT EXISTS faculty
                 (id INTEGER PRIMARY KEY, name TEXT, username TEXT UNIQUE, password TEXT)''')
    
    # Create Attendance table
    c.execute('''CREATE TABLE IF NOT EXISTS attendance
                 (id INTEGER PRIMARY KEY, ht_number TEXT, date TEXT, period TEXT, status TEXT, 
                  faculty TEXT, subject TEXT, lesson_plan TEXT)''')
    
    # Create Section-Subject-Mapping table with modified structure
    c.execute('''DROP TABLE IF EXISTS section_subject_mapping''')
    c.execute('''CREATE TABLE section_subject_mapping
                 (id INTEGER PRIMARY KEY, section TEXT, subject_name TEXT)''')
    
    conn.commit()
    conn.close()

def import_excel_to_sqlite(excel_file):
    # First initialize the database
    init_db()
    
    # Connect to SQLite database
    conn = sqlite3.connect('attendance.db')
    cursor = conn.cursor()
    
    try:
        # Import Faculty data
        df_faculty = pd.read_excel(excel_file, sheet_name='Faculty')
        faculty_data = []
        for _, row in df_faculty.iterrows():
            if pd.notna(row['Username']) and pd.notna(row['Password']):
                faculty_data.append({
                    'name': str(row['Faculty Name']),
                    'username': str(row['Username']).strip(),
                    'password': str(row['Password']).strip()
                })
        
        # Import Students data
        df_students = pd.read_excel(excel_file, sheet_name='Students')
        students_data = []
        for _, row in df_students.iterrows():
            if pd.notna(row['HT Number']):
                students_data.append({
                    'ht_number': str(row['HT Number']),
                    'name': str(row['Student Name']),
                    'original_section': str(row['Original Section']),
                    'merged_section': str(row['Merged Section'])
                })
        
        # Import Section-Subject-Mapping data with proper newline handling
        df_mapping = pd.read_excel(excel_file, sheet_name='Section-Subject-Mapping')
        mapping_data = []
        for _, row in df_mapping.iterrows():
            if pd.notna(row['Section']) and pd.notna(row['Subject Names']):
                section = str(row['Section'])
                # Split subjects by newline and create separate entries
                subjects = str(row['Subject Names']).split('\n')
                for subject in subjects:
                    subject = subject.strip()
                    if subject:  # Only add non-empty subjects
                        mapping_data.append({
                            'section': section,
                            'subject_name': subject
                        })
        
        # Insert data into tables
        cursor.executemany("""
            INSERT OR REPLACE INTO faculty (name, username, password)
            VALUES (:name, :username, :password)
        """, faculty_data)
        
        cursor.executemany("""
            INSERT OR REPLACE INTO students (ht_number, name, original_section, merged_section)
            VALUES (:ht_number, :name, :original_section, :merged_section)
        """, students_data)
        
        # Insert mapping data with individual subject entries
        cursor.executemany("""
            INSERT INTO section_subject_mapping (section, subject_name)
            VALUES (:section, :subject_name)
        """, mapping_data)
        
        conn.commit()
        print(f"Successfully imported:")
        print(f"- {len(faculty_data)} faculty records")
        print(f"- {len(students_data)} student records")
        print(f"- {len(mapping_data)} individual subject mappings")
        
        # Print sample records for verification
        print("\nSample Section-Subject Mappings:")
        cursor.execute("SELECT * FROM section_subject_mapping LIMIT 5")
        print(cursor.fetchall())
        
    except Exception as e:
        print(f"Error importing data: {str(e)}")
        conn.rollback()
    finally:
        conn.close()

# Also update the get_section_subjects function in your app.py
def get_section_subjects(section):
    conn = sqlite3.connect('attendance.db')
    cursor = conn.cursor()
    cursor.execute("SELECT subject_name FROM section_subject_mapping WHERE section = ?", (section,))
    subjects = cursor.fetchall()
    conn.close()
    return [subject[0] for subject in subjects]

if __name__ == "__main__":
    excel_file = "attendance.xlsx"
    import_excel_to_sqlite(excel_file)