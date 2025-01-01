import sqlite3
from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker
from app import Base, Student, Faculty, Section, Subject, Attendance

# Database connection
SQLALCHEMY_DATABASE_URL = "sqlite:///./attendance.db"
engine = create_engine(SQLALCHEMY_DATABASE_URL)
SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)

def view_table_data(table_name):
    db = SessionLocal()
    try:
        if table_name == "students":
            data = db.query(Student).all()
            for student in data:
                print(f"HT Number: {student.ht_number}, Name: {student.name}, Course: {student.course}, "
                      f"Original Section: {student.original_section}, Merged Section: {student.merged_section}")
        elif table_name == "faculty":
            data = db.query(Faculty).all()
            for faculty in data:
                print(f"ID: {faculty.id}, Name: {faculty.name}, Username: {faculty.username}")
        elif table_name == "sections":
            data = db.query(Section).all()
            for section in data:
                print(f"ID: {section.id}, Name: {section.name}")
        elif table_name == "subjects":
            data = db.query(Subject).all()
            for subject in data:
                print(f"ID: {subject.id}, Name: {subject.name}")
        elif table_name == "attendance":
            data = db.query(Attendance).limit(50).all()  # Limiting to 50 to avoid overwhelming output
            for attendance in data:
                print(f"ID: {attendance.id}, Student: {attendance.student_id}, Date: {attendance.date}, "
                      f"Period: {attendance.period}, Status: {attendance.status}, Subject: {attendance.subject}")
        else:
            print(f"Table '{table_name}' not recognized.")
    finally:
        db.close()

def main():
    while True:
        print("\nAvailable tables: students, faculty, sections, subjects, attendance")
        table_name = input("Enter the name of the table you want to view (or 'quit' to exit): ").lower()
        
        if table_name == 'quit':
            break
        
        view_table_data(table_name)

if __name__ == "__main__":
    main()