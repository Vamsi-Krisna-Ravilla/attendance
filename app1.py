import streamlit as st
import pandas as pd
import sqlite3
from datetime import datetime, timedelta

# Database Operations
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
    
    # Create Section-Subject-Mapping table
    c.execute('''CREATE TABLE IF NOT EXISTS section_subject_mapping
                 (id INTEGER PRIMARY KEY, section TEXT, subject_name TEXT)''')
    
    # Add default admin user if not exists
    c.execute("INSERT OR IGNORE INTO faculty (username, password, name) VALUES (?, ?, ?)", 
             ('admin', 'admin', 'Administrator'))
    
    conn.commit()
    conn.close()

def add_student(ht_number, name, original_section, merged_section):
    conn = sqlite3.connect('attendance.db')
    c = conn.cursor()
    c.execute("INSERT OR REPLACE INTO students VALUES (?, ?, ?, ?)", 
              (ht_number, name, original_section, merged_section))
    conn.commit()
    conn.close()

def add_faculty(name, username, password):
    conn = sqlite3.connect('attendance.db')
    c = conn.cursor()
    c.execute("INSERT OR REPLACE INTO faculty (name, username, password) VALUES (?, ?, ?)", 
              (name, username, password))
    conn.commit()
    conn.close()

def mark_attendance(ht_number, date, period, status, faculty, subject, lesson_plan):
    conn = sqlite3.connect('attendance.db')
    c = conn.cursor()
    c.execute("INSERT INTO attendance (ht_number, date, period, status, faculty, subject, lesson_plan) VALUES (?, ?, ?, ?, ?, ?, ?)", 
              (ht_number, date, period, status, faculty, subject, lesson_plan))
    conn.commit()
    conn.close()

def get_students(section=None):
    conn = sqlite3.connect('attendance.db')
    if section:
        df = pd.read_sql_query("SELECT * FROM students WHERE merged_section = ?", conn, params=(section,))
    else:
        df = pd.read_sql_query("SELECT * FROM students", conn)
    conn.close()
    return df

def get_faculty():
    conn = sqlite3.connect('attendance.db')
    df = pd.read_sql_query("SELECT * FROM faculty", conn)
    conn.close()
    return df

def get_attendance(start_date, end_date=None, section=None):
    conn = sqlite3.connect('attendance.db')
    query = "SELECT * FROM attendance WHERE date >= ?"
    params = [start_date]
    if end_date:
        query += " AND date <= ?"
        params.append(end_date)
    if section:
        query += " AND ht_number IN (SELECT ht_number FROM students WHERE merged_section = ?)"
        params.append(section)
    df = pd.read_sql_query(query, conn, params=params)
    conn.close()
    return df

def get_section_subjects(section):
    conn = sqlite3.connect('attendance.db')
    cursor = conn.cursor()
    cursor.execute("SELECT DISTINCT subject_name FROM section_subject_mapping WHERE section = ? ORDER BY subject_name", (section,))
    subjects = cursor.fetchall()
    conn.close()
    return [subject[0] for subject in subjects]

# Page Functions
def login():
    st.title("Login")
    
    # Center the login form
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.image("https://rvit.edu.in/rvitlogo_f.jpg", width=200)
        login_type = st.radio("Select Login Type", ["Faculty", "Admin"])
        username = st.text_input("Username").strip()  # Remove whitespace
        password = st.text_input("Password", type="password").strip()  # Remove whitespace
        
        if st.button("Login", type="primary", use_container_width=True):
            conn = sqlite3.connect('attendance.db')
            cursor = conn.cursor()
            
            # Use parameterized query to prevent SQL injection
            cursor.execute("""
                SELECT * FROM faculty 
                WHERE username = ? AND password = ?
            """, (username, password))
            
            user = cursor.fetchone()
            conn.close()
            
            if user:
                st.session_state.logged_in = True
                st.session_state.username = username
                st.session_state.faculty_name = user[1]  # faculty name is in index 1
                
                # Check if the user is an admin
                if login_type == "Admin" and username == "rvk":
                    st.session_state.is_admin = True
                else:
                    st.session_state.is_admin = False
                
                st.rerun()
            else:
                st.error("Invalid credentials")


def mark_attendance_page():
    st.subheader("Mark Attendance")
    
    # Get unique sections
    students_df = get_students()
    if students_df.empty:
        st.warning("No students found in the database. Please add students first.")
        return
        
    sections = students_df['merged_section'].unique()
    
    col1, col2, col3 = st.columns(3)
    with col1:
        section = st.selectbox("Select Section", sections)
    with col2:
        period = st.selectbox("Select Period", ['P1', 'P2', 'P3', 'P4', 'P5', 'P6'])
    with col3:
        subject = st.selectbox("Select Subject", get_section_subjects(section))
    
    if section and period and subject:
        students = get_students(section)
        
        # Add "Mark All Present" button
        col1, col2 = st.columns(2)
        with col1:
            mark_all = st.button("Mark All Present", type="primary")
        with col2:
            mark_none = st.button("Mark All Absent", type="secondary")
        
        attendance_data = {}
        
        # Create a container for student checkboxes
        with st.container():
            for _, student in students.iterrows():
                col1, col2 = st.columns([3, 1])
                with col1:
                    st.write(f"{student['name']} ({student['ht_number']})")
                with col2:
                    # Set default value based on "Mark All" buttons
                    default_value = mark_all if mark_all else (not mark_none if mark_none else False)
                    status = st.checkbox("Present", value=default_value, key=f"attendance_{student['ht_number']}")
                    attendance_data[student['ht_number']] = 'P' if status else 'A'
        
        lesson_plan = st.text_area("Lesson Plan", placeholder="Enter the topic covered in this class...")
        
        if st.button("Submit Attendance", type="primary"):
            if not lesson_plan.strip():
                st.error("Please enter a lesson plan before submitting attendance.")
            else:
                date = datetime.now().strftime('%Y-%m-%d')
                for ht_number, status in attendance_data.items():
                    mark_attendance(ht_number, date, period, status, 
                                 st.session_state.faculty_name, subject, lesson_plan)
                st.success("Attendance marked successfully!")
                # Clear the form
                st.rerun()


# Add these functions after your existing database functions

def get_subject_analysis(start_date, end_date, section=None):
    conn = sqlite3.connect('attendance.db')
    query = """
        SELECT 
            subject,
            COUNT(DISTINCT date || period) as total_classes,
            SUM(CASE WHEN status = 'P' THEN 1 ELSE 0 END) as present_count,
            COUNT(*) as total_attendance,
            ROUND(CAST(SUM(CASE WHEN status = 'P' THEN 1 ELSE 0 END) AS FLOAT) / COUNT(*) * 100, 2) as attendance_percentage
        FROM attendance 
        WHERE date BETWEEN ? AND ?
    """
    params = [start_date, end_date]
    
    if section:
        query += " AND ht_number IN (SELECT ht_number FROM students WHERE merged_section = ?)"
        params.append(section)
    
    query += " GROUP BY subject"
    
    df = pd.read_sql_query(query, conn, params=params)
    conn.close()
    return df

def get_student_report(ht_number, start_date, end_date):
    conn = sqlite3.connect('attendance.db')
    
    # Get student details
    student_query = "SELECT * FROM students WHERE ht_number = ?"
    student_df = pd.read_sql_query(student_query, conn, params=[ht_number])
    
    # Get attendance details
    attendance_query = """
        SELECT 
            subject,
            COUNT(DISTINCT date || period) as total_classes,
            SUM(CASE WHEN status = 'P' THEN 1 ELSE 0 END) as present_count,
            ROUND(CAST(SUM(CASE WHEN status = 'P' THEN 1 ELSE 0 END) AS FLOAT) / COUNT(*) * 100, 2) as attendance_percentage
        FROM attendance 
        WHERE ht_number = ? AND date BETWEEN ? AND ?
        GROUP BY subject
    """
    attendance_df = pd.read_sql_query(attendance_query, conn, params=[ht_number, start_date, end_date])
    
    # Get daily attendance
    daily_query = """
        SELECT date, period, subject, status, faculty, lesson_plan
        FROM attendance 
        WHERE ht_number = ? AND date BETWEEN ? AND ?
        ORDER BY date DESC, period ASC
    """
    daily_df = pd.read_sql_query(daily_query, conn, params=[ht_number, start_date, end_date])
    
    conn.close()
    return student_df, attendance_df, daily_df

# Modify the view_statistics_page function to include the new tabs
def view_statistics_page():
    st.subheader("View Statistics")
    
    tab1, tab2, tab3 = st.tabs(["Overall Statistics", "Subject Analysis", "Student Reports"])
    
    col1, col2 = st.columns(2)
    with col1:
        start_date = st.date_input("Start Date")
    with col2:
        end_date = st.date_input("End Date")
    
    students_df = get_students()
    if students_df.empty:
        st.warning("No students found in the database.")
        return
        
    section = st.selectbox("Select Section", ['All'] + list(students_df['merged_section'].unique()))
    
    with tab1:
        if st.button("Generate Overall Report", type="primary"):
            attendance_df = get_attendance(start_date, end_date, None if section == 'All' else section)
            
            if not attendance_df.empty:
                # Calculate statistics
                total_students = len(attendance_df['ht_number'].unique())
                total_classes = len(attendance_df['date'].unique())
                present_count = len(attendance_df[attendance_df['status'] == 'P'])
                total_count = len(attendance_df)
                avg_attendance = (present_count / total_count * 100) if total_count > 0 else 0
                
                # Display metrics
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric("Total Students", total_students)
                with col2:
                    st.metric("Total Classes", total_classes)
                with col3:
                    st.metric("Average Attendance", f"{avg_attendance:.2f}%")
                
                # Display detailed report
                st.write("### Detailed Attendance Report")
                st.dataframe(attendance_df)
                
                # Download option
                csv = attendance_df.to_csv(index=False)
                st.download_button(
                    label="Download Report",
                    data=csv,
                    file_name=f"attendance_report_{datetime.now().strftime('%Y%m%d')}.csv",
                    mime="text/csv"
                )
            else:
                st.info("No attendance records found for the selected criteria.")
    
    with tab2:
        if st.button("Generate Subject Analysis", type="primary"):
            subject_df = get_subject_analysis(start_date, end_date, None if section == 'All' else section)
            
            if not subject_df.empty:
                st.write("### Subject-wise Attendance Analysis")
                
                # Bar chart for subject-wise attendance
                st.bar_chart(subject_df.set_index('subject')['attendance_percentage'])
                
                # Detailed metrics
                st.write("### Detailed Subject Statistics")
                st.dataframe(subject_df)
                
                # Download option
                csv = subject_df.to_csv(index=False)
                st.download_button(
                    label="Download Subject Analysis",
                    data=csv,
                    file_name=f"subject_analysis_{datetime.now().strftime('%Y%m%d')}.csv",
                    mime="text/csv"
                )
            else:
                st.info("No attendance records found for subject analysis.")
    
    with tab3:
        student_ht = st.selectbox("Select Student", 
                                students_df['ht_number'].tolist(),
                                format_func=lambda x: f"{x} - {students_df[students_df['ht_number'] == x]['name'].iloc[0]}")
        
        if st.button("Generate Student Report", type="primary"):
            student_df, attendance_df, daily_df = get_student_report(student_ht, start_date, end_date)
            
            if not student_df.empty:
                # Student Information
                st.write("### Student Information")
                st.write(f"Name: {student_df['name'].iloc[0]}")
                st.write(f"Section: {student_df['merged_section'].iloc[0]}")
                
                # Overall Attendance Statistics
                if not attendance_df.empty:
                    st.write("### Subject-wise Attendance")
                    
                    # Create a bar chart for subject-wise attendance
                    st.bar_chart(attendance_df.set_index('subject')['attendance_percentage'])
                    
                    # Detailed subject-wise statistics
                    st.dataframe(attendance_df)
                
                # Daily Attendance Log
                if not daily_df.empty:
                    st.write("### Daily Attendance Log")
                    st.dataframe(daily_df)
                    
                    # Download options
                    csv = daily_df.to_csv(index=False)
                    st.download_button(
                        label="Download Student Report",
                        data=csv,
                        file_name=f"student_report_{student_ht}_{datetime.now().strftime('%Y%m%d')}.csv",
                        mime="text/csv"
                    )
                else:
                    st.info("No attendance records found for the selected date range.")
            else:
                st.error("Student not found in the database.")


def admin_page():
    st.title("Admin Dashboard")
    
    tab1, tab2, tab3 = st.tabs(["Manage Students", "Manage Faculty", "View Data"])
    
    with tab1:
        st.subheader("Add/Edit Student")
        with st.form("student_form"):
            ht_number = st.text_input("HT Number")
            name = st.text_input("Student Name")
            original_section = st.text_input("Original Section")
            merged_section = st.text_input("Merged Section")
            
            if st.form_submit_button("Add Student"):
                if all([ht_number, name, original_section, merged_section]):
                    add_student(ht_number, name, original_section, merged_section)
                    st.success("Student added successfully!")
                else:
                    st.error("All fields are required!")
    
    with tab2:
        st.subheader("Add/Edit Faculty")
        with st.form("faculty_form"):
            name = st.text_input("Faculty Name")
            username = st.text_input("Username")
            password = st.text_input("Password", type="password")
            
            if st.form_submit_button("Add Faculty"):
                if all([name, username, password]):
                    add_faculty(name, username, password)
                    st.success("Faculty added successfully!")
                else:
                    st.error("All fields are required!")
    
    with tab3:
        st.subheader("View Database")
        if st.button("View All Students"):
            st.write(get_students())
        if st.button("View All Faculty"):
            faculty_df = get_faculty()
            # Hide password column for security
            if not faculty_df.empty:
                faculty_df = faculty_df.drop('password', axis=1)
            st.write(faculty_df)

def main():
    # Initialize the database
    init_db()
    
    # Set page config
    st.set_page_config(
        page_title="RVIT - Attendance Management System",
        page_icon="https://rvit.edu.in/rvitlogo_f.jpg",
        layout="wide",
        initial_sidebar_state="collapsed"
    )
    
    # Add custom CSS
    st.markdown("""
        <style>
        .stApp {
            max-width: 100%;
            padding: 1rem;
        }
        .stButton button {
            width: 100%;
            padding: 0.8rem !important;
            border-radius: 10px !important;
            margin: 0.5rem 0 !important;
        }
        .student-row {
            background-color: #f0f2f6;
            padding: 1rem;
            border-radius: 10px;
            margin: 0.5rem 0;
        }
        </style>
    """, unsafe_allow_html=True)
    
    # Session state management
    if 'logged_in' not in st.session_state:
        st.session_state.logged_in = False
    
    # Main application logic
    if not st.session_state.logged_in:
        login()
    else:
        # Sidebar for navigation
        with st.sidebar:
            st.image("https://rvit.edu.in/rvitlogo_f.jpg", width=100)
            st.write(f"Welcome, {st.session_state.faculty_name}")
            st.divider()
            
            if st.session_state.is_admin:
                page = "Admin Dashboard"
            else:
                page = st.radio("Navigation", ["Mark Attendance", "View Statistics"])
            
            if st.button("Logout", type="primary"):
                for key in list(st.session_state.keys()):
                    del st.session_state[key]
                st.rerun()
        
        # Main content
        if page == "Admin Dashboard":
            admin_page()
        elif page == "Mark Attendance":
            mark_attendance_page()
        elif page == "View Statistics":
            view_statistics_page()

if __name__ == "__main__":
    main()